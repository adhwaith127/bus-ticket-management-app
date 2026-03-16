#!/usr/bin/env python3
import argparse
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import date
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError


@dataclass
class Result:
    index: int
    status_code: int | None
    body: str
    ok: bool
    error: str = ""


def build_payload(
    device_id: str,
    company_code: str,
    ticket_number: int,
    transaction_id: int,
    reference_number: str,
    ticket_date: str,
    ticket_time: str,
) -> str:
    fields = [
        "Ticket",                  # 0 request_type
        device_id,                 # 1 device_id
        "1",                       # 2 trip_number
        str(ticket_number),        # 3 ticket_number
        ticket_date,               # 4 ticket_date
        ticket_time,               # 5 ticket_time
        "1",                       # 6 from_stage
        "2",                       # 7 to_stage
        "1",                       # 8 full_count
        "0",                       # 9 half_count
        "0",                       # 10 st_count
        "0",                       # 11 phy_count
        "0",                       # 12 lugg_count
        "1.00",                    # 13 ticket_amount
        "0.00",                    # 14 lugg_amount
        "1",                       # 15 ticket_type
        "0.00",                    # 16 adjust_amount
        "0",                       # 17 pass_id
        "0",                       # 18 warrant_amount
        "0",                       # 19 refund_status
        "0.00",                    # 20 refund_amount
        "0",                       # 21 ladies_count
        "0",                       # 22 senior_count
        str(transaction_id),       # 23 transaction_id
        "0",                       # 24 ticket_status (cash)
        reference_number,          # 25 reference_number
        company_code,              # 26 company_code
        "",                        # trailing field to preserve final delimiter
    ]
    return "|".join(fields)


def send_one(base_url: str, payload: str, index: int, timeout: float) -> Result:
    query = urlencode({"fn": payload})
    url = f"{base_url}?{query}"
    request = Request(url, method="GET")
    try:
        with urlopen(request, timeout=timeout) as response:
            body = response.read().decode("utf-8", errors="replace").strip()
            status_code = response.getcode()
            return Result(index=index, status_code=status_code, body=body, ok=200 <= status_code < 300)
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace").strip()
        return Result(index=index, status_code=exc.code, body=body, ok=False, error=f"HTTP {exc.code}")
    except URLError as exc:
        return Result(index=index, status_code=None, body="", ok=False, error=str(exc.reason))
    except Exception as exc:  # noqa: BLE001
        return Result(index=index, status_code=None, body="", ok=False, error=str(exc))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Burst test the getTicket API with synthetic ticket requests.")
    parser.add_argument(
        "--url",
        default="http://192.168.0.61:8001/ticket-app/getTicket",
        help="Full getTicket endpoint URL.",
    )
    parser.add_argument("--count", type=int, default=100, help="Total number of requests to send.")
    parser.add_argument(
        "--concurrency",
        type=int,
        default=20,
        help="Number of concurrent requests. Start smaller locally, then increase carefully.",
    )
    parser.add_argument("--device-id", default="12345", help="Device ID to place in the payload.")
    parser.add_argument("--company-code", default="385", help="Company code expected by the backend.")
    parser.add_argument("--ticket-start", type=int, default=1, help="Starting ticket number.")
    parser.add_argument(
        "--transaction-start",
        type=int,
        default=5524758110,
        help="Starting transaction id. Each request increments this value.",
    )
    parser.add_argument(
        "--ticket-date",
        default=date.today().isoformat(),
        help="Ticket date to place in the payload. Defaults to today.",
    )
    parser.add_argument("--ticket-time", default="04:00:00", help="Ticket time to place in the payload.")
    parser.add_argument("--timeout", type=float, default=10.0, help="Per-request timeout in seconds.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.count < 1:
        print("count must be >= 1", file=sys.stderr)
        return 1
    if args.concurrency < 1:
        print("concurrency must be >= 1", file=sys.stderr)
        return 1

    request_plan = []
    for index, (ticket_number, transaction_id) in enumerate(
        zip(
            range(args.ticket_start, args.ticket_start + args.count),
            range(args.transaction_start, args.transaction_start + args.count),
        ),
        start=1,
    ):
        reference_number = f"{ticket_number:06d}00507AA{args.device_id}"
        payload = build_payload(
            device_id=args.device_id,
            company_code=args.company_code,
            ticket_number=ticket_number,
            transaction_id=transaction_id,
            reference_number=reference_number,
            ticket_date=args.ticket_date,
            ticket_time=args.ticket_time,
        )
        request_plan.append((index, payload))

    started_at = time.perf_counter()
    results: list[Result] = []

    with ThreadPoolExecutor(max_workers=args.concurrency) as executor:
        futures = [
            executor.submit(send_one, args.url, payload, index, args.timeout)
            for index, payload in request_plan
        ]
        for future in as_completed(futures):
            results.append(future.result())

    elapsed = time.perf_counter() - started_at
    results.sort(key=lambda item: item.index)

    success = [result for result in results if result.ok]
    failed = [result for result in results if not result.ok]
    grouped_statuses: dict[str, int] = {}

    for result in results:
        key = str(result.status_code) if result.status_code is not None else "ERR"
        grouped_statuses[key] = grouped_statuses.get(key, 0) + 1

    print(f"URL: {args.url}")
    print(f"Requests sent: {args.count}")
    print(f"Concurrency: {args.concurrency}")
    print(f"Elapsed: {elapsed:.2f}s")
    print(f"Success: {len(success)}")
    print(f"Failed: {len(failed)}")
    print("Status counts:", ", ".join(f"{code}={count}" for code, count in sorted(grouped_statuses.items())))

    if failed:
        print("\nSample failures:")
        for result in failed[:10]:
            detail = result.error or result.body
            print(f"  #{result.index}: status={result.status_code} detail={detail}")

    if success:
        print("\nSample successes:")
        for result in success[:5]:
            print(f"  #{result.index}: status={result.status_code} body={result.body}")

    return 0 if not failed else 2


if __name__ == "__main__":
    raise SystemExit(main())
