"""
ETM Device Registry Views
=========================
Handles the full lifecycle of physical ETM and Android app devices:

  POST   /etm-devices/register          — Device self-registers on first boot (no auth required)
  GET    /etm-devices                   — List all devices (superadmin/executive: all;
                                          dealer_admin: their companies' devices;
                                          company_admin: own company's devices)
  GET    /etm-devices/pending           — Pending approval queue (superadmin/executive only)
  POST   /etm-devices/<id>/assign       — Assign device to a company (superadmin/executive)
  POST   /etm-devices/<id>/approve      — Approve + call license server (superadmin/executive)
  POST   /etm-devices/<id>/check-status — Refresh licence expiry from license server
  POST   /etm-devices/<id>/revoke       — Set Inactive (superadmin/executive)
  GET    /etm-devices/summary           — Counts per company/dealer for dashboards

Superadmin visibility rule:
  Superadmin only sees devices belonging to companies that were created
  by (i.e. mapped to) a dealer. Companies that self-registered are NOT
  visible to the superadmin — only to executives.
"""

import logging
import requests
from datetime import date, datetime

from django.conf import settings
from django.utils import timezone
from django.db.models import Count, Q
from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response

from ..models import ETMDevice, Company, Dealer, DealerCustomerMapping
from ..serializers import ETMDeviceSerializer
from .auth_views import get_user_from_cookie
from .utils import (
    _is_superadmin,
    _is_executive,
    _is_dealer_admin,
    _is_company_admin,
    _is_superadmin_or_executive,
    _can_manage_devices,
)

logger = logging.getLogger(__name__)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _call_device_registration(company, device):
    """
    Call license server DeviceRegistration endpoint.
    Returns (device_registration_id, product_type_id, error_message).
    """
    if not company.product_registration_id:
        return None, None, "Company has no product_registration_id. Validate company license first."

    payload = {
        "ProductRegistrationId": company.product_registration_id,
        "UniqueIdentifier": company.unique_identifier or "",
        "MacAddress": device.mac_address or device.serial_number,
        "DeviceType": 1 if device.device_type == ETMDevice.DeviceType.ETM else 2,
        "SerialNumber": device.serial_number,
        "ProjectName": getattr(settings, "PROJECT_NAME", "Bus Ticketing System"),
    }

    url = getattr(settings, "LICENSE_SERVER_BASE_URL", "") + "/DeviceRegistration"

    try:
        logger.info(f"Calling DeviceRegistration for device {device.serial_number}")
        resp = requests.post(url, json=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        if data.get("status") == "Success" or data.get("DeviceRegistrationId"):
            return data.get("DeviceRegistrationId", ""), data.get("ProductTypeId", ""), None

        return None, None, data.get("message", "License server returned failure")

    except requests.exceptions.Timeout:
        return None, None, "License server timeout"
    except requests.exceptions.ConnectionError:
        return None, None, "Cannot connect to license server"
    except Exception as exc:
        logger.exception(f"DeviceRegistration error: {exc}")
        return None, None, str(exc)


def _call_check_device_status(company, device):
    """
    Call license server CheckDeviceStatus endpoint.
    Returns (licence_active_to_date, error_message).
    """
    if not company.product_registration_id or not device.device_registration_id:
        return None, "Missing product_registration_id or device_registration_id"

    payload = {
        "ProductRegistrationId": company.product_registration_id,
        "DeviceRegistrationId": device.device_registration_id,
        "ProjectName": getattr(settings, "PROJECT_NAME", "Bus Ticketing System"),
    }

    url = getattr(settings, "LICENSE_SERVER_BASE_URL", "") + "/CheckDeviceStatus"

    try:
        logger.info(f"Calling CheckDeviceStatus for device {device.serial_number}")
        resp = requests.post(url, json=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        raw_date = data.get("LicenceActiveTo") or data.get("licence_active_to")
        if raw_date:
            try:
                parsed = datetime.strptime(raw_date[:10], "%Y-%m-%d").date()
                return parsed, None
            except ValueError:
                return None, f"Cannot parse expiry date: {raw_date}"

        return None, data.get("message", "No expiry date in response")

    except requests.exceptions.Timeout:
        return None, "License server timeout"
    except requests.exceptions.ConnectionError:
        return None, "Cannot connect to license server"
    except Exception as exc:
        logger.exception(f"CheckDeviceStatus error: {exc}")
        return None, str(exc)


def _device_qs_for_user(user):
    """
    Return the base ETMDevice queryset scoped to what this user is allowed to see.

    - superadmin   → only devices whose company was onboarded via a dealer
    - executive    → devices belonging to their mapped companies only
    - dealer_admin → devices belonging to companies mapped to their dealer
    - company_admin→ own company's devices only
    """
    qs = ETMDevice.objects.select_related("company", "depot", "dealer", "approved_by")

    if _is_superadmin(user):
        dealer_company_ids = DealerCustomerMapping.objects.filter(
            is_active=True
        ).values_list("company_id", flat=True)
        # Include unassigned (company=None) so superadmin can see and assign pending devices
        return qs.filter(Q(company_id__in=dealer_company_ids) | Q(company__isnull=True))

    if _is_executive(user):
        # Executive sees only devices belonging to their mapped companies.
        from ..models import ExecutiveCompanyMapping
        exec_company_ids = ExecutiveCompanyMapping.objects.filter(
            executive_user_id=user.id, is_active=True
        ).values_list("company_id", flat=True)
        return qs.filter(company_id__in=exec_company_ids)

    if _is_dealer_admin(user):
        if not user.dealer_id:
            return qs.none()
        company_ids = DealerCustomerMapping.objects.filter(
            dealer_id=user.dealer_id, is_active=True
        ).values_list("company_id", flat=True)
        return qs.filter(company_id__in=company_ids)

    if _is_company_admin(user):
        if not user.company_id:
            return qs.none()
        return qs.filter(company_id=user.company_id)

    return qs.none()


def _get_scoped_device(user, device_id):
    """
    Fetch a single ETMDevice within the user's allowed scope.
    Returns (device, None) on success, (None, Response) on failure.
    Returns 404 for both "not found" and "out of scope" — prevents enumeration.
    """
    try:
        return _device_qs_for_user(user).select_related("company").get(pk=device_id), None
    except ETMDevice.DoesNotExist:
        return None, Response({"error": "Device not found"}, status=status.HTTP_404_NOT_FOUND)


# ── Views ─────────────────────────────────────────────────────────────────────

@api_view(["POST"])
def register_device(request):
    """
    Called by the ETM device / Android app on first boot.
    No authentication needed — device sends its serial number and MAC.
    Creates a Pending record. If serial already exists, returns existing status.
    """
    serial = (request.data.get("serial_number") or "").strip()
    mac = (request.data.get("mac_address") or "").strip()
    device_type = request.data.get("device_type", ETMDevice.DeviceType.ETM)
    display_name = (request.data.get("display_name") or "").strip()

    if not serial:
        return Response({"error": "serial_number is required"}, status=status.HTTP_400_BAD_REQUEST)

    if device_type not in dict(ETMDevice.DeviceType.choices):
        device_type = ETMDevice.DeviceType.ETM

    # Idempotent — if device re-registers just return current status
    existing = ETMDevice.objects.filter(serial_number=serial).first()
    if existing:
        return Response({
            "message": "Device already registered",
            "data": {
                "serial_number": existing.serial_number,
                "licence_status": existing.licence_status,
                "is_active": existing.is_active,
            }
        }, status=status.HTTP_200_OK)

    device = ETMDevice.objects.create(
        serial_number=serial,
        mac_address=mac,
        device_type=device_type,
        display_name=display_name or serial,
        licence_status=ETMDevice.LicenceStatus.PENDING,
        is_active=False,
    )

    logger.info(f"New device registered: {serial} ({device_type})")
    return Response({
        "message": "Device registered. Waiting for admin approval.",
        "data": {
            "id": device.id,
            "serial_number": device.serial_number,
            "device_type": device.device_type,
            "licence_status": device.licence_status,
        }
    }, status=status.HTTP_201_CREATED)


@api_view(["GET"])
def list_devices(request):
    """
    List devices scoped to the requesting user's role.
    Optional query params: ?status=Pending|Active|Inactive|Expired  ?company=<id>
    """
    user = get_user_from_cookie(request)
    if not user:
        return Response({"error": "Authentication required"}, status=status.HTTP_401_UNAUTHORIZED)

    qs = _device_qs_for_user(user)
    if qs is None:
        return Response({"error": "Unauthorized"}, status=status.HTTP_403_FORBIDDEN)

    # Filters
    filter_status = request.query_params.get("status")
    if filter_status:
        qs = qs.filter(licence_status=filter_status)

    filter_company = request.query_params.get("company")
    if filter_company:
        qs = qs.filter(company_id=filter_company)

    filter_dealer = request.query_params.get("dealer")
    if filter_dealer and _is_superadmin_or_executive(user):
        company_ids = DealerCustomerMapping.objects.filter(
            dealer_id=filter_dealer, is_active=True
        ).values_list("company_id", flat=True)
        qs = qs.filter(company_id__in=company_ids)

    qs = qs.order_by("-created_at")
    serializer = ETMDeviceSerializer(qs, many=True)
    return Response({"message": "Success", "data": serializer.data}, status=status.HTTP_200_OK)


@api_view(["GET"])
def pending_devices(request):
    """
    Returns only Pending devices. Superadmin/executive only.
    """
    user = get_user_from_cookie(request)
    if not user:
        return Response({"error": "Authentication required"}, status=status.HTTP_401_UNAUTHORIZED)
    if not _can_manage_devices(user):
        return Response({"error": "Unauthorized"}, status=status.HTTP_403_FORBIDDEN)

    qs = _device_qs_for_user(user).filter(licence_status=ETMDevice.LicenceStatus.PENDING)
    serializer = ETMDeviceSerializer(qs, many=True)
    return Response({"message": "Success", "data": serializer.data}, status=status.HTTP_200_OK)


@api_view(["POST"])
def assign_device(request, device_id):
    """
    Assign a device to a company (and optionally a depot/dealer).
    Only superadmin and executive can do this.
    Body: { company: <id>, depot: <id|null>, dealer: <id|null>, display_name: "" }
    """
    user = get_user_from_cookie(request)
    if not user:
        return Response({"error": "Authentication required"}, status=status.HTTP_401_UNAUTHORIZED)
    if not _can_manage_devices(user):
        return Response({"error": "Unauthorized"}, status=status.HTTP_403_FORBIDDEN)

    device, err = _get_scoped_device(user, device_id)
    if err:
        return err

    company_id = request.data.get("company")
    if not company_id:
        return Response({"error": "company is required"}, status=status.HTTP_400_BAD_REQUEST)

    try:
        company = Company.objects.get(pk=company_id)
    except Company.DoesNotExist:
        return Response({"error": "Company not found"}, status=status.HTTP_404_NOT_FOUND)

    device.company = company

    # Auto-resolve dealer from company's active dealer mapping
    dealer_mapping = DealerCustomerMapping.objects.filter(
        company=company, is_active=True
    ).select_related("dealer").first()
    device.dealer = dealer_mapping.dealer if dealer_mapping else None

    depot_id = request.data.get("depot")
    if depot_id:
        from ..models import Depot
        try:
            depot = Depot.objects.get(pk=depot_id, company=company)
            device.depot = depot
        except Exception:
            return Response({"error": "Depot not found or does not belong to this company"}, status=status.HTTP_400_BAD_REQUEST)

    display_name = request.data.get("display_name", "").strip()
    if display_name:
        device.display_name = display_name

    device.save()
    serializer = ETMDeviceSerializer(device)
    return Response({"message": "Device assigned successfully", "data": serializer.data}, status=status.HTTP_200_OK)


@api_view(["POST"])
def approve_device(request, device_id):
    """
    Approve a device: calls license server DeviceRegistration, then marks Active.
    Superadmin/executive only.
    The device must already be assigned to a company.
    """
    user = get_user_from_cookie(request)
    if not user:
        return Response({"error": "Authentication required"}, status=status.HTTP_401_UNAUTHORIZED)
    if not _can_manage_devices(user):
        return Response({"error": "Unauthorized"}, status=status.HTTP_403_FORBIDDEN)

    device, err = _get_scoped_device(user, device_id)
    if err:
        return err

    if not device.company:
        return Response(
            {"error": "Device must be assigned to a company before approval"},
            status=status.HTTP_400_BAD_REQUEST,
        )

    if device.licence_status == ETMDevice.LicenceStatus.ACTIVE:
        return Response({"error": "Device is already active"}, status=status.HTTP_400_BAD_REQUEST)

    # Call license server
    reg_id, _, error = _call_device_registration(device.company, device)

    if error:
        # Log but still allow admin to force-approve internally
        logger.warning(f"License server error for device {device.serial_number}: {error}")
        # Mark as Active internally even if license server unreachable,
        # admin can re-run check-status later to sync expiry
        device.device_registration_id = ""
    else:
        device.device_registration_id = str(reg_id or "")

    device.licence_status = ETMDevice.LicenceStatus.ACTIVE
    device.is_active = True
    device.approved_by = user
    device.approved_at = timezone.now()
    device.save()

    # Try to fetch expiry immediately (best-effort, don't fail if it errors)
    if device.device_registration_id:
        expiry_date, _ = _call_check_device_status(device.company, device)
        if expiry_date:
            device.licence_active_to = expiry_date
            device.save(update_fields=["licence_active_to"])

    serializer = ETMDeviceSerializer(device)
    warning_msg = f" (license server warning: {error})" if error else ""
    return Response({
        "message": f"Device approved successfully{warning_msg}",
        "data": serializer.data,
    }, status=status.HTTP_200_OK)


@api_view(["POST"])
def check_device_status(request, device_id):
    """
    Refresh licence expiry from license server. Superadmin/executive only.
    Updates licence_active_to and flips Expired if past due.
    """
    user = get_user_from_cookie(request)
    if not user:
        return Response({"error": "Authentication required"}, status=status.HTTP_401_UNAUTHORIZED)
    if not _can_manage_devices(user):
        return Response({"error": "Unauthorized"}, status=status.HTTP_403_FORBIDDEN)

    device, err = _get_scoped_device(user, device_id)
    if err:
        return err

    if not device.company:
        return Response({"error": "Device has no company assigned"}, status=status.HTTP_400_BAD_REQUEST)

    expiry_date, error = _call_check_device_status(device.company, device)

    if error:
        return Response({"error": f"License server error: {error}"}, status=status.HTTP_502_BAD_GATEWAY)

    device.licence_active_to = expiry_date
    today = date.today()
    if expiry_date and today > expiry_date:
        device.licence_status = ETMDevice.LicenceStatus.EXPIRED
        device.is_active = False
    elif device.licence_status == ETMDevice.LicenceStatus.EXPIRED:
        # Renewed — set back to active
        device.licence_status = ETMDevice.LicenceStatus.ACTIVE
        device.is_active = True

    device.save()
    serializer = ETMDeviceSerializer(device)
    return Response({"message": "Device status refreshed", "data": serializer.data}, status=status.HTTP_200_OK)


@api_view(["POST"])
def revoke_device(request, device_id):
    """
    Set device to Inactive. Superadmin/executive only.
    """
    user = get_user_from_cookie(request)
    if not user:
        return Response({"error": "Authentication required"}, status=status.HTTP_401_UNAUTHORIZED)
    if not _can_manage_devices(user):
        return Response({"error": "Unauthorized"}, status=status.HTTP_403_FORBIDDEN)

    device, err = _get_scoped_device(user, device_id)
    if err:
        return err

    device.licence_status = ETMDevice.LicenceStatus.INACTIVE
    device.is_active = False
    device.save()

    serializer = ETMDeviceSerializer(device)
    return Response({"message": "Device revoked", "data": serializer.data}, status=status.HTTP_200_OK)


@api_view(["GET"])
def device_summary(request):
    """
    Aggregated device counts. Used by dashboards.

    Response shape:
    {
      "total": 42,
      "by_status": { "Pending": 3, "Active": 35, "Inactive": 2, "Expired": 2 },
      "by_company": [ { "company_id": 1, "company_name": "...", "total": 5, "active": 4 }, ... ],
      "by_dealer":  [ { "dealer_id": 1, "dealer_name": "...", "total": 10, "active": 8 }, ... ]
    }
    """
    user = get_user_from_cookie(request)
    if not user:
        return Response({"error": "Authentication required"}, status=status.HTTP_401_UNAUTHORIZED)

    qs = _device_qs_for_user(user)

    total = qs.count()

    by_status = {}
    for s in ETMDevice.LicenceStatus.values:
        by_status[s] = qs.filter(licence_status=s).count()

    by_company = []
    company_agg = (
        qs.filter(company__isnull=False)
        .values("company__id", "company__company_name")
        .annotate(
            total=Count("id"),
            active=Count("id", filter=Q(licence_status=ETMDevice.LicenceStatus.ACTIVE)),
        )
        .order_by("-total")
    )
    for row in company_agg:
        by_company.append({
            "company_id": row["company__id"],
            "company_name": row["company__company_name"],
            "total": row["total"],
            "active": row["active"],
        })

    by_dealer = []
    if _is_superadmin_or_executive(user):
        dealer_agg = (
            qs.filter(dealer__isnull=False)
            .values("dealer__id", "dealer__dealer_name")
            .annotate(
                total=Count("id"),
                active=Count("id", filter=Q(licence_status=ETMDevice.LicenceStatus.ACTIVE)),
            )
            .order_by("-total")
        )
        for row in dealer_agg:
            by_dealer.append({
                "dealer_id": row["dealer__id"],
                "dealer_name": row["dealer__dealer_name"],
                "total": row["total"],
                "active": row["active"],
            })

    return Response({
        "message": "Success",
        "data": {
            "total": total,
            "by_status": by_status,
            "by_company": by_company,
            "by_dealer": by_dealer,
        }
    }, status=status.HTTP_200_OK)
