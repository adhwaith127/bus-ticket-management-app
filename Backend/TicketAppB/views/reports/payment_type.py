from django.http import JsonResponse
from django.db.models import Sum
from rest_framework.decorators import api_view
from ...models import TransactionData
from ..auth_views import get_user_from_cookie

PAYMENT_LABELS = {0: 'Cash', 1: 'UPI'}


# GET /ticket-app/reports/payment-type
# Returns collection breakdown by payment type (Cash/UPI) for a bus over a date range.
# Params: bus_no (bus reg number), from_date (YYYY-MM-DD), to_date (YYYY-MM-DD),
#         payment_mode (all | cash | upi) — defaults to all
@api_view(['GET'])
def payment_type_report(request):
    user = get_user_from_cookie(request)
    if not user:
        return JsonResponse({'error': 'Unauthorized'}, status=401)

    bus_no = request.GET.get('bus_no')
    from_date = request.GET.get('from_date')
    to_date = request.GET.get('to_date')
    payment_mode = request.GET.get('payment_mode', 'all').lower()  # all | cash | upi

    if not bus_no or not from_date or not to_date:
        return JsonResponse({'error': 'bus_no, from_date and to_date are required'}, status=400)

    qs = TransactionData.objects.filter(
        company_code=user.company,
        device_id=bus_no,
        ticket_date__range=[from_date, to_date],
    )

    if payment_mode == 'cash':
        qs = qs.filter(ticket_status=0)
    elif payment_mode == 'upi':
        qs = qs.filter(ticket_status=1)

    rows_qs = qs.values('ticket_date', 'device_id', 'ticket_status').annotate(
        collection=Sum('ticket_amount')
    ).order_by('ticket_date', 'ticket_status')

    rows = [
        {
            'sl_no': i + 1,
            'date': str(r['ticket_date']),
            'bus_no': r['device_id'],
            'payment_type': PAYMENT_LABELS.get(r['ticket_status'], 'Unknown'),
            'collection': str(r['collection'] or '0.00'),
        }
        for i, r in enumerate(rows_qs)
    ]

    # Totals per payment type
    totals_qs = qs.values('ticket_status').annotate(total=Sum('ticket_amount'))
    totals = {PAYMENT_LABELS.get(t['ticket_status'], 'Unknown'): str(t['total'] or '0.00') for t in totals_qs}

    return JsonResponse({
        'bus_no': bus_no,
        'from_date': from_date,
        'to_date': to_date,
        'payment_mode': payment_mode,
        'rows': rows,
        'totals': totals,
    })
