from django.http import JsonResponse
from django.db.models import Sum, Count
from rest_framework.decorators import api_view
from ...models import TransactionData
from ..auth_views import get_user_from_cookie


# GET /ticket-app/reports/farewise
# Returns two sections:
#   1. Fare-wise ticket count and revenue (grouped by fare amount)
#   2. Passenger type count per trip (Full, Half, ST, PHY, Luggage)
# Params: bus_no (bus reg number), from_date (YYYY-MM-DD), to_date (YYYY-MM-DD)
@api_view(['GET'])
def farewise_report(request):
    user = get_user_from_cookie(request)
    if not user:
        return JsonResponse({'error': 'Unauthorized'}, status=401)

    bus_no = request.GET.get('bus_no')
    from_date = request.GET.get('from_date')
    to_date = request.GET.get('to_date')

    if not bus_no or not from_date or not to_date:
        return JsonResponse({'error': 'bus_no, from_date and to_date are required'}, status=400)

    qs = TransactionData.objects.filter(
        company_code=user.company,
        device_id=bus_no,
        ticket_date__range=[from_date, to_date],
    )

    # Part 1: Fare-wise ticket count and revenue
    fare_rows = qs.values('ticket_amount').annotate(
        ticket_count=Count('id'),
        revenue=Sum('ticket_amount'),
    ).order_by('ticket_amount')

    fares = [
        {
            'sl_no': i + 1,
            'fare': str(r['ticket_amount']),
            'ticket_count': r['ticket_count'],
            'revenue': str(r['revenue'] or '0.00'),
        }
        for i, r in enumerate(fare_rows)
    ]

    # Part 2: Passenger count per trip
    trip_rows = qs.values('trip_number').annotate(
        full=Sum('full_count'),
        half=Sum('half_count'),
        st=Sum('st_count'),
        phy=Sum('phy_count'),
        lugg=Sum('lugg_count'),
    ).order_by('trip_number')

    passenger_counts = [
        {
            'trip': r['trip_number'],
            'full': r['full'] or 0,
            'half': r['half'] or 0,
            'st': r['st'] or 0,
            'phy': r['phy'] or 0,
            'lugg': r['lugg'] or 0,
        }
        for r in trip_rows
    ]

    return JsonResponse({
        'bus_no': bus_no,
        'from_date': from_date,
        'to_date': to_date,
        'fares': fares,
        'passenger_counts': passenger_counts,
    })
