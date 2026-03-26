from django.http import JsonResponse
from rest_framework.decorators import api_view
from ...models import TripCloseData
from ..auth_views import get_user_from_cookie


# GET /ticket-app/reports/passenger-info
# Returns list of all trips for a device on a specific date,
# with cash/UPI breakdown and route info per trip.
# Params: device_id (bus reg number / palmtec ID), date (YYYY-MM-DD)
@api_view(['GET'])
def passenger_info(request):
    user = get_user_from_cookie(request)
    if not user:
        return JsonResponse({'error': 'Unauthorized'}, status=401)

    device_id = request.GET.get('device_id')
    date = request.GET.get('date')  # YYYY-MM-DD

    if not device_id or not date:
        return JsonResponse({'error': 'device_id and date are required'}, status=400)

    trips = TripCloseData.objects.filter(
        company_code=user.company,
        palmtec_id=device_id,
        start_date=date,
    ).order_by('-trip_no').values(
        'trip_no',
        'route_code',
        'up_down_trip',
        'start_time',
        'end_time',
        'total_collection',
        'total_cash_amount',
        'upi_ticket_amount',
        'total_tickets',
        'upi_ticket_count',
        'total_cash_tickets',
    )

    trip_list = [
        {
            'trip_no': t['trip_no'],
            'route_name': t['route_code'],
            'direction': 'Up' if t['up_down_trip'] == 'U' else 'Down',
            'start_time': str(t['start_time']) if t['start_time'] else None,
            'end_time': str(t['end_time']) if t['end_time'] else None,
            'total_collection': str(t['total_collection']),
            'cash_amount': str(t['total_cash_amount']),
            'upi_amount': str(t['upi_ticket_amount']),
            'total_tickets': t['total_tickets'],
            'upi_tickets': t['upi_ticket_count'],
            'cash_tickets': t['total_cash_tickets'],
        }
        for t in trips
    ]

    return JsonResponse({
        'device_id': device_id,
        'date': date,
        'trips': trip_list,
    })
