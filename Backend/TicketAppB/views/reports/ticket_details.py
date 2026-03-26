from django.http import JsonResponse
from rest_framework.decorators import api_view
from ...models import TransactionData, Stage
from ..auth_views import get_user_from_cookie

PAYMENT_LABELS = {0: 'Cash', 1: 'UPI'}


# GET /ticket-app/reports/ticket-details
# Returns all individual tickets issued in a specific trip,
# with stage names resolved and a summary of passenger type totals.
# Params: device_id (bus reg number), trip_no (integer), date (YYYY-MM-DD)
@api_view(['GET'])
def ticket_details(request):
    user = get_user_from_cookie(request)
    if not user:
        return JsonResponse({'error': 'Unauthorized'}, status=401)

    device_id = request.GET.get('device_id')
    trip_no = request.GET.get('trip_no')
    date = request.GET.get('date')  # YYYY-MM-DD

    if not device_id or not trip_no or not date:
        return JsonResponse({'error': 'device_id, trip_no and date are required'}, status=400)

    tickets = TransactionData.objects.filter(
        company_code=user.company,
        device_id=device_id,
        trip_number=trip_no,
        ticket_date=date,
    ).order_by('ticket_time').values(
        'ticket_number',
        'ticket_time',
        'from_stage',
        'to_stage',
        'ticket_type',
        'ticket_amount',
        'ticket_status',
        'full_count',
        'half_count',
        'st_count',
        'phy_count',
    )

    # Stage name lookup
    stages = Stage.objects.filter(company=user.company, is_deleted=False).values('stage_code', 'stage_name')
    stage_map = {}
    for s in stages:
        try:
            stage_map[int(s['stage_code'])] = s['stage_name']
        except (ValueError, TypeError):
            pass

    ticket_list = [
        {
            'ticket_number': t['ticket_number'],
            'time': str(t['ticket_time']) if t['ticket_time'] else None,
            'from_stage': t['from_stage'],
            'from_stage_name': stage_map.get(t['from_stage'], str(t['from_stage']) if t['from_stage'] else None),
            'to_stage': t['to_stage'],
            'to_stage_name': stage_map.get(t['to_stage'], str(t['to_stage']) if t['to_stage'] else None),
            'ticket_type': t['ticket_type'],
            'amount': str(t['ticket_amount']),
            'payment_type': PAYMENT_LABELS.get(t['ticket_status'], 'Unknown'),
            'full_count': t['full_count'],
            'half_count': t['half_count'],
            'st_count': t['st_count'],
            'phy_count': t['phy_count'],
        }
        for t in tickets
    ]

    # Summary counts
    total_full = sum(t['full_count'] for t in ticket_list)
    total_half = sum(t['half_count'] for t in ticket_list)
    total_st = sum(t['st_count'] for t in ticket_list)
    total_phy = sum(t['phy_count'] for t in ticket_list)

    return JsonResponse({
        'device_id': device_id,
        'trip_no': trip_no,
        'date': date,
        'summary': {
            'full_count': total_full,
            'half_count': total_half,
            'st_count': total_st,
            'phy_count': total_phy,
        },
        'tickets': ticket_list,
    })
