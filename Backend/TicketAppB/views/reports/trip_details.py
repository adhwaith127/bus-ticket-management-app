from django.http import JsonResponse
from django.db.models import Sum
from rest_framework.decorators import api_view
from ...models import TransactionData, TripCloseData, Stage
from ..auth_views import get_user_from_cookie


def _build_stage_map(company):
    """Returns {stage_code_int: stage_name} for the company."""
    stages = Stage.objects.filter(company=company, is_deleted=False).values('stage_code', 'stage_name')
    stage_map = {}
    for s in stages:
        try:
            stage_map[int(s['stage_code'])] = s['stage_name']
        except (ValueError, TypeError):
            pass
    return stage_map


# GET /ticket-app/reports/trip-details
# Returns stage-wise boarded/deboarded passenger table for a specific trip,
# along with a summary header (total collection, ticket counts per type).
# Params: device_id (bus reg number), trip_no (integer), date (YYYY-MM-DD)
@api_view(['GET'])
def trip_details(request):
    user = get_user_from_cookie(request)
    if not user:
        return JsonResponse({'error': 'Unauthorized'}, status=401)

    device_id = request.GET.get('device_id')
    trip_no = request.GET.get('trip_no')
    date = request.GET.get('date')  # YYYY-MM-DD

    if not device_id or not trip_no or not date:
        return JsonResponse({'error': 'device_id, trip_no and date are required'}, status=400)

    # Trip summary from TripCloseData
    try:
        trip = TripCloseData.objects.get(
            company_code=user.company,
            palmtec_id=device_id,
            trip_no=int(trip_no),
            start_date=date,
        )
        summary = {
            'trip_no': trip.trip_no,
            'route_code': trip.route_code,
            'direction': 'Up' if trip.up_down_trip == 'U' else 'Down',
            'total_collection': str(trip.total_collection),
            'full_count': trip.full_count,
            'half_count': trip.half_count,
            'st_count': trip.st1_count,
            'phy_count': trip.physical_count,
            'pass_count': trip.pass_count,
            'total_tickets': trip.total_tickets,
        }
    except TripCloseData.DoesNotExist:
        summary = None

    # Stage-wise boarding from TransactionData
    qs = TransactionData.objects.filter(
        company_code=user.company,
        device_id=device_id,
        trip_number=trip_no,
        ticket_date=date,
    )

    boarded = qs.values('from_stage').annotate(
        full=Sum('full_count'),
        half=Sum('half_count'),
        st=Sum('st_count'),
        phy=Sum('phy_count'),
    )

    deboarded = qs.values('to_stage').annotate(
        full=Sum('full_count'),
        half=Sum('half_count'),
        st=Sum('st_count'),
        phy=Sum('phy_count'),
    )

    stage_map = _build_stage_map(user.company)

    boarded_map = {
        r['from_stage']: {'f': r['full'] or 0, 'h': r['half'] or 0, 'st': r['st'] or 0, 'ph': r['phy'] or 0}
        for r in boarded if r['from_stage'] is not None
    }
    deboarded_map = {
        r['to_stage']: {'f': r['full'] or 0, 'h': r['half'] or 0, 'st': r['st'] or 0, 'ph': r['phy'] or 0}
        for r in deboarded if r['to_stage'] is not None
    }

    all_stages = sorted(set(boarded_map.keys()) | set(deboarded_map.keys()))

    stage_table = [
        {
            'stage_code': stage_code,
            'stage_name': stage_map.get(stage_code, str(stage_code)),
            'boarded': boarded_map.get(stage_code, {'f': 0, 'h': 0, 'st': 0, 'ph': 0}),
            'deboarded': deboarded_map.get(stage_code, {'f': 0, 'h': 0, 'st': 0, 'ph': 0}),
        }
        for stage_code in all_stages
    ]

    return JsonResponse({
        'device_id': device_id,
        'trip_no': trip_no,
        'date': date,
        'summary': summary,
        'stage_table': stage_table,
    })
