from django.http import JsonResponse
from rest_framework.decorators import api_view
from ...models import TripCloseData, CrewAssignment
from ..auth_views import get_user_from_cookie


# GET /ticket-app/reports/duty
# Returns all trips made by a bus on a specific date, with conductor name.
# Params: device_id (bus reg number), date (YYYY-MM-DD)
@api_view(['GET'])
def duty_report(request):
    user = get_user_from_cookie(request)
    if not user:
        return JsonResponse({'error': 'Unauthorized'}, status=401)

    device_id = request.GET.get('device_id')
    date = request.GET.get('date')  # YYYY-MM-DD

    if not device_id or not date:
        return JsonResponse({'error': 'device_id and date are required'}, status=400)

    trips = TripCloseData.objects.filter(
        palmtec_id=device_id,
        start_date=date,
        company_code=user.company,
    ).order_by('trip_no').values(
        'trip_no', 'start_time', 'start_ticket_no', 'end_ticket_no', 'total_collection'
    )

    # Conductor name from crew assignment (matched by bus reg number = device_id)
    conductor_name = None
    try:
        assignment = CrewAssignment.objects.select_related('conductor', 'vehicle').filter(
            vehicle__bus_reg_num=device_id,
            company=user.company,
        ).first()
        if assignment and assignment.conductor:
            conductor_name = assignment.conductor.employee_name
    except Exception:
        pass

    trip_list = [
        {
            'trip_no': t['trip_no'],
            'time': str(t['start_time']) if t['start_time'] else None,
            'start_ticket': t['start_ticket_no'],
            'end_ticket': t['end_ticket_no'],
            'collection': str(t['total_collection']),
        }
        for t in trips
    ]

    return JsonResponse({
        'bus_no': device_id,
        'date': date,
        'conductor': conductor_name,
        'trips': trip_list,
    })
