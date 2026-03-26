from django.http import JsonResponse
from django.db.models import Sum
from rest_framework.decorators import api_view
from ...models import TripCloseData
from ..auth_views import get_user_from_cookie


# GET /ticket-app/reports/bus-summary
# Returns total revenue per day for a bus over a date range.
# Params: bus_no (bus reg number), from_date (YYYY-MM-DD), to_date (YYYY-MM-DD)
@api_view(['GET'])
def bus_summary_report(request):
    user = get_user_from_cookie(request)
    if not user:
        return JsonResponse({'error': 'Unauthorized'}, status=401)

    bus_no = request.GET.get('bus_no')
    from_date = request.GET.get('from_date')
    to_date = request.GET.get('to_date')

    if not bus_no or not from_date or not to_date:
        return JsonResponse({'error': 'bus_no, from_date and to_date are required'}, status=400)

    qs = TripCloseData.objects.filter(
        company_code=user.company,
        palmtec_id=bus_no,
        start_date__range=[from_date, to_date],
    ).values('start_date').annotate(
        revenue=Sum('total_collection')
    ).order_by('start_date')

    rows = [
        {
            'sl_no': i + 1,
            'date': str(row['start_date']),
            'bus_no': bus_no,
            'revenue': str(row['revenue'] or '0.00'),
        }
        for i, row in enumerate(qs)
    ]

    total = sum(float(r['revenue']) for r in rows)

    return JsonResponse({
        'bus_no': bus_no,
        'from_date': from_date,
        'to_date': to_date,
        'rows': rows,
        'total': f'{total:.2f}',
    })
