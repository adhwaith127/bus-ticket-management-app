import logging
from decimal import Decimal
from datetime import datetime
from rest_framework import status
from ..models import TransactionData
from django.http import JsonResponse
from .auth_views import get_user_from_cookie
from rest_framework.response import Response
from ..serializers import TicketDataSerializer
from rest_framework.decorators import api_view
from django.db import IntegrityError, transaction
from django.views.decorators.csrf import csrf_exempt
from django.contrib.auth import get_user_model
from django.http import HttpResponse


User=get_user_model()
logger = logging.getLogger(__name__)

# used by machine
@csrf_exempt
def getTransactionDataFromDevice(request):
    if request.method != "GET":
        return HttpResponse("METHOD_NOT_ALLOWED",status=status.HTTP_405_METHOD_NOT_ALLOWED,content_type="text/plain")

    raw = request.GET.get("fn")

    if not raw:
        return HttpResponse("NO_DATA",status=status.HTTP_400_BAD_REQUEST,content_type="text/plain")

    logger.info("Transaction from device: %s", raw)

    parts = raw.split("|")

    try:
        transaction = TransactionData.objects.create(
            request_type = parts[0] if len(parts) > 0 else None,
            device_id    = parts[1] if len(parts) > 1 else None,
            trip_number  = parts[2] if len(parts) > 2 else None,
            ticket_number= parts[3] if len(parts) > 3 else None,

            ticket_date = datetime.strptime(parts[4], "%Y-%m-%d").date()
                          if len(parts) > 4 and parts[4] else None,
            ticket_time = datetime.strptime(parts[5], "%H:%M:%S").time()
                          if len(parts) > 5 and parts[5] else None,

            from_stage = int(parts[6]) if len(parts) > 6 and parts[6] else 0,
            to_stage   = parts[7] or 0,

            full_count = parts[8] or 0,
            half_count = parts[9] or 0,
            st_count   = parts[10] or 0,
            phy_count  = parts[11] or 0,
            lugg_count = parts[12] or 0,

            ticket_amount = Decimal(parts[13]) if len(parts) > 13 and parts[13] else Decimal("0.00"),
            lugg_amount   = parts[14] or 0,

            ticket_type   = parts[15] if len(parts) > 15 else None,
            adjust_amount = parts[16] or 0,

            pass_id        = parts[17] if len(parts) > 17 else None,
            warrant_amount = parts[18] or 0,

            refund_status = parts[19] if len(parts) > 19 else None,
            refund_amount = parts[20] or 0,

            ladies_count = parts[21] or 0,
            senior_count = parts[22] or 0,

            transaction_id   = parts[23] if len(parts) > 23 else None,
            ticket_status    = parts[24] if len(parts) > 24 else None,
            reference_number = parts[25] if len(parts) > 25 else None,
            company_code     = parts[26] if len(parts) > 26 else None,

            raw_payload = raw
        )

    except IntegrityError:
        return HttpResponse("DUPLICATE", status=status.HTTP_200_OK,content_type="text/plain")

    except Exception:
        logger.exception("Transaction parsing failed")
        return HttpResponse("ERROR",status=status.HTTP_500_INTERNAL_SERVER_ERROR,content_type="text/plain")

    response_chars=raw[0:32]
    device_response=f'OK#SUCCESS#fn={response_chars}#'

    return HttpResponse(device_response, content_type="text/plain", status=status.HTTP_201_CREATED)

@api_view(['GET'])
def get_all_transaction_data(request):
    user = get_user_from_cookie(request)
    if not user:
        return Response({'error': 'Authentication required'}, status=status.HTTP_401_UNAUTHORIZED)
    
    try:
        # Filter by user's company if user has a company assigned
        if user.company:
            ticketdata = TransactionData.objects.filter(company_code=user.company).order_by('created_at')
        else:
            # If no company assigned, return empty data
            ticketdata = TransactionData.objects.none()
        
        serializer = TicketDataSerializer(ticketdata, many=True)
        return Response({"message": "success", "data": serializer.data}, status=status.HTTP_200_OK)
    except Exception as e:
        return Response({"message": "Data fetching failed"}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
    


@csrf_exempt
def some_function(request):
    if request.method != "GET":
        return JsonResponse({"error": "Method not allowed"}, status=status.HTTP_405_METHOD_NOT_ALLOWED)