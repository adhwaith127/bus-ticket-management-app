from rest_framework import status
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from datetime import datetime
import logging
from ..models import TransactionData
from ..serializers import TicketDataSerializer
from rest_framework.decorators import api_view
from rest_framework.response import Response
from django.db import IntegrityError, transaction
from decimal import Decimal


logger = logging.getLogger(__name__)

@csrf_exempt
def getTransactionDataFromDevice(request):
    raw = request.GET.get("fn")

    if not raw:
        return JsonResponse({"status": "error","message": "No input data"}, status=status.HTTP_400_BAD_REQUEST)

    logger.info("Transaction from device: %s", raw)

    parts = raw.split("|")

    try:
        # this is for keeping the transaction either complete all or nothing at all
        # with transaction.atomic():
        transaction = TransactionData.objects.create(
            request_type     = parts[0] if len(parts) > 0 else None,
            device_id        = parts[1] if len(parts) > 1 else None,
            trip_number      = parts[2] if len(parts) > 2 else None,
            ticket_number    = parts[3] if len(parts) > 3 else None,

            ticket_date = datetime.strptime(parts[4], "%Y-%m-%d").date() if len(parts) > 4 and parts[4] else None,
            ticket_time = datetime.strptime(parts[5], "%H:%M:%S").time() if len(parts) > 5 and parts[5] else None,

            from_stage = int(parts[6]) if len(parts) > 6 and parts[6] else 0,
            to_stage     = parts[7] or 0,

            full_count   = parts[8] or 0,
            half_count   = parts[9] or 0,
            st_count     = parts[10] or 0,
            phy_count    = parts[11] or 0,
            lugg_count   = parts[12] or 0,

            ticket_amount = Decimal(parts[13]) if len(parts) > 13 and parts[13] else Decimal("0.00"),
            lugg_amount   = parts[14] or 0,

            ticket_type   = parts[15] if len(parts) > 15 else None,
            adjust_amount = parts[16] or 0,

            pass_id        = parts[17] if len(parts) > 17 else None,
            warrant_amount = parts[18] or 0,

            refund_status = parts[19] if len(parts) > 19 else None,
            refund_amount = parts[20] or 0,

            ladies_count  = parts[21] or 0,
            senior_count  = parts[22] or 0,

            transaction_id   = parts[23] if len(parts) > 23 else None,
            ticket_status    = parts[24] if len(parts) > 24 else None,
            reference_number = parts[25] if len(parts) > 25 else None,
            company_code     = parts[26] if len(parts) > 26 else None,

            raw_payload = raw
        )
    
    except IntegrityError:
        return JsonResponse({"status": "duplicate", "message": "Transaction already exists"},status=status.HTTP_400_BAD_REQUEST)

    except Exception as e:
        logger.exception("Transaction parsing failed")
        return JsonResponse({"status": "error","message": str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    return JsonResponse({"status": "success","transaction_id": transaction.id},status=status.HTTP_201_CREATED)


@api_view(['GET'])
def get_all_transaction_data(request):
    try:
        ticketdata=TransactionData.objects.all().order_by('created_at')
        serializer=TicketDataSerializer(ticketdata,many=True)
        
        return Response({"message": "success","data": serializer.data},status=status.HTTP_200_OK)
    
    except Exception as e:
        return Response({"message":"Data fetching failed"},status=status.HTTP_500_INTERNAL_SERVER_ERROR)