import logging
from decimal import Decimal,InvalidOperation
from datetime import datetime
from rest_framework import status
from ..models import TransactionData,TripCloseData
from django.http import HttpResponse
from django.http import JsonResponse
from .auth_views import get_user_from_cookie
from rest_framework.response import Response
from django.contrib.auth import get_user_model
from ..serializers import TicketDataSerializer
from rest_framework.decorators import api_view
from django.db import IntegrityError, transaction
from django.views.decorators.csrf import csrf_exempt


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

    # get first 32 chars for response
    response_chars=raw[0:32]

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
        device_response=f'OK#SUCCESS#fn={response_chars}#'
        return HttpResponse(device_response, content_type="text/plain", status=status.HTTP_200_OK)

    except Exception:
        logger.exception("Transaction parsing failed")
        return HttpResponse("ERROR",status=status.HTTP_500_INTERNAL_SERVER_ERROR,content_type="text/plain")

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
def getTripCloseDataFromDevice(request):   
    # Check request method
    if request.method != 'GET':
        return HttpResponse("METHOD_NOT_ALLOWED",status=status.HTTP_405_METHOD_NOT_ALLOWED,content_type="text/plain")

    try:
        # Extract raw data from request
        raw_payload = request.GET.get('fn', '')

        if not raw_payload:
            return HttpResponse("NO_DATA",status=status.HTTP_400_BAD_REQUEST,content_type="text/plain")

        # Split data by delimiter
        parts = raw_payload.split('|')

        # Check if we have minimum required parts
        if len(parts) < 33:
            return HttpResponse("MISSING_DATA",status=status.HTTP_400_BAD_REQUEST,content_type="text/plain")

        # Validate request type
        if parts[0] != 'TrpCl':
            return HttpResponse(f"INVALID",status=status.HTTP_400_BAD_REQUEST,content_type="text/plain")
        
        # Create TripCloseData instance
        try:
            trip_data = TripCloseData.objects.create(
                # Device information
                palmtec_id=parts[1],
                license_code=parts[2],

                # Trip identification
                schedule=int(parts[3]) if parts[3] else 0,
                trip_no=int(parts[4]) if parts[4] else 0,
                route_code=parts[31],
                up_down_trip=parts[32] if len(parts) > 32 else '',

                # Trip timing - parse and combine date
                start_datetime=datetime.strptime(f"{parts[5]} {parts[6]}", "%Y-%m-%d %H:%M:%S"),
                end_datetime=datetime.strptime(f"{parts[7]} {parts[8]}", "%Y-%m-%d %H:%M:%S"),

                # Ticket range
                start_ticket_no=int(parts[9]) if parts[9] else 0,
                end_ticket_no=int(parts[10]) if parts[10] else 0,

                # Passenger counts
                full_count=int(parts[11]) if parts[11] else 0,
                half_count=int(parts[12]) if parts[12] else 0,
                st1_count=int(parts[13]) if parts[13] else 0,
                luggage_count=int(parts[14]) if parts[14] else 0,
                physical_count=int(parts[15]) if parts[15] else 0,
                pass_count=int(parts[16]) if parts[16] else 0,
                ladies_count=int(parts[17]) if parts[17] else 0,
                senior_count=int(parts[18]) if parts[18] else 0,

                # Collection amounts - convert to Decimal
                full_collection=Decimal(str(parts[19])) if parts[19] else Decimal('0.00'),
                half_collection=Decimal(str(parts[20])) if parts[20] else Decimal('0.00'),
                st_collection=Decimal(str(parts[21])) if parts[21] else Decimal('0.00'),
                luggage_collection=Decimal(str(parts[22])) if parts[22] else Decimal('0.00'),
                physical_collection=Decimal(str(parts[23])) if parts[23] else Decimal('0.00'),
                ladies_collection=Decimal(str(parts[24])) if parts[24] else Decimal('0.00'),
                senior_collection=Decimal(str(parts[25])) if parts[25] else Decimal('0.00'),

                # Other financial data
                adjust_collection=Decimal(str(parts[26])) if parts[26] else Decimal('0.00'),
                expense_amount=Decimal(str(parts[27])) if parts[27] else Decimal('0.00'),
                total_collection=Decimal(str(parts[28])) if parts[28] else Decimal('0.00'),

                # UPI payment data
                upi_ticket_count=int(parts[29]) if parts[29] else 0,
                upi_ticket_amount=Decimal(str(parts[30])) if parts[30] else Decimal('0.00'),
            )
        
            # Return success response to device
            # Extract first 32 characters for response
            response_chars = raw_payload[0:32]
            device_response = f'OK#SUCCESS#fn={response_chars}#'
            return HttpResponse(device_response,status=status.HTTP_201_CREATED,content_type="text/plain")

        # Handle duplicate entry (unique_together constraint)
        except IntegrityError:
            response_chars = raw_payload[0:32]
            device_response = f'OK#SUCCESS#fn={response_chars}#'
            return HttpResponse(device_response,status=status.HTTP_200_OK,content_type="text/plain")
    
    # Handle conversion errors (int/Decimal/datetime)
    except ValueError as e:
        logger.exception("Transaction parsing failed, ValueError")
        return HttpResponse(f"ERROR",status=status.HTTP_400_BAD_REQUEST,content_type="text/plain")
    
    # Handle any other unexpected errors
    except Exception as e:
        logger.exception("Transaction parsing failed")
        return HttpResponse(f"ERROR",status=status.HTTP_500_INTERNAL_SERVER_ERROR,content_type="text/plain")
    


@api_view(['GET'])
def get_all_trip_close_data(request):
    user = get_user_from_cookie(request)
    if not user:
        return Response({'error': 'Authentication required'}, status=status.HTTP_401_UNAUTHORIZED)
    try:
        pass
    except Exception as e:
        return Response({"message": "Data fetching failed"}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)