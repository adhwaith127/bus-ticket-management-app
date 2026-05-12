import logging
from datetime import datetime
from rest_framework import status
from ...models import TransactionData, TripCloseData
from django.http import JsonResponse
from .auth import get_user_from_cookie
from rest_framework.response import Response
from ...serializers.transactions import TicketDataSerializer, TripCloseDataSerializer
from rest_framework.decorators import api_view
from django.db import OperationalError
from django.utils.dateparse import parse_datetime
import pytz

logger = logging.getLogger('ticket.transactions')


@api_view(['GET'])
def get_all_transaction_data(request):
    """
    Fetch transaction data with support for cursor-based polling.
    
    Query Parameters:
    - from_date: Start date (YYYY-MM-DD) - required
    - to_date: End date (YYYY-MM-DD) - required
    - since: ISO timestamp for incremental updates (optional)
    """
    user = get_user_from_cookie(request)
    if not user:
        return Response({'error': 'Authentication required'}, status=status.HTTP_401_UNAUTHORIZED)

    try:
        # Get date range from query parameters
        from_date = request.GET.get('from_date')
        to_date = request.GET.get('to_date')
        since_timestamp = request.GET.get('since')  # For polling updates
        
        # Validate required parameters
        if not from_date or not to_date:
            return Response(
                {'error': 'from_date and to_date are required'}, 
                status=status.HTTP_400_BAD_REQUEST
            )
        
        # Base queryset filtered by user's company and date range
        if user.company:
            queryset = TransactionData.objects.filter(
                company_code=user.company,
                ticket_date__gte=from_date,
                ticket_date__lte=to_date
            )
        else:
            queryset = TransactionData.objects.none()
        
        # If 'since' parameter provided, filter for polling updates
        if since_timestamp:
            try:
                # Parse the timestamp - handle different formats
                # Format from DB: 2026-01-07 05:16:06.134
                since_dt = parse_datetime(since_timestamp)
                
                if since_dt is None:
                    # Try alternative parsing
                    since_dt = datetime.fromisoformat(since_timestamp.replace('Z', '+00:00'))
                
                if since_dt:
                    # Make sure we're comparing timezone-aware datetimes
                    if since_dt.tzinfo is None:
                        since_dt = pytz.UTC.localize(since_dt)
                    
                    # Filter for records created AFTER the cursor timestamp
                    queryset = queryset.filter(created_at__gt=since_dt)
                    
                    logger.info(f"Polling query: since={since_timestamp}")
                else:
                    logger.warning(f"Could not parse since timestamp: {since_timestamp}")
                    return Response({"message": "success", "data": []}, status=status.HTTP_200_OK)
                    
            except (ValueError, TypeError) as e:
                logger.warning(f"Invalid since timestamp: {since_timestamp}, error: {e}")
                return Response({"message": "success", "data": []}, status=status.HTTP_200_OK)
        
        # Order by created_at descending (newest first) for consistency
        # Frontend expects newest first
        queryset = queryset.order_by('-created_at')
        
        # Limit results to prevent huge responses
        queryset = queryset[:500]
        
        # Serialize data
        serializer = TicketDataSerializer(queryset, many=True)
        
        return Response({
            "message": "success", 
            "data": serializer.data,
            "count": len(serializer.data)
        }, status=status.HTTP_200_OK)
    
    except OperationalError:
        return Response({"message": "Error fetching data", "error": str(e)},status=status.HTTP_503_SERVICE_UNAVAILABLE)

    except Exception as e:
        logger.exception("Error fetching transaction data")
        return Response({"message": "Data fetching failed", "error": str(e)},status=status.HTTP_500_INTERNAL_SERVER_ERROR)


@api_view(['GET'])
def get_all_trip_close_data(request):
    """
    Fetch trip close data with support for cursor-based polling.
    
    Query Parameters:
    - from_date: Start date (YYYY-MM-DD) - required
    - to_date: End date (YYYY-MM-DD) - required
    - since: ISO timestamp for incremental updates (optional)
    """
    user = get_user_from_cookie(request)
    if not user:
        return JsonResponse({"error": "Authentication required"}, status=status.HTTP_401_UNAUTHORIZED)

    try:
        # Get date range from query parameters
        from_date = request.GET.get('from_date')
        to_date = request.GET.get('to_date')
        since_timestamp = request.GET.get('since')  # For polling updates
        
        # Validate required parameters
        if not from_date or not to_date:
            return JsonResponse(
                {'error': 'from_date and to_date are required'}, 
                status=status.HTTP_400_BAD_REQUEST
            )
        
        # Base queryset filtered by user's company and date range
        if user.company:
            queryset = TripCloseData.objects.filter(
                company_code=user.company,
                start_date__gte=from_date,
                start_date__lte=to_date
            )
        else:
            queryset = TripCloseData.objects.none()
        
        # If 'since' parameter provided, filter for polling updates
        if since_timestamp:
            try:
                # Parse the timestamp
                since_dt = parse_datetime(since_timestamp)

                if since_dt is None:
                    # Try alternative parsing
                    since_dt = datetime.fromisoformat(since_timestamp.replace('Z', '+00:00'))

                if since_dt:
                    # Make sure we're comparing timezone-aware datetimes
                    if since_dt.tzinfo is None:
                        since_dt = pytz.UTC.localize(since_dt)

                    # Filter for records created AFTER the cursor timestamp
                    queryset = queryset.filter(created_at__gt=since_dt)
                    
                    logger.info(f"Trip polling query: since={since_timestamp}")
                else:
                    logger.warning(f"Could not parse since timestamp: {since_timestamp}")
                    return JsonResponse({"message": "success", "data": []}, status=status.HTTP_200_OK)
                    
            except (ValueError, TypeError) as e:
                logger.warning(f"Invalid since timestamp: {since_timestamp}, error: {e}")
                return JsonResponse({"message": "success", "data": []}, status=status.HTTP_200_OK)
        
        # Order by created_at descending (newest first)
        queryset = queryset.order_by('-created_at')
        
        # Limit results to prevent huge responses
        queryset = queryset[:500]
        
        # Serialize data
        serializer = TripCloseDataSerializer(queryset, many=True)
        
        return JsonResponse({
            "message": "success", 
            "data": serializer.data,
            "count": len(serializer.data)
        }, status=status.HTTP_200_OK)

    except OperationalError:
        return JsonResponse({"message": "Error fetching data"}, status=status.HTTP_503_SERVICE_UNAVAILABLE)

    except Exception as e:
        logger.exception("Error fetching trip close data")
        return JsonResponse({"message": f"{e}"}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)