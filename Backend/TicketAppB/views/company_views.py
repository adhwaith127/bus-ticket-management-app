from rest_framework import status
from rest_framework.response import Response
from rest_framework.decorators import api_view
from ..models import Company
from ..serializers import CompanySerializer
from django.contrib.auth import get_user_model
from .auth_views import get_user_from_cookie
import requests
import time
from django.conf import settings


User=get_user_model()


# Build payload for license server registration
def build_license_registration_payload(company):

    return {
        "CustomerName": company.company_name,
        "PhoneNumber": company.contact_number,
        "CustomerEmail": company.company_email,
        "GSTNumber": company.gst_number or "",
        "CustomerContactPerson": company.contact_person,
        "CustomerContact": company.contact_number,
        "CustomerAddress": company.address,
        "CustomerAddress2": company.address_2 or "",
        "CustomerState": company.state,
        "CustomerCity": company.city,
        "DeviceModel": "Windows",
        "DeviceIdentifier1": company.company_name,
        "DeviceType": 1,
        "Version": getattr(settings, 'APP_VERSION', '1.0.0'),
        "ProjectName": getattr(settings, 'PROJECT_NAME', 'Bus Ticketing System')
    }


# Register company with external license server
def register_with_license_server(company):

    payload = build_license_registration_payload(company)
    
    try:
        response = requests.post(
            settings.PRODUCT_REGISTRATION_URL,
            json=payload,
            timeout=30
        )
        response.raise_for_status()
        
        data = response.json()
        
        # Returns customer_id on success
        if data.get('status') == 'Success' and data.get('CustomerId'):
            return {
                'success': True,
                'customer_id': data['CustomerId']
            }
        else:
            return {
                'success': False,
                'error': 'Registration failed: Invalid response from license server'
            }
    
    except requests.exceptions.RequestException as e:
        return {
            'success': False,
            'error': f'License server connection error: {str(e)}'
        }



# Poll license server for authentication approval
def poll_license_authentication(customer_id, timeout_seconds=300, interval_seconds=3):

    payload = {"CustomerId": customer_id}
    start_time = time.time()
    
    # Checks every 3 seconds for up to 5 minutes
    while time.time() - start_time < timeout_seconds:
        try:
            response = requests.post(
                settings.PRODUCT_AUTH_URL,
                json=payload,
                timeout=30
            )
            response.raise_for_status()
            
            # Returns authentication data when approved
            data = response.json()
            auth_status = data.get('Authenticationstatus', '')
            
            # Success case
            if auth_status == 'Approve':
                return {
                    'success': True,
                    'status': 'Approve',
                    'data': data
                }
            
            # Expired license
            if 'expired' in auth_status.lower():
                return {
                    'success': True,
                    'status': 'Expired',
                    'data': data
                }
            
            # Blocked
            if auth_status == 'Block':
                return {
                    'success': True,
                    'status': 'Block',
                    'data': data
                }
            
            # Still waiting - continue polling
            if 'waiting' in auth_status.lower() or auth_status == 'Pending':
                time.sleep(interval_seconds)
                continue
            
            # Unknown status - treat as error
            return {
                'success': False,
                'error': f'Unexpected authentication status: {auth_status}'
            }
        
        except requests.exceptions.RequestException as e:
            return {
                'success': False,
                'error': f'Authentication polling error: {str(e)}'
            }
    
    # Timeout
    return {
        'success': False,
        'error': 'Authentication timeout - license approval took too long'
    }



@api_view(['POST'])
def validate_company_license(request, pk):
    """
    Main endpoint for license validation
    
    Flow:
    1. Check if company exists
    2. Register with license server (if no company_id)
    3. Save company_id to database
    4. Poll license server for authentication
    5. Update company with license details
    6. Return updated company data
    """
    user = get_user_from_cookie(request)
    if not user:
        return Response(
            {'error': 'Authentication required'}, 
            status=status.HTTP_401_UNAUTHORIZED
        )
    
    # Step 1: Get company
    try:
        company = Company.objects.get(pk=pk)
    except Company.DoesNotExist:
        return Response(
            {"message": "Company not found"}, 
            status=status.HTTP_404_NOT_FOUND
        )
    
    # Step 2: Register with license server (if needed)
    if not company.company_id:
        registration_result = register_with_license_server(company)
        
        if not registration_result['success']:
            return Response(
                {
                    "message": "License registration failed",
                    "error": registration_result['error']
                },
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )
        
        # Step 3: Save company_id
        company.company_id = registration_result['customer_id']
        company.save()
    
    # Step 4: Poll for authentication
    auth_result = poll_license_authentication(company.company_id)
    
    if not auth_result['success']:
        return Response(
            {
                "message": "License authentication failed",
                "error": auth_result['error']
            },
            status=status.HTTP_500_INTERNAL_SERVER_ERROR
        )
    
    # Step 5: Update company with license details
    auth_data = auth_result.get('data', {})
    auth_status = auth_result['status']
    
    # Map authentication status to our model
    if auth_status == 'Approve':
        company.authentication_status = Company.AuthStatus.APPROVED
    elif auth_status == 'Expired':
        company.authentication_status = Company.AuthStatus.EXPIRED
    elif auth_status == 'Block':
        company.authentication_status = Company.AuthStatus.BLOCKED
    
    # Update license details (only if approved)
    if auth_status == 'Approve':
        company.product_registration_id = auth_data.get('ProductRegistrationId')
        company.unique_identifier = auth_data.get('UniqueIDentifier')
        company.product_from_date = auth_data.get('ProductFromDate')
        company.product_to_date = auth_data.get('ProductToDate')
        company.project_code = auth_data.get('ProjectCode')
        
        # Calculate device_count and branch_count if needed
        # You can customize this logic based on your requirements
        company.device_count = auth_data.get('TotalCount', 0)
        company.branch_count = auth_data.get('OutletCount', 0)
    
    company.save()
    
    # Step 6: Return updated data
    serializer = CompanySerializer(company)
    
    return Response(
        {
            "message": f"License validation completed - Status: {auth_status}",
            "data": serializer.data
        },
        status=status.HTTP_200_OK
    )



@api_view(['GET'])
def all_company_data(request):
    # verify user
    user = get_user_from_cookie(request)
    if not user:
        return Response(
            {'error': 'Authentication required'}, 
            status=status.HTTP_401_UNAUTHORIZED
        )
    
    companies = Company.objects.all().order_by('-id')
    serializer = CompanySerializer(companies, many=True)
    
    return Response(
        {
            "message": "Success",
            "data": serializer.data
        },
        status=status.HTTP_200_OK
    )



@api_view(['POST'])
def create_company(request):
    # verify user
    user = get_user_from_cookie(request)
    if not user:
        return Response(
            {'error': 'Authentication required'}, 
            status=status.HTTP_401_UNAUTHORIZED
        )
    
    if not request.data:
        return Response(
            {"message": "No input received"},
            status=status.HTTP_400_BAD_REQUEST
        )
    
    serializer = CompanySerializer(data=request.data)
    
    # Initial authentication_status will be 'Pending'
    if serializer.is_valid():
        company = serializer.save()
        return Response(
            {
                "message": "Successfully added company",
                "data": serializer.data
            },
            status=status.HTTP_201_CREATED
        )
    
    return Response(
        {
            "message": "Validation failed",
            "errors": serializer.errors
        },
        status=status.HTTP_400_BAD_REQUEST
    )


# Update existing company details
# Cannot update license-related fields directly (use validate_license endpoint)
@api_view(['PUT'])
def update_company_details(request, pk):
    # verify user
    user = get_user_from_cookie(request)
    if not user:
        return Response(
            {'error': 'Authentication required'}, 
            status=status.HTTP_401_UNAUTHORIZED
        )
    
    try:
        company = Company.objects.get(pk=pk)
    except Company.DoesNotExist:
        return Response(
            {"message": "Company not found"}, 
            status=status.HTTP_404_NOT_FOUND
        )
    
    serializer = CompanySerializer(company, data=request.data, partial=True)
    
    if serializer.is_valid():
        serializer.save()
        return Response(
            {
                "message": "Company updated successfully", 
                "data": serializer.data
            },
            status=status.HTTP_200_OK
        )
    
    return Response(
        {
            "message": "Validation failed", 
            "errors": serializer.errors
        },
        status=status.HTTP_400_BAD_REQUEST
    )