import logging
from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response

from ..models import (
    BusType, EmployeeType, Employee,
    Stage, Route, VehicleType,
    Currency, Settings, CrewAssignment,RouteStage,RouteBusType,Fare
)
from ..serializers import (
    BusTypeSerializer, EmployeeTypeSerializer, EmployeeSerializer,
    StageSerializer, RouteSerializer, VehicleTypeSerializer,
    CurrencySerializer, SettingsSerializer, CrewAssignmentSerializer
)
from .auth_views import get_user_from_cookie


logger = logging.getLogger(__name__)


# SHARED HELPERS
# These two functions are used at the top of every single view.
# Keeping them here avoids repeating the same 10 lines everywhere.
def _get_authenticated_company_admin(request):
    """
    Returns (user, company, error_response).
    If anything is wrong, error_response is set and the view should return it immediately.
    If everything is fine, user and company are set and error_response is None.

    Usage in every view:
        user, company, err = _get_authenticated_company_admin(request)
        if err:
            return err
    """
    user = get_user_from_cookie(request)
    if not user:
        return None, None, Response(
            {'error': 'Authentication required'},
            status=status.HTTP_401_UNAUTHORIZED
        )
    if user.role != 'company_admin':
        return None, None, Response(
            {'error': 'Only company admins can access this.'},
            status=status.HTTP_403_FORBIDDEN
        )
    company = user.company
    if not company:
        return None, None, Response(
            {'error': 'No company mapped to this user.'},
            status=status.HTTP_400_BAD_REQUEST
        )
    return user, company, None


def _get_object_or_404(model, pk, company):
    """
    Tries to fetch a single object by pk that belongs to this company.
    Returns (object, error_response).
    If found: object is set, error_response is None.
    If not found: object is None, error_response is a 404 Response.

    The company filter here is important — it prevents one company admin
    from editing another company's data just by guessing a pk.
    """
    try:
        obj = model.objects.get(pk=pk, company=company)
        return obj, None
    except model.DoesNotExist:
        return None, Response(
            {'error': f'{model.__name__} not found.'},
            status=status.HTTP_404_NOT_FOUND
        )


# BUS TYPE
# Simple CRUD. No FK dependencies from frontend side.
@api_view(['GET'])
def get_bus_types(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    bus_types = BusType.objects.filter(company=company).order_by('id')
    serializer = BusTypeSerializer(bus_types, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_bus_type(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    serializer = BusTypeSerializer(data=request.data)
    if serializer.is_valid():
        serializer.save(company=company, created_by=user)
        return Response({'message': 'Bus type created successfully', 'data': serializer.data}, status=status.HTTP_201_CREATED)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['PUT'])
def update_bus_type(request, pk):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    obj, err = _get_object_or_404(BusType, pk, company)
    if err:
        return err

    serializer = BusTypeSerializer(obj, data=request.data, partial=True)
    if serializer.is_valid():
        serializer.save(updated_by=user)
        return Response({'message': 'Bus type updated successfully', 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# EMPLOYEE TYPE
# Same simple pattern as BusType.
@api_view(['GET'])
def get_employee_types(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    emp_types = EmployeeType.objects.filter(company=company).order_by('id')
    serializer = EmployeeTypeSerializer(emp_types, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_employee_type(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    serializer = EmployeeTypeSerializer(data=request.data)
    if serializer.is_valid():
        serializer.save(company=company, created_by=user)
        return Response({'message': 'Employee type created successfully', 'data': serializer.data}, status=status.HTTP_201_CREATED)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['PUT'])
def update_employee_type(request, pk):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    obj, err = _get_object_or_404(EmployeeType, pk, company)
    if err:
        return err

    serializer = EmployeeTypeSerializer(obj, data=request.data, partial=True)
    if serializer.is_valid():
        serializer.save(updated_by=user)
        return Response({'message': 'Employee type updated successfully', 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# EMPLOYEE
# Has a FK to EmployeeType.
# wo things to notice here:
#   1. context={'company': company} is passed to the serializer so that
#      validate_emp_type() can confirm the chosen type belongs to this company.
#   2. is_deleted is a soft delete flag — no actual DELETE endpoint.
#      To "delete" an employee the admin sets is_deleted=True via the update endpoint.
@api_view(['GET'])
def get_employees(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    # By default, exclude soft-deleted employees from the listing
    # Frontend can pass ?show_deleted=true to include them if needed
    show_deleted = request.query_params.get('show_deleted', 'false').lower() == 'true'
    qs = Employee.objects.filter(company=company)
    if not show_deleted:
        qs = qs.filter(is_deleted=False)

    qs = qs.select_related('emp_type').order_by('id')
    serializer = EmployeeSerializer(qs, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_employee(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    # Pass company in context so the serializer's validate_emp_type can use it
    serializer = EmployeeSerializer(data=request.data, context={'company': company})
    if serializer.is_valid():
        serializer.save(company=company, created_by=user)
        return Response({'message': 'Employee created successfully', 'data': serializer.data}, status=status.HTTP_201_CREATED)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['PUT'])
def update_employee(request, pk):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    obj, err = _get_object_or_404(Employee, pk, company)
    if err:
        return err

    serializer = EmployeeSerializer(obj, data=request.data, partial=True, context={'company': company})
    if serializer.is_valid():
        serializer.save(updated_by=user)
        return Response({'message': 'Employee updated successfully', 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# STAGE
# Simple CRUD. is_deleted is a soft delete flag here too.
@api_view(['GET'])
def get_stages(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    show_deleted = request.query_params.get('show_deleted', 'false').lower() == 'true'
    qs = Stage.objects.filter(company=company)
    if not show_deleted:
        qs = qs.filter(is_deleted=False)

    qs = qs.order_by('id')
    serializer = StageSerializer(qs, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_stage(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    serializer = StageSerializer(data=request.data)
    if serializer.is_valid():
        serializer.save(company=company, created_by=user)
        return Response({'message': 'Stage created successfully', 'data': serializer.data}, status=status.HTTP_201_CREATED)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['PUT'])
def update_stage(request, pk):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    obj, err = _get_object_or_404(Stage, pk, company)
    if err:
        return err

    serializer = StageSerializer(obj, data=request.data, partial=True)
    if serializer.is_valid():
        serializer.save(updated_by=user)
        return Response({'message': 'Stage updated successfully', 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# ROUTE
# Has a FK to BusType (company-validated via serializer context).
# s_deleted soft delete same as Employee and Stage.
# RouteStage inline management is skipped for now — ready to add later.

# UPDATED Route Views - REPLACE in masterdata_views.py
# Now handles both RouteStage AND RouteBusType inline
@api_view(['GET'])
def get_routes(request):
    """
    Fetch routes with nested route_stages and route_bus_types.
    """
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    show_deleted = request.query_params.get('show_deleted', 'false').lower() == 'true'
    qs = Route.objects.filter(company=company)
    if not show_deleted:
        qs = qs.filter(is_deleted=False)

    qs = qs.select_related('bus_type').prefetch_related(
        'route_stages__stage',
        'route_bus_types__bus_type'  # NEW: prefetch allowed bus types
    ).order_by('id')
    
    serializer = RouteSerializer(qs, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_route(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    serializer = RouteSerializer(data=request.data, context={'company': company})
    if serializer.is_valid():
        route = serializer.save(company=company, created_by=user)
        
        # ── Handle nested route_stages ──────────────────────────────────
        route_stages_data = request.data.get('route_stages', [])
        if route_stages_data:
            _save_route_stages(route, route_stages_data, company, user)
        
        # ── Handle nested allowed_bus_types ─────────────────────────────
        allowed_bus_type_ids = request.data.get('allowed_bus_types', [])
        if allowed_bus_type_ids:
            _save_route_bus_types(route, allowed_bus_type_ids, company, user)
        
        # Re-fetch route with stages and bus types to return complete data
        route_with_nested = Route.objects.prefetch_related(
            'route_stages__stage',
            'route_bus_types__bus_type'
        ).get(pk=route.id)
        return_serializer = RouteSerializer(route_with_nested)
        
        return Response(
            {'message': 'Route created successfully', 'data': return_serializer.data},
            status=status.HTTP_201_CREATED
        )

    return Response(
        {'message': 'Validation failed', 'errors': serializer.errors},
        status=status.HTTP_400_BAD_REQUEST
    )


@api_view(['PUT'])
def update_route(request, pk):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    obj, err = _get_object_or_404(Route, pk, company)
    if err:
        return err

    serializer = RouteSerializer(obj, data=request.data, partial=True, context={'company': company})
    if serializer.is_valid():
        route = serializer.save(updated_by=user)
        
        # ── Handle nested route_stages ──────────────────────────────────
        if 'route_stages' in request.data:
            route_stages_data = request.data.get('route_stages', [])
            RouteStage.objects.filter(route=route).delete()
            if route_stages_data:
                _save_route_stages(route, route_stages_data, company, user)
        
        # ── Handle nested allowed_bus_types ─────────────────────────────
        if 'allowed_bus_types' in request.data:
            allowed_bus_type_ids = request.data.get('allowed_bus_types', [])
            RouteBusType.objects.filter(route=route).delete()
            if allowed_bus_type_ids:
                _save_route_bus_types(route, allowed_bus_type_ids, company, user)
        
        # Re-fetch route with nested data
        route_with_nested = Route.objects.prefetch_related(
            'route_stages__stage',
            'route_bus_types__bus_type'
        ).get(pk=route.id)
        return_serializer = RouteSerializer(route_with_nested)
        
        return Response(
            {'message': 'Route updated successfully', 'data': return_serializer.data},
            status=status.HTTP_200_OK
        )

    return Response(
        {'message': 'Validation failed', 'errors': serializer.errors},
        status=status.HTTP_400_BAD_REQUEST
    )


# ── Helper function for RouteStage (EXISTING - no changes) ──────────────────
def _save_route_stages(route, stages_data, company, user):
    """
    Helper to bulk create RouteStage records.
    Validates that all stage IDs belong to this company.
    """
    from ..models import Stage
    
    stage_ids = [s['stage'] for s in stages_data if 'stage' in s]
    valid_stages = Stage.objects.filter(
        id__in=stage_ids,
        company=company
    ).values_list('id', flat=True)
    valid_stage_set = set(valid_stages)
    
    route_stages_to_create = []
    for stage_data in stages_data:
        stage_id = stage_data.get('stage')
        if not stage_id or stage_id not in valid_stage_set:
            continue
        
        route_stages_to_create.append(
            RouteStage(
                route=route,
                stage_id=stage_id,
                sequence_no=stage_data.get('sequence_no', 0),
                distance=stage_data.get('distance', 0),
                stage_local_lang=stage_data.get('stage_local_lang', ''),
                company=company,
                created_by=user
            )
        )
    
    if route_stages_to_create:
        RouteStage.objects.bulk_create(route_stages_to_create)


# ── NEW: Helper function for RouteBusType ────────────────────────────────────
def _save_route_bus_types(route, bus_type_ids, company, user):
    """
    Helper to bulk create RouteBusType records.
    Validates that all bus_type IDs belong to this company.
    
    Expected bus_type_ids format:
    [1, 3, 5]  # List of BusType IDs
    """
    from ..models import BusType
    
    # Validate all bus types belong to this company
    valid_bus_types = BusType.objects.filter(
        id__in=bus_type_ids,
        company=company
    ).values_list('id', flat=True)
    valid_bus_type_set = set(valid_bus_types)
    
    # Create RouteBusType records
    route_bus_types_to_create = []
    for bus_type_id in bus_type_ids:
        if bus_type_id not in valid_bus_type_set:
            continue
        
        route_bus_types_to_create.append(
            RouteBusType(
                route=route,
                bus_type_id=bus_type_id,
                company=company,
                created_by=user
            )
        )
    
    # Bulk create for efficiency
    if route_bus_types_to_create:
        RouteBusType.objects.bulk_create(route_bus_types_to_create)


# ── Dropdown endpoints (keep existing get_stages_dropdown) ───────────────────
# No changes needed to get_stages_dropdown or get_bus_types_dropdown

# VEHICLE TYPE
# Has a FK to BusType (company-validated via serializer context).
# s_deleted soft delete same as above.
@api_view(['GET'])
def get_vehicles(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    show_deleted = request.query_params.get('show_deleted', 'false').lower() == 'true'
    qs = VehicleType.objects.filter(company=company)
    if not show_deleted:
        qs = qs.filter(is_deleted=False)

    qs = qs.select_related('bus_type').order_by('id')
    serializer = VehicleTypeSerializer(qs, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_vehicle(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    serializer = VehicleTypeSerializer(data=request.data, context={'company': company})
    if serializer.is_valid():
        serializer.save(company=company, created_by=user)
        return Response({'message': 'Vehicle created successfully', 'data': serializer.data}, status=status.HTTP_201_CREATED)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['PUT'])
def update_vehicle(request, pk):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    obj, err = _get_object_or_404(VehicleType, pk, company)
    if err:
        return err

    serializer = VehicleTypeSerializer(obj, data=request.data, partial=True, context={'company': company})
    if serializer.is_valid():
        serializer.save(updated_by=user)
        return Response({'message': 'Vehicle updated successfully', 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# CURRENCY
# Simple CRUD. No FK dependencies.
@api_view(['GET'])
def get_currencies(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    currencies = Currency.objects.filter(company=company).order_by('id')
    serializer = CurrencySerializer(currencies, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_currency(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    serializer = CurrencySerializer(data=request.data)
    if serializer.is_valid():
        serializer.save(company=company, created_by=user)
        return Response({'message': 'Currency created successfully', 'data': serializer.data}, status=status.HTTP_201_CREATED)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['PUT'])
def update_currency(request, pk):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    obj, err = _get_object_or_404(Currency, pk, company)
    if err:
        return err

    serializer = CurrencySerializer(obj, data=request.data, partial=True)
    if serializer.is_valid():
        serializer.save(updated_by=user)
        return Response({'message': 'Currency updated successfully', 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# SETTINGS
# Special case — OneToOne with Company means there is exactly one Settings
# ecord per company. So there's no list, no create button.
#
# GET: Try to fetch the settings. If none exists yet, return empty defaults.
# PUT: Use get_or_create to either update existing or create fresh on first save.
#
# _get_object_or_404 is NOT used here because we want graceful empty defaults
# instead of a 404 when settings haven't been set up yet.
@api_view(['GET', 'PUT'])
def get_settings(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    if request.method == 'GET':
        try:
            settings_obj = Settings.objects.get(company=company)
            serializer = SettingsSerializer(settings_obj)
            return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)
        except Settings.DoesNotExist:
            # No settings yet — return empty so frontend shows defaults
            return Response({'message': 'No settings found', 'data': None}, status=status.HTTP_200_OK)

    # PUT: get_or_create to either update existing or create fresh on first save.
    settings_obj, created = Settings.objects.get_or_create(
        company=company,
        defaults={'created_by': user}
    )

    serializer = SettingsSerializer(settings_obj, data=request.data, partial=True)
    if serializer.is_valid():
        serializer.save(updated_by=user)
        msg = 'Settings created successfully' if created else 'Settings updated successfully'
        return Response({'message': msg, 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['PUT'])
def update_settings(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    # get_or_create: if settings exist, update them. If not, create them now.
    # created is a boolean Django returns telling us which path it took.
    settings_obj, created = Settings.objects.get_or_create(
        company=company,
        defaults={'created_by': user}   # only applied on first creation
    )

    serializer = SettingsSerializer(settings_obj, data=request.data, partial=True)
    if serializer.is_valid():
        serializer.save(updated_by=user)
        msg = 'Settings created successfully' if created else 'Settings updated successfully'
        return Response({'message': msg, 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# CREW ASSIGNMENT
# Links driver + conductor + cleaner + vehicle.
# ey difference from other views: before saving, we validate that:
#   - driver's emp_type code is 'DRIVER'
#   - conductor's emp_type code is 'CONDUCTOR' (if provided)
#   - cleaner's emp_type code is 'CLEANER' (if provided)
# This validation lives in the VIEW (not serializer) because it requires
# fetching related objects and checking a nested field — easier to read here.
# All 4 FK objects are also confirmed to belong to this company.
def _validate_crew_member(employee_id, expected_type_code, company, field_label):
    """
    Fetches an Employee by ID, confirms it belongs to this company,
    and confirms its emp_type_code matches the expected role.
    Returns (employee_object, error_message_or_None).
    """
    try:
        emp = Employee.objects.select_related('emp_type').get(pk=employee_id, company=company)
    except Employee.DoesNotExist:
        return None, f'{field_label} not found in your company.'

    if emp.emp_type.emp_type_code != expected_type_code:
        return None, f'Selected {field_label} is not of type {expected_type_code}.'

    return emp, None


@api_view(['GET'])
def get_crew_assignments(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    qs = CrewAssignment.objects.filter(company=company).select_related(
        'driver', 'conductor', 'cleaner', 'vehicle'
    ).order_by('id')
    serializer = CrewAssignmentSerializer(qs, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_crew_assignment(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    # --- Validate driver (required) ---
    driver_id = request.data.get('driver')
    if not driver_id:
        return Response({'error': 'Driver is required.'}, status=status.HTTP_400_BAD_REQUEST)
    driver, err_msg = _validate_crew_member(driver_id, 'DRIVER', company, 'Driver')
    if err_msg:
        return Response({'error': err_msg}, status=status.HTTP_400_BAD_REQUEST)

    # --- Validate conductor (optional) ---
    conductor = None
    conductor_id = request.data.get('conductor')
    if conductor_id:
        conductor, err_msg = _validate_crew_member(conductor_id, 'CONDUCTOR', company, 'Conductor')
        if err_msg:
            return Response({'error': err_msg}, status=status.HTTP_400_BAD_REQUEST)

    # --- Validate cleaner (optional) ---
    cleaner = None
    cleaner_id = request.data.get('cleaner')
    if cleaner_id:
        cleaner, err_msg = _validate_crew_member(cleaner_id, 'CLEANER', company, 'Cleaner')
        if err_msg:
            return Response({'error': err_msg}, status=status.HTTP_400_BAD_REQUEST)

    # --- Validate vehicle (required, must belong to this company) ---
    vehicle_id = request.data.get('vehicle')
    if not vehicle_id:
        return Response({'error': 'Vehicle is required.'}, status=status.HTTP_400_BAD_REQUEST)
    try:
        vehicle = VehicleType.objects.get(pk=vehicle_id, company=company)
    except VehicleType.DoesNotExist:
        return Response({'error': 'Vehicle not found in your company.'}, status=status.HTTP_400_BAD_REQUEST)

    serializer = CrewAssignmentSerializer(data=request.data)
    if serializer.is_valid():
        serializer.save(company=company, created_by=user)
        return Response({'message': 'Crew assignment created successfully', 'data': serializer.data}, status=status.HTTP_201_CREATED)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['PUT'])
def update_crew_assignment(request, pk):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    obj, err = _get_object_or_404(CrewAssignment, pk, company)
    if err:
        return err

    # Re-validate any crew members that are being changed in this update
    driver_id = request.data.get('driver')
    if driver_id:
        _, err_msg = _validate_crew_member(driver_id, 'DRIVER', company, 'Driver')
        if err_msg:
            return Response({'error': err_msg}, status=status.HTTP_400_BAD_REQUEST)

    conductor_id = request.data.get('conductor')
    if conductor_id:
        _, err_msg = _validate_crew_member(conductor_id, 'CONDUCTOR', company, 'Conductor')
        if err_msg:
            return Response({'error': err_msg}, status=status.HTTP_400_BAD_REQUEST)

    cleaner_id = request.data.get('cleaner')
    if cleaner_id:
        _, err_msg = _validate_crew_member(cleaner_id, 'CLEANER', company, 'Cleaner')
        if err_msg:
            return Response({'error': err_msg}, status=status.HTTP_400_BAD_REQUEST)

    vehicle_id = request.data.get('vehicle')
    if vehicle_id:
        try:
            VehicleType.objects.get(pk=vehicle_id, company=company)
        except VehicleType.DoesNotExist:
            return Response({'error': 'Vehicle not found in your company.'}, status=status.HTTP_400_BAD_REQUEST)

    serializer = CrewAssignmentSerializer(obj, data=request.data, partial=True)
    if serializer.is_valid():
        serializer.save(updated_by=user)
        return Response({'message': 'Crew assignment updated successfully', 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# DROPDOWN DATA ENDPOINTS
# These are lightweight GET-only endpoints used by the frontend to populate
# ropdown menus in forms. For example:
#   - Employee form needs a dropdown of EmployeeTypes
#   - Vehicle form needs a dropdown of BusTypes
#   - CrewAssignment form needs dropdowns for drivers, conductors, vehicles
# All filtered to this company, minimal fields only.
@api_view(['GET'])
def get_bus_types_dropdown(request):
    """Minimal bus type list for dropdown menus."""
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    data = list(
        BusType.objects.filter(company=company, is_active=True)
        .values('id', 'bustype_code', 'name')
        .order_by('name')
    )
    return Response({'message': 'Success', 'data': data}, status=status.HTTP_200_OK)


@api_view(['GET'])
def get_employee_types_dropdown(request):
    """Minimal employee type list for dropdown menus."""
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    data = list(
        EmployeeType.objects.filter(company=company)
        .values('id', 'emp_type_code', 'emp_type_name')
        .order_by('emp_type_name')
    )
    return Response({'message': 'Success', 'data': data}, status=status.HTTP_200_OK)


@api_view(['GET'])
def get_employees_by_type_dropdown(request):
    """
    Returns employees filtered by emp_type_code query param.
    Used by CrewAssignment form to load drivers, conductors, cleaners separately.
    Example: /masterdata/dropdowns/employees/?type=DRIVER
    """
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    emp_type_code = request.query_params.get('type')
    qs = Employee.objects.filter(company=company, is_deleted=False).select_related('emp_type')
    if emp_type_code:
        qs = qs.filter(emp_type__emp_type_code=emp_type_code)

    data = list(qs.values('id', 'employee_code', 'employee_name').order_by('employee_name'))
    return Response({'message': 'Success', 'data': data}, status=status.HTTP_200_OK)


@api_view(['GET'])
def get_vehicles_dropdown(request):
    """Minimal vehicle list for dropdown menus."""
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    data = list(
        VehicleType.objects.filter(company=company, is_deleted=False)
        .values('id', 'bus_reg_num')
        .order_by('bus_reg_num')
    )
    return Response({'message': 'Success', 'data': data}, status=status.HTTP_200_OK)


@api_view(['GET'])
def get_routes(request):
    """
    No changes needed - the nested route_stages field in serializer
    automatically includes stages when fetching.
    """
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    show_deleted = request.query_params.get('show_deleted', 'false').lower() == 'true'
    qs = Route.objects.filter(company=company)
    if not show_deleted:
        qs = qs.filter(is_deleted=False)

    qs = qs.select_related('bus_type').prefetch_related('route_stages__stage').order_by('id')
    serializer = RouteSerializer(qs, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_route(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    serializer = RouteSerializer(data=request.data, context={'company': company})
    if serializer.is_valid():
        route = serializer.save(company=company, created_by=user)
        
        # ── Handle nested route_stages ──────────────────────────────────
        route_stages_data = request.data.get('route_stages', [])
        
        if route_stages_data:
            _save_route_stages(route, route_stages_data, company, user)
        
        # Re-fetch route with stages to return complete data
        route_with_stages = Route.objects.prefetch_related('route_stages__stage').get(pk=route.id)
        return_serializer = RouteSerializer(route_with_stages)
        
        return Response(
            {'message': 'Route created successfully', 'data': return_serializer.data},
            status=status.HTTP_201_CREATED
        )

    return Response(
        {'message': 'Validation failed', 'errors': serializer.errors},
        status=status.HTTP_400_BAD_REQUEST
    )


@api_view(['PUT'])
def update_route(request, pk):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    obj, err = _get_object_or_404(Route, pk, company)
    if err:
        return err

    serializer = RouteSerializer(obj, data=request.data, partial=True, context={'company': company})
    if serializer.is_valid():
        route = serializer.save(updated_by=user)
        
        # ── Handle nested route_stages ──────────────────────────────────
        # Only update stages if the frontend sends the route_stages field
        if 'route_stages' in request.data:
            route_stages_data = request.data.get('route_stages', [])
            
            # Delete existing stages and recreate
            # (Simpler than complex diffing logic for now)
            RouteStage.objects.filter(route=route).delete()
            
            if route_stages_data:
                _save_route_stages(route, route_stages_data, company, user)
        
        # Re-fetch route with stages
        route_with_stages = Route.objects.prefetch_related('route_stages__stage').get(pk=route.id)
        return_serializer = RouteSerializer(route_with_stages)
        
        return Response(
            {'message': 'Route updated successfully', 'data': return_serializer.data},
            status=status.HTTP_200_OK
        )

    return Response(
        {'message': 'Validation failed', 'errors': serializer.errors},
        status=status.HTTP_400_BAD_REQUEST
    )


# ── Helper function ──────────────────────────────────────────────────────
def _save_route_stages(route, stages_data, company, user):
    """
    Helper to bulk create RouteStage records.
    Validates that all stage IDs belong to this company.
    
    Expected stages_data format:
    [
        {"stage": 1, "sequence_no": 1, "distance": 0, "stage_local_lang": ""},
        {"stage": 2, "sequence_no": 2, "distance": 5.5, "stage_local_lang": ""},
        ...
    ]
    """
    from ..models import Stage
    
    # Validate all stages belong to this company
    stage_ids = [s['stage'] for s in stages_data if 'stage' in s]
    valid_stages = Stage.objects.filter(
        id__in=stage_ids,
        company=company
    ).values_list('id', flat=True)
    valid_stage_set = set(valid_stages)
    
    # Create RouteStage records
    route_stages_to_create = []
    for stage_data in stages_data:
        stage_id = stage_data.get('stage')
        
        # Skip if stage not valid
        if not stage_id or stage_id not in valid_stage_set:
            continue
        
        route_stages_to_create.append(
            RouteStage(
                route=route,
                stage_id=stage_id,
                sequence_no=stage_data.get('sequence_no', 0),
                distance=stage_data.get('distance', 0),
                stage_local_lang=stage_data.get('stage_local_lang', ''),
                company=company,
                created_by=user
            )
        )
    
    # Bulk create for efficiency
    if route_stages_to_create:
        RouteStage.objects.bulk_create(route_stages_to_create)


@api_view(['GET'])
def get_stages_dropdown(request):
    """Minimal stage list for dropdown menus in route forms."""
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    data = list(
        Stage.objects.filter(company=company, is_deleted=False)
        .values('id', 'stage_code', 'stage_name')
        .order_by('stage_name')
    )
    return Response({'message': 'Success', 'data': data}, status=status.HTTP_200_OK)


# @api_view(['GET'])
# def get_fare_editor(request, route_id):
#     """
#     Get fare data for a specific route.
#     Returns different structure based on fare_type:
#     - fare_type=1 (Table): Returns 1D array (row=1, col=1..N)
#     - fare_type=2 (Graph): Returns 2D matrix (row=1..N, col=1..N upper triangular)
#     """
#     user, company, err = _get_authenticated_company_admin(request)
#     if err:
#         return err
    
#     # Get route
#     route, err = _get_object_or_404(Route, route_id, company)
#     if err:
#         return err
    
#     # Get route stages (for both modes)
#     stages = route.route_stages.select_related('stage').order_by('sequence_no')
#     stage_list = [{
#         'sequence_no': rs.sequence_no,
#         'stage_id': rs.stage.id,
#         'stage_code': rs.stage.stage_code,
#         'stage_name': rs.stage.stage_name,
#     } for rs in stages]
    
#     n_stages = len(stage_list)
    
#     # Get existing fares
#     fares = Fare.objects.filter(route=route).order_by('row', 'col')
    
#     # Build response based on fare_type
#     if route.fare_type == 1:
#         # TABLE FARE (1D) - row=1, col represents number of stages traveled
#         fare_dict = {f.col: f.fare_amount for f in fares if f.row == 1}
        
#         # Build 1D array: fare_list[i] = fare for traveling (i+1) stages
#         fare_list = [fare_dict.get(i+1, 0) for i in range(n_stages)]
        
#         return Response({
#             'message': 'Success',
#             'data': {
#                 'route': {
#                     'id': route.id,
#                     'route_code': route.route_code,
#                     'route_name': route.route_name,
#                     'fare_type': route.fare_type,
#                 },
#                 'stages': stage_list,
#                 'fare_type_name': 'Table Fare (Distance-Based)',
#                 'fare_list': fare_list,  # 1D array for Table Fare
#             }
#         }, status=status.HTTP_200_OK)
    
#     else:  # fare_type == 2
#         # GRAPH FARE (2D Matrix) - row=origin, col=destination
#         fare_dict = {(f.row, f.col): f.fare_amount for f in fares}
        
#         # Build 2D matrix (upper triangular)
#         fare_matrix = []
#         for row_idx in range(n_stages):
#             row_data = []
#             for col_idx in range(n_stages):
#                 row_seq = stage_list[row_idx]['sequence_no']
#                 col_seq = stage_list[col_idx]['sequence_no']
#                 fare_amount = fare_dict.get((row_seq, col_seq), 0)
#                 row_data.append(fare_amount)
#             fare_matrix.append(row_data)
        
#         return Response({
#             'message': 'Success',
#             'data': {
#                 'route': {
#                     'id': route.id,
#                     'route_code': route.route_code,
#                     'route_name': route.route_name,
#                     'fare_type': route.fare_type,
#                 },
#                 'stages': stage_list,
#                 'fare_type_name': 'Graph Fare (Point-to-Point)',
#                 'fare_matrix': fare_matrix,  # 2D matrix for Graph Fare
#             }
#         }, status=status.HTTP_200_OK)


# @api_view(['POST'])
# def update_fare_table(request, route_id):
#     """
#     Bulk update fares for a route.
#     Accepts different payload based on fare_type:
#     - fare_type=1 (Table): { "fare_list": [10, 20, 30, ...] }
#     - fare_type=2 (Graph): { "fare_matrix": [[0,10,20],[10,0,12],[20,12,0]] }
#     """
#     user, company, err = _get_authenticated_company_admin(request)
#     if err:
#         return err
    
#     route, err = _get_object_or_404(Route, route_id, company)
#     if err:
#         return err
    
#     # Get route stages for validation
#     stages = route.route_stages.order_by('sequence_no')
#     n_stages = stages.count()
#     stage_list = list(stages.values_list('sequence_no', flat=True))
    
#     # Delete existing fares for this route
#     Fare.objects.filter(route=route).delete()
    
#     fares_to_create = []
    
#     # ── Handle Table Fare (fare_type=1) ─────────────────────────────────────
#     if route.fare_type == 1:
#         fare_list = request.data.get('fare_list', [])
        
#         if not fare_list or not isinstance(fare_list, list):
#             return Response(
#                 {'message': 'Invalid fare_list format. Expected 1D array.'},
#                 status=status.HTTP_400_BAD_REQUEST
#             )
        
#         if len(fare_list) != n_stages:
#             return Response(
#                 {'message': f'fare_list must have {n_stages} entries (number of stages).'},
#                 status=status.HTTP_400_BAD_REQUEST
#             )
        
#         # Create fare records: row=1, col=stage_count, fare_amount=fare
#         for col_idx, fare_amount in enumerate(fare_list):
#             if fare_amount == 0:
#                 continue  # Skip zero fares
            
#             fares_to_create.append(
#                 Fare(
#                     route=route,
#                     row=1,  # Always row=1 for Table Fare
#                     col=col_idx + 1,  # col = number of stages traveled (1, 2, 3, ...)
#                     fare_amount=int(fare_amount),
#                     route_name=route.route_name,
#                     company=company,
#                     created_by=user
#                 )
#             )
    
#     # ── Handle Graph Fare (fare_type=2) ─────────────────────────────────────
#     else:  # fare_type == 2
#         fare_matrix = request.data.get('fare_matrix', [])
        
#         if not fare_matrix or not isinstance(fare_matrix, list):
#             return Response(
#                 {'message': 'Invalid fare_matrix format. Expected 2D array.'},
#                 status=status.HTTP_400_BAD_REQUEST
#             )
        
#         if len(fare_matrix) != n_stages:
#             return Response(
#                 {'message': f'fare_matrix must have {n_stages} rows (number of stages).'},
#                 status=status.HTTP_400_BAD_REQUEST
#             )
        
#         for i, row in enumerate(fare_matrix):
#             if len(row) != n_stages:
#                 return Response(
#                     {'message': f'Row {i} must have {n_stages} columns.'},
#                     status=status.HTTP_400_BAD_REQUEST
#                 )
        
#         # Create fare records: row=from_stage, col=to_stage, fare_amount=fare
#         for i, row in enumerate(fare_matrix):
#             for j, fare_amount in enumerate(row):
#                 if fare_amount == 0:
#                     continue  # Skip zero fares
                
#                 row_seq = stage_list[i]
#                 col_seq = stage_list[j]
                
#                 fares_to_create.append(
#                     Fare(
#                         route=route,
#                         row=row_seq,  # Origin stage sequence
#                         col=col_seq,  # Destination stage sequence
#                         fare_amount=int(fare_amount),
#                         route_name=route.route_name,
#                         company=company,
#                         created_by=user
#                     )
#                 )
    
#     # Bulk create
#     if fares_to_create:
#         Fare.objects.bulk_create(fares_to_create)
    
#     fare_type_name = 'Table Fare' if route.fare_type == 1 else 'Graph Fare'
#     return Response(
#         {'message': f'{fare_type_name} updated successfully. {len(fares_to_create)} fare records created.'},
#         status=status.HTTP_200_OK
#     )




# =============================================================================
# KEY CHANGES:
# =============================================================================
# 
# 1. In get_fare_editor() for Graph Fare:
#    - Changed: fare_dict.get((row_seq, col_seq), 0)
#    - To:      fare_dict.get((row_idx + 1, col_idx + 1), 0)
#    - Why:     Fare row/col are always 1-indexed, regardless of sequence_no
# 
# 2. In update_fare_table() for Graph Fare:
#    - Changed: row=row_seq, col=col_seq
#    - To:      row=i + 1, col=j + 1
#    - Why:     Always store fares with 1-based indexing for consistency
# 
# 3. Added empty stages check to prevent crashes
# 
# =============================================================================


@api_view(['GET'])
def get_fare_editor(request, route_id):
    """
    Get fare data for a specific route.
    
    CRITICAL FIX: RouteStage.sequence_no can start at 0, but Fare row/col typically start at 1.
    We normalize by using the index position in the stage list (0-based) and adding 1
    when looking up fares.
    """
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err
    
    route, err = _get_object_or_404(Route, route_id, company)
    if err:
        return err
    
    stages = route.route_stages.select_related('stage').order_by('sequence_no')
    
    stage_list = [{
        'sequence_no': rs.sequence_no,
        'stage_id': rs.stage.id,
        'stage_code': rs.stage.stage_code,
        'stage_name': rs.stage.stage_name,
    } for rs in stages]
    
    # CRITICAL: Print stages to debug sequence_no and stage IDs
    print(stages)
    print(stage_list)
    
    n_stages = len(stage_list)
    
    if n_stages == 0:
        return Response({
            'message': 'No stages defined for this route',
            'data': {
                'route': {
                    'id': route.id,
                    'route_code': route.route_code,
                    'route_name': route.route_name,
                    'fare_type': route.fare_type,
                },
                'stages': [],
                'fare_type_name': 'Table Fare' if route.fare_type == 1 else 'Graph Fare',
                'fare_list': [],
                'fare_matrix': [],
            }
        }, status=status.HTTP_200_OK)
    
    fares = Fare.objects.filter(route=route).order_by('row', 'col')
    
    # Debug: Print fares to check row/col values and amounts
    print(fares)  
    
    # ── Table Fare (fare_type=1) ────────────────────────────────────────────
    if route.fare_type == 1:
        fare_dict = {f.col: f.fare_amount for f in fares if f.row == 1}
        
        # Build 1D array: fare_list[i] = fare for traveling (i+1) stages
        # col values in DB are 1-indexed (1, 2, 3, ...)
        fare_list = [fare_dict.get(i + 1, 0) for i in range(n_stages)]
        
        return Response({
            'message': 'Success',
            'data': {
                'route': {
                    'id': route.id,
                    'route_code': route.route_code,
                    'route_name': route.route_name,
                    'fare_type': route.fare_type,
                },
                'stages': stage_list,
                'fare_type_name': 'Table Fare (Distance-Based)',
                'fare_list': fare_list,
            }
        }, status=status.HTTP_200_OK)
    
    # ── Graph Fare (fare_type=2) ────────────────────────────────────────────
    else:
        fare_dict = {(f.row, f.col): f.fare_amount for f in fares}
        
        # Build 2D matrix
        # CRITICAL: Fare row/col are 1-indexed (1,1), (1,2), (2,2), ...
        # We use (i+1, j+1) to look them up
        fare_matrix = []
        for row_idx in range(n_stages):
            row_data = []
            for col_idx in range(n_stages):
                # Use 1-based indexing for fare lookup
                fare_amount = fare_dict.get((row_idx + 1, col_idx + 1), 0)
                row_data.append(fare_amount)
            fare_matrix.append(row_data)
        
        return Response({
            'message': 'Success',
            'data': {
                'route': {
                    'id': route.id,
                    'route_code': route.route_code,
                    'route_name': route.route_name,
                    'fare_type': route.fare_type,
                },
                'stages': stage_list,
                'fare_type_name': 'Graph Fare (Point-to-Point)',
                'fare_matrix': fare_matrix,
            }
        }, status=status.HTTP_200_OK)


@api_view(['POST'])
def update_fare_table(request, route_id):
    """
    Bulk update fares for a route.
    
    CRITICAL FIX: Always store Fare records with row/col starting at 1
    (even though RouteStage.sequence_no might start at 0).
    """
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err
    
    route, err = _get_object_or_404(Route, route_id, company)
    if err:
        return err
    
    stages = route.route_stages.order_by('sequence_no')
    n_stages = stages.count()
    
    if n_stages == 0:
        return Response(
            {'message': 'No stages defined for this route. Add stops before creating fares.'},
            status=status.HTTP_400_BAD_REQUEST
        )
    
    # Delete existing fares
    Fare.objects.filter(route=route).delete()
    
    fares_to_create = []
    
    # ── Handle Table Fare (fare_type=1) ─────────────────────────────────────
    if route.fare_type == 1:
        fare_list = request.data.get('fare_list', [])
        
        if not fare_list or not isinstance(fare_list, list):
            return Response(
                {'message': 'Invalid fare_list format. Expected 1D array.'},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        if len(fare_list) != n_stages:
            return Response(
                {'message': f'fare_list must have {n_stages} entries (number of stages).'},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        # Create fare records: row=1, col=1,2,3,...
        for col_idx, fare_amount in enumerate(fare_list):
            if fare_amount == 0:
                continue
            
            fares_to_create.append(
                Fare(
                    route=route,
                    row=1,
                    col=col_idx + 1,  # 1-indexed: 1, 2, 3, ...
                    fare_amount=int(fare_amount),
                    route_name=route.route_name,
                    company=company,
                    created_by=user
                )
            )
    
    # ── Handle Graph Fare (fare_type=2) ─────────────────────────────────────
    else:
        fare_matrix = request.data.get('fare_matrix', [])
        
        if not fare_matrix or not isinstance(fare_matrix, list):
            return Response(
                {'message': 'Invalid fare_matrix format. Expected 2D array.'},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        if len(fare_matrix) != n_stages:
            return Response(
                {'message': f'fare_matrix must have {n_stages} rows.'},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        for i, row in enumerate(fare_matrix):
            if len(row) != n_stages:
                return Response(
                    {'message': f'Row {i} must have {n_stages} columns.'},
                    status=status.HTTP_400_BAD_REQUEST
                )
        
        # Create fare records: row=1..N, col=1..N (1-indexed)
        for i, row in enumerate(fare_matrix):
            for j, fare_amount in enumerate(row):
                if fare_amount == 0:
                    continue
                
                fares_to_create.append(
                    Fare(
                        route=route,
                        row=i + 1,  # 1-indexed: 1, 2, 3, ...
                        col=j + 1,  # 1-indexed: 1, 2, 3, ...
                        fare_amount=int(fare_amount),
                        route_name=route.route_name,
                        company=company,
                        created_by=user
                    )
                )
    
    # Bulk create
    if fares_to_create:
        Fare.objects.bulk_create(fares_to_create)
    
    fare_type_name = 'Table Fare' if route.fare_type == 1 else 'Graph Fare'
    return Response(
        {'message': f'{fare_type_name} updated successfully. {len(fares_to_create)} fare records created.'},
        status=status.HTTP_200_OK
    )