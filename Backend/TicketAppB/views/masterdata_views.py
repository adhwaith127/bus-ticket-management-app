import logging
from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response

from ..models import (
    BusType, EmployeeType, Employee,
    Stage, Route, VehicleType,
    Currency, Settings, CrewAssignment,
)
from ..serializers import (
    BusTypeSerializer, EmployeeTypeSerializer, EmployeeSerializer,
    StageSerializer, RouteSerializer, VehicleTypeSerializer,
    CurrencySerializer, SettingsSerializer, CrewAssignmentSerializer,
)
from .auth_views import get_user_from_cookie

logger = logging.getLogger(__name__)


# =============================================================================
# SHARED HELPERS
# These two functions are used at the top of every single view.
# Keeping them here avoids repeating the same 10 lines everywhere.
# =============================================================================

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


# =============================================================================
# SECTION 1 — BUS TYPE
# Simple CRUD. No FK dependencies from frontend side.
# =============================================================================

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


# =============================================================================
# SECTION 2 — EMPLOYEE TYPE
# Same simple pattern as BusType.
# =============================================================================

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


# =============================================================================
# SECTION 3 — EMPLOYEE
# Has a FK to EmployeeType.
# Two things to notice here:
#   1. context={'company': company} is passed to the serializer so that
#      validate_emp_type() can confirm the chosen type belongs to this company.
#   2. is_deleted is a soft delete flag — no actual DELETE endpoint.
#      To "delete" an employee the admin sets is_deleted=True via the update endpoint.
# =============================================================================

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


# =============================================================================
# SECTION 4 — STAGE
# Simple CRUD. is_deleted is a soft delete flag here too.
# =============================================================================

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


# =============================================================================
# SECTION 5 — ROUTE
# Has a FK to BusType (company-validated via serializer context).
# is_deleted soft delete same as Employee and Stage.
# RouteStage inline management is skipped for now — ready to add later.
# =============================================================================

@api_view(['GET'])
def get_routes(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    show_deleted = request.query_params.get('show_deleted', 'false').lower() == 'true'
    qs = Route.objects.filter(company=company)
    if not show_deleted:
        qs = qs.filter(is_deleted=False)

    qs = qs.select_related('bus_type').order_by('id')
    serializer = RouteSerializer(qs, many=True)
    return Response({'message': 'Success', 'data': serializer.data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_route(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    serializer = RouteSerializer(data=request.data, context={'company': company})
    if serializer.is_valid():
        serializer.save(company=company, created_by=user)
        return Response({'message': 'Route created successfully', 'data': serializer.data}, status=status.HTTP_201_CREATED)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


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
        serializer.save(updated_by=user)
        return Response({'message': 'Route updated successfully', 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# =============================================================================
# SECTION 6 — VEHICLE TYPE
# Has a FK to BusType (company-validated via serializer context).
# is_deleted soft delete same as above.
# =============================================================================

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


# =============================================================================
# SECTION 7 — CURRENCY
# Simple CRUD. No FK dependencies.
# =============================================================================

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


# =============================================================================
# SECTION 8 — SETTINGS
# Special case — OneToOne with Company means there is exactly one Settings
# record per company. So there's no list, no create button.
#
# GET: Try to fetch the settings. If none exists yet, return empty defaults.
# PUT: Use get_or_create to either update existing or create fresh on first save.
#
# _get_object_or_404 is NOT used here because we want graceful empty defaults
# instead of a 404 when settings haven't been set up yet.
# =============================================================================

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


# =============================================================================
# SECTION 9 — CREW ASSIGNMENT
# Links driver + conductor + cleaner + vehicle.
# Key difference from other views: before saving, we validate that:
#   - driver's emp_type code is 'DRIVER'
#   - conductor's emp_type code is 'CONDUCTOR' (if provided)
#   - cleaner's emp_type code is 'CLEANER' (if provided)
# This validation lives in the VIEW (not serializer) because it requires
# fetching related objects and checking a nested field — easier to read here.
# All 4 FK objects are also confirmed to belong to this company.
# =============================================================================

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


# =============================================================================
# SECTION 10 — DROPDOWN DATA ENDPOINTS
# These are lightweight GET-only endpoints used by the frontend to populate
# dropdown menus in forms. For example:
#   - Employee form needs a dropdown of EmployeeTypes
#   - Vehicle form needs a dropdown of BusTypes
#   - CrewAssignment form needs dropdowns for drivers, conductors, vehicles
# All filtered to this company, minimal fields only.
# =============================================================================

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
