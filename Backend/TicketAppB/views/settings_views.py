import logging
from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response

from ..models import Currency, Settings, ETMDevice, DeviceSettings, SettingsProfile
from ..serializers import CurrencySerializer, SettingsSerializer, DeviceSettingsSerializer, SettingsProfileSerializer
from .utils import _get_authenticated_company_admin, _get_object_or_404


logger = logging.getLogger(__name__)


# ── Currency ──────────────────────────────────────────────────────────────────

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


# ── Settings ──────────────────────────────────────────────────────────────────

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
            return Response({'message': 'No settings found', 'data': None}, status=status.HTTP_200_OK)

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


# ── Device Settings ───────────────────────────────────────────────────────────

@api_view(['GET'])
def list_company_devices(request):
    """GET /masterdata/device-settings/devices — device picker list."""
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    devices = ETMDevice.objects.filter(company=company).order_by('id')
    data = [
        {
            'id':             d.id,
            'serial_number':  d.serial_number,
            'display_name':   d.display_name or d.serial_number,
            'device_type':    d.device_type,
            'licence_status': d.licence_status,
            'has_settings':   DeviceSettings.objects.filter(device=d).exists(),
        }
        for d in devices
    ]
    return Response({'message': 'Success', 'data': data}, status=status.HTTP_200_OK)


@api_view(['GET', 'PUT'])
def get_device_settings(request, device_id):
    """
    GET /masterdata/device-settings/<id>  — fetch device settings
    PUT /masterdata/device-settings/<id>  — create or update device settings
    """
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err

    try:
        device = ETMDevice.objects.get(id=device_id, company=company)
    except ETMDevice.DoesNotExist:
        return Response({'message': 'Device not found or not assigned to your company.'}, status=status.HTTP_404_NOT_FOUND)

    if request.method == 'GET':
        try:
            obj = DeviceSettings.objects.get(device=device)
            return Response({'message': 'Success', 'data': DeviceSettingsSerializer(obj).data}, status=status.HTTP_200_OK)
        except DeviceSettings.DoesNotExist:
            return Response({'message': 'No settings for this device yet.', 'data': None}, status=status.HTTP_200_OK)

    obj, created = DeviceSettings.objects.get_or_create(
        device=device,
        defaults={'company': company, 'created_by': user},
    )
    serializer = DeviceSettingsSerializer(obj, data=request.data, partial=True)
    if serializer.is_valid():
        serializer.save(updated_by=user)
        msg = 'Device settings created.' if created else 'Device settings updated.'
        return Response({'message': msg, 'data': serializer.data}, status=status.HTTP_200_OK)

    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


# ── Settings Profiles ──────────────────────────────────────────────────────────

_PROFILE_FIELDS = [
    'user_pwd', 'master_pwd',
    'half_per', 'con_per', 'phy_per', 'round_amt', 'luggage_unit_rate',
    'main_display', 'main_display2',
    'header1', 'header2', 'header3', 'footer1', 'footer2',
    'language_option', 'report_font',
    'st_fare_edit', 'st_max_amt', 'st_ratio', 'st_min_amt',
    'st_roundoff_enable', 'st_roundoff_amt',
    'roundoff', 'round_up', 'remove_ticket_flag', 'stage_font_flag',
    'next_fare_flag', 'odometer_entry', 'ticket_no_big_font',
    'crew_check', 'tripsend_enable', 'schedulesend_enable',
    'inspect_rpt', 'multiple_pass', 'simple_report', 'inspector_sms',
    'auto_shut_down', 'userpswd_enable', 'exp_enable',
    'stage_updation_msg', 'default_stage',
]


@api_view(['GET'])
def list_profiles(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err
    profiles = SettingsProfile.objects.filter(company=company).order_by('name')
    return Response({'message': 'Success', 'data': SettingsProfileSerializer(profiles, many=True).data}, status=status.HTTP_200_OK)


@api_view(['POST'])
def create_profile(request):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err
    serializer = SettingsProfileSerializer(data=request.data)
    if serializer.is_valid():
        serializer.save(company=company, created_by=user)
        return Response({'message': 'Profile created', 'data': serializer.data}, status=status.HTTP_201_CREATED)
    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['PUT', 'DELETE'])
def profile_detail(request, profile_id):
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err
    try:
        profile = SettingsProfile.objects.get(id=profile_id, company=company)
    except SettingsProfile.DoesNotExist:
        return Response({'message': 'Profile not found.'}, status=status.HTTP_404_NOT_FOUND)

    if request.method == 'DELETE':
        profile.delete()
        return Response({'message': 'Profile deleted.'}, status=status.HTTP_200_OK)

    serializer = SettingsProfileSerializer(profile, data=request.data, partial=True)
    if serializer.is_valid():
        serializer.save(updated_by=user)
        return Response({'message': 'Profile updated', 'data': serializer.data}, status=status.HTTP_200_OK)
    return Response({'message': 'Validation failed', 'errors': serializer.errors}, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
def apply_profile_to_device(request, profile_id, device_id):
    """Copy a profile's settings onto a device's DeviceSettings."""
    user, company, err = _get_authenticated_company_admin(request)
    if err:
        return err
    try:
        profile = SettingsProfile.objects.get(id=profile_id, company=company)
    except SettingsProfile.DoesNotExist:
        return Response({'message': 'Profile not found.'}, status=status.HTTP_404_NOT_FOUND)
    try:
        device = ETMDevice.objects.get(id=device_id, company=company)
    except ETMDevice.DoesNotExist:
        return Response({'message': 'Device not found.'}, status=status.HTTP_404_NOT_FOUND)

    obj, _ = DeviceSettings.objects.get_or_create(
        device=device,
        defaults={'company': company, 'created_by': user},
    )
    for field in _PROFILE_FIELDS:
        setattr(obj, field, getattr(profile, field))
    obj.updated_by = user
    obj.save()

    return Response(
        {'message': f'Profile "{profile.name}" applied to device.', 'data': DeviceSettingsSerializer(obj).data},
        status=status.HTTP_200_OK,
    )
