import struct
import logging
from django.http import HttpResponse, JsonResponse
from ..models import Settings, Route, Employee, VehicleType, ExpenseMaster
from .auth_views import get_user_from_cookie

logger = logging.getLogger(__name__)


# ─── Binary packing helpers ────────────────────────────────────────────────────

def _s(val, size):
    """Fixed-width ASCII string, null-padded, truncated to fit."""
    return (val or '').encode('ascii', errors='replace')[:size].ljust(size, b'\x00')

def _b(val):
    """Single unsigned byte (0–255)."""
    try:
        return bytes([max(0, min(255, int(float(val or 0))))])
    except (ValueError, TypeError):
        return b'\x00'

def _bool(val):
    return b'\x01' if val else b'\x00'

def _f(val):
    """4-byte little-endian float (IEEE 754)."""
    try:
        return struct.pack('<f', float(val or 0))
    except (ValueError, TypeError):
        return struct.pack('<f', 0.0)

def _i(val):
    """2-byte little-endian signed short."""
    try:
        return struct.pack('<h', int(val or 0))
    except (ValueError, TypeError):
        return struct.pack('<h', 0)


# ─── File packers ──────────────────────────────────────────────────────────────

def _pack_busdat(s):
    """
    Build BUS.DAT binary from Settings model.
    Structure: SETUP (~640 bytes) + HARDWARE_SETUP (64 bytes) = 704 bytes total.
    Matches VB6 Type SETUP + HARDWARE_SETUP definitions in mdFunctions.bas.
    """
    # ── SETUP section ──────────────────────────────────────────────────────────
    data  = _s(s.main_display, 18)
    data += _s(s.main_display2, 23)
    data += _s(s.header1, 32)
    data += _s(s.header2, 32)
    data += _s(s.header3, 32)
    data += _s(s.footer1, 32)
    data += _s(s.footer2, 32)
    data += b'\x00'                          # PaperFeed
    data += _s(s.palmtec_id, 6)
    data += b'\x00'                          # DefaultFull
    data += _b(s.half_per)
    data += _b(s.con_per)
    data += _f(s.st_max_amt)
    data += _f(s.st_min_con)
    data += _b(s.phy_per)
    data += b'\x00'                          # LuggageUnitRateEdit
    data += _f(s.luggage_unit_rate)
    data += _b(s.stage_updation_msg)
    data += b'\x00'                          # StageDisplayFont
    data += b'\x00'                          # UseDuplicate
    data += b'\x00'                          # UseDup1
    data += _bool(s.roundoff)
    data += _bool(s.round_up)
    data += _s(s.currency, 8)
    data += _i(s.round_amt)
    data += b'\x00'                          # ucbAdjust
    data += b'\x00'                          # ucbReviewPasswd
    data += b'\x00'                          # ucbReportPasswd
    data += b'\x00'                          # ucbSTFromStage
    data += _bool(s.st_fare_edit)
    data += _s(s.master_pwd, 11)             # cMasterClearPassword
    data += _b(s.report_flag)
    data += _bool(s.next_fare_flag)
    data += _b(s.stage_updation_msg)         # UpdateStageMsg
    data += _bool(s.remove_ticket_flag)
    data += _bool(s.stage_font_flag)
    data += b'\x00'                          # EnableStageDefault
    data += b'\x00'                          # PrinterSel
    data += _bool(s.odometer_entry)
    data += _bool(s.ticket_no_big_font)
    data += _bool(s.crew_check)
    data += _s(s.ph_no2, 13)                 # PhNo
    data += b'\x00'                          # TripSMS
    data += _bool(s.schedulesend_enable)     # ScheduleSMS
    data += b'\x00'                          # TicketRpt
    data += b'\x00'                          # Busno
    data += b'\x00'                          # Driver
    data += b'\x00'                          # Conductor
    data += _bool(s.inspect_rpt)
    data += b'\x00'                          # RepeatST
    data += b'\x01' if s.sendbill_enable == '1' else b'\x00'
    data += _bool(s.tripsend_enable)
    data += _bool(s.schedulesend_enable)
    data += _bool(s.sendpend)
    data += _s(s.ph_no2, 13)                 # PhNo2
    data += _s(s.access_point, 24)
    data += _s(s.dest_adds, 32)
    data += _s(s.username, 16)
    data += _s(s.password, 16)
    data += _s(s.uploadpath, 32)
    data += _s(s.downloadpath, 32)
    data += _s(s.http_url, 64)
    data += _bool(s.gprs_enable)
    data += b'\x00'                          # MsgPrompt
    data += b'\x01' if s.exp_enable == '1' else b'\x00'
    data += b'\x01' if s.smart_card == '1' else b'\x00'
    data += b'\x00'                          # Modomon
    data += b'\x01' if s.ftp_enable == '1' else b'\x00'
    data += _s('', 11)                       # RemovePswd
    data += b'\x00'                          # StageReport_E_D
    data += _bool(s.st_roundoff_enable)
    data += _i(s.st_roundoff_amt)
    data += _bool(s.simple_report)
    data += _b(s.report_font)
    data += _bool(s.multiple_pass)
    data += _bool(s.inspector_sms)
    data += b'\x00'                          # StageEntry
    data += _s(s.ph_no3, 13)
    data += _bool(s.auto_shut_down)
    data += _bool(s.userpswd_enable)
    data += b'\x00'                          # DieselEntryEnable
    data += b'\x00'                          # TripTimeEnable
    data += b'\x00'                          # TripCloseReport
    data += b'\x00'                          # ucPaperFeed
    data += b'\x00'                          # refund
    data += b'\x00'                          # shedule_close_rpt
    data += b'\x00'                          # ladis_per
    data += b'\x00'                          # seniar_per
    data += b'\x00' * 70                     # ucTemp padding

    # ── HARDWARE_SETUP section (64 bytes) ──────────────────────────────────────
    # Ptime (4 bytes): Hour, Min, Sec, Hundredths — zeroed (device sets its own clock)
    data += b'\x00' * 4
    # Pdate (4 bytes): Day, Month, Year(2 bytes) — zeroed
    data += b'\x00' * 4
    data += _s(s.master_pwd, 11)             # MSR_PSWD
    data += _s(s.user_pwd, 11)              # USR_PSWD
    data += _s('', 11)                       # SPR_PSWD (supervisor)
    data += b'\x80'                          # val_contrast  (default mid)
    data += b'\x80'                          # val_brightness
    data += b'\x00'                          # screensaver_onoff
    data += b'\x1E'                          # backlit_timer (30s default)
    data += b'\x00'                          # keyhitdelay
    data += b'\x00'                          # boarder_en
    data += b'\x00'                          # dooropen_alert
    data += b'\x00'                          # paperout_alert
    data += b'\x00'                          # ucHalfPagePrinter
    data += b'\x01'                          # buzz_onoff (on)
    data += b'\x00'                          # rs232_baud
    data += b'\x00'                          # ir_baud
    data += b'\x00'                          # rf_baud
    data += b'\x00'                          # connecting_medium
    data += b'\x00'                          # footer_stat
    data += _b(s.language_option)           # select_language
    data += b'\x00'                          # login_mode
    data += b'\x00'                          # ucKPLight_opt
    data += _i(0)                            # usShuntdownTime
    data += _b(s.language_option)           # LangNo
    data += b'\x00' * 2                     # ucTemp

    return data  # 704 bytes total


def _pack_routelst(route):
    """
    Build RouteLST binary record for one route (64 bytes).
    Matches VB6 Type RouteLST in mdFunctions.bas.
    """
    stage_count = route.route_stages.count()
    data  = _s(route.route_code, 5)
    data += _s(route.route_name, 25)
    data += _b(stage_count)
    data += _f(route.min_fare)
    data += _b(route.fare_type)
    data += _bool(route.half)
    data += _bool(route.conc)
    data += _bool(route.ph)
    data += _bool(route.luggage)
    data += _bool(route.adjust)
    data += _b(route.start_from)
    data += _b(route.bus_type.pk % 256)     # BusType byte ID
    data += _s(route.bus_type.name, 16)
    data += _f(0)                            # OptedKM
    data += _bool(route.pass_allow)
    return data  # 64 bytes


def _pack_stagelst(route_stages):
    """
    Build STAGE.LST binary for ordered route stages (16 bytes per stage).
    Matches VB6 Type STAGEDETAILS in mdFunctions.bas.
    """
    data = b''
    for rs in route_stages:
        data += _s(rs.stage.stage_name, 12)
        data += _f(rs.distance)
    return data  # 16 bytes × stage_count


def _pack_crewdat(employees):
    """
    Build CREW.DAT binary (32 bytes per employee).
    Matches VB6 Type CREWDET in mdFunctions.bas.
    """
    type_map = {'driver': 1, 'conductor': 2, 'cleaner': 3, 'inspector': 4}
    data = b''
    for emp in employees:
        type_byte = type_map.get(emp.emp_type.emp_type_name.lower(), 0)
        data += _s(emp.employee_name, 16)
        data += _s(emp.employee_code, 8)
        data += bytes([type_byte])
        data += _s(emp.password, 7)
    return data  # 32 bytes × employee_count


def _pack_expensedet(expenses):
    """
    Build EXPENSEDET.DAT binary (31 bytes per expense).
    expense_code(5) + expense_name(25) + palmtec_id byte(1)
    """
    data = b''
    for exp in expenses:
        try:
            palmtec_byte = max(0, min(255, int(exp.palmtec_id or 0)))
        except (ValueError, TypeError):
            palmtec_byte = 0
        data += _s(exp.expense_code, 5)
        data += _s(exp.expense_name, 25)
        data += bytes([palmtec_byte])
    return data  # 31 bytes × expense_count


def _pack_vehicledat(vehicles):
    """
    Build VEHICLE.DAT binary (33 bytes per vehicle).
    bus_reg_num(12) + bus_type_name(16) + bus_type_code(5)
    """
    data = b''
    for v in vehicles:
        data += _s(v.bus_reg_num, 12)
        data += _s(v.bus_type.name, 16)
        data += _s(v.bus_type.bustype_code, 5)
    return data  # 33 bytes × vehicle_count


# ─── Auth helper ───────────────────────────────────────────────────────────────

def _get_company(request):
    """Return company from JWT cookie, or None if unauthenticated."""
    try:
        user = get_user_from_cookie(request)
        if user and hasattr(user, 'company') and user.company:
            return user.company
        return None
    except Exception:
        return None


# ─── Views ─────────────────────────────────────────────────────────────────────

def get_routes_list(request):
    """
    GET /device/routes/
    Returns JSON list of {route_code, route_name} for APK route selection popup.
    """
    if request.method != 'GET':
        return JsonResponse({'message': 'Method not allowed'}, status=405)

    company = _get_company(request)
    if not company:
        return JsonResponse({'message': 'Unauthorized'}, status=401)

    routes = (
        Route.objects
        .filter(company=company, is_deleted=False)
        .values('route_code', 'route_name')
        .order_by('route_code')
    )
    return JsonResponse({'routes': list(routes)})


def get_schedule_file(request):
    """
    GET /device/schedule/?route_code=R01
    Returns binary: RouteLST record (64 bytes) + STAGE records (16 bytes each).
    APK sends this to device after operator selects a route.
    """
    if request.method != 'GET':
        return HttpResponse('METHOD_NOT_ALLOWED', status=405)

    company = _get_company(request)
    if not company:
        return HttpResponse('UNAUTHORIZED', status=401)

    route_code = request.GET.get('route_code', '').strip()
    if not route_code:
        return HttpResponse('MISSING_ROUTE_CODE', status=400)

    try:
        route = (
            Route.objects
            .select_related('bus_type')
            .prefetch_related('route_stages__stage')
            .get(company=company, route_code=route_code, is_deleted=False)
        )
    except Route.DoesNotExist:
        return HttpResponse('ROUTE_NOT_FOUND', status=404)

    route_stages = route.route_stages.select_related('stage').order_by('sequence_no')
    binary = _pack_routelst(route) + _pack_stagelst(route_stages)

    response = HttpResponse(binary, content_type='application/octet-stream')
    response['Content-Disposition'] = f'attachment; filename="schedule_{route_code}.bin"'
    return response


def get_settings_file(request):
    """
    GET /device/settings/
    Returns BUS.DAT binary (704 bytes).
    """
    if request.method != 'GET':
        return HttpResponse('METHOD_NOT_ALLOWED', status=405)

    company = _get_company(request)
    if not company:
        return HttpResponse('UNAUTHORIZED', status=401)

    try:
        s = Settings.objects.get(company=company)
    except Settings.DoesNotExist:
        return HttpResponse('SETTINGS_NOT_FOUND', status=404)

    binary = _pack_busdat(s)
    response = HttpResponse(binary, content_type='application/octet-stream')
    response['Content-Disposition'] = 'attachment; filename="BUS.DAT"'
    return response


def get_crew_file(request):
    """
    GET /device/crew/
    Returns CREW.DAT binary (32 bytes per employee).
    """
    if request.method != 'GET':
        return HttpResponse('METHOD_NOT_ALLOWED', status=405)

    company = _get_company(request)
    if not company:
        return HttpResponse('UNAUTHORIZED', status=401)

    employees = (
        Employee.objects
        .filter(company=company, is_deleted=False)
        .select_related('emp_type')
        .order_by('employee_code')
    )
    binary = _pack_crewdat(employees)
    response = HttpResponse(binary, content_type='application/octet-stream')
    response['Content-Disposition'] = 'attachment; filename="CREW.DAT"'
    return response


def get_vehicles_file(request):
    """
    GET /device/vehicles/
    Returns VEHICLE.DAT binary (33 bytes per vehicle).
    """
    if request.method != 'GET':
        return HttpResponse('METHOD_NOT_ALLOWED', status=405)

    company = _get_company(request)
    if not company:
        return HttpResponse('UNAUTHORIZED', status=401)

    vehicles = (
        VehicleType.objects
        .filter(company=company, is_deleted=False)
        .select_related('bus_type')
        .order_by('bus_reg_num')
    )
    binary = _pack_vehicledat(vehicles)
    response = HttpResponse(binary, content_type='application/octet-stream')
    response['Content-Disposition'] = 'attachment; filename="VEHICLE.DAT"'
    return response


def get_expenses_file(request):
    """
    GET /device/expenses/
    Returns EXPENSEDET.DAT binary (31 bytes per expense).
    """
    if request.method != 'GET':
        return HttpResponse('METHOD_NOT_ALLOWED', status=405)

    company = _get_company(request)
    if not company:
        return HttpResponse('UNAUTHORIZED', status=401)

    expenses = (
        ExpenseMaster.objects
        .filter(company=company)
        .order_by('expense_code')
    )
    binary = _pack_expensedet(expenses)
    response = HttpResponse(binary, content_type='application/octet-stream')
    response['Content-Disposition'] = 'attachment; filename="EXPENSEDET.DAT"'
    return response
