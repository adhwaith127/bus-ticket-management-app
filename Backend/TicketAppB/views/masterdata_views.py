# masterdata_views.py
# This module re-exports from the split view files.
# URLs continue to reference masterdata_views.xxx with no changes needed.

from .transport_views import (
    get_bus_types, create_bus_type, update_bus_type,
    get_stages, create_stage, update_stage,
    get_routes, create_route, update_route,
    get_vehicles, create_vehicle, update_vehicle,
    get_bus_types_dropdown, get_stages_dropdown, get_vehicles_dropdown,
    get_fare_editor, update_fare_table,
)

from .crew_views import (
    get_employee_types, create_employee_type, update_employee_type,
    get_employees, create_employee, update_employee,
    get_crew_assignments, create_crew_assignment, update_crew_assignment, delete_crew_assignment,
    get_employee_types_dropdown, get_employees_by_type_dropdown,
)

from .settings_views import (
    get_currencies, create_currency, update_currency,
    get_settings,
)
