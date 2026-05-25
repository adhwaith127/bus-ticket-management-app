from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('TicketAppB', '0061_add_trip_schedule_fk_to_odometer_expense'),
    ]

    operations = [
        migrations.RunSQL(
            sql=[
                "DROP TABLE IF EXISTS schedule_close_data;",
                "DROP TABLE IF EXISTS schedule_open_data;",
                "DROP TABLE IF EXISTS trip_close_data;",
                "DROP TABLE IF EXISTS trip_open_data;",
            ],
            reverse_sql=migrations.RunSQL.noop,
        ),
        migrations.SeparateDatabaseAndState(
            state_operations=[
                migrations.DeleteModel(name='ScheduleCloseData'),
                migrations.DeleteModel(name='ScheduleOpenData'),
                migrations.DeleteModel(name='TripCloseData'),
                migrations.DeleteModel(name='TripOpenData'),
            ],
            database_operations=[],
        ),
    ]
