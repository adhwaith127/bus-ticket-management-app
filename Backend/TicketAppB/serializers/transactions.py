from rest_framework import serializers
from ..models import TransactionData, TripCloseData


class TicketDataSerializer(serializers.ModelSerializer):
    TICKET_TYPE_BITS = {
        1:  'Full',
        2:  'Half',
        4:  'Luggage',
        8:  'PH',
        16: 'Student',
    }

    ticket_type_display = serializers.SerializerMethodField()
    formatted_ticket_date = serializers.SerializerMethodField()

    class Meta:
        model = TransactionData
        fields = [
            'id',
            'palmtec_id',
            'trip_number',
            'ticket_number',
            'ticket_date',
            'formatted_ticket_date',
            'ticket_time',
            'from_stage',
            'to_stage',
            'ticket_type',
            'ticket_type_display',
            'full_count',
            'half_count',
            'st_count',
            'phy_count',
            'lugg_count',
            'total_tickets',
            'ticket_amount',
            'lugg_amount',
            'adjust_amount',
            'pass_id',
            'warrant_amount',
            'refund_status',
            'refund_amount',
            'ladies_count',
            'senior_count',
            'transaction_id',
            'ticket_status',
            'reference_number',
            'created_at',
        ]

    def get_ticket_type_display(self, obj):
        try:
            val = int(obj.ticket_type)
        except (TypeError, ValueError):
            return obj.ticket_type or 'Unknown'
        labels = [label for bit, label in self.TICKET_TYPE_BITS.items() if val & bit]
        return ' + '.join(labels) if labels else 'Unknown'

    def get_formatted_ticket_date(self, obj):
        if obj.ticket_date:
            return obj.ticket_date.strftime('%d-%m-%Y')
        return None


class TripCloseDataSerializer(serializers.ModelSerializer):
    total_passengers = serializers.SerializerMethodField()
    total_tickets_issued = serializers.SerializerMethodField()
    formatted_start_date = serializers.SerializerMethodField()
    formatted_end_date = serializers.SerializerMethodField()

    class Meta:
        model = TripCloseData
        fields = [
            "id",
            "palmtec_id",
            "schedule",
            "trip_no",
            "route_id",
            "up_down_trip",
            "start_date",
            "start_time",
            "end_date",
            "end_time",
            "formatted_start_date",
            "formatted_end_date",
            "start_datetime",
            "end_datetime",
            "start_ticket_no",
            "end_ticket_no",
            "full_count",
            "half_count",
            "st1_count",
            "luggage_count",
            "physical_count",
            "pass_count",
            "ladies_count",
            "senior_count",
            "total_tickets",
            "total_cash_tickets",
            "upi_ticket_count",
            "full_collection",
            "half_collection",
            "st_collection",
            "luggage_collection",
            "physical_collection",
            "ladies_collection",
            "senior_collection",
            "adjust_collection",
            "expense_amount",
            "total_collection",
            "total_cash_amount",
            "upi_ticket_amount",
            "total_passengers",
            "total_tickets_issued",
            "received_at",
            "created_at",
        ]

    def get_total_passengers(self, obj):
        return obj.get_total_passengers()

    def get_total_tickets_issued(self, obj):
        return obj.get_total_tickets_issued()

    def get_formatted_start_date(self, obj):
        if obj.start_date:
            return obj.start_date.strftime('%d-%m-%Y')
        return None

    def get_formatted_end_date(self, obj):
        if obj.end_date:
            return obj.end_date.strftime('%d-%m-%Y')
        return None
