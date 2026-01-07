from rest_framework import serializers
from .models import Company,CustomUser,TransactionData,TripCloseData,Branch


class CompanySerializer(serializers.ModelSerializer):
    # Read-only fields that are computed or set by system
    is_validated = serializers.ReadOnlyField()
    needs_validation = serializers.ReadOnlyField()
    created_by = serializers.PrimaryKeyRelatedField(read_only=True)
    
    class Meta:
        model = Company
        fields = [
            'id',
            'company_id',
            'company_name',
            'company_email',
            'gst_number',
            'contact_person',
            'contact_number',
            'address',
            'address_2',
            'city',
            'state',
            'zip_code',
            'number_of_licence',
            'authentication_status',
            'product_registration_id',
            'unique_identifier',
            'product_from_date',
            'product_to_date',
            'project_code',
            'device_count',
            'branch_count',
            'created_at',
            'updated_at',
            'is_validated',
            'needs_validation',
            'created_by',
        ]
        read_only_fields = [
            'id',
            'company_id',
            'authentication_status',
            'product_registration_id',
            'unique_identifier',
            'product_from_date',
            'product_to_date',
            'project_code',
            'device_count',
            'branch_count',
            'created_at',
            'updated_at',
            'created_by',
        ]
    
    def validate_company_email(self, value):
        """Ensure email is unique (except for current instance in update)"""
        if self.instance:
            # Update case: allow same email for current instance
            if Company.objects.exclude(pk=self.instance.pk).filter(company_email=value).exists():
                raise serializers.ValidationError("A company with this email already exists.")
        else:
            # Create case: ensure email doesn't exist
            if Company.objects.filter(company_email=value).exists():
                raise serializers.ValidationError("A company with this email already exists.")
        return value
    
    def validate_number_of_licence(self, value):
        """Ensure license count is positive"""
        if value < 1:
            raise serializers.ValidationError("Number of licenses must be at least 1.")
        return value
    
    def validate_contact_number(self, value):
        """Basic phone number validation"""
        # Remove common formatting characters
        cleaned = value.replace('-', '').replace(' ', '').replace('(', '').replace(')', '')
        if not cleaned.isdigit():
            raise serializers.ValidationError("Contact number must contain only digits and basic formatting characters.")
        if len(cleaned) < 10:
            raise serializers.ValidationError("Contact number must be at least 10 digits.")
        return value


class UserSerializer(serializers.ModelSerializer):
    class Meta:
        model=CustomUser
        fields='__all__'


class TicketDataSerializer(serializers.ModelSerializer):
    # Ticket type mapping for display
    TICKET_TYPE_MAPPING = {
        '0': 'Full',
        '1': 'Half',
        '2': 'Student',
        '3': 'Physical',
        '4': 'Luggage'
    }

    # Add display fields
    payment_mode_display = serializers.SerializerMethodField()
    ticket_type_display = serializers.SerializerMethodField()
    formatted_ticket_date = serializers.SerializerMethodField()

    class Meta:
        model = TransactionData
        fields = [
            'id',
            'request_type',
            'device_id',
            'trip_number',
            'ticket_number',
            'ticket_date',  # Original date field
            'formatted_ticket_date',  # Display format DD-MM-YYYY
            'ticket_time',
            'from_stage',
            'to_stage',
            'ticket_type',  # Original field (stores "0", "1", etc.)
            'ticket_type_display',  # Display format ("Full", "Half", etc.)
            'full_count',
            'half_count',
            'st_count',
            'phy_count',
            'lugg_count',
            'total_tickets',  # Total tickets count
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
            'ticket_status',  # Original field (stores 0 or 1)
            'payment_mode_display',  # Display format ("Cash" or "UPI")
            'reference_number',
            'branch_code',  # Will be null for now
            'created_at',
        ]
        # EXCLUDED: raw_payload, company_code
    
    def get_payment_mode_display(self, obj):
        """
        Convert ticket_status integer to display string.
        0 -> "Cash"
        1 -> "UPI"
        None/Other -> "Unknown"
        """
        if obj.ticket_status == 0:
            return "Cash"
        elif obj.ticket_status == 1:
            return "UPI"
        return "Unknown"
    
    def get_ticket_type_display(self, obj):
        """
        Convert ticket_type code to display string.
        Maps: "0" -> "Full", "1" -> "Half", etc.
        For unknown values (future combinations), returns the raw value.
        """
        if obj.ticket_type and obj.ticket_type in self.TICKET_TYPE_MAPPING:
            return self.TICKET_TYPE_MAPPING[obj.ticket_type]
        elif obj.ticket_type:
            # Return raw value for unknown/future ticket types
            return obj.ticket_type
        return "Unknown"
    
    def get_formatted_ticket_date(self, obj):
        """
        Format date as DD-MM-YYYY for frontend display.
        Returns None if ticket_date is None.
        """
        if obj.ticket_date:
            return obj.ticket_date.strftime('%d-%m-%Y')
        return None


class TripCloseDataSerializer(serializers.ModelSerializer):
    total_passengers = serializers.SerializerMethodField()
    total_tickets_issued = serializers.SerializerMethodField()

    class Meta:
        model = TripCloseData
        fields = [
            "id",
            "palmtec_id",
            "company_code",
            "schedule",
            "trip_no",
            "route_code",
            "up_down_trip",
            "start_datetime",
            "end_datetime",
            "start_ticket_no",
            "end_ticket_no",

            # Passenger counts
            "full_count",
            "half_count",
            "st1_count",
            "luggage_count",
            "physical_count",
            "pass_count",
            "ladies_count",
            "senior_count",

            # Collections
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

            # UPI
            "upi_ticket_count",
            "upi_ticket_amount",

            # Derived fields
            "total_passengers",
            "total_tickets_issued",

            # Timestamps
            "received_at",
            "created_at",
        ]

    def get_total_passengers(self, obj):
        return obj.get_total_passengers()

    def get_total_tickets_issued(self, obj):
        return obj.get_total_tickets_issued()
    

class BranchSerializer(serializers.ModelSerializer):
    company = serializers.PrimaryKeyRelatedField(read_only=True)
    created_by = serializers.PrimaryKeyRelatedField(read_only=True)
    
    class Meta:
        model=Branch
        fields=[
            'id',
            'company',
            'branch_code',
            'branch_name',
            'address',
            'city',
            'state',
            'zip_code',
            'is_active',
            'created_by'
        ]
        read_only_fields=[
            'id',
            'company',
            'is_active',
            'created_by',
        ]