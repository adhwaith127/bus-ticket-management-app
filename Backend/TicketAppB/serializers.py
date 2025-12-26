from rest_framework import serializers
from .models import Company,CustomUser,TransactionData,TripCloseData


class CompanySerializer(serializers.ModelSerializer):
    # Read-only fields that are computed or set by system
    is_validated = serializers.ReadOnlyField()
    needs_validation = serializers.ReadOnlyField()
    
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
    class Meta:
        model=TransactionData
        fields='__all__'



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