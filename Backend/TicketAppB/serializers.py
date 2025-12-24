from rest_framework import serializers
from .models import Company,CustomUser,TransactionData


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