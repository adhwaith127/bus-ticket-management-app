from rest_framework import serializers
from ..models import Company, Depot


class CompanySerializer(serializers.ModelSerializer):
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
            'device_count',
            'depot_count',
            'mobile_device_count',
            'created_at',
            'updated_at',
            'is_validated',
            'needs_validation',
            'client_type',
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
            'created_at',
            'updated_at',
            'created_by',
        ]

    def validate_company_email(self, value):
        if self.instance:
            if Company.objects.exclude(pk=self.instance.pk).filter(company_email=value).exists():
                raise serializers.ValidationError("A company with this email already exists.")
        else:
            if Company.objects.filter(company_email=value).exists():
                raise serializers.ValidationError("A company with this email already exists.")
        return value

    def validate_number_of_licence(self, value):
        if value < 0:
            raise serializers.ValidationError("Number of licenses cannot be negative.")
        return value

    def validate_contact_number(self, value):
        cleaned = value.replace('-', '').replace(' ', '').replace('(', '').replace(')', '')
        if not cleaned.isdigit():
            raise serializers.ValidationError("Contact number must contain only digits and basic formatting characters.")
        if len(cleaned) < 10:
            raise serializers.ValidationError("Contact number must be at least 10 digits.")
        return value


class DepotSerializer(serializers.ModelSerializer):
    company = serializers.PrimaryKeyRelatedField(read_only=True)
    created_by = serializers.PrimaryKeyRelatedField(read_only=True)

    class Meta:
        model = Depot
        fields = [
            'id',
            'company',
            'depot_code',
            'depot_name',
            'address',
            'city',
            'state',
            'zip_code',
            'is_active',
            'created_by',
        ]
        read_only_fields = [
            'id',
            'company',
            'is_active',
            'created_by',
        ]
