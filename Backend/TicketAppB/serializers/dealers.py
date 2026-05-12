from rest_framework import serializers
from ..models import Dealer, DealerCustomerMapping


class DealerSerializer(serializers.ModelSerializer):
    created_by = serializers.PrimaryKeyRelatedField(read_only=True)

    class Meta:
        model = Dealer
        fields = [
            'id',
            'dealer_code',
            'dealer_name',
            'contact_person',
            'contact_number',
            'email',
            'address',
            'city',
            'state',
            'zip_code',
            'gst_number',
            'is_active',
            'number_of_licence',
            'created_by',
            'created_at',
            'updated_at',
        ]
        read_only_fields = [
            'id',
            'number_of_licence',
            'created_by',
            'created_at',
            'updated_at',
        ]


class DealerCustomerMappingSerializer(serializers.ModelSerializer):
    created_by = serializers.PrimaryKeyRelatedField(read_only=True)

    class Meta:
        model = DealerCustomerMapping
        fields = [
            'id',
            'dealer',
            'company',
            'is_active',
            'created_by',
            'created_at',
            'updated_at',
        ]
        read_only_fields = [
            'id',
            'created_by',
            'created_at',
            'updated_at',
        ]

    def validate(self, attrs):
        dealer = attrs.get('dealer') or getattr(self.instance, 'dealer', None)
        company = attrs.get('company') or getattr(self.instance, 'company', None)
        if dealer and company:
            existing = DealerCustomerMapping.objects.filter(dealer=dealer, company=company)
            if self.instance:
                existing = existing.exclude(pk=self.instance.pk)
            if existing.exists():
                raise serializers.ValidationError("This dealer is already mapped to the selected company.")
        return attrs
