from rest_framework import serializers
from .models import Company,CustomUser,TransactionData

class CompanySerializer(serializers.ModelSerializer):
    class Meta:
        model = Company
        fields = '__all__'


class UserSerializer(serializers.ModelSerializer):
    class Meta:
        model=CustomUser
        fields='__all__'


class TicketDataSerializer(serializers.ModelSerializer):
    class Meta:
        model=TransactionData
        fields='__all__'