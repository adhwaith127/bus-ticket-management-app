from rest_framework import status
from rest_framework.response import Response
from rest_framework.decorators import api_view
from ..models import Company,CustomUser
from ..serializers import CompanySerializer,UserSerializer
from django.contrib.auth import get_user_model

User=get_user_model()


@api_view(['POST'])
def create_user(request):
    try:        
        if not request.data:
            return Response({'error': 'Invalid Input'}, status=status.HTTP_400_BAD_REQUEST)
        
        username=request.data.get('username')
        email=request.data.get('email')
        role=request.data.get('role')
        company=request.data.get('company')
        password=request.data.get('password')

        try:
            company_instance = Company.objects.get(company_name=company)
        except Company.DoesNotExist:
            return Response({"message": "Invalid Company Given"},status=status.HTTP_400_BAD_REQUEST)
        
        user = User.objects.create_user(username=username, email=email, password=password,company=company_instance,role=role)
        user.save()

        return Response({"message":"User added successfully"},status=status.HTTP_201_CREATED)
    
    except Exception as e:
        return Response({"message":"User creation failed"},status=status.HTTP_400_BAD_REQUEST)



@api_view(['GET'])
def get_all_users(request):
    users=CustomUser.objects.all().order_by('-id')

    serializer = UserSerializer(users, many=True)

    return Response(
        {"message": "Success","data": serializer.data},status=status.HTTP_200_OK)

@api_view(['GET'])
def all_company_data(request):
    companies = Company.objects.all().order_by('-id')

    serializer = CompanySerializer(companies, many=True)

    return Response(
        {"message": "Success","data": serializer.data},status=status.HTTP_200_OK)


@api_view(['POST'])
def create_company(request):
    if not request.data:
        return Response({"message": "No input received"},status=status.HTTP_400_BAD_REQUEST)

    serializer = CompanySerializer(data=request.data)

    if serializer.is_valid():
        company = serializer.save()
        return Response(
            {"message": "Successfully added company","data": serializer.data},status=status.HTTP_201_CREATED)

    return Response({"message": "Validation failed","errors": serializer.errors},status=status.HTTP_400_BAD_REQUEST)