from rest_framework import status
from rest_framework.response import Response
from rest_framework.decorators import api_view
from ..models import Company,CustomUser
from ..serializers import UserSerializer
from django.contrib.auth import get_user_model
from .auth_views import get_user_from_cookie


User=get_user_model()


@api_view(['POST'])
def create_user(request):
    # verify user
    user=get_user_from_cookie(request)
    if not user:
        return Response({'error': 'Authentication required'}, status=status.HTTP_401_UNAUTHORIZED)
    
    try:        
        if not request.data:
            return Response({'error': 'Invalid Input'}, status=status.HTTP_400_BAD_REQUEST)
        
        username=request.data.get('username')
        email=request.data.get('email')
        role=request.data.get('role')
        company=request.data.get('company_id')
        password=request.data.get('password')

        try:
            company_instance = Company.objects.get(id=company)
        except Company.DoesNotExist:
            return Response({"message": "Invalid Company Given"},status=status.HTTP_400_BAD_REQUEST)
        
        user = User.objects.create_user(username=username, email=email, password=password,company=company_instance,role=role)
        user.save()

        return Response({"message":"User added successfully"},status=status.HTTP_201_CREATED)
    
    except Exception as e:
        return Response({"message":"User creation failed"},status=status.HTTP_500_INTERNAL_SERVER_ERROR)



@api_view(['GET'])
def get_all_users(request):
    # verify user
    user=get_user_from_cookie(request)
    if not user:
        return Response({'error': 'Authentication required'}, status=status.HTTP_401_UNAUTHORIZED)

    users=CustomUser.objects.all().order_by('id')

    serializer = UserSerializer(users, many=True)

    return Response({"message": "Success","data": serializer.data},status=status.HTTP_200_OK)