from django.urls import path
from .views import auth_views,company_views,user_views,data_views


urlpatterns = [
    # authentication
    path('signup/', auth_views.signup_view, name='signup'),
    path('login/', auth_views.login_view, name='login'),
    path('token/refresh/', auth_views.refresh_token_view, name='token_refresh'),
    path('logout/', auth_views.logout_view, name='logout'),
    path('protected/', auth_views.protected_view, name='protected'),
    path('verify-auth/', auth_views.verify_auth, name='verify_auth'),


    # user data
    path('create_user/',user_views.create_user,name='create-user'),
    path('get_users/',user_views.get_all_users,name='get_all_users'),


    # company data
    path('customer-data/', company_views.all_company_data, name='company_data'),
    path('create-company/', company_views.create_company, name='create_company'),
    path('update-company-details/<int:pk>/', company_views.update_company_details, name='update_company'),
    path('register-company-license/<int:pk>/', company_views.register_company_with_license_server, name='register_company_license'),  # NEW
    path('validate-company-license/<int:pk>/', company_views.validate_company_license, name='validate_company_license'),


    # ticket data
    path('getTransactionDataFromDevice/',data_views.getTransactionDataFromDevice,name='get_transaction_data'),
    path('get_all_transaction_data/',data_views.get_all_transaction_data,name='get_all_transaction_data'),
    path('getTripCloseDataFromDevice/',data_views.getTripCloseDataFromDevice,name='get_trip_close_data'),
    path('get_all_trip_close_data/',data_views.get_all_trip_close_data,name='get_all_trip_close_data'),
]