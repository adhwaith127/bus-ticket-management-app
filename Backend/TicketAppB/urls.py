from django.urls import path
from .views import auth_views,admin_views


urlpatterns = [
    # authentication
    path('signup/', auth_views.signup_view, name='signup'),
    path('login/', auth_views.login_view, name='login'),
    path('token/refresh/', auth_views.refresh_token_view, name='token_refresh'),
    path('logout/', auth_views.logout_view, name='logout'),
    path('protected/', auth_views.protected_view, name='protected'),
    path('verify-auth/', auth_views.verify_auth, name='verify_auth'),

    # company data
    path('create_user',admin_views.create_user,name='create-user'),
    path('get_users/',admin_views.get_all_users,name='get_all_users'),
    path('customer-data/',admin_views.all_company_data, name='company_data'),
    path('create-company/',admin_views.create_company,name='create_company')
]