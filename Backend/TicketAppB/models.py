from django.db import models
from django.contrib.auth.models import AbstractUser


class Company(models.Model):
    company_id = models.CharField(max_length=100, unique=True, null=True, blank=True)
    company_name = models.CharField(max_length=100)
    company_email = models.EmailField(unique=True)
    gst_number = models.CharField(max_length=20, null=True, blank=True)
    contact_person = models.CharField(max_length=100)
    contact_number = models.CharField(max_length=20)
    address = models.TextField()
    address_2 = models.TextField(blank=True, null=True)
    city = models.CharField(max_length=100)
    state = models.CharField(max_length=100)
    zip_code = models.CharField(max_length=20)
    number_of_licence = models.IntegerField(default=1)

    authentication_status = models.CharField(max_length=100, null=True, blank=True,default='waiting-for-response')
    product_registration_id = models.IntegerField(null=True, blank=True)
    unique_identifier = models.CharField(max_length=255, null=True, blank=True)
    product_from_date = models.DateField(null=True, blank=True)
    product_to_date = models.DateField(null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'company'

    def __str__(self):
        return self.company_name


class CustomUser(AbstractUser):
    role = models.CharField(max_length=32, blank=True, null=True,default='user')
    is_verified = models.BooleanField(default=False)
    company=models.ForeignKey(to=Company,on_delete=models.CASCADE,null=True,blank=True)
    
    class Meta:
        db_table = 'custom_user'
    
    def __str__(self):
        return self.username