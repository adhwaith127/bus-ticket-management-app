from django.db import models
from django.contrib.auth.models import AbstractUser


from django.db import models

class Company(models.Model):
    # Authentication Status Choices
    class AuthStatus(models.TextChoices):
        PENDING = 'Pending', 'Pending'
        APPROVED = 'Approve', 'Approved'
        EXPIRED = 'Expired', 'Expired'
        BLOCKED = 'Block', 'Blocked'
    
    # Basic Company Information
    company_id = models.CharField(max_length=100, unique=True, null=True, blank=True)
    company_name = models.CharField(max_length=100)
    company_email = models.EmailField(unique=True)
    gst_number = models.CharField(max_length=20, null=True, blank=True)
    
    # Contact Information
    contact_person = models.CharField(max_length=100)
    contact_number = models.CharField(max_length=20)
    
    # Address Information
    address = models.TextField()
    address_2 = models.TextField(blank=True, null=True)
    city = models.CharField(max_length=100)
    state = models.CharField(max_length=100)
    zip_code = models.CharField(max_length=20)
    
    # License Information
    number_of_licence = models.IntegerField(default=1)
    authentication_status = models.CharField(
        max_length=20,
        choices=AuthStatus.choices,
        default=AuthStatus.PENDING,
        null=True,
        blank=True
    )
    
    # License Server Fields
    product_registration_id = models.IntegerField(null=True, blank=True)
    unique_identifier = models.CharField(max_length=255, null=True, blank=True)
    product_from_date = models.DateField(null=True, blank=True)
    product_to_date = models.DateField(null=True, blank=True)
    
    # Additional License Fields (NEW)
    project_code = models.CharField(max_length=100, null=True, blank=True)
    device_count = models.IntegerField(null=True, blank=True)
    branch_count = models.IntegerField(null=True, blank=True)
    
    # Timestamps
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        db_table = 'company'
        verbose_name = 'Company'
        verbose_name_plural = 'Companies'
    
    def __str__(self):
        return self.company_name
    
    @property
    def is_validated(self):
        """Check if company license is validated"""
        return self.authentication_status == self.AuthStatus.APPROVED
    
    @property
    def needs_validation(self):
        """Check if company needs license validation"""
        return self.authentication_status == self.AuthStatus.PENDING

class CustomUser(AbstractUser):
    role = models.CharField(max_length=32, blank=True, null=True,default='user')
    is_verified = models.BooleanField(default=False)
    company=models.ForeignKey(to=Company,on_delete=models.CASCADE,null=True,blank=True)
    
    class Meta:
        db_table = 'custom_user'
    
    def __str__(self):
        return self.username
 


class TransactionData(models.Model):
    request_type      = models.CharField(max_length=20, null=True, blank=True)
    device_id         = models.CharField(max_length=20, null=True, blank=True)
    trip_number       = models.CharField(max_length=20, null=True, blank=True)
    ticket_number     = models.CharField(max_length=20, null=True, blank=True)
    ticket_date       = models.DateField(null=True, blank=True)
    ticket_time       = models.TimeField(null=True, blank=True)

    from_stage        = models.IntegerField(null=True, blank=True)
    to_stage          = models.IntegerField(null=True, blank=True)

    full_count        = models.IntegerField(default=0)
    half_count        = models.IntegerField(default=0)
    st_count          = models.IntegerField(default=0)
    phy_count         = models.IntegerField(default=0)
    lugg_count        = models.IntegerField(default=0)

    ticket_amount     = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    lugg_amount       = models.DecimalField(max_digits=10, decimal_places=2, default=0)

    ticket_type       = models.CharField(max_length=10, null=True, blank=True)
    adjust_amount     = models.DecimalField(max_digits=10, decimal_places=2, default=0)

    pass_id           = models.CharField(max_length=20, null=True, blank=True)
    warrant_amount    = models.DecimalField(max_digits=10, decimal_places=2, default=0)

    refund_status     = models.CharField(max_length=5, null=True, blank=True)
    refund_amount     = models.DecimalField(max_digits=10, decimal_places=2, default=0)

    ladies_count      = models.IntegerField(default=0)
    senior_count      = models.IntegerField(default=0)

    transaction_id    = models.CharField(max_length=50, null=True, blank=True)
    ticket_status     = models.CharField(max_length=10, null=True, blank=True)
    reference_number  = models.CharField(max_length=50, null=True, blank=True)

    company_code      = models.CharField(max_length=10, null=True, blank=True)

    # full string
    raw_payload       = models.TextField()  

    created_at        = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = "transaction_data"
        indexes = [
            models.Index(fields=["device_id", "ticket_date"]),
            models.Index(fields=["company_code"]),
        ]
        constraints = [
            models.UniqueConstraint(
                fields=[
                    'device_id',
                    'trip_number',
                    'ticket_number',
                    'ticket_date',
                    'ticket_time',
                ],
                name='uniq_device_trip_ticket_datetime'
            )
        ]

    def __str__(self):
        return f"{self.ticket_number} - {self.device_id}"