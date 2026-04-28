from decimal import Decimal
from django.db import models
from django.contrib.auth.models import AbstractUser
from django.core.validators import MinValueValidator
from django.conf import settings

# Company & Depot models
class Company(models.Model):
    # Authentication Status Choices
    class AuthStatus(models.TextChoices):
        PENDING = 'Pending', 'Pending'
        # for showing in UI that license validation is undergoing
        VALIDATING = 'Validating', 'Validating'
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
    number_of_licence = models.IntegerField(default=0)
    authentication_status = models.CharField(
        max_length=20,
        choices=AuthStatus.choices,
        default=AuthStatus.PENDING,
        null=True,
        blank=True
    )

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='companies_created'
    )
    
    # License Server Fields
    product_registration_id = models.IntegerField(null=True, blank=True)
    unique_identifier = models.CharField(max_length=255, null=True, blank=True)
    product_from_date = models.DateField(null=True, blank=True)
    product_to_date = models.DateField(null=True, blank=True)
    
    # Additional License Fields
    device_count = models.IntegerField(default=0, null=True, blank=True)       # NoOfUPIDevice
    depot_count = models.IntegerField(default=0, null=True, blank=True)       # NoOfDepot
    mobile_device_count = models.IntegerField(default=2, null=True, blank=True) # NoOfMobileDevice
    
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
    
    @property
    def is_validating(self):
        """Check if validation is in progress"""
        return self.authentication_status == self.AuthStatus.VALIDATING



class Depot(models.Model):
    company = models.ForeignKey(
        Company,
        on_delete=models.CASCADE,
        related_name='depots'
    )

    depot_code = models.CharField(
        max_length=50,
        unique=True
    )

    depot_name = models.CharField(
        max_length=100
    )

    address = models.TextField()
    city = models.CharField(max_length=100)
    state = models.CharField(max_length=100)
    zip_code = models.CharField(max_length=20)

    # who created this depot
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='depots_created'
    )

    is_active = models.BooleanField(default=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'depot'
        unique_together = ['company', 'depot_code']
        indexes = [
            models.Index(fields=['company', 'depot_code']),
        ]

    def __str__(self):
        return f"{self.depot_name} ({self.company.company_name})"


class Dealer(models.Model):
    dealer_code = models.CharField(max_length=50, unique=True)
    dealer_name = models.CharField(max_length=150)
    contact_person = models.CharField(max_length=100)
    contact_number = models.CharField(max_length=20)
    email = models.EmailField(unique=True)
    address = models.TextField()
    city = models.CharField(max_length=100)
    state = models.CharField(max_length=100)
    zip_code = models.CharField(max_length=20)
    gst_number = models.CharField(max_length=20, null=True, blank=True)
    is_active = models.BooleanField(default=True)

    # License pool allocated by superadmin at dealer creation/edit
    # number_of_licence = total (ETM + Android); device_count = ETM; mobile_device_count = Android
    allocated_licence_count = models.IntegerField(default=0)         # total licences allocated
    allocated_device_count = models.IntegerField(default=0)          # ETM licences allocated
    allocated_mobile_device_count = models.IntegerField(default=0)   # Android licences allocated

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='dealers_created'
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'dealer'
        indexes = [
            models.Index(fields=['dealer_code']),
            models.Index(fields=['dealer_name']),
            models.Index(fields=['is_active']),
        ]

    def __str__(self):
        return f"{self.dealer_code} - {self.dealer_name}"


class DealerCustomerMapping(models.Model):
    dealer = models.ForeignKey(
        Dealer,
        on_delete=models.CASCADE,
        related_name='company_mappings'
    )
    company = models.ForeignKey(
        Company,
        on_delete=models.CASCADE,
        related_name='dealer_mappings'
    )
    is_active = models.BooleanField(default=True)

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='dealer_company_mappings_created'
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'dealer_customer_mapping'
        unique_together = ['dealer', 'company']
        indexes = [
            models.Index(fields=['dealer', 'company']),
            models.Index(fields=['company']),
            models.Index(fields=['is_active']),
        ]

    def __str__(self):
        return f"{self.dealer.dealer_name} -> {self.company.company_name}"


class ETMDevice(models.Model):
    """
    Registry of physical ETM (Electronic Ticket Machine) and Android app devices.
    Each device must be registered and approved before it can push ticket data.

    Lifecycle:
      1. Device boots and POSTs to /etm-devices/register  → status = Pending
      2. Superadmin/executive sees it in the pending queue and assigns it to a Company
      3. Admin clicks Approve → calls license server DeviceRegistration API
      4. License server returns device_registration_id → status = Active
      5. Admin can later call Check Status to refresh expiry from license server
    """

    class DeviceType(models.TextChoices):
        ETM = 'ETM', 'ETM (Electronic Ticket Machine)'
        ANDROID = 'ANDROID', 'Android App Device'

    class LicenceStatus(models.TextChoices):
        PENDING = 'Pending', 'Pending'
        ACTIVE = 'Active', 'Active'
        INACTIVE = 'Inactive', 'Inactive'
        EXPIRED = 'Expired', 'Expired'

    serial_number = models.CharField(max_length=100, unique=True)
    display_name = models.CharField(max_length=100, blank=True)
    device_type = models.CharField(
        max_length=20,
        choices=DeviceType.choices,
        default=DeviceType.ETM,
    )
    mac_address = models.CharField(max_length=100, blank=True)

    # Ownership — all nullable so a freshly-registered device has no owner yet
    company = models.ForeignKey(
        Company,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='etm_devices',
    )
    depot = models.ForeignKey(
        Depot,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='etm_devices',
    )
    # Which dealer supplied this device (set automatically from company's dealer mapping)
    dealer = models.ForeignKey(
        Dealer,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='etm_devices',
    )

    # License server binding
    device_registration_id = models.CharField(max_length=255, blank=True)

    licence_status = models.CharField(
        max_length=20,
        choices=LicenceStatus.choices,
        default=LicenceStatus.PENDING,
    )
    licence_active_to = models.DateField(null=True, blank=True)
    is_active = models.BooleanField(default=False)

    # Approval audit
    approved_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='approved_etm_devices',
    )
    approved_at = models.DateTimeField(null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'etm_device'
        indexes = [
            models.Index(fields=['serial_number']),
            models.Index(fields=['company', 'licence_status']),
            models.Index(fields=['dealer']),
            models.Index(fields=['licence_status']),
        ]

    def __str__(self):
        company_name = self.company.company_name if self.company else 'Unassigned'
        return f"{self.serial_number} ({self.device_type}) — {company_name}"

    @property
    def is_expired(self):
        from datetime import date
        if self.licence_active_to:
            return date.today() > self.licence_active_to
        return False

    @property
    def days_until_expiry(self):
        from datetime import date
        if self.licence_active_to:
            return (self.licence_active_to - date.today()).days
        return None


class ExecutiveCompanyMapping(models.Model):
    executive_user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name='executive_company_mappings'
    )
    company = models.ForeignKey(
        Company,
        on_delete=models.CASCADE,
        related_name='executive_mappings'
    )
    is_active = models.BooleanField(default=True)

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='executive_company_mappings_created'
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'executive_company_mapping'
        unique_together = ['executive_user', 'company']
        indexes = [
            models.Index(fields=['executive_user', 'company']),
            models.Index(fields=['company']),
            models.Index(fields=['is_active']),
        ]

    def __str__(self):
        return f"{self.executive_user.username} -> {self.company.company_name}"