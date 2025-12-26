from decimal import Decimal
from django.db import models
from django.contrib.auth.models import AbstractUser
from django.core.validators import MinValueValidator

from django.db import models

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
    
    # Additional License Fields
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
    
    @property
    def is_validating(self):
        """Check if validation is in progress"""
        return self.authentication_status == self.AuthStatus.VALIDATING


# Rest of the models remain the same...
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
    # values are 0 and 1 . 0 for cash 1 for upi
    ticket_status     = models.CharField(max_length=10, null=True, blank=True)
    reference_number  = models.CharField(max_length=50, null=True, blank=True)

    company_code      = models.CharField(max_length=10, null=True, blank=True)

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


class TripCloseData(models.Model):  
    palmtec_id = models.CharField(
        max_length=50,
        db_index=True,
        help_text="Device identifier (PalmtecID)"
    )
    
    company_code = models.CharField(
        max_length=100,
        help_text="Company code"
    )
    
    schedule = models.IntegerField(
        help_text="Schedule number"
    )
    
    trip_no = models.IntegerField(
        help_text="Trip number"
    )
    
    route_code = models.CharField(
        max_length=50,
        db_index=True,
        help_text="Route code"
    )
    
    up_down_trip = models.CharField(
        max_length=1,
        help_text="Trip direction indicator (U/D)"
    )
    
    start_datetime = models.DateTimeField(
        db_index=True,
        help_text="Trip start date and time"
    )
    
    end_datetime = models.DateTimeField(
        help_text="Trip end date and time"
    )
    
    start_ticket_no = models.BigIntegerField(
        help_text="Starting ticket number (lSTicketNo)"
    )
    
    end_ticket_no = models.BigIntegerField(
        help_text="Ending ticket number (lETicketNo)"
    )
    
    full_count = models.IntegerField(
        default=0,
        validators=[MinValueValidator(0)],
        help_text="Full fare passengers (sFull + uFull)"
    )
    
    half_count = models.IntegerField(
        default=0,
        validators=[MinValueValidator(0)],
        help_text="Half fare passengers (sHalf + uChild)"
    )
    
    st1_count = models.IntegerField(
        default=0,
        validators=[MinValueValidator(0)],
        help_text="ST1 type passengers (sST1 + uSTCount)"
    )
    
    luggage_count = models.IntegerField(
        default=0,
        validators=[MinValueValidator(0)],
        help_text="Luggage count (sLugg + uLugg)"
    )
    
    physical_count = models.IntegerField(
        default=0,
        validators=[MinValueValidator(0)],
        help_text="Physical handicap passengers (sPhy + uPhy)"
    )
    
    pass_count = models.IntegerField(
        default=0,
        validators=[MinValueValidator(0)],
        help_text="Pass holders (sPass)"
    )
    
    ladies_count = models.IntegerField(
        default=0,
        validators=[MinValueValidator(0)],
        help_text="Ladies passengers (sLadies + uLadies)"
    )
    
    senior_count = models.IntegerField(
        default=0,
        validators=[MinValueValidator(0)],
        help_text="Senior citizen passengers (sSenior + uSenior)"
    )
    
    full_collection = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="Full fare collection (fFullColl + uFullColl)"
    )
    
    half_collection = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="Half fare collection (fHalfColl + uChildColl)"
    )
    
    st_collection = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="ST collection (fSTColl + uSTColl)"
    )
    
    luggage_collection = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="Luggage collection (fLuggageColl + uLuggColl)"
    )
    
    physical_collection = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="Physical handicap collection (fPhyColl + uPhyColl)"
    )
    
    ladies_collection = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="Ladies collection (fLadiColl + uLadiesColl)"
    )
    
    senior_collection = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="Senior collection (fSeniorColl + uSeniorColl)"
    )
    
    adjust_collection = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        help_text="Adjustment collection (fAdjustColl) - can be negative"
    )
    
    expense_amount = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="Expense amount (fExpenseAmount)"
    )
    
    total_collection = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="Total collection (fTotalColl)"
    )
    
    upi_ticket_count = models.IntegerField(
        default=0,
        validators=[MinValueValidator(0)],
        help_text="UPI ticket count (sUpiTicketCount)"
    )
    
    upi_ticket_amount = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal('0.00'),
        validators=[MinValueValidator(Decimal('0.00'))],
        help_text="UPI ticket amount (fUPITicketAmount)"
    )
    
    received_at = models.DateTimeField(
        auto_now_add=True,
        help_text="When server received this data"
    )
    
    created_at = models.DateTimeField(
        auto_now_add=True,
        help_text="Record creation timestamp"
    )
    
    updated_at = models.DateTimeField(
        auto_now=True,
        help_text="Record last update timestamp"
    )
    
    class Meta:
        db_table = 'trip_close_data'
        verbose_name = 'Trip Close Data'
        verbose_name_plural = 'Trip Close Datas'
        
        indexes = [
            models.Index(fields=['palmtec_id', 'start_datetime']),
            models.Index(fields=['route_code', 'start_datetime']),
            models.Index(fields=['start_datetime']),
        ]
        
        unique_together = [
            ['palmtec_id', 'schedule', 'trip_no', 'start_datetime']
        ]
        
        ordering = ['-start_datetime']
    
    def __str__(self):
        return f"Trip {self.trip_no} - {self.route_code} - {self.palmtec_id} ({self.start_datetime})"
    
    def get_total_passengers(self):
        return (
            self.full_count + 
            self.half_count + 
            self.st1_count + 
            self.physical_count + 
            self.pass_count + 
            self.ladies_count + 
            self.senior_count
        )
    
    def get_total_tickets_issued(self):
        if self.end_ticket_no and self.start_ticket_no:
            return self.end_ticket_no - self.start_ticket_no + 1
        return 0