from django.db import models
from django.contrib.auth import get_user_model
# from django.contrib.postgres.fields import ArrayField # Optional, if using Postgres
from office.models import (
    Station, CSIUnit, SectionalOfficer, Designation, 
    Manufacturer, FailureReason, IPSModuleType, IPSCompany
)

User = get_user_model()

class BaseReport(models.Model):
    """
    Abstract base class for all reports to ensure consistent metadata.
    """
    class Status(models.TextChoices):
        OPEN = 'Open', 'Open'
        RESOLVED = 'Resolved', 'Resolved'
        DRAFT = 'Draft', 'Draft'

    # Tracking
    submitted_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    status = models.CharField(max_length=20, choices=Status.choices, default=Status.OPEN)

    # Snapshot of Personal Details at time of submission 
    # (In case user changes designation/posting later, report remains accurate)
    reporter_name = models.CharField(max_length=100)
    reporter_designation = models.ForeignKey(Designation, on_delete=models.SET_NULL, null=True)
    
    # Hierarchy Context
    sectional_officer = models.ForeignKey(SectionalOfficer, on_delete=models.PROTECT)
    csi_unit = models.ForeignKey(CSIUnit, on_delete=models.PROTECT)

    class Meta:
        abstract = True

# -------------------------------------------------------------------------
# 1. Failure Report (Fire Alarm)
# -------------------------------------------------------------------------
class FailureReport(BaseReport):
    """
    Corresponds to ReportForm.tsx / FORM-FA-01
    """
    report_date = models.DateField(help_text="Date of reporting")
    
    # Location Details
    posting_station = models.ForeignKey(Station, on_delete=models.PROTECT, related_name='failure_postings')
    affected_location = models.ForeignKey(Station, on_delete=models.PROTECT, related_name='failure_locations')
    route = models.CharField(max_length=50) # e.g., "A", "B", "D-SPL"

    # Failure Specifics
    make = models.ForeignKey(Manufacturer, on_delete=models.PROTECT)
    failure_datetime = models.DateTimeField()
    
    # Using M2M allows selecting multiple reasons from master data
    reasons = models.ManyToManyField(FailureReason, related_name='reports')
    
    remarks = models.TextField(blank=True)
    is_under_amc = models.BooleanField(default=False)
    is_under_warranty = models.BooleanField(default=False)

    def __str__(self):
        return f"Fail-{self.id} | {self.posting_station} | {self.report_date}"

# -------------------------------------------------------------------------
# 2. Relay Room Log
# -------------------------------------------------------------------------
class RelayRoomLog(BaseReport):
    """
    Corresponds to InspectionForm.tsx / FORM-RR-02
    """
    log_date = models.DateField()
    location = models.ForeignKey(Station, on_delete=models.PROTECT)
    
    opening_time = models.TimeField()
    closing_time = models.TimeField()
    
    # Auth Details
    opening_code = models.CharField(max_length=50) # Dropdown value
    
    # Generated serial numbers
    sn_opening = models.CharField(max_length=50)
    sn_closing = models.CharField(max_length=50)
    
    reason_remarks = models.TextField(help_text="Reason for opening")

    def __str__(self):
        return f"Relay-{self.id} | {self.location} | {self.log_date}"

# -------------------------------------------------------------------------
# 3. Maintenance Report
# -------------------------------------------------------------------------
class MaintenanceReport(models.Model):
    MAINTENANCE_CHOICES = [
        ('Point Maintenance', 'Point Maintenance'),
        ('Signal Maintenance', 'Signal Maintenance'),
        ('IPS/Power Supply', 'IPS/ Power Supply including battery maintenance'),
        ('LC Gate Maintenance', 'LC gate Maintenance'),
    ]

    # Header Details
    submission_date = models.DateField()
    created_at = models.DateTimeField(auto_now_add=True)
    
    # Staff Details
    # Storing names as strings for simplicity in this demo, 
    # but could be ForeignKeys to a Staff model if you have one.
    staff_name = models.CharField(max_length=150, blank=True) 
    
    # Hierarchy (Using your existing Master Data models)
    sectional_officer = models.ForeignKey(SectionalOfficer, on_delete=models.PROTECT)
    csi_unit = models.ForeignKey(CSIUnit, on_delete=models.PROTECT)
    designation = models.ForeignKey(Designation, on_delete=models.SET_NULL, null=True)
    station = models.ForeignKey(Station, on_delete=models.PROTECT)

    # Activity Details
    maintenance_type = models.CharField(max_length=100, choices=MAINTENANCE_CHOICES)
    
    # Conditional Fields based on type
    # For "Point Maintenance done (Nos.)"
    asset_count = models.IntegerField(null=True, blank=True, help_text="Number of points maintained")
    # For "LC gate Maintenance(LC gate no.)"
    asset_id = models.CharField(max_length=50, null=True, blank=True, help_text="LC Gate Number")

    section = models.CharField(max_length=50, help_text="e.g., ADI-SHB")
    remarks = models.TextField(blank=True)

    def __str__(self):
        return f"{self.maintenance_type} at {self.station} on {self.submission_date}"

# -------------------------------------------------------------------------
# 4. IPS Report (Complex Matrix)
# -------------------------------------------------------------------------
class IPSReport(BaseReport):
    """
    Corresponds to IPSModuleForm.tsx / FORM-IPS-03
    Header table for the weekly report.
    """
    submission_date = models.DateField(help_text="Must be Monday")
    week_from = models.DateField()
    week_to = models.DateField()
    remarks = models.TextField(blank=True, help_text="Action Plan / Non-AMC")

    class Meta:
        # Prevent duplicate reports for same CSI in same week?
        unique_together = ('csi_unit', 'week_from')

    def __str__(self):
        return f"IPS-{self.id} | {self.csi_unit} | {self.week_from}"

class IPSEntry(models.Model):
    """
    Individual rows in the IPS Report Matrix.
    One Report -> Many Entries.
    """
    report = models.ForeignKey(IPSReport, on_delete=models.CASCADE, related_name='entries')
    
    module_type = models.ForeignKey(IPSModuleType, on_delete=models.PROTECT)
    company = models.ForeignKey(IPSCompany, on_delete=models.PROTECT)
    
    # The Counts
    qty_defective = models.PositiveIntegerField(default=0)
    qty_spare = models.PositiveIntegerField(default=0)
    qty_spare_amc = models.PositiveIntegerField(default=0)
    qty_defective_amc = models.PositiveIntegerField(default=0)

    class Meta:
        # Ensure only one entry per Module+Company per Report
        unique_together = ('report', 'module_type', 'company')