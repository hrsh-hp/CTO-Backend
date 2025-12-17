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
    log_date = models.DateField(db_column='date') # mapped to 'date' in frontend
    
    # We use SlugRelatedField in serializer, so ForeignKey is fine here.
    location = models.ForeignKey(Station, on_delete=models.PROTECT, related_name='relay_logs')
    
    opening_time = models.TimeField()
    closing_time = models.TimeField()
    
    # Auth Details
    opening_code = models.CharField(max_length=50) 
    
    # Serial numbers (Manual input from frontend)
    sn_opening = models.CharField(max_length=50)
    sn_closing = models.CharField(max_length=50)
    
    remarks = models.TextField(help_text="Reason for opening")

    def __str__(self):
        return f"Relay-{self.id} | {self.location} | {self.log_date}"

# -------------------------------------------------------------------------
# 3. Maintenance Report
# -------------------------------------------------------------------------
class MaintenanceReport(models.Model):
    TYPE_CHOICES = [
        ('Point Maintenance', 'Point Maintenance'),
        ('Signal Maintenance', 'Signal Maintenance'),
        ('IPS/Power Supply', 'IPS/Power Supply'),
        ('LC Gate Maintenance', 'LC Gate Maintenance'),
    ]

    submission_date = models.DateField()
    
    # Relationships
    sectional_officer = models.ForeignKey(SectionalOfficer, on_delete=models.PROTECT)
    csi_unit = models.ForeignKey(CSIUnit, on_delete=models.PROTECT)

    maintenance_type = models.CharField(max_length=100, choices=TYPE_CHOICES)

    # Updated to match Frontend "assetNumbers" (e.g. "Pt-101, Pt-102")
    asset_numbers = models.CharField(
        max_length=255, 
        null=True, 
        blank=True, 
        help_text="List of Asset IDs (e.g. Pt-101, LC-42)"
    )

    section = models.CharField(max_length=100, help_text="Section Code (e.g. ADI-SHB)")
    
    # Added to match Frontend "workDescription"
    work_description = models.TextField(blank=True, help_text="Details of work carried out")
    
    remarks = models.TextField(blank=True)

    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.maintenance_type} - {self.section} ({self.submission_date})"

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

class ACFailureReport(BaseReport):
    """
    Corresponds to ACReportForm.tsx / FORM-AC-04
    """
    # Location Details
    location_code = models.CharField(max_length=50) # Station/Location Code
    
    # Asset Details
    total_ac_units = models.IntegerField(default=0)
    ac_type = models.CharField(max_length=20, choices=[('Split', 'Split'), ('Window', 'Window')])
    
    # Failure Details
    total_fail_count = models.CharField(max_length=20) # '1', '2', ... 'All'
    failure_date_time = models.DateTimeField()
    
    # Status
    under_warranty = models.CharField(max_length=3, choices=[('Yes', 'Yes'), ('No', 'No')])
    under_amc = models.CharField(max_length=3, choices=[('Yes', 'Yes'), ('No', 'No')])
    
    remarks = models.TextField(blank=True, null=True)

    def __str__(self):
        return f"AC-Fail-{self.id} | {self.location_code}"
    
class MovementReport(models.Model):
    # Existing fields
    name = models.CharField(max_length=100)
    date = models.DateField()
    sectional_officer = models.CharField(max_length=100, blank=True, null=True)
    csi = models.CharField(max_length=100)
    
    # Updated Designations to hold combined string like "SSE/SIG/ADI"
    designation = models.CharField(max_length=100) 

    # NEW FIELDS matching the Excel format
    move_from = models.CharField(max_length=100, help_text="Location moving from or Leave status")
    move_to = models.CharField(max_length=100, help_text="Location moving to")
    work_done = models.TextField(help_text="Description of work done")

    # Meta fields
    submitted_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.name} - {self.date}"
    

class JPCReport(models.Model):
    # Location & Date
    station = models.CharField(max_length=100) # From dropdown list
    jpc_date = models.DateField()
    
    # Point Data
    total_points = models.IntegerField()
    inspected_today = models.IntegerField()
    total_inspected_cum = models.IntegerField()
    pending_points = models.IntegerField()
    
    # Authority
    inspection_by = models.CharField(max_length=10, choices=[('SI', 'SI'), ('CSI', 'CSI')])
    inspector_name = models.CharField(max_length=100)
    
    # Metadata
    submitted_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-jpc_date', '-submitted_at']
        verbose_name = "JPC Inspection Report"

    def __str__(self):
        return f"JPC {self.station} - {self.jpc_date}"