from django.db import models

class SectionalOfficer(models.Model):
    """
    Represents the top-level hierarchy (e.g., 'DSTE I', 'ADSTE ADI').
    """
    name = models.CharField(max_length=100, unique=True)
    code = models.CharField(max_length=20, unique=True, blank=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name

class CSIUnit(models.Model):
    """
    Represents a CSI Unit (e.g., 'ADI', 'SBI', 'VG').
    Linked to a Sectional Officer.
    """
    name = models.CharField(max_length=100) # e.g. "ADI"
    sectional_officer = models.ForeignKey(
        SectionalOfficer, 
        on_delete=models.CASCADE,
        related_name='csi_units'
    )
    
    class Meta:
        verbose_name = "CSI Unit"
        unique_together = ('name', 'sectional_officer')

    def __str__(self):
        return f"{self.name} (under {self.sectional_officer})"

class SIUnit(models.Model):
    """
    Represents a Sectional Inspector Unit (e.g., 'SI VTA').
    Linked to a CSI Unit.
    """
    name = models.CharField(max_length=100)
    csi_unit = models.ForeignKey(
        CSIUnit,
        on_delete=models.CASCADE,
        related_name='si_units'
    )

    class Meta:
        verbose_name = "SI Unit"
        unique_together = ('name', 'csi_unit')

    def __str__(self):
        return f"{self.name} (under {self.csi_unit.name})"

class Station(models.Model):
    """
    Physical Railway Stations (e.g., 'NDLS', 'ADI').
    Linked to a CSI Unit.
    """
    code = models.CharField(max_length=50, unique=True) # e.g. "ADI"
    name = models.CharField(max_length=100)
    csi_unit = models.ForeignKey(
        CSIUnit, 
        on_delete=models.CASCADE,
        related_name='stations'
    )
    si_unit = models.ForeignKey(
        SIUnit,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='stations'
    )

    def __str__(self):
        return f"{self.code} - {self.name}"

class Designation(models.Model):
    """
    Dropdown options for Staff Designation (e.g., 'SSE', 'JE').
    """
    title = models.CharField(max_length=50, unique=True)
    rank_order = models.PositiveIntegerField(default=0, help_text="For sorting purposes")

    class Meta:
        ordering = ['rank_order']

    def __str__(self):
        return self.title

# --- Failure Report Dropdowns ---

class Manufacturer(models.Model):
    """Dropdown for 'Make' (e.g., 'Ravel', 'Vighnharta')"""
    name = models.CharField(max_length=100, unique=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name

class FailureReason(models.Model):
    """Dropdown options for 'Reason' checkboxes"""
    text = models.CharField(max_length=255, unique=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.text

# --- IPS Report Dropdowns ---

class IPSModuleType(models.Model):
    """Dropdown for IPS Modules (e.g., 'Inverter', 'SMR')"""
    name = models.CharField(max_length=100, unique=True)
    
    def __str__(self):
        return self.name

class IPSCompany(models.Model):
    """Dropdown for IPS Companies (e.g., 'STATCON', 'HBL')"""
    name = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.name