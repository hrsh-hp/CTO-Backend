from rest_framework import serializers
from .models import ACFailureReport, FailureReport, IPSReport, IPSEntry, IPSModuleType, IPSCompany, JPCReport, MaintenanceReport, MovementReport, RelayRoomLog
from office.models import CSIUnit, SectionalOfficer, Designation, Station

class IPSEntrySerializer(serializers.ModelSerializer):
    # Allow frontend to send string names ("SMR", "HBL") instead of IDs
    module_type = serializers.SlugRelatedField(
        slug_field='name', 
        queryset=IPSModuleType.objects.all()
    )
    company = serializers.SlugRelatedField(
        slug_field='name', 
        queryset=IPSCompany.objects.all()
    )

    class Meta:
        model = IPSEntry
        fields = [
            'id', 'module_type', 'company', 
            'qty_defective', 'qty_spare', 
            'qty_spare_amc', 'qty_defective_amc'
        ]

class IPSReportSerializer(serializers.ModelSerializer):
    entries = IPSEntrySerializer(many=True)
    
    # Lookup CSI by name ("ADI") instead of ID
    csi = serializers.SlugRelatedField(
        source='csi_unit', # Map 'csi' in JSON to 'csi_unit' in Model
        slug_field='name', 
        queryset=CSIUnit.objects.all()
    )

    class Meta:
        model = IPSReport
        fields = [
            'id', 'submission_date', 'week_from', 'week_to', 
            'csi', 'remarks', 'entries', 'created_at'
        ]

    def create(self, validated_data):
        """
        Handle the creation of the Report AND its nested Entries in one transaction.
        """
        entries_data = validated_data.pop('entries')
        
        # Auto-populate sectional_officer from the selected CSI Unit
        csi_unit = validated_data.get('csi_unit')
        if csi_unit:
            validated_data['sectional_officer'] = csi_unit.sectional_officer
            
            # Auto-populate reporter_name if not provided
            # User requested: "CSI-<Station>" format
            if 'reporter_name' not in validated_data:
                validated_data['reporter_name'] = f"CSI-{csi_unit.name}"
                # reporter_designation is a ForeignKey, so we can't assign a string. 
                # Leaving it null is fine as it is nullable.

        # Auto-populate reporter details if available in context
        request = self.context.get('request')
        if request and request.user and request.user.is_authenticated:
            validated_data['submitted_by'] = request.user
        
        # Fallback for reporter_name if still missing (e.g. if csi_unit was somehow None)
        if 'reporter_name' not in validated_data:
             validated_data['reporter_name'] = "System/Unknown"
        

        # 1. Create the Parent Report
        report = IPSReport.objects.create(**validated_data)
        
        # 2. Create Child Entries
        # We loop through entries and link them to the new report
        entry_instances = []
        for entry_data in entries_data:
            entry_instances.append(
                IPSEntry(report=report, **entry_data)
            )
        
        # Bulk create is faster than saving one by one
        IPSEntry.objects.bulk_create(entry_instances)
        
        return report
    
class MaintenanceReportSerializer(serializers.ModelSerializer):
    # Map frontend 'sectionalOfficer' (string name) to the ID automatically
    sectional_officer = serializers.SlugRelatedField(
        slug_field='name', 
        queryset=SectionalOfficer.objects.all()
    )
    
    # Map frontend 'csi' (string name) to model field 'csi_unit'
    csi = serializers.SlugRelatedField(
        source='csi_unit', 
        slug_field='name', 
        queryset=CSIUnit.objects.all()
    )

    class Meta:
        model = MaintenanceReport
        fields = [
            'id', 
            'submission_date', 
            'sectional_officer', 
            'csi', 
            'maintenance_type', 
            'asset_numbers', 
            'section', 
            'work_description', 
            'remarks',
            'created_at'
        ]

class RelayRoomLogSerializer(serializers.ModelSerializer):
    # Lookup the Station object by its 'code' field based on the input string
    location = serializers.SlugRelatedField(
        slug_field='code', 
        queryset=Station.objects.all()
    )
    
    # Map frontend 'date' -> backend 'log_date'
    date = serializers.DateField(source='log_date')
    
    # Map frontend fields to backend fields
    name = serializers.CharField(source='reporter_name')
    designation = serializers.SlugRelatedField(
        source='reporter_designation', 
        slug_field='title', 
        queryset=Designation.objects.all(), 
        allow_null=True
    )
    csi = serializers.SlugRelatedField(
        source='csi_unit', 
        slug_field='name', 
        queryset=CSIUnit.objects.all()
    )
    sectional_officer = serializers.SlugRelatedField(
        slug_field='name', 
        queryset=SectionalOfficer.objects.all()
    )
    submitted_at = serializers.DateTimeField(source='created_at', read_only=True)

    class Meta:
        model = RelayRoomLog
        fields = [
            'id', 'name', 'designation', 
            'sectional_officer', 'csi', 'date', 'location', 
            'opening_time', 'closing_time', 'sn_opening', 
            'sn_closing', 'opening_code', 'remarks', 'submitted_at'
        ]

class ACFailureReportSerializer(serializers.ModelSerializer):
    # Map frontend 'csi' -> backend 'csi_unit'
    csi = serializers.SlugRelatedField(
        source='csi_unit',
        slug_field='name',
        queryset=CSIUnit.objects.all()
    )
    
    # Map frontend 'sectional_officer' (name) -> backend ID
    sectional_officer = serializers.SlugRelatedField(
        slug_field='name',
        queryset=SectionalOfficer.objects.all()
    )

    # If frontend sends 'reporter_designation' as string title
    reporter_designation = serializers.SlugRelatedField(
        slug_field='title',
        queryset=Designation.objects.all(),
        allow_null=True,
        required=False
    )

    class Meta:
        model = ACFailureReport
        fields = [
            'id', 'location_code', 'total_ac_units', 'ac_type',
            'total_fail_count', 'failure_date_time', 'under_warranty',
            'under_amc', 'remarks', 'csi', 'sectional_officer',
            'reporter_name', 'reporter_designation', 'created_at', 'status'
        ]

    def create(self, validated_data):
        # Auto-populate reporter details if available in context
        request = self.context.get('request')
        if request and request.user and request.user.is_authenticated:
            validated_data['submitted_by'] = request.user
            
        return super().create(validated_data)
    
class FailureReportSerializer(serializers.ModelSerializer):
    # Handle 'reason' as a list of strings for the frontend
    reason = serializers.JSONField() 

    class Meta:
        model = FailureReport
        fields = '__all__'

class MovementReportSerializer(serializers.ModelSerializer):
    class Meta:
        model = MovementReport
        fields = '__all__'
        # submitted_at is read-only by default

class JPCReportSerializer(serializers.ModelSerializer):
    class Meta:
        model = JPCReport
        fields = '__all__'