from rest_framework import serializers
from .models import (
    SectionalOfficer, CSIUnit, Station, Designation,
    Manufacturer, FailureReason, IPSModuleType, IPSCompany, SIUnit
)

# --- Basic Serializers ---

class StationSerializer(serializers.ModelSerializer):
    class Meta:
        model = Station
        fields = ['id', 'code', 'name']

class DesignationSerializer(serializers.ModelSerializer):
    class Meta:
        model = Designation
        fields = ['id', 'title', 'rank_order']

class ManufacturerSerializer(serializers.ModelSerializer):
    class Meta:
        model = Manufacturer
        fields = ['id', 'name']

class FailureReasonSerializer(serializers.ModelSerializer):
    class Meta:
        model = FailureReason
        fields = ['id', 'text']

class SIUnitSerializer(serializers.ModelSerializer):
    class Meta:
        model = SIUnit
        fields = ['id', 'name']

# --- Hierarchy / Nested Serializers ---

class CSIUnitNestedSerializer(serializers.ModelSerializer):
    # Nesting Stations inside CSI
    stations = StationSerializer(many=True, read_only=True)
    si_units = SIUnitSerializer(many=True, read_only=True)

    class Meta:
        model = CSIUnit
        fields = ['id', 'name', 'stations', 'si_units']

class HierarchySerializer(serializers.ModelSerializer):
    """
    Serializes the full tree: Officer -> CSIs -> Stations.
    Renamed 'csi_units' to 'csis' to match your Frontend Interface.
    """
    csis = CSIUnitNestedSerializer(many=True, read_only=True, source='csi_units')

    class Meta:
        model = SectionalOfficer
        fields = ['id', 'name', 'csis']