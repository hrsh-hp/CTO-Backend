from django.shortcuts import render

# Create your views here.
from rest_framework import viewsets, generics
from rest_framework.response import Response
from rest_framework.decorators import action
from .models import (
    SectionalOfficer, Designation, Manufacturer, FailureReason, Station
)
from .serializers import (
    HierarchySerializer, DesignationSerializer, 
    ManufacturerSerializer, FailureReasonSerializer,
    StationSerializer
)

class HierarchyViewSet(viewsets.ReadOnlyModelViewSet):
    """
    GET /api/office/hierarchy/
    Returns the full nested JSON tree for the dropdowns.
    """
    queryset = SectionalOfficer.objects.prefetch_related('csi_units__stations').filter(is_active=True)
    serializer_class = HierarchySerializer
    pagination_class = None  # Disable pagination to get the full tree at once

class MasterDataView(viewsets.ViewSet):
    """
    A unified endpoint to fetch simple lists (Designations, Makes, Reasons).
    Useful for loading all static dropdowns in one request.
    GET /api/office/master-data/
    """
    def list(self, request):
        return Response({
            "designations": DesignationSerializer(Designation.objects.all(), many=True).data,
            "makes": ManufacturerSerializer(Manufacturer.objects.all(), many=True).data,
            "reasons": FailureReasonSerializer(FailureReason.objects.all(), many=True).data,
        })

class StationViewSet(viewsets.ReadOnlyModelViewSet):
    """
    GET /api/office/stations/?csi_unit=5
    Allows filtering stations dynamically if needed.
    """
    queryset = Station.objects.all()
    serializer_class = StationSerializer
    
    def get_queryset(self):
        queryset = super().get_queryset()
        csi_id = self.request.query_params.get('csi_unit')
        if csi_id:
            queryset = queryset.filter(csi_unit_id=csi_id)
        return queryset