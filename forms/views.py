from django.shortcuts import render

# Create your views here.
from rest_framework import viewsets, filters
from django_filters.rest_framework import DjangoFilterBackend
from .models import IPSReport
from .serializers import IPSReportSerializer

class IPSReportViewSet(viewsets.ModelViewSet):
    # Prefetch related fields to prevent N+1 SQL query issues
    queryset = IPSReport.objects.all().select_related('csi_unit').prefetch_related(
        'entries__module_type', 
        'entries__company'
    )
    serializer_class = IPSReportSerializer
    
    # Filtering capabilities
    filter_backends = [DjangoFilterBackend, filters.OrderingFilter]
    filterset_fields = ['csi_unit__name', 'submission_date'] # Filter by CSI name or date
    ordering_fields = ['submission_date', 'created_at'] # Allow ordering by date fields

# class IPSReportViewSet(viewsets.ModelViewSet):
#     # Important: prefetch_related reduces 50+ SQL queries to just 2 or 3
#     queryset = IPSReport.objects.all().select_related('csi_unit').prefetch_related(
#         'entries__module_type', 
#         'entries__company'
#     ).order_by('-submission_date')
    
#     serializer_class = IPSReportSerializer