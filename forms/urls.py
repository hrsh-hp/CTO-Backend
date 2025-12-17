from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import ACFailureReportViewSet, FailureReportViewSet, IPSReportViewSet, JPCReportViewSet, MaintenanceReportViewSet, MovementReportViewSet,export_movement_excel, RelayRoomLogViewSet

router = DefaultRouter()
router.register(r'ips-reports', IPSReportViewSet, basename='ips-report')
router.register(r'maintenance-reports', MaintenanceReportViewSet)
router.register(r'relay-logs', RelayRoomLogViewSet)
router.register(r'ac-reports', ACFailureReportViewSet)
router.register(r'failure-reports', FailureReportViewSet)
router.register(r'movement-reports', MovementReportViewSet)
router.register(r'jpc-reports', JPCReportViewSet)

urlpatterns = [
    path('movement-reports/export-excel/', export_movement_excel, name='export_movement_excel'),
    path('', include(router.urls)),
]