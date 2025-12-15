from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import IPSReportViewSet, MaintenanceReportViewSet, RelayRoomLogViewSet

router = DefaultRouter()
router.register(r'ips-reports', IPSReportViewSet, basename='ips-report')
router.register(r'maintenance-reports', MaintenanceReportViewSet)
router.register(r'relay-logs', RelayRoomLogViewSet)


urlpatterns = [
    path('', include(router.urls)),
]