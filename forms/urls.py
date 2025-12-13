from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import IPSReportViewSet

router = DefaultRouter()
router.register(r'ips-reports', IPSReportViewSet, basename='ips-report')

urlpatterns = [
    path('', include(router.urls)),
]