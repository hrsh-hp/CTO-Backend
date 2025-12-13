from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import HierarchyViewSet, StationViewSet, MasterDataView

router = DefaultRouter()
router.register(r'hierarchy', HierarchyViewSet, basename='hierarchy')
router.register(r'stations', StationViewSet, basename='stations')
router.register(r'master-data', MasterDataView, basename='master-data')

urlpatterns = [
    path('', include(router.urls)),
]