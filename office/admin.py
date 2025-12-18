from django.contrib import admin
from office.models import Designation, Station, CSIUnit, SectionalOfficer, Manufacturer, FailureReason, IPSModuleType, IPSCompany
# Register your models here.

admin.site.register(Designation)
admin.site.register(Station)
admin.site.register(CSIUnit)
admin.site.register(SectionalOfficer)
admin.site.register(Manufacturer)
admin.site.register(FailureReason)
admin.site.register(IPSModuleType)
admin.site.register(IPSCompany)