from django.contrib import admin
from  forms.models import FailureReport,ACFailureReport, BaseReport,RelayRoomLog,MaintenanceReport,IPSReport,IPSEntry
# Register your models here.

admin.site.register(FailureReport)
admin.site.register(RelayRoomLog)
admin.site.register(MaintenanceReport)
admin.site.register(IPSReport)
admin.site.register(ACFailureReport)
admin.site.register(IPSEntry)