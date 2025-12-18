import pandas as pd
import os
import json
from django.core.management.base import BaseCommand
from django.contrib.auth import get_user_model
from forms.models import IPSReport, MaintenanceReport, FailureReport, RelayRoomLog, ACFailureReport
from office.models import (
    SectionalOfficer, CSIUnit, Station, Designation, 
    Manufacturer, FailureReason, SIUnit
)

User = get_user_model()

class Command(BaseCommand):
    help = 'Seeds the database with Office Hierarchy and Test Users'

    def handle(self, *args, **kwargs):
        self.stdout.write("Deleting old data...")
        # Delete dependent reports first
        IPSReport.objects.all().delete()
        MaintenanceReport.objects.all().delete()
        FailureReport.objects.all().delete()
        RelayRoomLog.objects.all().delete()
        ACFailureReport.objects.all().delete()

        Station.objects.all().delete()
        SIUnit.objects.all().delete()
        CSIUnit.objects.all().delete()
        SectionalOfficer.objects.all().delete()
        Designation.objects.all().delete()
        Manufacturer.objects.all().delete()
        FailureReason.objects.all().delete()
        # Optional: Clean up users except superuser if needed
        User.objects.exclude(is_superuser=True).delete()

        self.stdout.write("Seeding Master Data...")

        # 1. Designations (Fixed 'rank' -> 'rank_order')
        designations = ['CSI', 'SSE', 'JE', 'ESM-I', 'ESM-II', 'ESM-III', 'Helper']
        desig_objs = {}
        for idx, title in enumerate(designations):
            # The field in your model is 'rank_order' based on the error trace
            obj, _ = Designation.objects.get_or_create(title=title, defaults={'rank_order': idx})
            desig_objs[title] = obj

        # 2. Failure Dropdowns
        makes = ['Ravel', 'Vighnharta']
        for m in makes: Manufacturer.objects.get_or_create(name=m)

        reasons = [
            'Aspirating System Defective', 'Heat & Smoke Multisensor Defective',
            'SMPS Defective', 'MCP Defective', 'Hooter/ Sounder Defective',
            'Mother Board Defective', 'Flame Sensor Defective', 'Battery Wiring Fault', 
            'Power Supply PCB Defective', 'Battery Defective', 'System Software Defective'
        ]
        for r in reasons: FailureReason.objects.get_or_create(text=r)

        # 3. Hierarchy from JSON
        json_file_path = '/home/harsh2006/Projects/Python/cto-app/cto/ai_studio_code.json'
        if not os.path.exists(json_file_path):
            self.stdout.write(self.style.ERROR(f'File not found: {json_file_path}'))
            return

        self.stdout.write(f"Reading {json_file_path}...")
        try:
            with open(json_file_path, 'r') as f:
                data = json.load(f)
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Error reading JSON: {e}'))
            return

        for officer_name, csi_data in data.items():
            # Create Sectional Officer
            officer_code = officer_name.replace(" ", "_").replace("/", "_").upper()
            officer_obj, _ = SectionalOfficer.objects.get_or_create(
                name=officer_name,
                defaults={'code': officer_code}
            )
            self.stdout.write(f"Processing Officer: {officer_name}")

            for csi_name, si_data in csi_data.items():
                # Create CSI Unit
                csi_obj, _ = CSIUnit.objects.get_or_create(
                    name=csi_name,
                    sectional_officer=officer_obj
                )

                for si_name, si_details in si_data.items():
                    # Create SI Unit
                    si_obj, _ = SIUnit.objects.get_or_create(
                        name=si_name,
                        csi_unit=csi_obj
                    )

                    stations_dict = si_details.get('stations', {})
                    for station_name, station_code in stations_dict.items():
                        # Create Station
                        # Using update_or_create to handle potential duplicates in JSON gracefully
                        # though we cleared the DB, so create should be fine if JSON is clean.
                        Station.objects.update_or_create(
                            code=station_code,
                            defaults={
                                'name': station_name,
                                'csi_unit': csi_obj,
                                'si_unit': si_obj
                            }
                        )

        self.stdout.write(self.style.SUCCESS('Successfully seeded database from JSON!'))