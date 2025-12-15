import pandas as pd
import os
from django.core.management.base import BaseCommand
from django.contrib.auth import get_user_model
from forms.models import IPSReport, MaintenanceReport, FailureReport, RelayRoomLog
from office.models import (
    SectionalOfficer, CSIUnit, Station, Designation, 
    Manufacturer, FailureReason
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

        Station.objects.all().delete()
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

        # 3. Hierarchy from Excel
        file_path = '/home/harsh2006/Projects/Python/cto-app/cto/Officer and Staff wise station details.xlsx'
        if not os.path.exists(file_path):
            self.stdout.write(self.style.ERROR(f'File not found: {file_path}'))
            return

        try:
            df = pd.read_excel(file_path, header=1)
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Error reading excel: {e}'))
            return
        
        # Forward fill Section Officer and CSI
        df['Section Officer'] = df['Section Officer'].ffill()
        df['CSI'] = df['CSI'].ffill()
        
        # Drop rows where Station is NaN
        df = df.dropna(subset=['Station'])

        for index, row in df.iterrows():
            officer_name = str(row['Section Officer']).strip()
            csi_name = str(row['CSI']).strip()
            station_name = str(row['Station']).strip()
            
            if officer_name == 'nan' or csi_name == 'nan':
                continue

            # Create or Get Sectional Officer
            officer_code = officer_name.replace(" ", "_").replace("/", "_").upper()
            officer_obj, _ = SectionalOfficer.objects.get_or_create(
                name=officer_name,
                defaults={'code': officer_code}
            )

            # Create or Get CSI Unit
            csi_obj, _ = CSIUnit.objects.get_or_create(
                name=csi_name,
                sectional_officer=officer_obj
            )

            # Create Station
            # Use station name as code, replacing spaces with underscores
            station_code = station_name.replace(" ", "_").upper()
            
            # Check if station exists (to avoid duplicates if excel has duplicates)
            if not Station.objects.filter(code=station_code).exists():
                Station.objects.create(
                    code=station_code,
                    name=station_name,
                    csi_unit=csi_obj
                )

        self.stdout.write(self.style.SUCCESS('Successfully seeded database from Excel!'))