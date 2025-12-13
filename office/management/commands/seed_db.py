from django.core.management.base import BaseCommand
from django.contrib.auth import get_user_model
from office.models import (
    SectionalOfficer, CSIUnit, Station, Designation, 
    Manufacturer, FailureReason
)

User = get_user_model()

class Command(BaseCommand):
    help = 'Seeds the database with Office Hierarchy and Test Users'

    def handle(self, *args, **kwargs):
        self.stdout.write("Deleting old data...")
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

        # 3. Hierarchy (Matches src/constants.ts)
        hierarchy = [
            {
                "officer": "DSTE I",
                "units": [
                    { "csi": "ADI", "stations": ["ADI", "MAN", "VAT", "GER", "BJD"] },
                    { "csi": "SBI", "stations": ["SBI", "KLL", "JKA", "KHDB", "CLDY", "CDK"] }
                ]
            },
            {
                "officer": "DSTE II",
                "units": [
                    { "csi": "VG", "stations": ["VG", "JTN", "SUNR"] },
                    { "csi": "MSH", "stations": ["MSH", "UJA", "KMLI"] }
                ]
            },
            {
                "officer": "ADSTE ADI",
                "units": [
                    { "csi": "GIM", "stations": ["GIM", "BCOB", "SIOB"] }
                ]
            },
            {
                "officer": "ADSTE DHG",
                "units": [
                    { "csi": "DHG", "stations": ["DHG"] },
                    { "csi": "MALB", "stations": ["MALB", "FL", "HALV"] },
                    { "csi": "PNU", "stations": ["PNU", "DISA"] }
                ]
            },
            {
                "officer": "ADSTE RDHP",
                "units": [
                    { "csi": "RDHP", "stations": ["RDHP", "BLDI"] }
                ]
            }
        ]

        for h in hierarchy:
            # Generate a code from the name (e.g. "DSTE I" -> "DSTE_I")
            code = h["officer"].replace(" ", "_").upper()
            officer_obj = SectionalOfficer.objects.create(name=h["officer"], code=code)
            
            for unit in h["units"]:
                csi_obj = CSIUnit.objects.create(
                    name=unit["csi"], 
                    sectional_officer=officer_obj
                )
                
                for stn_code in unit["stations"]:
                    Station.objects.create(
                        code=stn_code, 
                        name=f"Station {stn_code}", 
                        csi_unit=csi_obj
                    )

        self.stdout.write(self.style.SUCCESS('Successfully seeded database!'))