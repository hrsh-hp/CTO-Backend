import os
import django
import random
from datetime import date, time, timedelta

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'cto.settings')
django.setup()

from forms.models import RelayRoomLog
from office.models import Station, CSIUnit, SectionalOfficer

def seed_relay_logs():
    print("Seeding Relay Room Logs...")
    
    stations = list(Station.objects.all())
    if not stations:
        print("No stations found. Please run seed_db first.")
        return

    # Create 10 dummy logs
    for i in range(10):
        station = random.choice(stations)
        csi_unit = station.csi_unit
        sectional_officer = csi_unit.sectional_officer
        
        log_date = date.today() - timedelta(days=random.randint(0, 30))
        opening_time = time(random.randint(8, 18), random.randint(0, 59))
        # Closing time 1-3 hours later
        closing_hour = min(23, opening_time.hour + random.randint(1, 3))
        closing_time = time(closing_hour, random.randint(0, 59))
        
        log = RelayRoomLog.objects.create(
            log_date=log_date,
            location=station,
            opening_time=opening_time,
            closing_time=closing_time,
            opening_code=f"CODE-{random.randint(100, 999)}",
            sn_opening=str(random.randint(1000, 9999)),
            sn_closing=str(random.randint(1000, 9999)),
            remarks=f"Routine Inspection {i+1}",
            csi_unit=csi_unit,
            sectional_officer=sectional_officer,
            reporter_name="System Seeder",
            status='Open'
        )
        print(f"Created log for {station.name} on {log_date}")

    print("Done seeding Relay Room Logs.")

if __name__ == '__main__':
    seed_relay_logs()
