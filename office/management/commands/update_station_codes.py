import pandas as pd
import os
from django.core.management.base import BaseCommand
from office.models import Station

class Command(BaseCommand):
    help = 'Updates Station codes from DATALOGGER AND RTU SHEET 2025.xlsx'

    def handle(self, *args, **kwargs):
        file_path = '/home/harsh2006/Projects/Python/cto-app/cto/DATALOGGER AND RTU SHEET 2025.xlsx'
        if not os.path.exists(file_path):
            self.stdout.write(self.style.ERROR(f'File not found: {file_path}'))
            return

        try:
            # Read Excel, header is at row 2 (index 1)
            df = pd.read_excel(file_path, header=1)
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Error reading excel: {e}'))
            return

        # Normalize column names
        df.columns = [str(col).strip().upper() for col in df.columns]
        
        if 'STATION NAME' not in df.columns or 'STATION CODE' not in df.columns:
             self.stdout.write(self.style.ERROR(f'Required columns STATION NAME and STATION CODE not found. Found: {df.columns}'))
             return

        updated_count = 0
        not_found_count = 0
        
        # Create a mapping of normalized DB station names to Station objects
        # We'll try to match loosely
        db_stations = Station.objects.all()
        
        # Helper to normalize strings for comparison
        def normalize(s):
            if not isinstance(s, str): return ""
            return s.lower().replace(" ", "").replace("'", "").replace("(", "").replace(")", "").replace("-", "").replace(".", "")

        db_map = {normalize(s.name): s for s in db_stations}

        for index, row in df.iterrows():
            excel_name = row['STATION NAME']
            excel_code = row['STATION CODE']
            
            if pd.isna(excel_name) or pd.isna(excel_code):
                continue
                
            excel_name = str(excel_name).strip()
            excel_code = str(excel_code).strip()
            
            norm_excel_name = normalize(excel_name)
            
            station = db_map.get(norm_excel_name)
            
            # Try some manual overrides/fuzzy logic if direct match fails
            if not station:
                # Try matching "SABARMATI BG" -> "SABARMATISBIB"
                if "SABARMATIBG" == norm_excel_name:
                     station = db_map.get("sabarmatisbib")
                elif "SABARMATIDCABIN" == norm_excel_name:
                     station = db_map.get("sabarmatid")
                elif "KANKARIYA" == norm_excel_name:
                     station = db_map.get("kankaria")
                elif "VIJAPUR" == norm_excel_name:
                     station = db_map.get("vijapurvjf")
                elif "CHANDLODIYAA" == norm_excel_name:
                     station = db_map.get("chandlodiyaa") # Assuming DB has CHANDLODIYA 'A' -> chandlodiyaa
                elif "CHANDLODIYAB" == norm_excel_name:
                     station = db_map.get("chandlodiyab")
                elif "SHAHIBAUG" == norm_excel_name:
                     # Check if we have something similar
                     pass
                elif "MAHESANA" == norm_excel_name:
                     station = db_map.get("mahesanajn") # Maybe?
                elif "VIRAMGAM" == norm_excel_name:
                     station = db_map.get("viramgamjn")
                elif "DHRANGADHRA" == norm_excel_name:
                     station = db_map.get("dhrangadhra") # Should have matched if exact
                elif "MALIYAMIYANA" == norm_excel_name:
                     station = db_map.get("maliyamiya") # Maybe truncated?
                elif "SAMAKHIYALI" == norm_excel_name:
                     station = db_map.get("samakhyali") # Spelling diff?
                elif "GANDHIDHAMBBLOCKCABIN" == norm_excel_name:
                     station = db_map.get("gandhidhamb")
                elif "BHUJOC" == norm_excel_name:
                     station = db_map.get("bhuj")
                elif "NALIYA" == norm_excel_name:
                     station = db_map.get("naliya")
                elif "RADHANPUR" == norm_excel_name:
                     station = db_map.get("radhanpur")
                elif "DIYODAR" == norm_excel_name:
                     station = db_map.get("diyodar")
                elif "BHILDI" == norm_excel_name:
                     station = db_map.get("bhildi")
                elif "MITHA" == norm_excel_name:
                     station = db_map.get("mitha")
                elif "DEVALIYA" == norm_excel_name:
                     station = db_map.get("devalia") # Spelling diff?
                elif "KATARIYA" == norm_excel_name:
                     station = db_map.get("kataria")
                elif "ASARVA" == norm_excel_name:
                     station = db_map.get("asarva")
                elif "NARODA" == norm_excel_name:
                     station = db_map.get("naroda")
                elif "DABHODA" == norm_excel_name:
                     station = db_map.get("dabhoda")
                elif "NANDOLDAHEGAM" == norm_excel_name:
                     station = db_map.get("nandoldehegam")
                elif "RAKHIYAL" == norm_excel_name:
                     station = db_map.get("rakhiyal")
                elif "TALOD" == norm_excel_name:
                     station = db_map.get("talod")
                elif "PRATIJ" == norm_excel_name:
                     station = db_map.get("prantij")
                elif "SONASAN" == norm_excel_name:
                     station = db_map.get("sonasan")
                elif "HIMMATNAGAR" == norm_excel_name:
                     station = db_map.get("himmatnagar")
                elif "KHEDBHRMA" == norm_excel_name:
                     station = db_map.get("khedbrahma")
                elif "IDAR" == norm_excel_name:
                     station = db_map.get("idar")
                elif "LC17A" == norm_excel_name:
                     pass

            if station:
                # Update the code
                # Check if code is already taken by another station (to avoid unique constraint error)
                if Station.objects.filter(code=excel_code).exclude(id=station.id).exists():
                    self.stdout.write(self.style.WARNING(f"Skipping {excel_name}: Code {excel_code} already in use."))
                    continue

                station.code = excel_code
                try:
                    station.save()
                    self.stdout.write(self.style.SUCCESS(f"Updated {station.name}: {excel_code}"))
                    updated_count += 1
                except Exception as e:
                    self.stdout.write(self.style.ERROR(f"Failed to save {station.name}: {e}"))
            else:
                self.stdout.write(self.style.WARNING(f"Station not found in DB: {excel_name}"))
                not_found_count += 1

        self.stdout.write(self.style.SUCCESS(f'Finished. Updated: {updated_count}, Not Found: {not_found_count}'))
