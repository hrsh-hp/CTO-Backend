import pandas as pd
import xlsxwriter

# Data extraction for all dates
# Structure: [SR, SSE, Section, A_D, A_A, A_N, B_D, B_A, B_N, C_D, C_A, C_N, D_D, D_A, D_N]
# Note: Totals are calculated automatically by the script to ensure accuracy.

data_sheets = {
    "01-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 0,0,0, 10,10,0, 0,0,0],
        [2, "ADI", "SI VTA", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 3,3,0, 2,2,0, 0,0,0],
        [2, "ADI", "SI HMT", 0,0,0, 2,2,0, 0,0,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 0,0,0, 5,5,0, 0,0,0],
        [3, "SBI", "SI SBI", 0,0,0, 0,0,0, 8,8,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 1,1,0, 5,5,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 2,2,0, 3,3,0, 0,0,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 1,1,0, 3,3,0, 1,1,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 2,2,0, 3,3,0, 0,0,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 2,2,0, 2,2,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 0,0,0, 2,2,0, 1,1,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI BLDI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI Br.GIM", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI MALB", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 2,2,0, 1,1,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [13, "VG", "SI BAJN", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 0,0,0, 2,2,0, 1,1,0],
    ],
    "02-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 1,1,0, 9,9,0, 0,0,0],
        [2, "ADI", "SI VTA", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 0,0,0, 7,7,0, 0,0,0],
        [2, "ADI", "SI HMT", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [3, "SBI", "SI SBI", 0,0,0, 0,0,0, 5,5,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 2,2,0, 4,4,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 2,2,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 3,3,0, 4,4,0, 0,0,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 2,2,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 0,0,0, 0,0,0, 1,1,0],
        [7, "PNU", "SI BLDI", 0,0,0, 0,0,0, 2,2,0, 1,1,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI Br.GIM", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 2,2,0, 1,1,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 0,0,0, 0,0,0, 1,1,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [12, "MALB", "SI MALB", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [13, "VG", "SI BAJN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
    ],
    "03-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 0,0,0, 7,7,0, 0,0,0],
        [2, "ADI", "SI VTA", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [2, "ADI", "SI HMT", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 2,2,0, 3,2,1, 0,0,0],
        [3, "SBI", "SI SBI", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 1,1,0, 4,4,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 1,1,0, 2,2,0, 1,1,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 2,2,0, 4,4,0, 0,0,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 2,2,0, 4,4,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 1,1,0, 1,1,0, 1,1,0],
        [7, "PNU", "SI BLDI", 0,0,0, 0,0,0, 2,2,0, 1,1,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 2,2,0, 2,2,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI Br.GIM", 0,0,0, 0,0,0, 0,0,0, 1,1,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 0,0,0, 1,1,0],
        [12, "MALB", "SI MALB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [13, "VG", "SI BAJN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
    ],
    "04-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 0,0,0, 9,9,0, 0,0,0],
        [2, "ADI", "SI VTA", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 2,2,0, 3,3,0, 0,0,0],
        [2, "ADI", "SI HMT", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 0,0,0, 3,3,0, 1,1,0],
        [3, "SBI", "SI SBI", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 1,1,0, 7,7,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 1,1,0, 4,4,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 2,2,0, 0,0,0, 1,1,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 3,3,0, 3,3,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 1,1,0, 2,2,0, 1,1,0],
        [7, "PNU", "SI BLDI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 2,2,0, 2,2,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 1,1,0],
        [9, "GIM", "SI Br.GIM", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 2,2,0, 6,6,0, 0,0,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [12, "MALB", "SI MALB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 2,2,0, 0,0,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [13, "VG", "SI BAJN", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
    ],
    "05-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 0,0,0, 8,8,0, 0,0,0],
        [2, "ADI", "SI VTA", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [2, "ADI", "SI HMT", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 1,1,0, 4,4,0, 0,0,0],
        [3, "SBI", "SI SBI", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 2,2,0, 3,3,0, 0,0,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 2,2,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 6,6,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI BLDI", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 0,0,0, 7,7,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 2,2,0, 0,0,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 1,1,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 1,1,0],
        [9, "GIM", "SI Br.GIM", 0,0,0, 0,0,0, 5,5,0, 0,0,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 0,0,0, 1,1,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [12, "MALB", "SI MALB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [13, "VG", "SI BAJN", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 2,2,0, 3,3,0, 0,0,0],
    ],
    "06-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 0,0,0, 6,6,0, 0,0,0],
        [2, "ADI", "SI VTA", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [2, "ADI", "SI HMT", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [3, "SBI", "SI SBI", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 2,2,0, 3,3,0, 0,0,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 4,4,0, 1,1,0, 0,0,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 2,2,0, 2,2,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI BLDI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 2,2,0, 4,4,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI Br.GIM", 1,1,0, 0,0,0, 2,2,0, 0,0,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 1,1,0, 4,4,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [12, "MALB", "SI MALB", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 2,2,0, 0,0,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 0,0,0, 0,0,0, 1,1,0],
        [13, "VG", "SI BAJN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
    ],
    "07-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 0,0,0, 7,7,0, 0,0,0],
        [2, "ADI", "SI VTA", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [2, "ADI", "SI HMT", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [3, "SBI", "SI SBI", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 1,1,0, 6,6,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 3,3,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 2,2,0, 2,2,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI BLDI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 2,2,0, 6,6,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 2,2,0, 5,5,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI Br.GIM", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 2,2,0, 3,3,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI MALB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 0,0,0, 3,3,0, 1,1,0],
        [13, "VG", "SI BAJN", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 0,0,0, 1,1,0, 1,1,0],
    ],
    "08-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 0,0,0, 10,10,0, 0,0,0],
        [2, "ADI", "SI VTA", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 0,0,0, 3,3,0, 1,1,0],
        [2, "ADI", "SI HMT", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [3, "SBI", "SI SBI", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 2,2,0, 6,6,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 2,2,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 3,3,0, 2,2,0, 0,0,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 2,2,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [7, "PNU", "SI BLDI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 3,3,0, 4,4,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI Br.GIM", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 2,2,0, 2,2,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [12, "MALB", "SI MALB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [13, "VG", "SI BAJN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
    ],
    "09-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 0,0,0, 8,8,0, 0,0,0],
        [2, "ADI", "SI VTA", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [2, "ADI", "SI HMT", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [3, "SBI", "SI SBI", 0,0,0, 1,1,0, 4,4,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 1,1,0, 5,5,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 2,2,0, 1,1,0, 0,0,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 3,3,0, 1,1,0, 0,0,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI BLDI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 0,0,0, 4,4,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI Br.GIM", 0,0,0, 2,2,0, 6,6,0, 0,0,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 2,2,0, 5,5,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [12, "MALB", "SI MALB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [13, "VG", "SI BAJN", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
    ],
    "10-12-2025": [
        [1, "ADI RRI", "ADI RRI", 0,0,0, 0,0,0, 9,9,0, 0,0,0],
        [2, "ADI", "SI VTA", 2,2,0, 1,1,0, 1,1,0, 0,0,0],
        [2, "ADI", "SI ASV", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [2, "ADI", "SI HMT", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [3, "SBI", "SI SHB", 0,0,0, 0,0,0, 5,5,0, 0,0,0],
        [3, "SBI", "SI SBI", 0,0,0, 0,0,0, 5,5,0, 0,0,0],
        [3, "SBI", "SI SAU", 0,0,0, 2,2,0, 5,5,0, 0,0,0],
        [3, "SBI", "SI VG", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI KLL", 0,0,0, 1,1,0, 3,3,0, 0,0,0],
        [4, "KLL", "SI GNC", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [4, "KLL", "SI UMN", 0,0,0, 1,1,0, 5,5,0, 0,0,0],
        [5, "Br MSH", "SI Br. MSH", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "MSH RRI", 0,0,0, 3,3,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI PTN", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [5, "Br MSH", "SI KTRD", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [6, "N MSH", "SI PNU", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [6, "N MSH", "SI SID", 0,0,0, 2,2,0, 3,3,0, 0,0,0],
        [6, "N MSH", "SI N MSH", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [7, "PNU", "PNU RRI", 0,0,0, 1,1,0, 2,2,0, 0,0,0],
        [7, "PNU", "SI Br. PNU", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI DEOR", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [7, "PNU", "SI BLDI", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [8, "RDHP", "SI RDHP", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [8, "RDHP", "SI SNLR", 0,0,0, 0,0,0, 6,6,0, 0,0,0],
        [8, "RDHP", "SI RRI", 0,0,0, 0,0,0, 2,2,0, 0,0,0],
        [9, "GIM", "SI AI", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI BHUJ", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [9, "GIM", "SI Br.GIM", 0,0,0, 2,2,0, 5,5,0, 0,0,0],
        [9, "GIM", "SI BCOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [10, "SIOB Br.", "SIOB/BR", 0,0,0, 2,2,0, 3,3,0, 0,0,0],
        [10, "SIOB Br.", "SI AAR", 0,0,0, 2,2,0, 1,1,0, 0,0,0],
        [11, "DHG", "SI DHG RRI", 0,0,0, 1,1,0, 1,1,0, 0,0,0],
        [11, "DHG", "SI Br.DHG", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [11, "DHG", "SI HVD", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI MALB", 0,0,0, 1,1,0, 0,0,0, 0,0,0],
        [12, "MALB", "SI SIOB", 0,0,0, 0,0,0, 0,0,0, 0,0,0],
        [13, "VG", "VG RRI", 0,0,0, 1,1,0, 4,4,0, 0,0,0],
        [13, "VG", "SI BAJN", 0,0,0, 0,0,0, 3,3,0, 0,0,0],
        [13, "VG", "SI JTX", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
        [13, "VG", "SI BKD", 0,0,0, 0,0,0, 1,1,0, 0,0,0],
    ],
}

def create_excel(data_dict, filename="Report.xlsx"):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    workbook = writer.book

    # Formats
    header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'border': 1
    })
    merge_format = workbook.add_format({
        'bold': True, 'valign': 'vcenter', 'align': 'center', 'border': 1
    })
    cell_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    total_row_format = workbook.add_format({
        'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })

    for date_str, rows in data_dict.items():
        # Calculate totals for each row
        processed_rows = []
        col_totals = [0] * 15 # D, A, N for A, B, C, D, Total

        for row in rows:
            # row: [SR, SSE, Section, A_D, A_A, A_N, B_D, B_A, B_N, C_D, C_A, C_N, D_D, D_A, D_N]
            # Calculate Total D, A, N
            tot_d = row[3] + row[6] + row[9] + row[12]
            tot_a = row[4] + row[7] + row[10] + row[13]
            tot_n = row[5] + row[8] + row[11] + row[14]
            
            full_row = row + [tot_d, tot_a, tot_n]
            processed_rows.append(full_row)

            # Add to column totals
            for i in range(3, 15):
                col_totals[i-3] += row[i]
        
        # Calculate grand totals
        grand_tot_d = col_totals[0] + col_totals[3] + col_totals[6] + col_totals[9]
        grand_tot_a = col_totals[1] + col_totals[4] + col_totals[7] + col_totals[10]
        grand_tot_n = col_totals[2] + col_totals[5] + col_totals[8] + col_totals[11]
        
        col_totals.extend([grand_tot_d, grand_tot_a, grand_tot_n])

        # Create DataFrame
        columns = [
            'SR NO', 'SSE', 'Section SI',
            'A_D', 'A_A', 'A_N',
            'B_D', 'B_A', 'B_N',
            'C_D', 'C_A', 'C_N',
            'D_D', 'D_A', 'D_N',
            'Total_D', 'Total_A', 'Total_N'
        ]
        df = pd.DataFrame(processed_rows, columns=columns)

        # Write to sheet
        sheet_name = date_str.replace("/", "-") # Excel doesn't like slashes in sheet names
        df.to_excel(writer, sheet_name=sheet_name, startrow=3, header=False, index=False)
        worksheet = writer.sheets[sheet_name]

        # Write Headers
        worksheet.merge_range('A1:R1', f"Date :-{date_str}", header_format)
        
        # Row 2 Headers
        worksheet.merge_range('A2:A3', "SR.\nNO", header_format)
        worksheet.merge_range('B2:B3', "SSE'S", header_format)
        worksheet.merge_range('C2:C3', "Section SI", header_format)
        
        worksheet.merge_range('D2:F2', "A (replacement of Gear)", header_format)
        worksheet.merge_range('G2:I2', "B (For Engg. Work )", header_format)
        worksheet.merge_range('J2:L2', "C ( For Maintenance)", header_format)
        worksheet.merge_range('M2:O2', "D (For Failure)", header_format)
        worksheet.merge_range('P2:R2', "Total", header_format)

        # Row 3 Sub-headers
        sub_headers = ['D', 'A', 'N'] * 5
        for idx, val in enumerate(sub_headers):
            worksheet.write(2, 3 + idx, val, header_format)

        # Merge SSE Column
        sse_groups = df.groupby('SSE', sort=False)
        start_row = 3
        for name, group in sse_groups:
            group_len = len(group)
            if group_len > 1:
                worksheet.merge_range(start_row, 1, start_row + group_len - 1, 1, name, merge_format)
                # Also merge SR NO if needed, though usually SR NO is per group
                sr_val = group['SR NO'].iloc[0]
                worksheet.merge_range(start_row, 0, start_row + group_len - 1, 0, sr_val, merge_format)
            else:
                worksheet.write(start_row, 1, name, merge_format)
                worksheet.write(start_row, 0, group['SR NO'].iloc[0], merge_format)
            start_row += group_len

        # Write Data Cells (replace 0 with empty string if desired, but images show 0 in totals)
        # Images show blank for 0 in data columns, but 0 in Total columns.
        for r_idx, row_data in enumerate(processed_rows):
            for c_idx, val in enumerate(row_data):
                if c_idx > 2: # Data columns
                    # Logic: If it's a data column (A,B,C,D) and value is 0, write blank.
                    # If it's a Total column and value is 0, write 0.
                    is_total_col = c_idx >= 15
                    display_val = val
                    if val == 0 and not is_total_col:
                        display_val = ""
                    worksheet.write(r_idx + 3, c_idx, display_val, cell_format)
                elif c_idx == 2: # Section SI
                    worksheet.write(r_idx + 3, c_idx, val, cell_format)

        # Write Total Row
        last_row = 3 + len(processed_rows)
        worksheet.merge_range(last_row, 0, last_row, 2, "Total", total_row_format)
        for idx, val in enumerate(col_totals):
            worksheet.write(last_row, 3 + idx, val, total_row_format)

        # Set Column Widths
        worksheet.set_column('A:A', 5)
        worksheet.set_column('B:B', 10)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('D:R', 4)

    writer.close()
    print(f"Excel file '{filename}' created successfully.")

if __name__ == "__main__":
    create_excel(data_sheets)