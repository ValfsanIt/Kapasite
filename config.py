import os
import pandas as pd
from datetime import datetime

server = '172.30.134.15'
database = 'VALFSAN604'
database_iot = 'VALFAN_IOT_001'
username = 'pythonreporter'
password = '1212casecase,,'

# Mail ayarları
# MAIL_MODE:
# - "o365": smtp.office365.com:587 + STARTTLS + AUTH
# - diğer: Yerel Exchange relay (mail.valfsan.com.tr:25, auth yok)
MAIL_MODE = "relay"  # "o365" veya "relay"

if MAIL_MODE == "o365":
    mail_server = "smtp.office365.com"
    mail_port = 587
    SMTP_USE_STARTTLS = True
    SMTP_REQUIRE_AUTH = True
    SMTP_USERNAME = "canias@valfsan.com.tr"
    SMTP_PASSWORD = ""
else:
    # Yerel Exchange relay
    mail_server = "mail.valfsan.com.tr"
    mail_port = 25
    SMTP_USE_STARTTLS = False
    SMTP_REQUIRE_AUTH = False
    
    SMTP_USERNAME = "canias@valfsan.com.tr"
    SMTP_PASSWORD = ""


RAPORLAMA_NOTIFY_TO = ["dayyildiz@valfsan.com.tr","planlama@valfsan.com.tr"]
RAPORLAMA_NOTIFY_CC = ["itstajyer@valfsan.com.tr"]





reengen_username = "takgun@valfsan.com.tr"
reengen_password = '1234Ta1234'
reengen_company = 'valfsan'



project_directory = os.path.dirname(os.path.abspath(__file__))

# Zamanlanmış Kapasite ZIP (report_job): tek güncel dosya.
# Öncelik: ortam değişkeni KAP_REPORT_OUT_DIR (tanımlıysa bu yol yerine o kullanılır).
# Her çalışmada aynı dosya adı üzerine yazılır; kök klasördeki diğer kapasite_raporlar*.zip silinir.
# Pazartesi: ek olarak SCHEDULED_REPORT_MONDAY_ARCHIVE_SUBDIR altına tarihli kopya yazılır (birikir; silinmez).
SCHEDULED_REPORT_OUTPUT_DIR = r"\\Vlffile01\data\16- ORTAK\KAPASITE_RAPOR"
SCHEDULED_REPORT_ZIP_BASENAME = "kapasite_raporlar_latest.zip"
SCHEDULED_REPORT_MONDAY_ARCHIVE_SUBDIR = "Pazartesi_raporlari"

directory = os.path.join(project_directory, "queries")
dirofquery = os.path.join(directory, "query2.txt")
dirofnewline = os.path.join(directory, "dateofstart.txt")
url_dyn = "http://172.30.134.16:20000/services/btstarter.aspx?tran_code=WSCANIAS&tran_param=VLFPYPORTAL,"

today = datetime.today()

if today.weekday() == 6:
    kb = 2
elif today.weekday() == 0:
    kb = 3
else:
    kb = 1
    
sleep_time = 2


def set_df_size(maxrow, maxcolumn, maxwidth, maxcolwidth):
    pd.set_option('display.max_rows', maxrow)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 500)
    pd.set_option('display.max_colwidth', maxcolwidth)
valftoreeg = {"Bütün": {"Analizörler": "Analizörler", "Bölümler": "Bölümler"},
              "CNC": {"CNCTO-01','CNCTO-02','CNCTO-04','CNCTO-05','CNCTO-06','CNCTO-07','CNCTO-08','CNCTO-09','CNCTO-10','CNCTO-11','CNCTO-12','CNCTO-13','CNCTO-14','CNCTO-15','CNCTO-16','CNC-07','CNC-08','CNC-26','CNC-28": 'valfsan_cnc_pano1',
                      "CNC-04', 'CNC-11', 'CNC-13', 'CNC-14', 'CNC-15', 'CNC-16', 'CNC-17', 'CNC-18', 'CNC-19', 'CNC-20', 'CNC-21', 'CNC-22', 'CNC-23', 'CNC-29', 'Z-01": 'valfsan_cnc_pano2',
                      "CNC-01', 'CNC-03', 'CNC-06', 'CNC-12', 'CNC-24', 'TB-01', 'STASLAMA', 'HONLAMA', 'ZIMPARA": 'valfsan_cnc_pano3',
                      "CNC-07": "valfsan_cncfreze_makina_cnc_07",
                      "CNC-04": "valfsan_cncfreze_trident_cnc_04",
                      "CNC-20": "valfsan_cncfreze_dmc_cnc_20",
                      "CNC-12": 'valfsan_cnctorna_starcnc_kayarotomat_cnc_12',
                      "CNC-24": 'valfsan_cncfreze_johnford_cnc_24',
                      "CNC-03": 'valfsan_cncfreze_haas_cnc_03',
                      "CNCTO-07": 'valfsan_cnctorna_dmc_cnc_to_07',
                      "CNCTO-15": 'valfsan_cnctorna_dmgmori_cnc_to_15',
                      "CNCTO-09": 'valfsan_cnctorna_spinner_cnc_to_09'},
              "ISIL ISLEM ": {"F-01": 'valfsan_isilislem_cinfirini',
                              "F-02": "valfsan_isilislem_hintfirini",
                              "F-03": "valfsan_isilislem_bakirkaynakfirini",
                              "MF-01": "valfsan_isilislem_menevisfirini_mf_01",
                              "MF-04": "valfsan_isilislem_menevisfirini_mf_04"},
              "DOGRULTMA": {"DM-02": "valfsan_dogrultma_dm_02"},
              "TASLAMA": {"DD-02": "valfsan_taslama_dd_02",
                          "DD-03": "valfsan_taslama_dd_03",
                          "DD-04": "valfsan_taslama_dd_04",
                          "DD-05": "valfsan_taslama_dd_05",
                          "DD-06": "valfsan_taslama_dd_06",
                          "DD-07": "valfsan_taslama_dd_07",
                          "DD-10": "valfsan_taslama_discus_dd_10",
                          "CT-03": "valfsan_taslama_ct_03",
                          "STAG-01": "valfsan_taslama_stag_01",
                          "STAG-02": "valfsan_taslama_stag_02",
                          "STAG-03": "valfsan_taslama_stag_03"},
              "PRESHANE": {'PRES - Pano 1': 'PRESHANE / PANO 1',
                           'P-06': 'PRESHANE / PANO 2',
                           'PRES - Pano 3': 'PRESH2hANE / PANO 3',
                           "P-16": "valfsan_pres16",
                           "P-17": "valfsan_pres17",
                           "P-20": "valfsan_pres20",
                           "P-60": "valfsan_pres20",
                           "P-21": "valfsan_pres21",
                           "P-29": "valfsan_pres29",
                           "P-31": "valfsan_pres31",
                           "P-36": "valfsan_pres36",
                           "P-69": "valfsan_pres_cinpresi",
                           "P-72": "valfsan_pres72",
                           "P-04', 'P-08', 'P-24', 'P-36', 'P-46', 'P-50', 'P-69', 'P-76":"valfsan_salter2",
                           "P-11', 'P-47', 'P-62', 'P-74":"valfsan_salter3",
                           "P-12', 'P-19', 'P-23', 'P-30',P-31', 'P-33', 'P-34', 'P-37', 'P-55', 'P-56', 'P-61', 'P-70', 'P-71":"valfsan_salter5",
                           "P-14', 'P-26', 'P-64', 'P-67','P-68', 'P-75', 'P-63', 'P-65', 'P-73', 'TC-05":"10 Pres (Salter 4-6)"},
              "KALIPHANE ": {'KALIPHANE - Pano 1': 'valfsan_kaliphane_pano1',
                             "MF-02": "valfsan_kaliphane_mf-02"},
              "YÜZEY İŞLEM": {
                          'S-03,V-01,V-02': 'valfsan_yuzeyislem',
                          "T-19','T-20','T-21','T-22','T-23','T-24','T-25','T-26','T-27','T-34','T-37','T-43','T-44','T-45": 'valfsan_yuzeyislem_2',
                          'OTEC-03,OTEC-04': "valfsan_otecsantrifuj",
                          'S-01': 'YÜZEY İŞLEM - Küçük Santrifüj S - 01',
                          'V-01': 'YÜZEY İŞLEM - Büyük Vibratör V - 01'},
              "KURUTMA": {"K-25": "KURUTMA - No 25",
                          "K-13','K-18','K-19','K-20','K-21','K-22','K-23','K-25','K-27','K-28": 'valfsan_pano1'},
              "YIKAMA": {'YKM-03': 'valfsan_yikama_protech_pano1'},
              "TEL EREZYON": {'(ER-03)': 'valfsan_telerezyon_pano1_er_03',
                              '(ER-04)': 'valfsan_telerezyon_pano2_er_04'},
              "KOMPRESÖR": {'KOMPRESÖR-ISIL': 'KOMPRESÖR / Isıl İşlem Yanı (9,5 bar - 22,5 kW)',
                            'KOM-04': 'valfsan_kompresor_04_mark',
                            'KOM-05': 'valfsan_kompresor_05_mark',
                            'KOM-06': 'valfsan_kompresor_06_mark',
                            'KOM-07': 'valfsan_kompresor_07_mark',
                            'KOM-08': 'valfsan_kompresor_08_mark',
                            'KOMPRESÖR-(3 adet)': 'valfsan_kompresor_yuksekbasinc'},
              "CHILLER SOGUTUCU": {'Chiller 1': 'valfsan_chillersogutucu1',
                                   'Chiller 3': 'valfsan_chillersogutucu1',
                                   'Chiller 4': 'valfsan_chillersogutucu1'},
              "KLIMA UNITELERI": {'KLİMA VRV': 'valfsan_klimavrvuniteleri_pano1',
                                  'KLİMA T.ODA': 'valfsan_temizodaklimasantrali_pano1'},
              "HAVALANDIRMA - FAN": {'PRESHANE / EGZOST EMİŞ FANI (3 adet)': 'valfsan_preshaneemisfani_11kw'},
              "TRAFO": {'TRAFO - 1600 kVA': 'valfsan_trafo'}}

today = datetime.today()

if today.weekday() == 6:
    kb = 2
elif today.weekday() == 0:
    kb = 3
else:
    kb = 1
