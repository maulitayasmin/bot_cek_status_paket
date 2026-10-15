import pandas as pd
import gspread
import requests
import logging
from datetime import datetime
import os
import time
import re

# Setup logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Konfigurasi
BOT_TOKEN = "8209310392:AAH31YDaFIwdy90mtDsAJcmWiQdVh1vu7lo"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1Ymg7Sjx5uT5y1vW6B3pxeV6FiaecXeKwYsVBkBcvI3I/edit?gid=0#gid=0"
SHEET_NAME = "TUR"
CREDENTIALS_FILE = "credentials.json"

class OrderCheckerBot:
    def __init__(self):
        self.gc = None
        self.worksheet = None
        self.setup_google_sheets()
    
    def setup_google_sheets(self):
        """Setup koneksi ke Google Sheets"""
        try:
            if not os.path.exists(CREDENTIALS_FILE):
                logger.error(f"File credentials '{CREDENTIALS_FILE}' tidak ditemukan!")
                return
            
            self.gc = gspread.service_account(filename=CREDENTIALS_FILE)
            spreadsheet = self.gc.open_by_url(SPREADSHEET_URL)
            self.worksheet = spreadsheet.worksheet(SHEET_NAME)
            logger.info("✅ Berhasil terhubung ke Google Sheets")
        except Exception as e:
            logger.error(f"❌ Error connecting to Google Sheets: {e}")
    
    def get_all_data(self):
        """Mengambil semua data dari spreadsheet"""
        try:
            if not self.worksheet:
                logger.error("Worksheet tidak tersedia")
                return pd.DataFrame()
            
            data = self.worksheet.get_all_records()
            logger.info(f"📊 Berhasil mengambil {len(data)} records")
            return pd.DataFrame(data)
        except Exception as e:
            logger.error(f"❌ Error getting data: {e}")
            return pd.DataFrame()

    def get_monthly_report(self, month_year=None):
        """Mendapatkan laporan bulanan - FIXED untuk format MM-DD-YY"""
        try:
            df = self.get_all_data()
            if df.empty:
                logger.error("❌ DataFrame is empty")
                return None
            
            # Jika tidak ada parameter, gunakan bulan ini
            if month_year is None:
                current_date = datetime.now()
                month_year = current_date.strftime('%m-%Y')
            
            logger.info(f"🔍 Processing monthly report for: {month_year}")
            
            # Parse bulan dan tahun (format: MM-YYYY)
            try:
                report_month, report_year = month_year.split('-')
                report_month, report_year = int(report_month), int(report_year)
            except:
                logger.error(f"❌ Format month_year salah: {month_year}")
                return None
            
            # Cari kolom tanggal
            date_columns = ['ORDER DATE', 'ORDER_DATE', 'TANGGAL', 'DATE', 'TANGGAL ORDER']
            date_col = None
            
            for col in date_columns:
                if col in df.columns:
                    date_col = col
                    logger.info(f"📅 Using date column: {date_col}")
                    break
            
            if not date_col:
                logger.error("❌ Kolom tanggal tidak ditemukan")
                return None
            
            # KONVERSI TANGGAL DARI SPREADSHEET (MM-DD-YY)
            logger.info("🔄 Mengkonversi tanggal dari spreadsheet (format: MM-DD-YY)...")
            
            # Coba format utama: MM-DD-YY
            df['CONVERTED_DATE'] = pd.to_datetime(
                df[date_col],
                format='%m-%d-%y',
                errors='coerce'
            )
            
            success_count = df['CONVERTED_DATE'].notna().sum()
            logger.info(f"✅ Berhasil mengkonversi {success_count}/{len(df)} records dengan MM-DD-YY")
            
            if success_count < len(df):
                # Coba format alternatif
                alt_formats = ['%m/%d/%y', '%m-%d-%Y', '%m/%d/%Y']
                
                for fmt in alt_formats:
                    mask = df['CONVERTED_DATE'].isna()
                    if mask.any():
                        parsed = pd.to_datetime(df.loc[mask, date_col], format=fmt, errors='coerce')
                        df.loc[mask, 'CONVERTED_DATE'] = parsed
                        
                        new_success = parsed.notna().sum()
                        if new_success > 0:
                            logger.info(f"  ✅ Format {fmt}: {new_success} records tambahan")
            
            # Ekstrak bulan dan tahun
            df['MONTH'] = df['CONVERTED_DATE'].dt.month
            df['YEAR'] = df['CONVERTED_DATE'].dt.year
            
            # Filter berdasarkan bulan dan tahun yang diminta
            monthly_data = df[(df['MONTH'] == report_month) & (df['YEAR'] == report_year)]
            
            logger.info(f"📊 Monthly data found: {len(monthly_data)} records")
            
            if monthly_data.empty:
                return {
                    'total_orders': 0,
                    'status_counts': {},
                    'sto_counts': {},
                    'month_year': f"{report_month:02d}-{report_year}",
                    'month_name': datetime(report_year, report_month, 1).strftime('%B %Y')
                }
            
            # Hitung berdasarkan status
            status_columns = ['STATUS', 'STATUS ORDER', 'KETERANGAN']
            status_col = None
            
            for col in status_columns:
                if col in df.columns:
                    status_col = col
                    break
            
            status_counts = {}
            if status_col:
                status_counts = monthly_data[status_col].value_counts().to_dict()
            
            # Hitung berdasarkan STO
            sto_columns = ['STO', 'NAMA STO', 'STO_NAME']
            sto_col = None
            
            for col in sto_columns:
                if col in df.columns:
                    sto_col = col
                    break
            
            sto_counts = {}
            if sto_col:
                sto_counts = monthly_data[sto_col].value_counts().to_dict()
            
            return {
                'total_orders': len(monthly_data),
                'status_counts': status_counts,
                'sto_counts': sto_counts,
                'month_year': f"{report_month:02d}-{report_year}",
                'month_name': datetime(report_year, report_month, 1).strftime('%B %Y'),
                'status_column': status_col,
                'sto_column': sto_col
            }
            
        except Exception as e:
            logger.error(f"❌ Error in get_monthly_report: {e}")
            return None

    def get_sto_pivot_report(self, month_year=None):
        """Mendapatkan laporan pivot per STO - FIXED VERSION"""
        try:
            df = self.get_all_data()
            if df.empty:
                logger.error("❌ DataFrame is empty")
                return None
            
            logger.info(f"📊 Total data rows: {len(df)}")
            logger.info(f"📋 Columns: {list(df.columns)}")
            
            # Jika tidak ada parameter, gunakan bulan ini
            if month_year is None:
                current_date = datetime.now()
                month_year = current_date.strftime('%m-%Y')
            
            logger.info(f"🔍 Processing STO pivot report for: {month_year}")
            
            # Parse bulan dan tahun
            month, year = month_year.split('-')
            month, year = int(month), int(year)
            
            # Cari kolom tanggal
            date_columns = ['ORDER DATE', 'ORDER_DATE', 'TANGGAL', 'DATE', 'TANGGAL ORDER']
            date_col = None
            
            for col in date_columns:
                if col in df.columns:
                    date_col = col
                    break
            
            if not date_col:
                logger.error("❌ Kolom tanggal tidak ditemukan")
                return None
            
            logger.info(f"📅 Using date column: {date_col}")
            
            # Konversi kolom tanggal dari spreadsheet (MM-DD-YY)
            df['CONVERTED_DATE'] = pd.to_datetime(
                df[date_col],
                format='%m-%d-%y',
                errors='coerce'
            )
            
            success_count = df['CONVERTED_DATE'].notna().sum()
            logger.info(f"✅ Berhasil mengkonversi {success_count}/{len(df)} records")
            
            if success_count < len(df):
                # Coba format alternatif
                alt_formats = ['%m/%d/%y', '%m-%d-%Y', '%m/%d/%Y']
                
                for fmt in alt_formats:
                    mask = df['CONVERTED_DATE'].isna()
                    if mask.any():
                        parsed = pd.to_datetime(df.loc[mask, date_col], format=fmt, errors='coerce')
                        df.loc[mask, 'CONVERTED_DATE'] = parsed
            
            # Filter berdasarkan bulan dan tahun
            df['MONTH'] = df['CONVERTED_DATE'].dt.month
            df['YEAR'] = df['CONVERTED_DATE'].dt.year
            
            monthly_data = df[(df['MONTH'] == month) & (df['YEAR'] == year)]
            
            logger.info(f"📊 Monthly data found: {len(monthly_data)} records")
            
            if monthly_data.empty:
                logger.warning(f"⚠️ No data found for {month_year}")
                return {
                    'total_orders': 0,
                    'sto_pivot': {},
                    'month_year': f"{month:02d}-{year}",
                    'month_name': datetime(year, month, 1).strftime('%B %Y')
                }
            
            # Cari kolom STO
            sto_columns = ['STO', 'NAMA STO', 'STO_NAME']
            sto_col = None
            
            for col in sto_columns:
                if col in monthly_data.columns:
                    sto_col = col
                    break
            
            if not sto_col:
                logger.error("❌ Kolom STO tidak ditemukan")
                return None
            
            logger.info(f"🏢 Using STO column: {sto_col}")
            
            # Cari kolom status
            status_columns = ['STATUS', 'STATUS ORDER', 'KETERANGAN']
            status_col = None
            
            for col in status_columns:
                if col in monthly_data.columns:
                    status_col = col
                    break
            
            if not status_col:
                logger.error("❌ Kolom status tidak ditemukan")
                return None
            
            logger.info(f"📊 Using status column: {status_col}")
            
            # Buat pivot per STO
            sto_pivot = {}
            unique_stos = monthly_data[sto_col].dropna().unique()
            
            logger.info(f"🏢 Processing {len(unique_stos)} unique STOs")
            
            for sto_name in unique_stos:
                if pd.isna(sto_name) or str(sto_name).strip() == '':
                    continue
                    
                sto_name_str = str(sto_name).strip()
                sto_data = monthly_data[monthly_data[sto_col] == sto_name]
                total_orders = len(sto_data)
                
                if total_orders == 0:
                    continue
                
                logger.info(f"   Processing {sto_name_str}: {total_orders} orders")
                
                # Hitung PS (Proses/Progress)
                ps_count = len(sto_data[sto_data[status_col].astype(str).str.lower().str.contains('ps|proses|progress', na=False)])
                
                # RE = Total semua status pada bulan yang dipilih
                re_count = total_orders
                
                # Hitung RE Rate
                re_rate = (ps_count / re_count * 100) if re_count > 0 else 0
                
                # Status breakdown
                status_breakdown = sto_data[status_col].value_counts().to_dict()
                
                # Ambil daftar SC
                sc_columns = ['SC', 'ORDER ID', 'ORDER_ID', 'NO ORDER']
                sc_col = None
                
                for col in sc_columns:
                    if col in sto_data.columns:
                        sc_col = col
                        break
                
                sc_list = []
                if sc_col:
                    sc_list = sto_data[sc_col].dropna().astype(str).tolist()
                
                sto_pivot[sto_name_str] = {
                    'total_orders': total_orders,
                    'ps_count': ps_count,
                    're_count': re_count,
                    're_rate': re_rate,
                    'status_breakdown': status_breakdown,
                    'sc_list': sc_list
                }
            
            logger.info(f"✅ STO pivot created for {len(sto_pivot)} STOs")
            
            return {
                'total_orders': len(monthly_data),
                'sto_pivot': sto_pivot,
                'month_year': f"{month:02d}-{year}",
                'month_name': datetime(year, month, 1).strftime('%B %Y'),
                'sto_column': sto_col,
                'status_column': status_col
            }
            
        except Exception as e:
            logger.error(f"❌ Error in get_sto_pivot_report: {e}")
            import traceback
            traceback.print_exc()
            return None

    def get_sto_pivot_details(self, month_year, sto_name):
        """Mendapatkan detail SC untuk STO tertentu dalam pivot - FIXED VERSION"""
        try:
            logger.info(f"🔍 Getting STO pivot details for: {month_year}, {sto_name}")
            
            # 1. Ambil semua data terlebih dahulu
            df = self.get_all_data()
            if df.empty:
                logger.error("❌ DataFrame is empty")
                return None
            
            # 2. Parse bulan dan tahun
            month, year = month_year.split('-')
            month, year = int(month), int(year)
            
            # 3. Cari kolom tanggal
            date_columns = ['ORDER DATE', 'ORDER_DATE', 'TANGGAL', 'DATE', 'TANGGAL ORDER']
            date_col = None
            
            for col in date_columns:
                if col in df.columns:
                    date_col = col
                    break
            
            if not date_col:
                logger.error("❌ Kolom tanggal tidak ditemukan")
                return None
            
            # 4. Konversi kolom tanggal dari spreadsheet (MM-DD-YY)
            df['CONVERTED_DATE'] = pd.to_datetime(
                df[date_col],
                format='%m-%d-%y',
                errors='coerce'
            )
            
            # Coba format alternatif jika gagal
            if df['CONVERTED_DATE'].isna().any():
                alt_formats = ['%m/%d/%y', '%m-%d-%Y', '%m/%d/%Y']
                for fmt in alt_formats:
                    mask = df['CONVERTED_DATE'].isna()
                    if mask.any():
                        parsed = pd.to_datetime(df.loc[mask, date_col], format=fmt, errors='coerce')
                        df.loc[mask, 'CONVERTED_DATE'] = parsed
            
            # 5. Filter berdasarkan bulan dan tahun
            df['MONTH'] = df['CONVERTED_DATE'].dt.month
            df['YEAR'] = df['CONVERTED_DATE'].dt.year
            
            monthly_data = df[(df['MONTH'] == month) & (df['YEAR'] == year)]
            
            logger.info(f"📊 Monthly data found: {len(monthly_data)} records for {month_year}")
            
            if monthly_data.empty:
                logger.warning(f"⚠️ No data found for {month_year}")
                return None
            
            # 6. Cari kolom STO
            sto_columns = ['STO', 'NAMA STO', 'STO_NAME']
            sto_col = None
            
            for col in sto_columns:
                if col in monthly_data.columns:
                    sto_col = col
                    break
            
            if not sto_col:
                logger.error("❌ Kolom STO tidak ditemukan")
                return None
            
            # 7. Filter untuk STO tertentu
            sto_data = monthly_data[monthly_data[sto_col].astype(str).str.contains(str(sto_name), case=False, na=False)]
            
            if sto_data.empty:
                logger.warning(f"⚠️ No data found for STO: {sto_name}")
                return None
            
            logger.info(f"✅ Found {len(sto_data)} records for STO: {sto_name}")
            
            # 8. Cari kolom status
            status_columns = ['STATUS', 'STATUS ORDER', 'KETERANGAN']
            status_col = None
            
            for col in status_columns:
                if col in sto_data.columns:
                    status_col = col
                    break
            
            # 9. Hitung statistik
            total_orders = len(sto_data)
            
            # Hitung PS (Proses/Progress)
            if status_col:
                ps_count = len(sto_data[sto_data[status_col].astype(str).str.lower().str.contains('ps|proses|progress', na=False)])
                status_breakdown = sto_data[status_col].value_counts().to_dict()
            else:
                ps_count = 0
                status_breakdown = {}
            
            # RE = Total semua status pada bulan yang dipilih
            re_count = total_orders
            
            # Hitung RE Rate
            re_rate = (ps_count / re_count * 100) if re_count > 0 else 0
            
            # 10. Ambil daftar SC
            sc_columns = ['SC', 'ORDER ID', 'ORDER_ID', 'NO ORDER']
            sc_col = None
            
            for col in sc_columns:
                if col in sto_data.columns:
                    sc_col = col
                    break
            
            sc_list = []
            if sc_col:
                sc_list = sto_data[sc_col].dropna().astype(str).tolist()
            
            return {
                'sto_name': sto_name,
                'total_orders': total_orders,
                'ps_count': ps_count,
                're_count': re_count,
                're_rate': re_rate,
                'status_breakdown': status_breakdown,
                'sc_list': sc_list,
                'month_name': datetime(year, month, 1).strftime('%B %Y'),
                'month_year': month_year
            }
            
        except Exception as e:
            logger.error(f"❌ Error in get_sto_pivot_details: {e}")
            import traceback
            traceback.print_exc()
            return None

    def search_by_order_id(self, order_id):
        """Mencari data berdasarkan Order ID"""
        try:
            df = self.get_all_data()
            if df.empty:
                return None
            
            order_id = str(order_id).strip().upper()
            
            if 'SC' in df.columns:
                result = df[df['SC'].astype(str).str.contains(order_id, case=False, na=False)]
            else:
                result = df[df.iloc[:, 1].astype(str).str.contains(order_id, case=False, na=False)]
            
            return result.to_dict('records') if not result.empty else None
        except Exception as e:
            logger.error(f"❌ Error in search_by_order_id: {e}")
            return None

    def search_by_date(self, date_str):
        """Mencari data berdasarkan tanggal - FIXED untuk format MM-DD-YY di spreadsheet"""
        try:
            df = self.get_all_data()
            if df.empty:
                logger.error("❌ DataFrame is empty")
                return None
            
            # Cari kolom yang berisi tanggal
            date_columns = ['ORDER DATE', 'ORDER_DATE', 'TANGGAL', 'DATE', 'TANGGAL ORDER']
            date_col = None
            
            for col in date_columns:
                if col in df.columns:
                    date_col = col
                    logger.info(f"📅 Menggunakan kolom tanggal: {date_col}")
                    break
            
            if not date_col:
                logger.error("❌ Kolom tanggal tidak ditemukan")
                return None
            
            # 1. PARSE INPUT USER (DD-MM-YYYY)
            try:
                # Input user: DD-MM-YYYY (contoh: 15-12-2024)
                day, month, year = map(int, date_str.split('-'))
                
                # Validasi
                if not (1 <= month <= 12):
                    logger.error(f"❌ Bulan tidak valid: {month}")
                    return None
                if not (1 <= day <= 31):
                    logger.error(f"❌ Hari tidak valid: {day}")
                    return None
                
                # Buat 4 digit tahun dari input user
                if year < 100:  # Jika tahun 2 digit
                    year = 2000 + year if year < 50 else 1900 + year
                
                # Target date dari input user
                target_date_user = datetime(year, month, day).date()
                logger.info(f"🔍 Input user: {date_str} -> {target_date_user}")
                
            except Exception as e:
                logger.error(f"❌ Error parsing tanggal input: {date_str}, Error: {e}")
                return None
            
            # 2. PARSE DATA SPREADSHEET (MM-DD-YY)
            logger.info("🔄 Memproses data spreadsheet...")
            
            # Konversi kolom tanggal dari spreadsheet
            # Data spreadsheet: MM-DD-YY (contoh: 12-15-24 untuk 15 Desember 2024)
            df['CONVERTED_DATE'] = pd.to_datetime(
                df[date_col],
                format='%m-%d-%y',  # Format spreadsheet: MM-DD-YY
                errors='coerce'
            ).dt.date
            
            # Cek jika parsing berhasil
            success_count = df['CONVERTED_DATE'].notna().sum()
            logger.info(f"✅ Berhasil mengkonversi {success_count}/{len(df)} records")
            
            if success_count == 0:
                # Coba format alternatif
                logger.info("🔄 Mencoba format alternatif...")
                alt_formats = [
                    '%m/%d/%y',    # 12/15/24
                    '%m-%d-%Y',    # 12-15-2024 (tahun 4 digit)
                    '%m/%d/%Y',    # 12/15/2024
                    '%d-%m-%y',    # 15-12-24 (jika ternyata DD-MM-YY)
                    '%d/%m/%y',    # 15/12/24
                ]
                
                for fmt in alt_formats:
                    df['CONVERTED_DATE'] = pd.to_datetime(
                        df[date_col],
                        format=fmt,
                        errors='coerce'
                    ).dt.date
                    
                    success = df['CONVERTED_DATE'].notna().sum()
                    if success > 0:
                        logger.info(f"✅ Format {fmt}: {success} records berhasil")
                        break
            
            # 3. FILTER BERDASARKAN TANGGAL
            result = df[df['CONVERTED_DATE'] == target_date_user]
            
            if not result.empty:
                logger.info(f"📊 Ditemukan {len(result)} records untuk tanggal {date_str}")
                return result.to_dict('records')
            
            # 4. FALLBACK: String matching
            logger.info("🔄 Mencoba fallback string matching...")
            
            # Karena data spreadsheet MM-DD-YY, kita perlu konversi input user ke format spreadsheet
            # Input user: DD-MM-YYYY -> konversi ke format spreadsheet: MM-DD-YY
            # Contoh: user input "15-12-2024" -> spreadsheet format "12-15-24"
            
            # Buat format spreadsheet dari input user
            spreadsheet_format_2digit = f"{month:02d}-{day:02d}-{str(year)[-2:]}"
            spreadsheet_format_4digit = f"{month:02d}-{day:02d}-{year}"
            
            search_patterns = [
                spreadsheet_format_2digit,  # 12-15-24
                spreadsheet_format_4digit,  # 12-15-2024
                f"{month:02d}/{day:02d}/{str(year)[-2:]}",  # 12/15/24
                f"{month}/{day}/{str(year)[-2:]}",  # 12/15/24 (tanpa leading zero)
            ]
            
            logger.info(f"🔍 Mencari pattern: {search_patterns}")
            
            for pattern in search_patterns:
                temp_result = df[df[date_col].astype(str).str.contains(
                    pattern,
                    case=False,
                    na=False,
                    regex=False
                )]
                
                if not temp_result.empty:
                    logger.info(f"✅ String matching ditemukan {len(temp_result)} dengan pattern: {pattern}")
                    return temp_result.to_dict('records')
            
            logger.info(f"📭 Tidak ditemukan data untuk tanggal {date_str}")
            return None
                
        except Exception as e:
            logger.error(f"❌ Error in search_by_date: {e}")
            import traceback
            traceback.print_exc()
            return None

    def search_by_sto(self, sto_name):
        """Mencari data berdasarkan STO"""
        try:
            df = self.get_all_data()
            if df.empty:
                return None
            
            sto_name = str(sto_name).strip().upper()
            
            if 'STO' in df.columns:
                result = df[df['STO'].astype(str).str.contains(sto_name, case=False, na=False)]
            else:
                # Cari di kolom lain yang mungkin berisi STO
                for col in df.columns:
                    if 'STO' in col.upper():
                        result = df[df[col].astype(str).str.contains(sto_name, case=False, na=False)]
                        break
                else:
                    result = pd.DataFrame()
            
            logger.info(f"🏢 Ditemukan {len(result)} records untuk STO: {sto_name}")
            return result.to_dict('records') if not result.empty else None
            
        except Exception as e:
            logger.error(f"❌ Error in search_by_sto: {e}")
            return None

    def get_sc_details(self, sc_number):
        """Mendapatkan detail lengkap untuk SC tertentu"""
        try:
            df = self.get_all_data()
            if df.empty:
                return None
            
            sc_number = str(sc_number).strip()
            
            # Cari kolom SC
            sc_columns = ['SC', 'ORDER ID', 'ORDER_ID', 'NO ORDER']
            sc_col = None
            
            for col in sc_columns:
                if col in df.columns:
                    sc_col = col
                    break
            
            if not sc_col:
                logger.error("❌ Kolom SC tidak ditemukan")
                return None
            
            # Filter data untuk SC yang dipilih
            result = df[df[sc_col].astype(str) == sc_number]
            
            if result.empty:
                logger.info(f"❌ Tidak ditemukan data untuk SC: {sc_number}")
                return None
            
            logger.info(f"✅ Ditemukan {len(result)} records untuk SC: {sc_number}")
            return result.to_dict('records')
            
        except Exception as e:
            logger.error(f"❌ Error in get_sc_details: {e}")
            return None
    
    def convert_date_formats(self, date_str, from_format, to_format):
        """Konversi format tanggal"""
        try:
            date_obj = datetime.strptime(date_str, from_format)
            return date_obj.strftime(to_format)
        except:
            return None
    
    def get_date_sample(self, count=5):
        """Ambil sample tanggal dari spreadsheet untuk debugging"""
        df = self.get_all_data()
        if df.empty:
            return []
        
        date_columns = ['ORDER DATE', 'ORDER_DATE', 'TANGGAL', 'DATE', 'TANGGAL ORDER']
        date_col = None
        
        for col in date_columns:
            if col in df.columns:
                date_col = col
                break
        
        if not date_col:
            return []
        
        samples = []
        for i in range(min(count, len(df))):
            raw_date = df.iloc[i][date_col]
            
            # Coba parse dengan berbagai format
            parsed_formats = []
            
            # Format MM-DD-YY
            try:
                dt = datetime.strptime(str(raw_date), '%m-%d-%y')
                parsed_formats.append(f"MM-DD-YY: {dt.date()}")
            except:
                pass
            
            # Format MM/DD/YY
            try:
                dt = datetime.strptime(str(raw_date), '%m/%d/%y')
                parsed_formats.append(f"MM/DD/YY: {dt.date()}")
            except:
                pass
            
            # Format DD-MM-YY
            try:
                dt = datetime.strptime(str(raw_date), '%d-%m-%y')
                parsed_formats.append(f"DD-MM-YY: {dt.date()}")
            except:
                pass
            
            samples.append({
                'raw': raw_date,
                'parsed': parsed_formats
            })
        
        return samples

class TelegramBot:
    def __init__(self):
        self.bot_token = BOT_TOKEN
        self.order_checker = OrderCheckerBot()
        self.base_url = f"https://api.telegram.org/bot{self.bot_token}/"
        # Dictionary untuk menyimpan state pencarian user
        self.user_search_state = {}
        logger.info("TelegramBot initialized")

    def get_dashboard_data(self, period="monthly"):
        """Mendapatkan data dashboard"""
        try:
            df = self.order_checker.get_all_data()
            if df.empty:
                logger.error("❌ DataFrame is empty for dashboard")
                return None
            
            # Cari kolom yang diperlukan
            date_columns = ['ORDER DATE', 'ORDER_DATE', 'TANGGAL', 'DATE', 'TANGGAL ORDER']
            date_col = None
            
            for col in date_columns:
                if col in df.columns:
                    date_col = col
                    break
            
            if not date_col:
                logger.error("❌ Kolom tanggal tidak ditemukan untuk dashboard")
                return None
            
            # Konversi tanggal
            df['CONVERTED_DATE'] = pd.to_datetime(
                df[date_col],
                format='%m-%d-%y',
                errors='coerce'
            )
            
            # Coba format alternatif jika gagal
            if df['CONVERTED_DATE'].isna().any():
                alt_formats = ['%m/%d/%y', '%m-%d-%Y', '%m/%d/%Y']
                for fmt in alt_formats:
                    mask = df['CONVERTED_DATE'].isna()
                    if mask.any():
                        parsed = pd.to_datetime(df.loc[mask, date_col], format=fmt, errors='coerce')
                        df.loc[mask, 'CONVERTED_DATE'] = parsed
            
            # Filter untuk bulan ini
            current_date = datetime.now()
            current_month = current_date.month
            current_year = current_date.year
            
            # Filter data bulan ini
            df['MONTH'] = df['CONVERTED_DATE'].dt.month
            df['YEAR'] = df['CONVERTED_DATE'].dt.year
            
            monthly_data = df[(df['MONTH'] == current_month) & (df['YEAR'] == current_year)]
            
            # Cari kolom status
            status_columns = ['STATUS', 'STATUS ORDER', 'KETERANGAN']
            status_col = None
            
            for col in status_columns:
                if col in df.columns:
                    status_col = col
                    break
            
            # Cari kolom STO
            sto_columns = ['STO', 'NAMA STO', 'STO_NAME']
            sto_col = None
            
            for col in sto_columns:
                if col in df.columns:
                    sto_col = col
                    break
            
            # Cari kolom teknisi
            tech_columns = ['TEKNISI', 'NAMA TEKNISI', 'TEKNISI NAME']
            tech_col = None
            
            for col in tech_columns:
                if col in df.columns:
                    tech_col = col
                    break
            
            # Hitung metrik dasar
            total_orders = len(df)
            monthly_orders = len(monthly_data)
            
            # Hitung status untuk bulan ini
            status_counts = {}
            if status_col:
                status_counts = monthly_data[status_col].value_counts().to_dict()
            
            # Hitung PS Rate bulan ini
            ps_count = 0
            if status_col:
                ps_count = len(monthly_data[monthly_data[status_col].astype(str).str.lower().str.contains('ps|proses|progress', na=False)])
            
            ps_rate = (ps_count / monthly_orders * 100) if monthly_orders > 0 else 0
            
            # Hitung top STO bulan ini
            sto_counts = {}
            if sto_col:
                sto_counts = monthly_data[sto_col].value_counts().to_dict()
            
            # Hitung top teknisi bulan ini
            tech_counts = {}
            if tech_col:
                tech_counts = monthly_data[tech_col].value_counts().to_dict()
            
            # Hitung trend (perbandingan dengan bulan lalu)
            last_month = (current_date.replace(day=1) - pd.DateOffset(months=1))
            last_month_data = df[(df['MONTH'] == last_month.month) & (df['YEAR'] == last_month.year)]
            last_month_orders = len(last_month_data)
            
            # Hitung perbedaan
            order_change = 0
            if last_month_orders > 0:
                order_change = ((monthly_orders - last_month_orders) / last_month_orders * 100)
            
            # Analisis kendala dari keterangan
            issue_analysis = {}
            if 'KETERANGAN' in monthly_data.columns:
                keterangan_data = monthly_data['KETERANGAN'].dropna()
                
                # Keyword patterns untuk kendala umum
                issue_patterns = {
                    'ODP FULL': ['odp full', 'full', 'kapasitas penuh'],
                    'TAT & TAJ': ['tat', 'taj', 'jarak', 'jauh'],
                    'DOUBLE INPUT': ['double input', 'dobel input', 'duplikat'],
                    'CANCEL PELANGGAN': ['cancel', 'batal', 'pelanggan batal'],
                    'TEKNICAL ISSUE': ['teknik', 'kendala teknik', 'masalah teknis'],
                    'SALAH TAGGING': ['salah tagging', 'wrong tag', 'tag salah']
                }
                
                for issue, keywords in issue_patterns.items():
                    count = 0
                    for keyword in keywords:
                        count += keterangan_data.astype(str).str.lower().str.contains(keyword, na=False).sum()
                    if count > 0:
                        issue_analysis[issue] = count
            
            return {
                'total_orders': total_orders,
                'monthly_orders': monthly_orders,
                'last_month_orders': last_month_orders,
                'order_change': order_change,
                'status_counts': status_counts,
                'ps_count': ps_count,
                'ps_rate': ps_rate,
                'sto_counts': sto_counts,
                'tech_counts': tech_counts,
                'issue_analysis': issue_analysis,
                'current_month': current_date.strftime('%B %Y'),
                'last_month': last_month.strftime('%B %Y'),
                'status_column': status_col,
                'sto_column': sto_col,
                'tech_column': tech_col
            }
            
        except Exception as e:
            logger.error(f"❌ Error in get_dashboard_data: {e}")
            import traceback
            traceback.print_exc()
            return None

    def format_dashboard_response(self, dashboard_data):
        """Memformat respons dashboard menjadi pesan yang menarik"""
        if not dashboard_data:
            return "❌ *Gagal mengambil data dashboard*"
        
        response = f"""
    📊 *DASHBOARD REAL-TIME*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    📅 *Periode:* {dashboard_data['current_month']}
    📈 *Total Order Sejauh Ini:* {dashboard_data['total_orders']:,}
    📋 *Order Bulan Ini:* {dashboard_data['monthly_orders']:,}

    """
        
        # Status ringkasan
        if dashboard_data['status_counts']:
            response += "┌─── *STATUS BULAN INI* ───\n"
            total_shown = 0
            for status, count in sorted(dashboard_data['status_counts'].items(), 
                                    key=lambda x: x[1], reverse=True)[:5]:
                percentage = (count / dashboard_data['monthly_orders'] * 100) if dashboard_data['monthly_orders'] > 0 else 0
                icon = "✅" if 'PS' in str(status).upper() else "❌" if 'CANCEL' in str(status).upper() else "⚠️"
                response += f"│ {icon} {status}: {count} ({percentage:.1f}%)\n"
                total_shown += count
            
            if len(dashboard_data['status_counts']) > 5:
                response += f"│ 📋 +{len(dashboard_data['status_counts']) - 5} status lainnya\n"
            
            response += f"└─────────────────────\n\n"
        
        # PS Rate
        response += f"🎯 *PS RATE:* {dashboard_data['ps_rate']:.1f}%\n"
        response += f"📊 *PS Count:* {dashboard_data['ps_count']:,}\n\n"
        
        # Trend vs bulan lalu
        change_icon = "📈" if dashboard_data['order_change'] > 0 else "📉" if dashboard_data['order_change'] < 0 else "➡️"
        response += f"🔄 *TREND VS {dashboard_data['last_month'].upper()}*\n"
        response += f"   {change_icon} Order: {dashboard_data['monthly_orders']:,} "
        if dashboard_data['order_change'] != 0:
            response += f"({dashboard_data['order_change']:+.1f}%)\n"
        else:
            response += "(tidak ada perubahan)\n"
        
        # Top STO
        if dashboard_data['sto_counts']:
            response += f"\n🏢 *TOP 5 STO BULAN INI*\n"
            top_stos = sorted(dashboard_data['sto_counts'].items(), 
                            key=lambda x: x[1], reverse=True)[:5]
            
            for i, (sto_name, count) in enumerate(top_stos, 1):
                percentage = (count / dashboard_data['monthly_orders'] * 100) if dashboard_data['monthly_orders'] > 0 else 0
                medal = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣"][i-1]
                response += f"   {medal} {sto_name}: {count} ({percentage:.1f}%)\n"
        
        # Top Teknisi
        if dashboard_data['tech_counts']:
            response += f"\n👷 *TOP 5 TEKNISI BULAN INI*\n"
            top_techs = sorted(dashboard_data['tech_counts'].items(), 
                            key=lambda x: x[1], reverse=True)[:5]
            
            for i, (tech_name, count) in enumerate(top_techs, 1):
                percentage = (count / dashboard_data['monthly_orders'] * 100) if dashboard_data['monthly_orders'] > 0 else 0
                medal = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣"][i-1]
                response += f"   {medal} {tech_name}: {count} ({percentage:.1f}%)\n"
        
        # Issue Analysis
        if dashboard_data['issue_analysis']:
            response += f"\n⚠️ *ANALISIS KENDALA TERBANYAK*\n"
            sorted_issues = sorted(dashboard_data['issue_analysis'].items(), 
                                key=lambda x: x[1], reverse=True)[:5]
            
            for issue, count in sorted_issues:
                percentage = (count / dashboard_data['monthly_orders'] * 100) if dashboard_data['monthly_orders'] > 0 else 0
                icon = "🔴" if 'FULL' in issue else "🟡" if 'TAT' in issue else "🔵"
                response += f"   {icon} {issue}: {count} ({percentage:.1f}%)\n"
        
        # Insight dan rekomendasi
        response += f"\n💡 *INSIGHT & REKOMENDASI*\n"
        
        if dashboard_data['ps_rate'] < 70:
            response += "   • ⚠️ PS Rate di bawah target 70%, perlu evaluasi\n"
        elif dashboard_data['ps_rate'] > 90:
            response += "   • ✅ PS Rate excellent! Pertahankan performa\n"
        
        if dashboard_data['issue_analysis']:
            top_issue = max(dashboard_data['issue_analysis'].items(), key=lambda x: x[1]) if dashboard_data['issue_analysis'] else None
            if top_issue:
                response += f"   • 🔧 Fokus perbaiki: {top_issue[0]} ({top_issue[1]} kasus)\n"
        
        if dashboard_data['monthly_orders'] < dashboard_data['last_month_orders']:
            response += "   • 📉 Penurunan order bulan ini, cek strategi sales\n"
        elif dashboard_data['monthly_orders'] > dashboard_data['last_month_orders'] * 1.2:
            response += "   • 📈 Peningkatan signifikan! Good job!\n"
        
        return response

    def create_dashboard_menu(self):
        """Membuat menu dashboard"""
        keyboard = {
            "inline_keyboard": [
                [
                    {"text": "📊 Dashboard Real-Time", "callback_data": "dashboard_main"},
                    {"text": "📈 Trend Analysis", "callback_data": "dashboard_trend"}
                ],
                [
                    {"text": "🏢 STO Performance", "callback_data": "dashboard_sto"},
                    {"text": "👷 Teknisi Ranking", "callback_data": "dashboard_tech"}
                ],
                [
                    {"text": "⚠️ Issue Analysis", "callback_data": "dashboard_issues"},
                    {"text": "🔍 Quick Actions", "callback_data": "dashboard_actions"}
                ],
                [
                    {"text": "🔄 Refresh", "callback_data": "dashboard_refresh"},
                    {"text": "🔙 Menu Utama", "callback_data": "main_menu"}
                ]
            ]
        }
        return keyboard

    def show_dashboard_main(self, chat_id, message_id=None):
        """Menampilkan dashboard utama"""
        loading_text = """
    🔄 *MEMPERSIAPKAN DASHBOARD...*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    📊 Mengumpulkan data real-time...
    📈 Menganalisis trend...
    ⚠️ Memindai kendala...
    ⏳ Mohon tunggu sebentar...
    """
        
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': loading_text,
                'parse_mode': 'Markdown'
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, loading_text, 'Markdown')
        
        # Ambil data dashboard
        dashboard_data = self.get_dashboard_data()
        
        if dashboard_data:
            response = self.format_dashboard_response(dashboard_data)
            reply_markup = self.create_dashboard_menu()
            
            if message_id:
                url = self.base_url + "editMessageText"
                payload = {
                    'chat_id': chat_id,
                    'message_id': message_id,
                    'text': response,
                    'parse_mode': 'Markdown',
                    'reply_markup': reply_markup
                }
                requests.post(url, json=payload)
            else:
                self.send_message(chat_id, response, 'Markdown', reply_markup)
        else:
            error_text = """
    ❌ *GAGAL MEMUAT DASHBOARD*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    Tidak dapat mengambil data dashboard saat ini.

    💡 *Kemungkinan penyebab:*
    • Koneksi ke spreadsheet terputus
    • Data tidak tersedia
    • Format spreadsheet berubah

    🔄 Silakan coba beberapa saat lagi.
    """
            if message_id:
                url = self.base_url + "editMessageText"
                payload = {
                    'chat_id': chat_id,
                    'message_id': message_id,
                    'text': error_text,
                    'parse_mode': 'Markdown',
                    'reply_markup': self.create_back_menu()
                }
                requests.post(url, json=payload)
            else:
                self.send_message(chat_id, error_text, 'Markdown', self.create_back_menu())

    def show_dashboard_trend(self, chat_id, message_id=None):
        """Menampilkan analisis trend"""
        try:
            df = self.order_checker.get_all_data()
            if df.empty:
                error_text = "❌ Tidak ada data untuk analisis trend"
                self.send_message(chat_id, error_text, 'Markdown', self.create_dashboard_menu())
                return
            
            # Cari kolom tanggal
            date_columns = ['ORDER DATE', 'ORDER_DATE', 'TANGGAL', 'DATE', 'TANGGAL ORDER']
            date_col = None
            
            for col in date_columns:
                if col in df.columns:
                    date_col = col
                    break
            
            if not date_col:
                error_text = "❌ Kolom tanggal tidak ditemukan"
                self.send_message(chat_id, error_text, 'Markdown', self.create_dashboard_menu())
                return
            
            # Konversi tanggal
            df['CONVERTED_DATE'] = pd.to_datetime(
                df[date_col],
                format='%m-%d-%y',
                errors='coerce'
            ).dt.date
            
            # Analisis 3 bulan terakhir
            current_date = datetime.now()
            months_data = []
            
            for i in range(3):
                month_date = current_date - pd.DateOffset(months=i)
                month_start = month_date.replace(day=1)
                month_end = (month_start + pd.DateOffset(months=1) - pd.DateOffset(days=1))
                
                month_data = df[
                    (df['CONVERTED_DATE'] >= month_start.date()) & 
                    (df['CONVERTED_DATE'] <= month_end.date())
                ]
                
                months_data.append({
                    'month': month_start.strftime('%B'),
                    'orders': len(month_data),
                    'data': month_data
                })
            
            # Format respons
            response = f"""
    📈 *TREND ANALYSIS - 3 BULAN TERAKHIR*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    """
            
            for i, month_info in enumerate(months_data):
                trend_icon = ""
                if i > 0:
                    prev_orders = months_data[i-1]['orders']
                    if month_info['orders'] > prev_orders:
                        trend_icon = "📈"
                    elif month_info['orders'] < prev_orders:
                        trend_icon = "📉"
                    else:
                        trend_icon = "➡️"
                
                response += f"📅 *{month_info['month']}:*\n"
                response += f"   📊 Total Order: {month_info['orders']:,} {trend_icon}\n"
                
                # Hitung PS rate jika ada kolom status
                status_columns = ['STATUS', 'STATUS ORDER', 'KETERANGAN']
                status_col = None
                
                for col in status_columns:
                    if col in month_info['data'].columns:
                        status_col = col
                        break
                
                if status_col:
                    ps_count = len(month_info['data'][
                        month_info['data'][status_col].astype(str).str.lower().str.contains('ps|proses|progress', na=False)
                    ])
                    ps_rate = (ps_count / month_info['orders'] * 100) if month_info['orders'] > 0 else 0
                    response += f"   🎯 PS Rate: {ps_rate:.1f}%\n"
                
                response += f"   ─────────────────\n"
            
            # Summary
            if len(months_data) >= 2:
                latest = months_data[0]['orders']
                previous = months_data[1]['orders']
                change = ((latest - previous) / previous * 100) if previous > 0 else 0
                
                response += f"\n📊 *SUMMARY TREND:*\n"
                response += f"   📅 Latest Month: {latest:,}\n"
                response += f"   📊 Previous Month: {previous:,}\n"
                response += f"   📈 Change: {change:+.1f}%\n"
                
                if change > 10:
                    response += f"   ✅ *Growth positif!* Pertahankan momentum\n"
                elif change < -10:
                    response += f"   ⚠️ *Perlu evaluasi* strategi\n"
                else:
                    response += f"   ➡️ *Stabil,* pertahankan konsistensi\n"
            
            reply_markup = {
                "inline_keyboard": [
                    [{"text": "📊 Dashboard Utama", "callback_data": "dashboard_main"}],
                    [{"text": "🔙 Menu Dashboard", "callback_data": "dashboard_main"}]
                ]
            }
            
            if message_id:
                url = self.base_url + "editMessageText"
                payload = {
                    'chat_id': chat_id,
                    'message_id': message_id,
                    'text': response,
                    'parse_mode': 'Markdown',
                    'reply_markup': reply_markup
                }
                requests.post(url, json=payload)
            else:
                self.send_message(chat_id, response, 'Markdown', reply_markup)
                
        except Exception as e:
            logger.error(f"❌ Error in show_dashboard_trend: {e}")
            error_text = "❌ Gagal memuat analisis trend"
            self.send_message(chat_id, error_text, 'Markdown', self.create_dashboard_menu())

    def show_dashboard_sto(self, chat_id, message_id=None):
        """Menampilkan performa STO"""
        dashboard_data = self.get_dashboard_data()
        
        if not dashboard_data or not dashboard_data.get('sto_counts'):
            error_text = "❌ Tidak ada data performa STO"
            self.send_message(chat_id, error_text, 'Markdown', self.create_dashboard_menu())
            return
        
        sto_counts = dashboard_data['sto_counts']
        total_orders = dashboard_data['monthly_orders']
        
        response = f"""
    🏢 *STO PERFORMANCE - {dashboard_data['current_month']}*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    📊 *Total Order Bulan Ini:* {total_orders:,}
    🏢 *Jumlah STO Aktif:* {len(sto_counts)}

    """
        
        # Urutkan STO berdasarkan jumlah order
        sorted_stos = sorted(sto_counts.items(), key=lambda x: x[1], reverse=True)
        
        # Tampilkan semua STO dengan ranking
        for i, (sto_name, count) in enumerate(sorted_stos, 1):
            percentage = (count / total_orders * 100) if total_orders > 0 else 0
            
            # Tentukan ranking icon
            if i == 1:
                rank_icon = "🥇"
            elif i == 2:
                rank_icon = "🥈"
            elif i == 3:
                rank_icon = "🥉"
            elif i <= 10:
                rank_icon = f"{i}️⃣"
            else:
                rank_icon = "📋"
            
            response += f"{rank_icon} *{sto_name}:*\n"
            response += f"   📊 Order: {count:,} ({percentage:.1f}%)\n"
            
            # Berikan insight berdasarkan performa
            if percentage > 20:
                response += f"   💡 *High Performer!* Kontributor utama\n"
            elif percentage < 5:
                response += f"   ⚠️ *Perlu boost* order\n"
            
            response += f"   ─────────────────\n"
        
        # Summary
        if sorted_stos:
            top_sto = sorted_stos[0]
            bottom_sto = sorted_stos[-1] if len(sorted_stos) > 1 else None
            
            response += f"\n📈 *PERFORMANCE SUMMARY:*\n"
            response += f"   🏆 *Top STO:* {top_sto[0]} ({top_sto[1]:,} order)\n"
            
            if bottom_sto:
                response += f"   📉 *Bottom STO:* {bottom_sto[0]} ({bottom_sto[1]:,} order)\n"
            
            # Hitung distribusi
            if len(sorted_stos) >= 3:
                top3_total = sum(count for _, count in sorted_stos[:3])
                top3_percentage = (top3_total / total_orders * 100) if total_orders > 0 else 0
                response += f"   📊 *Top 3 STO:* {top3_percentage:.1f}% dari total order\n"
        
        reply_markup = {
            "inline_keyboard": [
                [{"text": "📋 Detail per STO", "callback_data": "sto_pivot"}],
                [{"text": "📊 Dashboard Utama", "callback_data": "dashboard_main"}]
            ]
        }
        
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': response,
                'parse_mode': 'Markdown',
                'reply_markup': reply_markup
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, response, 'Markdown', reply_markup)

    def show_dashboard_tech(self, chat_id, message_id=None):
        """Menampilkan ranking teknisi"""
        dashboard_data = self.get_dashboard_data()
        
        if not dashboard_data or not dashboard_data.get('tech_counts'):
            error_text = "❌ Tidak ada data ranking teknisi"
            self.send_message(chat_id, error_text, 'Markdown', self.create_dashboard_menu())
            return
        
        tech_counts = dashboard_data['tech_counts']
        total_orders = dashboard_data['monthly_orders']
        
        response = f"""
    👷 *TEKNISI RANKING - {dashboard_data['current_month']}*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    📊 *Total Order Bulan Ini:* {total_orders:,}
    👷 *Jumlah Teknisi Aktif:* {len(tech_counts)}

    """
        
        # Urutkan teknisi berdasarkan jumlah order
        sorted_techs = sorted(tech_counts.items(), key=lambda x: x[1], reverse=True)[:10]  # Top 10 saja
        
        # Tampilkan ranking teknisi
        for i, (tech_name, count) in enumerate(sorted_techs, 1):
            percentage = (count / total_orders * 100) if total_orders > 0 else 0
            
            # Tentukan ranking icon
            if i == 1:
                rank_icon = "👑"
            elif i == 2:
                rank_icon = "🥈"
            elif i == 3:
                rank_icon = "🥉"
            elif i <= 10:
                rank_icon = f"{i}️⃣"
            else:
                rank_icon = "📋"
            
            # Bersihkan nama teknisi jika ada format kompleks
            clean_name = str(tech_name).split('/')[0].split('-')[0].strip()
            
            response += f"{rank_icon} *{clean_name}:*\n"
            response += f"   📊 Order: {count:,} ({percentage:.1f}%)\n"
            
            # Hitung rata-rata order per teknisi
            avg_order = total_orders / len(tech_counts) if len(tech_counts) > 0 else 0
            
            if count > avg_order * 1.5:
                response += f"   ⭐ *High Performer!* Di atas rata-rata\n"
            elif count < avg_order * 0.5:
                response += f"   ⚠️ *Perlu peningkatan* produktivitas\n"
            
            response += f"   ─────────────────\n"
        
        # Summary
        if sorted_techs:
            top_tech = sorted_techs[0]
            clean_top_name = str(top_tech[0]).split('/')[0].split('-')[0].strip()
            
            response += f"\n🏆 *PERFORMANCE SUMMARY:*\n"
            response += f"   👑 *Top Teknisi:* {clean_top_name} ({top_tech[1]:,} order)\n"
            
            # Hitung rata-rata
            avg_order = total_orders / len(tech_counts) if len(tech_counts) > 0 else 0
            response += f"   📊 *Rata-rata order per teknisi:* {avg_order:.1f}\n"
            
            # Hitung jumlah teknisi di atas rata-rata
            above_avg = sum(1 for _, count in tech_counts.items() if count > avg_order)
            above_avg_percentage = (above_avg / len(tech_counts) * 100) if len(tech_counts) > 0 else 0
            response += f"   📈 *Teknisi di atas rata-rata:* {above_avg} ({above_avg_percentage:.1f}%)\n"
        
        reply_markup = {
            "inline_keyboard": [
                [{"text": "📊 Dashboard Utama", "callback_data": "dashboard_main"}]
            ]
        }
        
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': response,
                'parse_mode': 'Markdown',
                'reply_markup': reply_markup
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, response, 'Markdown', reply_markup)

    def show_dashboard_issues(self, chat_id, message_id=None):
        """Menampilkan analisis kendala"""
        dashboard_data = self.get_dashboard_data()
        
        if not dashboard_data or not dashboard_data.get('issue_analysis'):
            error_text = """
    ⚠️ *ANALISIS KENDALA*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    📊 *Tidak ada data kendala yang terdeteksi* untuk bulan ini.

    💡 *Kemungkinan:*
    • Tidak ada kendala yang tercatat
    • Data keterangan kosong
    • Format keterangan berbeda
    """
            self.send_message(chat_id, error_text, 'Markdown', self.create_dashboard_menu())
            return
        
        issue_analysis = dashboard_data['issue_analysis']
        total_issues = sum(issue_analysis.values())
        monthly_orders = dashboard_data['monthly_orders']
        
        response = f"""
    ⚠️ *ISSUE ANALYSIS - {dashboard_data['current_month']}*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    📊 *Total Order Bulan Ini:* {monthly_orders:,}
    🔧 *Total Kendala Terdeteksi:* {total_issues}
    📈 *Issue Rate:* {(total_issues/monthly_orders*100) if monthly_orders > 0 else 0:.1f}%

    """
        
        # Urutkan kendala berdasarkan frekuensi
        sorted_issues = sorted(issue_analysis.items(), key=lambda x: x[1], reverse=True)
        
        for issue, count in sorted_issues:
            percentage = (count / monthly_orders * 100) if monthly_orders > 0 else 0
            issue_percentage = (count / total_issues * 100) if total_issues > 0 else 0
            
            # Tentukan icon berdasarkan jenis kendala
            if 'FULL' in issue:
                icon = "🔴"
            elif 'TAT' in issue or 'TAJ' in issue:
                icon = "🟡"
            elif 'DOUBLE' in issue:
                icon = "🟠"
            elif 'CANCEL' in issue:
                icon = "🔵"
            elif 'TEKNIK' in issue:
                icon = "⚡"
            else:
                icon = "⚠️"
            
            response += f"{icon} *{issue}:*\n"
            response += f"   📊 Frekuensi: {count:,}\n"
            response += f"   📈 Terjadi pada: {percentage:.1f}% order\n"
            response += f"   🎯 {issue_percentage:.1f}% dari total kendala\n"
            
            # Berikan rekomendasi berdasarkan jenis kendala
            if 'ODP FULL' in issue:
                response += f"   💡 *Rekomendasi:* Survey penambahan ODP baru\n"
            elif 'TAT' in issue or 'TAJ' in issue:
                response += f"   💡 *Rekomendasi:* Review tag pelanggan\n"
            elif 'DOUBLE' in issue:
                response += f"   💡 *Rekomendasi:* Validasi input data\n"
            
            response += f"   ─────────────────\n"
        
        # Summary dan insight
        if sorted_issues:
            top_issue = sorted_issues[0]
            
            response += f"\n🎯 *INSIGHT UTAMA:*\n"
            response += f"   🔥 *Kendala dominan:* {top_issue[0]} ({top_issue[1]} kasus)\n"
            
            # Rekomendasi prioritas
            response += f"\n🚀 *REKOMENDASI PRIORITAS:*\n"
            
            if top_issue[0] == 'ODP FULL':
                response += f"   1️⃣ Survey penambahan ODP baru\n"
                response += f"   2️⃣ Optimasi kapasitas ODP existing\n"
                response += f"   3️⃣ Redistribusi order ke ODP terdekat\n"
            elif 'TAT' in top_issue[0]:
                response += f"   1️⃣ Validasi tagging sebelum dispatch\n"
                response += f"   2️⃣ Training teknis untuk tim survey\n"
                response += f"   3️⃣ Update database ODP available\n"
            elif 'CANCEL' in top_issue[0]:
                response += f"   1️⃣ Follow up lebih intensif ke pelanggan\n"
                response += f"   2️⃣ Validasi kelayakan sebelum dispatch\n"
                response += f"   3️⃣ Analisis pattern cancel per area\n"
        
        reply_markup = {
            "inline_keyboard": [
                [{"text": "📊 Dashboard Utama", "callback_data": "dashboard_main"}]
            ]
        }
        
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': response,
                'parse_mode': 'Markdown',
                'reply_markup': reply_markup
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, response, 'Markdown', reply_markup)

    def show_dashboard_actions(self, chat_id, message_id=None):
        """Menampilkan quick actions berdasarkan dashboard data"""
        dashboard_data = self.get_dashboard_data()
        
        if not dashboard_data:
            error_text = "❌ Tidak dapat mengambil data untuk quick actions"
            self.send_message(chat_id, error_text, 'Markdown', self.create_dashboard_menu())
            return
        
        response = f"""
    🚀 *QUICK ACTIONS - {dashboard_data['current_month']}*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    💡 *Berdasarkan analisis data, berikut aksi cepat yang bisa dilakukan:*

    """
        
        actions = []
        
        # Aksi berdasarkan PS Rate
        if dashboard_data['ps_rate'] < 70:
            actions.append("📊 *Review Order Low PS:* Cek order dengan status non-PS")
        elif dashboard_data['ps_rate'] > 90:
            actions.append("🏆 *Apresiasi Tim:* PS Rate excellent!")
        
        # Aksi berdasarkan issue analysis
        if dashboard_data.get('issue_analysis'):
            top_issue = max(dashboard_data['issue_analysis'].items(), key=lambda x: x[1]) if dashboard_data['issue_analysis'] else None
            if top_issue and top_issue[1] > 5:  # Jika lebih dari 5 kasus
                actions.append(f"🔧 *Tackle {top_issue[0]}:* {top_issue[1]} kasus terdeteksi")
        
        # Aksi berdasarkan trend
        if dashboard_data['order_change'] < -10:
            actions.append("📉 *Boost Order:* Penurunan signifikan vs bulan lalu")
        elif dashboard_data['order_change'] > 20:
            actions.append("📈 *Maintain Momentum:* Peningkatan excellent!")
        
        # Aksi berdasarkan STO performance
        if dashboard_data.get('sto_counts'):
            sorted_stos = sorted(dashboard_data['sto_counts'].items(), key=lambda x: x[1])
            if sorted_stos:
                bottom_sto = sorted_stos[0]
                if bottom_sto[1] < 5:  # Jika kurang dari 5 order
                    actions.append(f"🏢 *Support {bottom_sto[0]}:* Hanya {bottom_sto[1]} order bulan ini")
        
        # Tampilkan actions
        if actions:
            for i, action in enumerate(actions, 1):
                response += f"{i}. {action}\n"
        else:
            response += "✅ *Tidak ada aksi khusus yang diperlukan*\n"
            response += "   Semua indikator dalam kondisi baik!\n"
        
        # Tombol aksi cepat
        quick_buttons = []
        
        if dashboard_data['ps_rate'] < 70:
            quick_buttons.append({"text": "📋 Cek Low PS", "callback_data": "action_low_ps"})
        
        if dashboard_data.get('issue_analysis'):
            quick_buttons.append({"text": "⚠️ Review Issues", "callback_data": "action_review_issues"})
        
        if dashboard_data.get('sto_counts'):
            quick_buttons.append({"text": "🏢 Monitor STO", "callback_data": "action_monitor_sto"})
        
        # Buat keyboard
        keyboard_rows = []
        
        if quick_buttons:
            # Batasi 2 tombol per baris
            for i in range(0, len(quick_buttons), 2):
                row = quick_buttons[i:i+2]
                keyboard_rows.append(row)
        
        # Tambahkan tombol navigasi
        keyboard_rows.append([{"text": "📊 Dashboard Utama", "callback_data": "dashboard_main"}])
        
        reply_markup = {"inline_keyboard": keyboard_rows}
        
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': response,
                'parse_mode': 'Markdown',
                'reply_markup': reply_markup
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, response, 'Markdown', reply_markup)
    
    def handle_callback(self, callback_query):
        """Handle callback query dari Telegram"""
        data = callback_query['data']  # ✅ Akses sebagai dictionary
        chat_id = callback_query['message']['chat']['id']  # ✅ Akses sebagai dictionary
        
        # =====================================================
        # 1. Pagination PIVOT: page_pivot_{month_year}_{sto}_{page}
        # =====================================================
        if data.startswith("page_pivot_"):
            try:
                parts = data.split("_")
                if len(parts) >= 4:
                    month_year = parts[2]
                    sto_name = parts[3].replace("_", " ")
                    page = int(parts[4])
                    self.show_sto_pivot_details(chat_id, month_year, sto_name, page)
            except Exception as e:
                self.send_message(chat_id, f"⚠️ Terjadi error pivot: {e}")
            return

        # =====================================================
        # 2. Pagination Search: page_{search_type}_{query}_{page}
        # =====================================================
        if data.startswith("page_"):
            try:
                parts = data.split("_")
                if len(parts) >= 4:
                    search_type = parts[1]
                    query = parts[2].replace("_", " ")
                    page = int(parts[3])
                    
                    if search_type == "date":
                        self.show_date_results(chat_id, query, page)
                    elif search_type == "sto":
                        self.search_and_show_sto(chat_id, query, page)
            except Exception as e:
                self.send_message(chat_id, f"⚠️ Terjadi error pagination: {e}")
            return

        # =====================================================
        # 3. Tombol Kembali ke Pivot: pivot_{month_year}
        # =====================================================
        if data.startswith("pivot_"):
            try:
                if data.startswith("pivot_all_"):
                    # Handle "pivot_all_" untuk menampilkan semua STO
                    month_year = data.replace("pivot_all_", "").replace("_", "-")
                    self.show_all_sto_pivot(chat_id, month_year)
                else:
                    # Handle "pivot_" untuk menampilkan ringkasan
                    month_year = data.replace("pivot_", "").replace("_", "-")
                    self.show_sto_pivot_report(chat_id, month_year)
                return
                
            except Exception as e:
                self.send_message(chat_id, f"⚠️ Terjadi error pivot: {e}")
                return

        # =====================================================
        # 4. Menu Utama dan callback lainnya
        # =====================================================
        if data == "main_menu":
            self.show_main_menu(chat_id)
            return

        elif data == "search_order":
            self.ask_for_order_id(chat_id)

        elif data == "search_date":
            self.ask_for_date(chat_id)

        elif data == "search_sto":
            self.show_sto_menu(chat_id)

        elif data == "monthly_report":
            self.show_monthly_report_menu(chat_id)

        elif data == "sto_pivot":
            self.show_sto_pivot_menu(chat_id)
            return

        elif data.startswith("pivot_"):
            # Handle pivot report untuk bulan tertentu
            month_year = data.replace("pivot_", "")
            self.show_sto_pivot_report(chat_id, month_year)
            return

        elif data == "custom_report_month":
            self.show_month_selection_menu(chat_id, None, "report")

        elif data == "custom_pivot_month":
            self.show_month_selection_menu(chat_id, None, "pivot")

        elif data.startswith("report_"):
            month_year = data.replace("report_", "")
            self.show_monthly_report(chat_id, month_year)

        elif data.startswith("pivot_detail_"):
            parts = data.split('_')
            if len(parts) >= 4:
                month_year = parts[2]
                sto_name = parts[3]
                self.show_sto_pivot_details(chat_id, month_year, sto_name)

        elif data.startswith("pivot_all_"):
            month_year = data.replace("pivot_all_", "")
            self.show_all_sto_pivot(chat_id, month_year)

        elif data.startswith("sc_detail_"):
            sc_number = data.replace("sc_detail_", "")
            self.show_sc_details(chat_id, sc_number)

        elif data == "help":
            self.show_help(chat_id)

        elif data.startswith("sto_"):
            sto_name = data.replace("sto_", "").upper()
            self.search_and_show_sto(chat_id, sto_name, page=1)

        
        if data == "dashboard":
            self.show_dashboard_main(chat_id, callback_query['message']['message_id'])
            return
    
        elif data == "dashboard_main":
            self.show_dashboard_main(chat_id, callback_query['message']['message_id'])
            return
        
        elif data == "dashboard_trend":
            self.show_dashboard_trend(chat_id, callback_query['message']['message_id'])
            return
        
        elif data == "dashboard_sto":
            self.show_dashboard_sto(chat_id, callback_query['message']['message_id'])
            return
        
        elif data == "dashboard_tech":
            self.show_dashboard_tech(chat_id, callback_query['message']['message_id'])
            return
        
        elif data == "dashboard_issues":
            self.show_dashboard_issues(chat_id, callback_query['message']['message_id'])
            return
        
        elif data == "dashboard_actions":
            self.show_dashboard_actions(chat_id, callback_query['message']['message_id'])
            return
        
        elif data == "dashboard_refresh":
            self.show_dashboard_main(chat_id, callback_query['message']['message_id'])
            return

        elif data == "action_low_ps":
            # Cari order dengan status non-PS
            self.show_low_ps_orders(chat_id, callback_query['message']['message_id'])
            return
    
        elif data == "action_review_issues":
            self.show_dashboard_issues(chat_id, callback_query['message']['message_id'])
            return
        
        elif data == "action_monitor_sto":
            self.show_dashboard_sto(chat_id, callback_query['message']['message_id'])
            return

    def show_low_ps_orders(self, chat_id, message_id=None):
        """Menampilkan order dengan status non-PS (low PS rate)"""
        try:
            df = self.order_checker.get_all_data()
            if df.empty:
                error_text = "❌ Tidak ada data order"
                self.send_message(chat_id, error_text, 'Markdown')
                return
            
            # Cari kolom status
            status_columns = ['STATUS', 'STATUS ORDER', 'KETERANGAN']
            status_col = None
            
            for col in status_columns:
                if col in df.columns:
                    status_col = col
                    break
            
            if not status_col:
                error_text = "❌ Kolom status tidak ditemukan"
                self.send_message(chat_id, error_text, 'Markdown')
                return
            
            # Filter order dengan status non-PS
            non_ps_data = df[~df[status_col].astype(str).str.lower().str.contains('ps|proses|progress', na=False)]
            
            # Filter untuk bulan ini
            date_columns = ['ORDER DATE', 'ORDER_DATE', 'TANGGAL', 'DATE', 'TANGGAL ORDER']
            date_col = None
            
            for col in date_columns:
                if col in df.columns:
                    date_col = col
                    break
            
            if date_col:
                # Konversi tanggal
                df['CONVERTED_DATE'] = pd.to_datetime(
                    df[date_col],
                    format='%m-%d-%y',
                    errors='coerce'
                )
                
                current_date = datetime.now()
                df['MONTH'] = df['CONVERTED_DATE'].dt.month
                df['YEAR'] = df['CONVERTED_DATE'].dt.year
                
                monthly_non_ps = df[
                    (df['MONTH'] == current_date.month) & 
                    (df['YEAR'] == current_date.year) &
                    (~df[status_col].astype(str).str.lower().str.contains('ps|proses|progress', na=False))
                ]
            else:
                monthly_non_ps = non_ps_data
            
            response = f"""
    📋 *LOW PS ORDERS - PERLU TINDAK LANJUT*
    ━━━━━━━━━━━━━━━━━━━━━━━━

    📊 *Total Non-PS Orders:* {len(non_ps_data):,}
    📅 *Non-PS Bulan Ini:* {len(monthly_non_ps):,}

    """
            
            if len(monthly_non_ps) > 0:
                response += "🔍 *TOP 10 ORDER NON-PS BULAN INI:*\n\n"
                
                for i, (_, row) in enumerate(monthly_non_ps.head(10).iterrows(), 1):
                    response += f"{i}. *{row.get('SC', 'N/A')}*\n"
                    response += f"   🏢 STO: {row.get('STO', 'N/A')}\n"
                    response += f"   📊 Status: {row.get(status_col, 'N/A')}\n"
                    response += f"   👷 Teknisi: {row.get('TEKNISI', 'N/A')}\n"
                    
                    if date_col:
                        response += f"   📅 Tanggal: {row.get(date_col, 'N/A')}\n"
                    
                    response += f"   ─────────────────\n"
                
                if len(monthly_non_ps) > 10:
                    response += f"\n📋 ...dan {len(monthly_non_ps) - 10} order lainnya\n"
            else:
                response += "✅ *Tidak ada order non-PS bulan ini!* Excellent work!\n"
            
            reply_markup = {
                "inline_keyboard": [
                    [{"text": "📊 Dashboard Utama", "callback_data": "dashboard_main"}],
                    [{"text": "🔙 Menu Utama", "callback_data": "main_menu"}]
                ]
            }
            
            if message_id:
                url = self.base_url + "editMessageText"
                payload = {
                    'chat_id': chat_id,
                    'message_id': message_id,
                    'text': response,
                    'parse_mode': 'Markdown',
                    'reply_markup': reply_markup
                }
                requests.post(url, json=payload)
            else:
                self.send_message(chat_id, response, 'Markdown', reply_markup)
                
        except Exception as e:
            logger.error(f"❌ Error in show_low_ps_orders: {e}")
            error_text = "❌ Gagal mengambil data low PS orders"
            self.send_message(chat_id, error_text, 'Markdown', self.create_back_menu())
    
    def send_message(self, chat_id, text, parse_mode=None, reply_markup=None):
        """Mengirim pesan ke Telegram"""
        url = self.base_url + "sendMessage"
        payload = {
            'chat_id': chat_id,
            'text': text
        }
        
        if parse_mode:
            payload['parse_mode'] = parse_mode
        if reply_markup:
            payload['reply_markup'] = reply_markup
        
        try:
            response = requests.post(url, json=payload, timeout=10)
            if response.status_code == 200:
                logger.info(f"✅ Message sent to chat_id: {chat_id}")
                return response.json()
            else:
                logger.error(f"❌ Failed to send message: {response.status_code} - {response.text}")
                return None
        except Exception as e:
            logger.error(f"❌ Error sending message: {e}")
            return None

    def create_main_menu(self):
        """Membuat menu utama dengan inline keyboard"""
        keyboard = {
            "inline_keyboard": [
                # [
                #     {"text": "📊 Dashboard", "callback_data": "dashboard"},  # ✅ TAMBAHKAN INI
                # ],
                [
                    {"text": "🔍 Cari Order ID", "callback_data": "search_order"},
                    {"text": "📅 Cari Tanggal", "callback_data": "search_date"}
                ],
                [
                    {"text": "📊 Laporan Bulanan", "callback_data": "monthly_report"},
                    {"text": "🏢 Pivot per STO", "callback_data": "sto_pivot"}
                ],
            ]
        }
        return keyboard

    def create_monthly_report_menu(self):
        """Membuat menu laporan bulanan"""
        current_date = datetime.now()
        current_month = current_date.strftime('%m-%Y')
        last_month = (current_date.replace(day=1) - pd.DateOffset(months=1)).strftime('%m-%Y')
        
        keyboard = {
            "inline_keyboard": [
                [
                    {"text": "📊 Bulan Ini", "callback_data": f"report_{current_month}"},
                    {"text": "📊 Bulan Lalu", "callback_data": f"report_{last_month}"}
                ],
                [
                    {"text": "📅 Pilih Bulan Lain", "callback_data": "custom_report_month"}
                ],
                [
                    {"text": "🔙 Menu Utama", "callback_data": "main_menu"}
                ]
            ]
        }
        return keyboard

    def create_sto_pivot_menu(self):
        """Membuat menu pivot STO"""
        current_date = datetime.now()
        current_month = current_date.strftime('%m-%Y')
        last_month = (current_date.replace(day=1) - pd.DateOffset(months=1)).strftime('%m-%Y')
        
        keyboard = {
            "inline_keyboard": [
                [
                    {"text": "📊 Bulan Ini", "callback_data": f"pivot_{current_month}"},
                    {"text": "📊 Bulan Lalu", "callback_data": f"pivot_{last_month}"}
                ],
                [
                    {"text": "📅 Pilih Bulan Lain", "callback_data": "custom_pivot_month"}
                ],
                [
                    {"text": "🔙 Menu Utama", "callback_data": "main_menu"}
                ]
            ]
        }
        return keyboard

    def create_sto_pivot_detail_menu(self, month_year, sto_pivot):
        """Membuat menu untuk melihat detail per STO dalam pivot"""
        keyboard_buttons = []
        
        # Tombol untuk setiap STO yang memiliki data
        for sto_name, sto_data in sto_pivot.items():
            if sto_data['total_orders'] > 0:
                keyboard_buttons.append([
                    {"text": f"🏢 {sto_name} ({sto_data['total_orders']})", "callback_data": f"pivot_detail_{month_year}_{sto_name}"}
                ])
        
        # Urutkan berdasarkan total orders (descending)
        keyboard_buttons.sort(key=lambda x: int(re.search(r'\((\d+)\)', x[0]['text']).group(1)), reverse=True)
        
        # Batasi maksimal 10 STO per halaman
        if len(keyboard_buttons) > 10:
            keyboard_buttons = keyboard_buttons[:10]
            keyboard_buttons.append([
                {"text": "📋 Lihat Semua STO", "callback_data": f"pivot_all_{month_year}"}
            ])
        
        keyboard_buttons.append([
            {"text": "🔙 Kembali", "callback_data": "sto_pivot"}
        ])
        
        keyboard = {
            "inline_keyboard": keyboard_buttons
        }
        return keyboard

    def create_sto_sc_list_menu(self, sc_list, month_year, sto_name, page=1, page_size=10):
        """Membuat menu daftar SC untuk STO tertentu - FIXED VERSION"""
        # Validasi input
        if not isinstance(sc_list, list):
            sc_list = []
        
        # Pagination
        total_sc = len(sc_list)
        total_pages = max(1, (total_sc + page_size - 1) // page_size)
        
        # Pastikan page dalam rentang yang valid
        page = max(1, min(page, total_pages))
        
        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size
        page_sc_list = sc_list[start_idx:end_idx]
        
        keyboard_buttons = []
        
        # Buat tombol untuk setiap SC dalam halaman ini
        for sc in page_sc_list:
            if sc and str(sc).strip():  # Pastikan SC tidak kosong
                # Pastikan SC aman untuk callback data
                safe_sc = str(sc).replace(" ", "_").replace("/", "_")
                keyboard_buttons.append([
                    {"text": f"📋 {sc}", "callback_data": f"sc_detail_{safe_sc}"}
                ])
        
        # Tombol navigasi halaman
        nav_buttons = []
        
        # Pastikan sto_name aman untuk callback
        safe_sto = str(sto_name).replace(" ", "_").replace("/", "_")
        
        if page > 1:
            nav_buttons.append({
                "text": "⬅️ Sebelumnya",
                "callback_data": f"page_pivot_{month_year}_{safe_sto}_{page-1}"
            })
        
        # Info halaman
        nav_buttons.append({
            "text": f"{page}/{total_pages}",
            "callback_data": "noop"
        })
        
        if page < total_pages:
            nav_buttons.append({
                "text": "Selanjutnya ➡️",
                "callback_data": f"page_pivot_{month_year}_{safe_sto}_{page+1}"
            })
        
        if nav_buttons:
            keyboard_buttons.append(nav_buttons)
        
        # Tombol kembali ke pivot utama
        keyboard_buttons.append([
            {"text": "🔙 Kembali ke Pivot", "callback_data": f"pivot_{month_year}"}
        ])
        
        keyboard = {"inline_keyboard": keyboard_buttons}
        return keyboard

    def create_month_selection_menu(self, callback_prefix="report"):
        """Membuat menu pemilihan bulan"""
        current_date = datetime.now()
        months = []
        
        # Buat pilihan untuk 6 bulan terakhir
        for i in range(12):
            date = current_date - pd.DateOffset(months=i)
            month_year = date.strftime('%m-%Y')
            month_name = date.strftime('%B %Y')
            months.append({
                "text": f"📅 {month_name}", 
                "callback_data": f"{callback_prefix}_{month_year}"
            })
        
        # Atur dalam grid 2 kolom
        keyboard_rows = []
        for i in range(0, len(months), 2):
            row = months[i:i+2]
            keyboard_rows.append(row)
        
        keyboard_rows.append([
            {"text": "🔙 Kembali", "callback_data": "monthly_report" if callback_prefix == "report" else "sto_pivot"}
        ])
        
        keyboard = {
            "inline_keyboard": keyboard_rows
        }
        return keyboard

    def format_monthly_report(self, report_data):
        """Memformat laporan bulanan menjadi pesan yang menarik"""
        if not report_data:
            return "❌ *Gagal mengambil data laporan*"
        
        month_name = report_data.get('month_name', report_data.get('month_year', 'N/A'))
        total_orders = report_data['total_orders']
        status_counts = report_data.get('status_counts', {})
        sto_counts = report_data.get('sto_counts', {})
        
        response = f"""
📊 *LAPORAN BULANAN*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Periode:* {month_name}
📈 *Total Order:* {total_orders} tiket

"""
        
        if status_counts:
            response += "┌─── *BREAKDOWN STATUS* ───\n"
            for status, count in sorted(status_counts.items(), key=lambda x: x[1], reverse=True)[:10]:
                percentage = (count / total_orders * 100) if total_orders > 0 else 0
                response += f"│ 📋 {status}: {count} ({percentage:.1f}%)\n"
            response += "└─────────────────────\n\n"
        
        if sto_counts:
            response += "┌─── *TOP STO* ───\n"
            for sto, count in sorted(sto_counts.items(), key=lambda x: x[1], reverse=True)[:5]:
                percentage = (count / total_orders * 100) if total_orders > 0 else 0
                response += f"│ 🏢 {sto}: {count} ({percentage:.1f}%)\n"
            response += "└────────────────\n"
        
        if total_orders == 0:
            response += "\n❌ *Tidak ada data untuk bulan ini*"
        
        return response

    def format_sto_pivot_report(self, pivot_data):
        """Memformat laporan pivot STO menjadi pesan yang menarik"""
        if not pivot_data:
            return "❌ *Gagal mengambil data pivot*"
        
        month_name = pivot_data.get('month_name', pivot_data.get('month_year', 'N/A'))
        total_orders = pivot_data['total_orders']
        sto_pivot = pivot_data['sto_pivot']
        
        if not sto_pivot:
            return f"""
📊 *PIVOT PER STO - TIDAK ADA DATA*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Periode:* {month_name}
📈 *Total Order:* {total_orders} tiket

❌ *Tidak ada data STO yang ditemukan*
"""
        
        response = f"""
📊 *PIVOT PER STO - PS/RE ANALYSIS*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Periode:* {month_name}
📈 *Total Order:* {total_orders} tiket
🏢 *Jumlah STO:* {len(sto_pivot)}

┌─── *RINGKASAN PER STO* ───
"""
        
        # Urutkan STO berdasarkan total orders (descending)
        sorted_sto = sorted(sto_pivot.items(), key=lambda x: x[1]['total_orders'], reverse=True)
        
        # Tampilkan hanya 5 STO teratas di ringkasan
        for sto_name, sto_data in sorted_sto[:5]:
            response += f"│ 🏢 *{sto_name}:*\n"
            response += f"│   📊 Total: {sto_data['total_orders']}\n"
            response += f"│   ✅ PS: {sto_data['ps_count']}\n"
            response += f"│   📈 RE: {sto_data['re_count']}\n"
            response += f"│   🎯 RE Rate: {sto_data['re_rate']:.1f}%\n"
            response += f"│   ─────────────────\n"
        
        if len(sorted_sto) > 5:
            response += f"│ ... dan {len(sorted_sto) - 5} STO lainnya\n"
        
        response += f"└─────────────────────\n\n"
        
        # Summary statistics
        total_ps = sum(sto_data['ps_count'] for sto_data in sto_pivot.values())
        total_re = sum(sto_data['re_count'] for sto_data in sto_pivot.values())
        overall_re_rate = (total_ps / total_re * 100) if total_re > 0 else 0
        
        response += f"📈 *STATISTIK KESELURUHAN:*\n"
        response += f"• ✅ Total PS: {total_ps}\n"
        response += f"• 📊 Total RE: {total_re}\n"
        response += f"• 🎯 Overall RE Rate: {overall_re_rate:.1f}%\n\n"
        
        return response

    def format_sto_pivot_details(self, sto_data, page=1, page_size=10):
        """Memformat detail pivot untuk STO tertentu - FIXED VERSION"""
        if not sto_data:
            return "❌ *Tidak ada data yang ditemukan*", 0
        
        sto_name = sto_data.get('sto_name', 'N/A')
        total_orders = sto_data.get('total_orders', 0)
        ps_count = sto_data.get('ps_count', 0)
        re_count = sto_data.get('re_count', 0)
        re_rate = sto_data.get('re_rate', 0)
        status_breakdown = sto_data.get('status_breakdown', {})
        sc_list = sto_data.get('sc_list', [])
        month_name = sto_data.get('month_name', 'N/A')
        month_year = sto_data.get('month_year', 'N/A')
        
        # Pagination
        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size
        page_sc_list = sc_list[start_idx:end_idx]
        total_pages = max(1, (len(sc_list) + page_size - 1) // page_size)
        
        response = f"""
🏢 *DETAIL PIVOT - {sto_name.upper()}*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Periode:* {month_name} ({month_year})
📊 *Total Order:* {total_orders}
✅ *PS Count:* {ps_count}
📈 *RE Count:* {re_count}
🎯 *RE Rate:* {re_rate:.1f}%

"""
        
        if status_breakdown:
            response += "┌─── *BREAKDOWN STATUS* ───\n"
            for status, count in status_breakdown.items():
                percentage = (count / total_orders * 100) if total_orders > 0 else 0
                response += f"│ 📋 {status}: {count} ({percentage:.1f}%)\n"
            response += "└─────────────────────\n\n"
        
        response += f"📋 *Daftar SC ({len(sc_list)}):*\n"
        
        if sc_list:
            for i, sc in enumerate(page_sc_list, start_idx + 1):
                response += f"{i}. `{sc}`\n"
            
            response += f"\n📄 *Halaman:* {page} dari {total_pages}\n"
            
            if len(sc_list) > 0:
                response += f"💡 *Klik salah satu SC di bawah untuk melihat detail lengkap*"
        else:
            response += "\n*Tidak ada data SC*"
        
        return response, total_pages
        
    def ask_for_order_id(self, chat_id, message_id=None):
        """Meminta user untuk memasukkan Order ID"""
        self.user_search_state[chat_id] = "waiting_for_order_id"
        
        text = """
🔍 *PENCARIAN ORDER ID*
━━━━━━━━━━━━━━━━━━━━━━━━

📝 *Mode aktif:* **Pencarian Order ID**

Silakan ketikkan *Order ID* yang ingin dicari:

📌 *Format:* `SC123456789`
📌 *Contoh:*
• `SC123456789`
• `SC987654321`
• `SC123` (cari sebagian)

💡 *Tips:*
• Bisa memasukkan **beberapa Order ID berturut-turut**
• Setelah hasil muncul, **langsung ketik Order ID berikutnya**
• Tekan `/start` untuk kembali ke menu utama
"""
        
        # Tombol sederhana hanya untuk ke menu utama
        keyboard = {
            "inline_keyboard": [
                [
                    {"text": "🏠 Menu Utama", "callback_data": "main_menu"}
                ]
            ]
        }
        
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': text,
                'parse_mode': 'Markdown',
                'reply_markup': keyboard
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, text, 'Markdown', keyboard)

    def ask_for_date(self, chat_id, message_id=None):
        """Meminta user untuk memasukkan tanggal"""
        self.user_search_state[chat_id] = "waiting_for_date"
        
        text = """
📅 *PENCARIAN BERDASARKAN TANGGAL*
━━━━━━━━━━━━━━━━━━━━━━━━

📝 *Mode aktif:* **Pencarian Tanggal**

Silakan ketikkan *tanggal* yang ingin dicari:

📌 *Format:* `DD-MM-YYYY` (Hari-Bulan-Tahun)
📌 *Contoh:*
• `15-12-2024` = 15 Desember 2024
• `01-01-2024` = 1 Januari 2024
• `31-03-2024` = 31 Maret 2024

💡 *Informasi Sistem:*
• Data di spreadsheet menggunakan format **MM-DD-YY**
• Contoh: `12-15-24` = 15 Desember 2024
• Sistem akan otomatis konversi format

💡 *Tips:*
• Bisa memasukkan **beberapa tanggal berturut-turut**
• Setelah hasil muncul, **langsung ketik tanggal berikutnya**
• Tekan `/start` untuk kembali ke menu utama
"""
        
        keyboard = {
            "inline_keyboard": [
                [
                    {"text": "🏠 Menu Utama", "callback_data": "main_menu"}
                ]
            ]
        }
        
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': text,
                'parse_mode': 'Markdown',
                'reply_markup': keyboard
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, text, 'Markdown', keyboard)

    def show_monthly_report_menu(self, chat_id, message_id=None):
        """Menampilkan menu laporan bulanan"""
        menu_text = """
📊 *LAPORAN BULANAN*
━━━━━━━━━━━━━━━━━━━━━━━━

📈 *Statistik bulanan meliputi:*
• Total order per bulan
• Breakdown status
• Distribusi per STO
• Persentase penyelesaian

Pilih periode laporan:
"""
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': menu_text,
                'parse_mode': 'Markdown',
                'reply_markup': self.create_monthly_report_menu()
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, menu_text, 'Markdown', self.create_monthly_report_menu())

    def show_sto_pivot_menu(self, chat_id, message_id=None):
        """Menampilkan menu pivot STO"""
        menu_text = """
🏢 *PIVOT PER STO - PS/RE ANALYSIS*
━━━━━━━━━━━━━━━━━━━━━━━━

📊 *Analisis Performa per STO:*

• 🏢 Detail setiap STO
• ✅ PS (Proses/Progress) Count
• 📈 RE (Total Status) Count  
• 🎯 RE Rate (PS/RE)
• 📋 Daftar SC per STO

Pilih periode laporan:
"""
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': menu_text,
                'parse_mode': 'Markdown',
                'reply_markup': self.create_sto_pivot_menu()
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, menu_text, 'Markdown', self.create_sto_pivot_menu())

    def show_monthly_report(self, chat_id, month_year):
        """Menampilkan laporan bulanan"""
        loading_text = f"""
🔄 *MEMPROSES LAPORAN BULANAN...*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Bulan:* {month_year}
📊 *Menganalisis data...*
⏳ Mohon tunggu sebentar...
"""
        self.send_message(chat_id, loading_text, 'Markdown')
        
        report_data = self.order_checker.get_monthly_report(month_year)
        
        if report_data:
            formatted_report = self.format_monthly_report(report_data)
            self.send_message(chat_id, formatted_report, 'Markdown', self.create_back_menu())
        else:
            error_text = f"""
❌ *GAGAL MENGAMBIL DATA*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Bulan:* {month_year}

Tidak dapat mengambil data laporan untuk bulan ini.

💡 *Kemungkinan penyebab:*
• Data tidak tersedia
• Format spreadsheet berubah
• Koneksi terputus
"""
            self.send_message(chat_id, error_text, 'Markdown', self.create_monthly_report_menu())

    def show_sto_pivot_report(self, chat_id, month_year):
        """Menampilkan laporan pivot STO"""
        loading_text = f"""
🔄 *MEMPROSES PIVOT STO...*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Bulan:* {month_year}
🏢 *Analisis per STO*
⏳ Mohon tunggu sebentar...
"""
        self.send_message(chat_id, loading_text, 'Markdown')
        
        pivot_data = self.order_checker.get_sto_pivot_report(month_year)
        
        if pivot_data:
            if pivot_data['sto_pivot']:
                formatted_report = self.format_sto_pivot_report(pivot_data)
                
                # Ganti bagian ini: langsung tampilkan menu pivot setelah laporan
                # Kirimkan callback_data dengan month_year yang sudah aman
                safe_month_year = month_year.replace("-", "_")
                
                keyboard = {
                    "inline_keyboard": [
                        [
                            # {"text": "📊 Lihat Ringkasan", "callback_data": f"pivot_{safe_month_year}"},
                            {"text": "📋 Detail per STO", "callback_data": f"pivot_all_{safe_month_year}"}
                        ],
                        [
                            {"text": "🔙 Kembali ke Menu", "callback_data": "sto_pivot"}
                        ]
                    ]
                }
                
                self.send_message(chat_id, formatted_report, 'Markdown', keyboard)
            else:
                error_text = f"""
📭 *TIDAK ADA DATA PIVOT*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Bulan:* {month_year}

Tidak ditemukan data pivot untuk bulan ini.

💡 *Kemungkinan penyebab:*
• Tidak ada data order di bulan tersebut
• Data kosong untuk periode ini
• Coba bulan lain
"""
                self.send_message(chat_id, error_text, 'Markdown', self.create_sto_pivot_menu())
        else:
            error_text = f"""
❌ *GAGAL MENGAMBIL DATA PIVOT*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Bulan:* {month_year}

Tidak dapat mengambil data pivot untuk bulan ini.

💡 *Kemungkinan penyebab:*
• Data tidak tersedia
• Format spreadsheet berubah
• Koneksi terputus
"""
            self.send_message(chat_id, error_text, 'Markdown', self.create_sto_pivot_menu())

    def show_sto_pivot_details(self, chat_id, month_year, sto_name, page=1):
        """Menampilkan detail pivot untuk STO tertentu - FIXED VERSION"""
        try:
            loading_text = f"""
🔄 *MENGAMBIL DETAIL STO...*
━━━━━━━━━━━━━━━━━━━━━━━━

🏢 *STO:* {sto_name}
📅 *Bulan:* {month_year}
⏳ Mohon tunggu sebentar...
"""
            self.send_message(chat_id, loading_text, 'Markdown')
            
            logger.info(f"🔍 Showing STO pivot details: {month_year}, {sto_name}, page {page}")
            
            # Ambil data detail STO
            sto_data = self.order_checker.get_sto_pivot_details(month_year, sto_name)
            
            if not sto_data:
                error_text = f"""
❌ *DATA TIDAK DITEMUKAN*
━━━━━━━━━━━━━━━━━━━━━━━━

🏢 *STO:* {sto_name}
📅 *Bulan:* {month_year}

Tidak ditemukan data untuk kombinasi STO dan bulan ini.

💡 *Kemungkinan penyebab:*
• STO tidak ada di bulan tersebut
• Data kosong untuk periode ini
• Format nama STO berbeda
"""
                self.send_message(chat_id, error_text, 'Markdown', self.create_back_menu())
                return
            
            # Format detail STO
            formatted_details, total_pages = self.format_sto_pivot_details(sto_data, page)
            
            # Buat menu dengan tombol SC yang bisa diklik
            reply_markup = self.create_sto_sc_list_menu(
                sto_data['sc_list'], 
                month_year, 
                sto_name, 
                page
            )
            
            self.send_message(chat_id, formatted_details, 'Markdown', reply_markup)
            
        except Exception as e:
            logger.error(f"❌ Error in show_sto_pivot_details: {e}")
            import traceback
            traceback.print_exc()
            error_text = f"""
⚠️ *TERJADI KESALAHAN*
━━━━━━━━━━━━━━━━━━━━━━━━

Gagal mengambil detail data STO.
Error: {str(e)}
"""
            self.send_message(chat_id, error_text, 'Markdown', self.create_back_menu())

    def show_month_selection_menu(self, chat_id, message_id=None, callback_prefix="report"):
        """Menampilkan menu pemilihan bulan"""
        menu_text = """
📅 *PILIH BULAN*
━━━━━━━━━━━━━━━━━━━━━━━━

Pilih bulan yang ingin dilihat laporannya:

"""
        reply_markup = self.create_month_selection_menu(callback_prefix)
        
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': menu_text,
                'parse_mode': 'Markdown',
                'reply_markup': reply_markup
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, menu_text, 'Markdown', reply_markup)

    def show_all_sto_pivot(self, chat_id, month_year):
        """Menampilkan semua STO dalam pivot (untuk data yang banyak)"""
        loading_text = f"""
🔄 *MEMPROSES SEMUA STO...*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Bulan:* {month_year}
🏢 *Mengambil semua STO*
⏳ Mohon tunggu sebentar...
"""
        self.send_message(chat_id, loading_text, 'Markdown')
        
        pivot_data = self.order_checker.get_sto_pivot_report(month_year)
        
        if pivot_data and pivot_data['sto_pivot']:
            # Format khusus untuk menampilkan semua STO
            month_name = pivot_data.get('month_name', month_year)
            total_orders = pivot_data['total_orders']
            sto_pivot = pivot_data['sto_pivot']
            
            response = f"""
📊 *SEMUA STO - PIVOT ANALYSIS*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Periode:* {month_name}
📈 *Total Order:* {total_orders} tiket
🏢 *Jumlah STO:* {len(sto_pivot)}

"""
            
            # Urutkan STO berdasarkan total orders (descending)
            sorted_sto = sorted(sto_pivot.items(), key=lambda x: x[1]['total_orders'], reverse=True)
            
            for sto_name, sto_data in sorted_sto:
                response += f"🏢 *{sto_name}:*\n"
                response += f"   📊 Total: {sto_data['total_orders']}\n"
                response += f"   ✅ PS: {sto_data['ps_count']}\n"
                response += f"   📈 RE: {sto_data['re_count']}\n"
                response += f"   🎯 RE Rate: {sto_data['re_rate']:.1f}%\n"
                
                # Tampilkan top 3 status
                top_statuses = sorted(sto_data['status_breakdown'].items(), 
                                    key=lambda x: x[1], reverse=True)[:3]
                if top_statuses:
                    response += f"   📋 Top Status: "
                    status_texts = []
                    for status, count in top_statuses:
                        status_texts.append(f"{status}({count})")
                    response += ", ".join(status_texts) + "\n"
                
                response += f"   ─────────────────\n"
            
            # Tombol kembali ke menu pivot
            safe_month_year = month_year.replace("-", "_")
            keyboard = {
                "inline_keyboard": [
                    [
                        {"text": "📊 Kembali ke Pivot", "callback_data": f"pivot_{safe_month_year}"}
                    ],
                    [
                        {"text": "🔙 Menu Pivot", "callback_data": "sto_pivot"}
                    ]
                ]
            }
            
            self.send_message(chat_id, response, 'Markdown', keyboard)
        else:
            error_text = f"""
❌ *TIDAK ADA DATA STO*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Bulan:* {month_year}

Tidak ditemukan data STO untuk bulan ini.
"""
            self.send_message(chat_id, error_text, 'Markdown', self.create_sto_pivot_menu())

    def show_sc_details(self, chat_id, sc_number):
        """Menampilkan detail SC"""
        loading_text = f"""
🔄 *MENGAMBIL DETAIL SC...*
━━━━━━━━━━━━━━━━━━━━━━━━

📝 *SC:* {sc_number}
⏳ Mohon tunggu sebentar...
"""
        self.send_message(chat_id, loading_text, 'Markdown')
        
        results = self.order_checker.get_sc_details(sc_number)
        
        if results:
            response = f"""
✅ *DETAIL ORDER DITEMUKAN*
━━━━━━━━━━━━━━━━━━━━━━━━

📝 *Order ID:* `{sc_number}`
📊 *Total Data:* {len(results)} record

"""
            for i, item in enumerate(results, 1):
                response += f"┌─── *DATA {i}* ───\n"
                for key, value in item.items():
                    if pd.notna(value) and str(value).strip() != '':
                        response += f"│ • *{key}:* {value}\n"
                response += f"└────────────────\n\n"
        else:
            response = f"""
❌ *DATA TIDAK DITEMUKAN*
━━━━━━━━━━━━━━━━━━━━━━━━

📝 *Order ID:* `{sc_number}`

Tidak ditemukan data untuk SC ini.
"""
        
        self.send_message(chat_id, response, 'Markdown', self.create_back_menu())

    def show_sto_menu(self, chat_id, message_id=None):
        """Menampilkan menu STO"""
        menu_text = """
🏢 *PENCARIAN BERDASARKAN STO*
━━━━━━━━━━━━━━━━━━━━━━━━

Pilih STO yang ingin dilihat datanya:

"""
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': menu_text,
                'parse_mode': 'Markdown',
                'reply_markup': self.create_sto_menu()
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, menu_text, 'Markdown', self.create_sto_menu())

    def show_pivot_menu(self, chat_id, month_year, message_id=None):
        """Menampilkan menu pivot untuk bulan tertentu"""
        menu_text = f"""
🏢 *PIVOT STO - {month_year}*
━━━━━━━━━━━━━━━━━━━━━━━━

Pilih aksi yang ingin dilakukan:
"""
        # Pastikan month_year aman untuk callback
        safe_month_year = month_year.replace("-", "_")
        
        keyboard = {
            "inline_keyboard": [
                [
                    {"text": "🔙 Kembali", "callback_data": "sto_pivot"},
                    # {"text": "📊 Lihat Ringkasan", "callback_data": f"pivot_{safe_month_year}"},
                    {"text": "📋 Detail per STO", "callback_data": f"pivot_all_{safe_month_year}"}
                ],
            ]
        }
        
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': menu_text,
                'parse_mode': 'Markdown',
                'reply_markup': keyboard
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, menu_text, 'Markdown', keyboard)

    def create_sto_menu(self):
        """Membuat menu STO"""
        keyboard = {
            "inline_keyboard": [
                [
                    {"text": "🏢 DS 01", "callback_data": "sto_ds01"},
                    {"text": "🏢 DS 02", "callback_data": "sto_ds02"},
                    {"text": "🏢 DS 03", "callback_data": "sto_ds03"}
                ],
                [
                    {"text": "🏢 DS 07", "callback_data": "sto_ds07"},
                    {"text": "🏢 DS 08", "callback_data": "sto_ds08"},
                    {"text": "🏢 DS 09", "callback_data": "sto_ds09"}
                ],
                [
                    {"text": "🏢 UP", "callback_data": "sto_up"},
                    {"text": "🔙 Menu Utama", "callback_data": "main_menu"}
                ]
            ]
        }
        return keyboard

    def process_message(self, message):
        """Memproses pesan yang diterima"""
        try:
            chat_id = message['chat']['id']
            text = message.get('text', '').strip()
            # Deteksi command dashboard
            if text.lower() in ['/dashboard', 'dashboard', '📊 dashboard']:
                self.show_dashboard_main(chat_id)
                return
            logger.info(f"📩 Received message from {chat_id}: {text}")
            
            user_state = self.user_search_state.get(chat_id)
            
            if text.startswith('/start'):
                self.reset_user_state(chat_id)
                self.show_main_menu(chat_id)
            
            elif text.startswith('/help'):
                self.reset_user_state(chat_id)
                self.show_help(chat_id)
            
            elif user_state == "waiting_for_order_id":
                
                loading_text = f"""
🔄 *MENCARI DATA...*
━━━━━━━━━━━━━━━━━━━━━━━━

📝 *Order ID:* `{text}`
⏳ Mohon tunggu sebentar...
"""
                self.send_message(chat_id, loading_text, 'Markdown')
                self.show_order_results(chat_id, text)
                
                # TETAPKAN KEMBALI STATE waiting_for_order_id
                # agar user bisa langsung ketik Order ID berikutnya
                self.user_search_state[chat_id] = "waiting_for_order_id"
            
            elif user_state == "waiting_for_date":
                
                if self.is_valid_date(text):
                    loading_text = f"""
🔄 *MENCARI DATA...*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Tanggal:* {text}
⏳ Mohon tunggu sebentar...
"""
                    self.send_message(chat_id, loading_text, 'Markdown')
                    self.show_date_results(chat_id, text, page=1)

                    # TETAPKAN KEMBALI STATE waiting_for_date
                    self.user_search_state[chat_id] = "waiting_for_date"
                else:
                    error_text = """
❌ *FORMAT TANGGAL SALAH*
━━━━━━━━━━━━━━━━━━━━━━━━

Format tanggal tidak sesuai!

📝 *Format yang benar:*
DD-MM-YYYY

💡 *Contoh:*
`01-12-2024`
`15-12-2024`
"""
                    self.send_message(chat_id, error_text, 'Markdown')
                    # Tetap dalam state waiting_for_date
                    self.user_search_state[chat_id] = "waiting_for_date"
            else:
                # Auto-detect input type
                if self.is_valid_date(text):
                    # Set mode ke waiting_for_date
                    self.user_search_state[chat_id] = "waiting_for_date"
                    loading_text = f"""
🔄 *MENCARI DATA...*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Tanggal:* {text}
⏳ Mohon tunggu sebentar...
"""
                    self.send_message(chat_id, loading_text, 'Markdown')
                    self.show_date_results(chat_id, text, page=1)
                
                elif self.is_likely_order_id(text):
                    # Set mode ke waiting_for_order_id
                    self.user_search_state[chat_id] = "waiting_for_order_id"
                    loading_text = f"""
🔄 *MENCARI DATA...*
━━━━━━━━━━━━━━━━━━━━━━━━

📝 *Order ID:* `{text}`
⏳ Mohon tunggu sebentar...
"""
                    self.send_message(chat_id, loading_text, 'Markdown')
                    self.show_order_results(chat_id, text)
                
                elif len(text) <= 10 and text.upper() in ['DS01', 'DS02', 'DS03', 'DS07', 'DS08', 'DS09', 'UP']:
                    self.search_and_show_sto(chat_id, text.upper(), page=1)
                
                else:
                    self.show_main_menu(chat_id)
        
        except Exception as e:
            logger.error(f"❌ Error in process_message: {e}")
            error_text = """
⚠️ *TERJADI KESALAHAN*
━━━━━━━━━━━━━━━━━━━━━━━━

Maaf, terjadi kesalahan sistem.

🔄 Silakan coba lagi beberapa saat.
"""
            self.send_message(chat_id, error_text, 'Markdown', self.create_back_menu())

    def is_valid_date(self, text):
        """Validasi format tanggal DD-MM-YYYY"""
        date_pattern = r'^\d{2}-\d{2}-\d{4}$'
        if re.match(date_pattern, text):
            try:
                datetime.strptime(text, '%d-%m-%Y')
                return True
            except ValueError:
                return False
        return False

    def is_likely_order_id(self, text):
        """Mendeteksi apakah teks kemungkinan adalah Order ID"""
        text = str(text).strip().upper()
        return text.startswith('SC') and len(text) > 2

    def show_order_results(self, chat_id, order_id):
        """Menampilkan hasil pencarian order ID"""
        results = self.order_checker.search_by_order_id(order_id)
        
        if results:
            header = f"""
✅ *DATA DITEMUKAN*
━━━━━━━━━━━━━━━━━━━━━━━━

📝 *Order ID:* `{order_id}`
📊 *Total Data:* {len(results)} hasil

"""
            response = header
            for i, item in enumerate(results, 1):
                response += f"┌─── *DATA {i}* ───────────────────────────\n"
                response += f"│ 📝 *Order ID:* `{item.get('SC', 'N/A')}`\n"
                response += f"│ 🏢 *STO:* {item.get('STO', 'N/A')}\n"
                response += f"│ 📊 *Status:* {item.get('STATUS', 'N/A')}\n"
                response += f"│ 👨‍💼 *Teknisi:* {item.get('TEKNISI', 'N/A')}\n"
                response += f"│ 📅 *Tanggal:* {item.get('ORDER DATE', 'N/A')}\n"
                # PROSES KETERANGAN KHUSUS
                keterangan = item.get('KETERANGAN', '')
                if keterangan:
                    ket_str = str(keterangan).strip()
                    
                    # Pisahkan berdasarkan koma TAPI jangan pisah angka desimal
                    parts = []
                    current_part = ""
                    inside_coord = False
                    
                    for char in ket_str:
                        if char == ',':
                            # Cek apakah koma ini bagian dari koordinat
                            if (current_part and current_part[-1].isdigit() and 
                                len(current_part) > 2 and '.' in current_part[-5:]):
                                # Kemungkinan bagian dari koordinat, lanjutkan
                                current_part += char
                                inside_coord = True
                            else:
                                # Koma sebagai pemisah
                                if current_part.strip():
                                    parts.append(current_part.strip())
                                current_part = ""
                                inside_coord = False
                        else:
                            current_part += char
                    
                    # Tambahkan bagian terakhir
                    if current_part.strip():
                        parts.append(current_part.strip())
                    
                    # Jika parsing kompleks gagal, gunakan split biasa
                    if not parts:
                        parts = [p.strip() for p in ket_str.split(',') if p.strip()]
                    
                    # Tampilkan hasil
                    if len(parts) == 1:
                        response += f"│ 📋 *Ket:* {parts[0]}\n"
                    elif len(parts) > 1:
                        response += f"│ 📋 *Keterangan:*\n"
                        for part in parts:
                            response += f"│   • {part}\n"
                response += f"└──────────────────────────────────────────\n\n"
                
                # Tambahkan instruksi untuk pencarian berikutnya
                if i == len(results):
                    response += """
💡 *Tips:*
• Ketik Order ID berikutnya untuk mencari lagi
• Tekan /start untuk kembali ke menu utama
"""
        else:
            response = f"""
❌ *DATA TIDAK DITEMUKAN*
━━━━━━━━━━━━━━━━━━━━━━━━

📝 *Order ID:* `{order_id}`

Tidak ditemukan order dengan ID tersebut.

💡 *Saran:*
• Periksa penulisan Order ID
• Pastikan Order ID diawali dengan 'SC'
• Coba dengan Order ID yang berbeda

📝 *Ketik Order ID berikutnya atau /start untuk menu*
"""
        
        # Kirim tanpa tombol khusus, karena user bisa langsung ketik
        self.send_message(chat_id, response, 'Markdown')

    def show_date_results(self, chat_id, date_str, page=1):
        """Menampilkan hasil pencarian tanggal dengan pagination"""
        results = self.order_checker.search_by_date(date_str)
        
        if results:
            page_size = 5
            total_pages = (len(results) + page_size - 1) // page_size
            
            header = f"""
✅ *DATA DITEMUKAN*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Tanggal:* {date_str}
📊 *Total Data:* {len(results)} tiket
📄 *Halaman:* {page} dari {total_pages}

"""
            formatted_results, page_results = self.format_results_page(results, page, page_size)
            response = header + formatted_results
            
            reply_markup = self.create_pagination_menu(page, total_pages, "date", date_str)
            self.send_message(chat_id, response, 'Markdown', reply_markup)
        else:
            response = f"""
❌ *DATA TIDAK DITEMUKAN*
━━━━━━━━━━━━━━━━━━━━━━━━

📅 *Tanggal:* {date_str}

Tidak ditemukan tiket untuk tanggal ini.

💡 *Saran:*
• Periksa format tanggal (DD-MM-YYYY)
• Pastikan data tersedia untuk tanggal tersebut
"""
            self.send_message(chat_id, response, 'Markdown', self.create_back_menu())

    def format_results_page(self, results, page, page_size=5):
        """Memformat hasil untuk halaman tertentu"""
        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size
        page_results = results[start_idx:end_idx]
        
        response = "📋 *HASIL PENCARIAN*\n\n"
        for i, item in enumerate(page_results, start_idx + 1):
            response += f"┌─── *DATA {i}* ──────────────────────────────\n"
            response += f"│ 📝 *Order ID:* `{item.get('SC', 'N/A')}`\n"
            response += f"│ 🏢 *STO:* {item.get('STO', 'N/A')}\n"
            response += f"│ 📊 *Status:* {item.get('STATUS', 'N/A')}\n"
            response += f"│ 👨‍💼 *Teknisi:* {item.get('TEKNISI', 'N/A')}\n"
            response += f"│ 📅 *Tanggal:* {item.get('ORDER DATE', 'N/A')}\n"
            response += f"└─────────────────────────────────────────────\n"
        
        return response, page_results

    def create_pagination_menu(self, current_page, total_pages, search_type, search_query):
        # Buat query aman
        safe_query = str(search_query).replace(" ", "_")

        keyboard_buttons = []

        if current_page > 1:
            keyboard_buttons.append({
                "text": "⬅️ Sebelumnya",
                "callback_data": f"page_{search_type}_{safe_query}_{current_page-1}"
            })

        if current_page < total_pages:
            keyboard_buttons.append({
                "text": "Selanjutnya ➡️",
                "callback_data": f"page_{search_type}_{safe_query}_{current_page+1}"
            })

        keyboard_buttons.append({
            "text": "🏠 Menu Utama",
            "callback_data": "main_menu"
        })

        keyboard = {"inline_keyboard": [keyboard_buttons]}
        return keyboard

    def search_and_show_sto(self, chat_id, sto_name, page=1):
        """Mencari dan menampilkan data STO dengan pagination"""
        loading_text = f"""
🔄 *MENCARI DATA...*
━━━━━━━━━━━━━━━━━━━━━━━━

🏢 *STO:* {sto_name}
⏳ Mohon tunggu sebentar...
"""
        self.send_message(chat_id, loading_text, 'Markdown')
        
        results = self.order_checker.search_by_sto(sto_name)
        
        if results:
            page_size = 5
            total_pages = (len(results) + page_size - 1) // page_size
            
            header = f"""
✅ *DATA DITEMUKAN*
━━━━━━━━━━━━━━━━━━━━━━━━

🏢 *STO:* {sto_name}
📊 *Total Data:* {len(results)} tiket
📄 *Halaman:* {page} dari {total_pages}

"""
            formatted_results, page_results = self.format_results_page(results, page, page_size)
            response = header + formatted_results
            
            # Tambahkan instruksi
            response += f"""
💡 *Tips:*
• Klik tombol di bawah untuk navigasi halaman
• Ketik STO lain (DS01, DS02, dll) untuk mencari lagi
• Tekan `/start` untuk kembali ke menu utama
"""
            
            reply_markup = self.create_pagination_menu(page, total_pages, "sto", sto_name)
            self.send_message(chat_id, response, 'Markdown', reply_markup)
        else:
            response = f"""
❌ *DATA TIDAK DITEMUKAN*
━━━━━━━━━━━━━━━━━━━━━━━━

🏢 *STO:* {sto_name}

Tidak ditemukan tiket untuk STO ini.

💡 *Saran:*
• Coba STO lainnya (DS01, DS02, UP, dll)
• Periksa penulisan nama STO
• Ketik STO lain untuk mencoba lagi
"""
            self.send_message(chat_id, response, 'Markdown', self.create_sto_menu())

    def show_help(self, chat_id, message_id=None):
        """Menampilkan bantuan"""
        help_text = """
ℹ️ *BANTUAN ORDER CHECKER BOT*
━━━━━━━━━━━━━━━━━━━━━━━━

📖 *CARA PENGGUNAAN:*
1️⃣ Pilih jenis pencarian di menu utama
2️⃣ Ketikkan data yang ingin dicari
3️⃣ Lihat hasil pencarian

🎯 *JENIS PENCARIAN:*
• 🔍 *Order ID* - Cari berdasarkan ID Order
• 📅 *Tanggal* - Cari berdasarkan tanggal
• 🏢 *STO* - Cari berdasarkan nama STO
• 📊 *Laporan* - Lihat statistik bulanan
• 🏢 *Pivot STO* - Analisis PS/RE per STO

📝 *FORMAT INPUT:*
• Order ID: `SC123456789`
• Tanggal: `DD-MM-YYYY`
• STO: `DS01`, `UP`, dll

💡 *FITUR BARU - PIVOT STO:*
• 🏢 Detail setiap STO
• ✅ PS Count (Proses/Progress)
• 📈 RE Count (Total Status)
• 🎯 RE Rate (PS/RE)
• 📋 Daftar SC per STO
"""
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': help_text,
                'parse_mode': 'Markdown',
                'reply_markup': self.create_back_menu()
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, help_text, 'Markdown', self.create_back_menu())

    def create_back_menu(self):
        """Membuat menu kembali"""
        keyboard = {
            "inline_keyboard": [
                [
                    {"text": "↩️ Kembali", "callback_data": "main_menu"}
                ]
            ]
        }
        return keyboard

    def show_main_menu(self, chat_id, message_id=None):
        """Menampilkan menu utama"""
        # Reset state user saat kembali ke menu utama
        self.reset_user_state(chat_id)
        
        welcome_text = """
✨ *ORDER CHECKER BOT* ✨
━━━━━━━━━━━━━━━━━━━━━━━━

🤖 *Digital Assistant untuk Cek Status Order*

Pilih jenis pencarian di bawah ini:
"""
        if message_id:
            url = self.base_url + "editMessageText"
            payload = {
                'chat_id': chat_id,
                'message_id': message_id,
                'text': welcome_text,
                'parse_mode': 'Markdown',
                'reply_markup': self.create_main_menu()
            }
            requests.post(url, json=payload)
        else:
            self.send_message(chat_id, welcome_text, 'Markdown', self.create_main_menu())

    def run(self):
        """Menjalankan bot"""
        print("✨ Order Checker Bot sedang berjalan...")
        print("🤖 Bot siap menerima perintah")
        print("🎯 Fitur: Cari Order ID, Tanggal, STO, Laporan Bulanan + Detail SC")
        print("🏢 FITUR BARU: Pivot per STO - PS/RE Analysis!")
        print("⏹️ Tekan Ctrl+C untuk menghentikan")
        print("━" * 50)
        
        if not self.test_bot_connection():
            print("❌ Gagal terhubung ke Telegram Bot API")
            return
        
        offset = None
        
        try:
            while True:
                logger.info("🔄 Checking for updates...")
                updates = self.get_updates(offset)
                
                if updates and 'result' in updates:
                    for update in updates['result']:
                        offset = update['update_id'] + 1
                        
                        if 'message' in update:
                            self.process_message(update['message'])
                        elif 'callback_query' in update:
                            # ✅ Gunakan handle_callback dengan mengirim dictionary
                            self.handle_callback(update['callback_query'])
                    
                    time.sleep(1)
                else:
                    time.sleep(5)
                
        except KeyboardInterrupt:
            print("\n🛑 Bot dihentikan oleh user")
        except Exception as e:
            logger.error(f"❌ Error in main loop: {e}")
            print(f"\n❌ Error: {e}")

    def test_bot_connection(self):
        """Test koneksi bot"""
        url = self.base_url + "getMe"
        try:
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if data.get('ok'):
                    bot_info = data['result']
                    logger.info(f"✅ Bot connected: {bot_info['first_name']} (@{bot_info['username']})")
                    return True
                else:
                    logger.error(f"❌ Bot connection failed: {data}")
                    return False
            else:
                logger.error(f"❌ HTTP Error: {response.status_code}")
                return False
        except Exception as e:
            logger.error(f"❌ Connection test failed: {e}")
            return False

    def get_updates(self, offset=None):
        """Mendapatkan update dari Telegram"""
        url = self.base_url + "getUpdates"
        params = {'timeout': 10, 'offset': offset}
        
        try:
            response = requests.get(url, params=params, timeout=15)
            if response.status_code == 200:
                return response.json()
            else:
                logger.error(f"❌ API Error: {response.status_code} - {response.text}")
                return None
        except Exception as e:
            logger.error(f"❌ Error getting updates: {e}")
            return None

    def reset_user_state(self, chat_id):
        """Reset state user"""
        if chat_id in self.user_search_state:
            del self.user_search_state[chat_id]

def main():
    print("🚀 Starting Order Checker Bot...")
    bot = TelegramBot()
    bot.run()

if __name__ == '__main__':
    main()