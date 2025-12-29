import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
import subprocess
import pandas as pd
import glob
import re

# ==============================================================================
# 1. BÖLÜM: BACKEND MANTIK (HİÇ DEĞİŞTİRİLMEDİ)
# ==============================================================================

OE_HEADERS = [
    'OE0001', 'Stno', 'XStno', 'Chipno', 'Database Id', 'First name', 'Surname', 'YB', 'S', 
    'Block', 'nc', 'Start', 'Finish', 'Time', 'Classifier', 'Credit -', 'Penalty +', 'Comment', 
    'Club no.', 'Cl.name', 'City', 'Nat', 'Location', 'Region', 'Cl. no.', 'Short', 'Long', 
    'Entry cl. No', 'Entry class (short)', 'Entry class (long)', 'Rank', 'Ranking points', 
    'Num1', 'Num2', 'Num3', 'Text1', 'Text2', 'Text3', 'Addr. surname', 'Addr. first name', 
    'Street', 'Line2', 'Zip', 'Addr. city', 'Phone', 'Mobile', 'Fax', 'EMail', 'Rented', 
    'Start fee', 'Paid', 'Team', 'Course no.', 'Course', 'km', 'm', 'Course controls'
]

detector = None

def init_gender_guesser():
    global detector
    try:
        import gender_guesser.detector as gender
        detector = gender.Detector()
    except ImportError:
        print("gender_guesser modülü bulunamadı. Otomatik yükleniyor...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "gender-guesser"])
            import gender_guesser.detector as gender
            detector = gender.Detector()
            print("Kurulum başarılı!")
        except Exception as e:
            print(f"Kurulum hatası: {e}")
            sys.exit(1)

def clean_for_cp1254(text):
    if pd.isna(text) or str(text).strip() == "":
        return ""
    text = str(text)
    return text.encode('cp1254', errors='ignore').decode('cp1254')

def title_case_tr(text):
    if pd.isna(text) or str(text).strip() == "":
        return ""
    text = clean_for_cp1254(text)
    words = text.split()
    new_words = []
    for word in words:
        if not word: continue
        first = word[0]
        if first == 'i': first = 'İ'
        elif first == 'ı': first = 'I'
        else: first = first.upper()
        rest = word[1:]
        rest = rest.replace('I', 'ı').replace('İ', 'i').lower()
        new_words.append(first + rest)
    return " ".join(new_words)

def identify_column_type(header):
    header_lower = header.lower().strip()
    name_keywords = ['ad', 'isim', 'isminiz']
    surname_keywords = ['soyad', 'soyisim', 'soyisimi', 'soyismi']
    club_keywords = ['kulüp', 'takım', 'club', 'okul', 'grup', 'kurum', 'takim']
    category_keywords = ['kategori', 'class', 'tür', 'hangi', 'tip']
    sex_keywords = ['cinsiyet', 'sex', 'gender', 'k/e', 'erkek', 'kadın', 'kadin']
    rent_keywords = ['kira', 'rent', 'ödünç']
    phone_keywords = ['tel', 'gsm', 'mobil']
    email_keywords = ['mail', 'posta']

    if any(kw in header_lower for kw in rent_keywords): return 'Rented'
    if re.search(r'\b(si|chip|çip|ident|si-card)\b', header_lower): return 'Chipno'
    if any(kw in header_lower for kw in club_keywords): return 'Cl.name'
    if any(kw in header_lower for kw in sex_keywords): return 'S'
    if any(kw in header_lower for kw in category_keywords): return 'Category_Source'
    if any(kw in header_lower for kw in phone_keywords): return 'Mobile'
    if any(kw in header_lower for kw in email_keywords): return 'EMail'
    
    has_surname = any(kw in header_lower for kw in surname_keywords)

    # İsim kontrolü yapmadan önce, metnin içindeki soyad kelimelerini siliyoruz.
    # Böylece "Soyad" kelimesinin içindeki "ad" harflerini yanlışlıkla isim sanmıyor.
    temp_header_for_name = header_lower
    for kw in surname_keywords:
        temp_header_for_name = temp_header_for_name.replace(kw, "")

    has_name = any(kw in temp_header_for_name for kw in name_keywords)

    if has_name and has_surname: return 'Full_Name_Source'
    if has_name: return 'First name'
    if has_surname: return 'Surname'

    return None

def guess_gender_oe(name):
    global detector
    if detector is None: init_gender_guesser() 
    
    if pd.isna(name) or str(name).strip() == "": return 'M'
    first_word = str(name).strip().split(' ')[0].capitalize()
    g = detector.get_gender(first_word)
    if g in ['female', 'mostly_female']: return 'F'
    else: return 'M'

def convert_forms_to_oe2010(input_path, output_path, log_callback=None):
    def log(msg):
        if log_callback: log_callback(msg)
        else: print(msg)

    # 1. Dosyayı Oku
    try:
        if input_path.endswith('.csv'):
            try:
                df = pd.read_csv(input_path, encoding='utf-8')
            except:
                try:
                    df = pd.read_csv(input_path, encoding='cp1254')
                except:
                    df = pd.read_csv(input_path, encoding='iso-8859-9')
        else:
            df = pd.read_excel(input_path)
        log(f"Dosya okundu: {os.path.basename(input_path)}")
    except PermissionError:
        return "Hata: Dosya açık. Lütfen kapatıp tekrar deneyin."
    except Exception as e:
        return f"Dosya okuma hatası: {e}"

    # 2. Sütun Eşleşmeleri
    col_map = {}
    for col in df.columns:
        detected_type = identify_column_type(col)
        if detected_type and detected_type not in col_map:
            col_map[detected_type] = col

    # 3. Çıktı Hazırla
    output_df = pd.DataFrame(columns=OE_HEADERS)
    for col in OE_HEADERS: output_df[col] = ""

    # A) İSİMLER
    if 'Full_Name_Source' in col_map:
        source_col = col_map['Full_Name_Source']
        full_names = df[source_col].astype(str).str.strip()
        split_data = full_names.str.rsplit(n=1, expand=True)
        output_df['First name'] = split_data[0].apply(title_case_tr)
        if split_data.shape[1] > 1:
            output_df['Surname'] = split_data[1].apply(title_case_tr)
    elif 'First name' in col_map:
        output_df['First name'] = df[col_map['First name']].apply(title_case_tr)
        if 'Surname' in col_map:
            output_df['Surname'] = df[col_map['Surname']].apply(title_case_tr)

    # B) CİNSİYET
    if 'S' in col_map:
        output_df['S'] = df[col_map['S']]
    else:
        log("UYARI | Cinsiyet sütunu bulunamadı. İsimlere göre tahmin ediliyor...")
        log("UYARI | Tahminler hatalı olabilir, KONTROL EDİNİZ!")
        output_df['S'] = output_df['First name'].apply(guess_gender_oe)

    # C) CLASS
    if 'Category_Source' in col_map:
        raw_cats = df[col_map['Category_Source']].fillna("Unknown").astype(str).str.strip()
        class_names = []
        short_names = []
        for i in range(len(output_df)):
            cat = raw_cats.iloc[i].replace(" ", "")
            if len(cat) > 10: short_cat = cat[-11:]
            else: short_cat = cat
            
            sex = output_df['S'].iloc[i]
            class_names.append(f"{cat}{sex}")
            short_names.append(f"{short_cat}{sex}")
        
        output_df['Short'] = short_names
        output_df['Long'] = class_names
        output_df['Entry class (short)'] = short_names
        output_df['Entry class (long)'] = class_names

        unique_classes = sorted(list(set(class_names)))
        class_id_map = {c: i+1 for i, c in enumerate(unique_classes)}
        output_df['Cl. no.'] = output_df['Long'].map(class_id_map)
        output_df['Entry cl. No'] = output_df['Long'].map(class_id_map)
        log(f"{len(unique_classes)} farklı kategori oluşturuldu.")
    else:
        output_df['Short'] = "Unknown"
        output_df['Cl. no.'] = 0

    # D) KULÜP
    if 'Cl.name' in col_map: output_df['City'] = df[col_map['Cl.name']]
    else: output_df['City'] = "Ferdi"

    club_names = output_df['City'].fillna("Ferdi").astype(str).str.strip()
    unique_clubs = sorted(list(set(club_names)))
    club_id_map = {club: i+1 for i, club in enumerate(unique_clubs)}
    output_df['Club no.'] = club_names.map(club_id_map)

    # E) ÇİP
    if 'Chipno' in col_map:
        output_df['Chipno'] = df[col_map['Chipno']].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
        output_df['Chipno'] = output_df['Chipno'].apply(lambda x: x if x.isdigit() else '')

    # F) SIRA NO
    output_df['Stno'] = range(1, len(df) + 1)
    #output_df['City'] = "Ankara"

    # G) KİRALIK ÇİP
    if 'Rented' in col_map:
        output_df['Rented'] = df[col_map['Rented']].fillna('').astype(str).str.lower().apply(
            lambda x: 'X' if any(word in x for word in ['evet', 'istiyor']) else '')

    # H) TELEFON & EMAIL
    if 'Mobile' in col_map:
        output_df['Mobile'] = df[col_map['Mobile']].fillna('').astype(str).str.strip()
    if 'EMail' in col_map:
        output_df['EMail'] = df[col_map['EMail']].fillna('').astype(str).str.strip()

    # 4. KAYDET
    output_df.to_csv(output_path, sep=',', index=False, encoding='iso-8859-9')
    
    return f"BAŞARILI! Dosya hazır:\n{output_path}"

# ==============================================================================
# 2. BÖLÜM: ARAYÜZ (GUI) KODU
# ==============================================================================

class OEConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Forms -> OE2010 Dönüştürücü")
        self.root.geometry("600x450")
        self.root.resizable(False, False)

        # Stil
        style = ttk.Style()
        style.theme_use('clam')

        # Başlık
        header_frame = tk.Frame(root, bg="#2c3e50", height=60)
        header_frame.pack(fill="x")
        lbl_header = tk.Label(header_frame, text="OE2010 Kayıt Dönüştürücü", 
                              font=("Segoe UI", 16, "bold"), fg="white", bg="#2c3e50")
        lbl_header.pack(pady=15)

        # Dosya Seçim Alanı
        frame_input = tk.Frame(root, pady=20)
        frame_input.pack(fill="x", padx=20)

        self.lbl_file = tk.Label(frame_input, text="Lütfen Excel veya CSV dosyasını seçin:", font=("Segoe UI", 10))
        self.lbl_file.pack(anchor="w")

        input_inner_frame = tk.Frame(frame_input)
        input_inner_frame.pack(fill="x", pady=5)

        self.entry_path = tk.Entry(input_inner_frame, width=50, font=("Segoe UI", 10))
        self.entry_path.pack(side="left", fill="x", expand=True)

        btn_browse = tk.Button(input_inner_frame, text="Gözat...", command=self.browse_file,
                               bg="#3498db", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10)
        btn_browse.pack(side="right", padx=5)

        # Dönüştür Butonu
        self.btn_convert = tk.Button(root, text="DÖNÜŞTÜR VE KAYDET", command=self.start_conversion,
                                     bg="#27ae60", fg="white", font=("Segoe UI", 12, "bold"), 
                                     height=2, relief="flat", state="disabled")
        self.btn_convert.pack(fill="x", padx=40, pady=10)

        # Log Alanı
        lbl_log = tk.Label(root, text="İşlem Kaydı:", font=("Segoe UI", 9, "bold"))
        lbl_log.pack(anchor="w", padx=20)

        self.text_log = tk.Text(root, height=10, state="disabled", font=("Consolas", 9), bg="#ecf0f1")
        self.text_log.pack(fill="both", padx=20, pady=5, expand=True)

        # Footer
        lbl_footer = tk.Label(root, text="Oryantiring Veri İşleme Aracı v1.0", fg="#7f8c8d", font=("Segoe UI", 8))
        lbl_footer.pack(pady=5)

        self.input_file_path = ""

        threading.Thread(target=init_gender_guesser, daemon=True).start()

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Dosya Seç",
            filetypes=(("Excel Dosyaları", "*.xlsx"), ("CSV Dosyaları", "*.csv"), ("Tüm Dosyalar", "*.*"))
        )
        if filename:
            self.input_file_path = filename
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, filename)
            self.btn_convert.config(state="normal")
            self.log(f"Dosya seçildi: {os.path.basename(filename)}")

    def log(self, message):
        self.text_log.config(state="normal")
        self.text_log.insert(tk.END, ">> " + message + "\n")
        self.text_log.see(tk.END)
        self.text_log.config(state="disabled")

    def start_conversion(self):
        if not self.input_file_path:
            messagebox.showerror("Hata", "Lütfen bir dosya seçin.")
            return
        
        # --- EKLENEN KISIM: KAYIT YERİ SEÇME ---
        # Kullanıcıya nereye kaydedeceğini soruyoruz.
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Dosyası", "*.csv"), ("Tüm Dosyalar", "*.*")],
            title="Dönüştürülen Dosyayı Nereye Kaydedelim?",
            initialfile="OE2010_Import_Final.csv"
        )
        
        # Eğer kullanıcı "İptal" derse işlemi durduruyoruz.
        if not save_path:
            return
        # ---------------------------------------

        self.btn_convert.config(state="disabled", text="İşleniyor...")
        
        # Seçilen kayıt yolunu (save_path) thread'e gönderiyoruz
        threading.Thread(target=self.run_process, args=(save_path,)).start()

    def run_process(self, output_path): # Argüman olarak output_path eklendi
        input_path = self.input_file_path
        
        # Eski otomatik dizin bulma kodlarını sildik, çünkü artık parametre olarak geliyor
        # directory = os.path.dirname(input_path)
        # output_filename = "OE2010_Import_Final.csv"
        # output_path = os.path.join(directory, output_filename)

        try:
            result = convert_forms_to_oe2010(input_path, output_path, log_callback=self.update_log_from_thread)
            
            self.root.after(0, lambda: self.finish_process(result))
            
        except Exception as e:
            self.root.after(0, lambda: self.show_error(str(e)))

    def update_log_from_thread(self, msg):
        self.root.after(0, lambda: self.log(msg))

    def finish_process(self, result_msg):
        self.log("-" * 30)
        if "Hata" in result_msg:
            messagebox.showerror("İşlem Başarısız", result_msg)
            self.log("HATA OLUŞTU.")
        else:
            self.log("İŞLEM TAMAMLANDI.")
            messagebox.showinfo("Başarılı", result_msg)
        
        self.btn_convert.config(state="normal", text="DÖNÜŞTÜR VE KAYDET")

    def show_error(self, error_msg):
        messagebox.showerror("Beklenmeyen Hata", error_msg)
        self.log(f"KRİTİK HATA: {error_msg}")
        self.btn_convert.config(state="normal", text="DÖNÜŞTÜR VE KAYDET")

if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    root = tk.Tk()
    app = OEConverterApp(root)
    root.mainloop()