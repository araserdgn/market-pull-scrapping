import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import threading
import pandas as pd
import webbrowser
import re
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime

# Koyu tema renkleri
DARK_BG = "#1e1e1e"        # Ana arka plan
DARK_SECONDARY = "#2d2d2d"  # İkincil arka plan
DARK_TEXT = "#ffffff"       # Ana metin rengi
ACCENT_COLOR = "#007acc"    # Vurgu rengi
BUTTON_BG = "#3c3c3c"       # Buton arka planı
BUTTON_HOVER = "#505050"    # Buton hover rengi
ENTRY_BG = "#3c3c3c"        # Giriş alanı arka planı
TABLE_BG = "#2d2d2d"        # Tablo arka planı
TABLE_SELECT = "#264f78"    # Tablo seçim rengi

class GoogleMapsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Haritalardan Pazar Araştırma Botu")
        self.root.geometry("1080x720")
        self.root.configure(bg=DARK_BG)

        # Üst frame
        self.frame_top = tk.Frame(self.root, bg=DARK_BG, bd=10, relief=tk.FLAT)
        self.frame_top.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.3)

        # Arama alanı
        self.label_search = tk.Label(self.frame_top, text="Aramak İstediğiniz Kelime:", 
                                   bg=DARK_BG, fg=DARK_TEXT, font=("Helvetica", 10))
        self.label_search.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        self.entry_search = tk.Entry(self.frame_top, font=("Helvetica", 10), 
                                   bd=5, relief=tk.FLAT, fg=DARK_TEXT, bg=ENTRY_BG)
        self.entry_search.grid(row=0, column=1, padx=5, pady=5)
        
        # Sayı girişi
        self.label_count = tk.Label(self.frame_top, text="Çekilecek İşletme Sayısı:", 
                                  bg=DARK_BG, fg=DARK_TEXT, font=("Helvetica", 10))
        self.label_count.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        self.entry_count = tk.Entry(self.frame_top, font=("Helvetica", 10), 
                                  bd=5, relief=tk.FLAT, fg=DARK_TEXT, bg=ENTRY_BG)
        self.entry_count.grid(row=1, column=1, padx=5, pady=5)

        # Buton frame'i
        self.button_frame = tk.Frame(self.frame_top, bg=DARK_BG)
        self.button_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky="w")

        # Butonlar
        self.button_start = tk.Button(self.button_frame, text="Verileri Çek", 
                                    command=self.start_scraping_thread, 
                                    font=("Helvetica", 10), bg=BUTTON_BG, 
                                    fg=DARK_TEXT, relief=tk.RAISED, padx=5, pady=5)
        self.button_start.pack(side=tk.LEFT, padx=5)

        self.button_export = tk.Button(self.button_frame, text="Excel'e Aktar", 
                                     command=self.export_to_excel, 
                                     font=("Helvetica", 10), bg=BUTTON_BG, 
                                     fg=DARK_TEXT, relief=tk.RAISED, padx=5, pady=5)
        self.button_export.pack(side=tk.LEFT, padx=5)

        self.button_analytics = tk.Button(self.button_frame, text="Analiz Raporu", 
                                        command=self.show_analytics, 
                                        font=("Helvetica", 10), bg=BUTTON_BG, 
                                        fg=DARK_TEXT, relief=tk.RAISED, padx=5, pady=5)
        self.button_analytics.pack(side=tk.LEFT, padx=5)

        # Filtreleme özellikleri
        self.add_search_filters()

        # Anlık arama
        self.add_live_search()

        # Tablo frame'i
        self.frame_bottom = tk.Frame(self.root, bg=DARK_SECONDARY, bd=10, relief=tk.FLAT)
        self.frame_bottom.place(relx=0.02, rely=0.35, relwidth=0.96, relheight=0.6)

        # Tablo oluşturma
        self.create_table()

        # Stil ayarları
        self.configure_styles()

    def add_search_filters(self):
        # Filtre frame'i
        self.filter_frame = tk.Frame(self.frame_top, bg=DARK_BG)
        self.filter_frame.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Şehir filtresi
        self.city_var = tk.StringVar()
        tk.Label(self.filter_frame, text="Şehir:", 
                bg=DARK_BG, fg=DARK_TEXT).pack(side=tk.LEFT, padx=5)
        self.city_filter = ttk.Combobox(self.filter_frame, textvariable=self.city_var, 
                                       values=["Tümü", "İstanbul", "Ankara", "İzmir"])
        self.city_filter.pack(side=tk.LEFT, padx=5)
        self.city_filter.set("Tümü")
        
        # Durum filtresi
        self.status_var = tk.StringVar()
        tk.Label(self.filter_frame, text="Mesaj Durumu:", 
                bg=DARK_BG, fg=DARK_TEXT).pack(side=tk.LEFT, padx=5)
        self.status_filter = ttk.Combobox(self.filter_frame, textvariable=self.status_var,
                                         values=["Tümü", "Gönderildi", "Gönderilmedi"])
        self.status_filter.pack(side=tk.LEFT, padx=5)
        self.status_filter.set("Tümü")
        
        # Filtreleme butonu
        self.filter_btn = tk.Button(self.filter_frame, text="Filtrele", 
                                  command=self.apply_filters,
                                  bg=BUTTON_BG, fg=DARK_TEXT)
        self.filter_btn.pack(side=tk.LEFT, padx=5)

    def add_live_search(self):
        # Anlık arama frame'i
        self.search_frame = tk.Frame(self.frame_top, bg=DARK_BG)
        self.search_frame.grid(row=4, column=0, columnspan=2, pady=5)
        
        # Anlık arama etiketi ve girişi
        tk.Label(self.search_frame, text="Anlık Arama:", 
                bg=DARK_BG, fg=DARK_TEXT).pack(side=tk.LEFT, padx=5)
        self.live_search_var = tk.StringVar()
        self.live_search_var.trace('w', self.on_live_search)
        self.live_search_entry = tk.Entry(self.search_frame, 
                                        textvariable=self.live_search_var,
                                        bg=ENTRY_BG, fg=DARK_TEXT)
        self.live_search_entry.pack(side=tk.LEFT, padx=5)

    def create_table(self):
        # Tablo oluşturma
        self.tree = ttk.Treeview(self.frame_bottom, 
                                columns=("İşletme Adı", "Adres", "İletişim No", 
                                        "Mesaj Atıldı Mı?", "Mesaj Gönder"),
                                show='headings', height=15)
        
        # Sütun başlıkları ve genişlikleri
        columns = {
            "İşletme Adı": 200,
            "Adres": 300,
            "İletişim No": 150,
            "Mesaj Atıldı Mı?": 120,
            "Mesaj Gönder": 120
        }
        
        # Sütunları yapılandır
        for col, width in columns.items():
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor='center')
        
        # Scrollbar
        self.scrollbar = ttk.Scrollbar(self.frame_bottom, orient="vertical", 
                                     command=self.tree.yview)
        self.tree.configure(yscroll=self.scrollbar.set)
        
        # Yerleşim
        self.scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Tablo tıklama olayı
        self.tree.bind("<Button-1>", self.on_tree_click)
        
        # Alternatif satır renkleri
        self.tree.tag_configure('oddrow', background=DARK_SECONDARY)
        self.tree.tag_configure('evenrow', background=TABLE_BG)

    def configure_styles(self):
        # Ttk stilleri
        style = ttk.Style()
        
        # Treeview stili
        style.configure("Treeview",
                    background=TABLE_BG,
                    foreground=DARK_TEXT,
                    fieldbackground=TABLE_BG,
                    borderwidth=0)
        
        # Başlık stili - Metin rengini siyah yapıyoruz
        style.configure("Treeview.Heading",
                    background=DARK_SECONDARY,
                    foreground='black',  # Başlık yazı rengini siyah yaptık
                    relief="flat",
                    font=('Helvetica', 10, 'bold'))
        
        style.map("Treeview",
                background=[('selected', TABLE_SELECT)],
                foreground=[('selected', DARK_TEXT)])
        
        # Combobox stili
        style.configure("TCombobox",
                    background=ENTRY_BG,
                    foreground=DARK_TEXT,
                    fieldbackground=ENTRY_BG,
                    arrowcolor=DARK_TEXT)
        
        # Combobox açılır menü rengi
        self.root.option_add('*TCombobox*Listbox.background', DARK_SECONDARY)
        self.root.option_add('*TCombobox*Listbox.foreground', DARK_TEXT)    

    def apply_filters(self):
        # Tüm öğeleri göster
        for item in self.tree.get_children():
            self.tree.item(item, tags=())
        
        selected_city = self.city_var.get()
        selected_status = self.status_var.get()
        
        # Sarı renk için tag tanımla
        self.tree.tag_configure('filtered', background='#FFD700')  # Altın sarısı
        
        # Tüm öğeleri kontrol et
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            if not values:  # Değerler boşsa atla
                continue
                
            address = str(values[1]) if len(values) > 1 else ""
            status = str(values[3]) if len(values) > 3 else ""
            
            show_item = True
            is_filtered = False
            
            # Şehir filtresi
            if selected_city != "Tümü" and address:
                city_found = False
                for city in address.split():
                    if city.lower() == selected_city.lower():
                        city_found = True
                        is_filtered = True  # Filtrelenmiş öğe
                        break
                if not city_found:
                    show_item = False
            
            # Durum filtresi
            if selected_status != "Tümü":
                if selected_status == "Gönderildi" and status == "Evet":
                    is_filtered = True  # Filtrelenmiş öğe
                elif selected_status == "Gönderilmedi" and status == "Hayır":
                    is_filtered = True  # Filtrelenmiş öğe
                else:
                    show_item = False
            
            # Öğeyi göster/gizle ve renklendir
            if not show_item:
                self.tree.item(item, tags=('hidden',))
            elif is_filtered:
                self.tree.item(item, tags=('filtered',))  # Filtrelenmiş öğeleri sarı yap
        
        # Tag stilleri
        self.tree.tag_configure('hidden', background=DARK_BG)
        self.tree.tag_configure('filtered', background='#FFD700')  # Altın sarısı

    def on_live_search(self, *args):
        search_term = self.live_search_var.get().lower()
        
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            found = False
            
            # Tüm değerlerde arama yap
            for value in values:
                if str(value).lower().find(search_term) != -1:
                    found = True
                    break
            
            # Bulunan/bulunmayan öğeleri işaretle
            if not found:
                self.tree.item(item, tags=('hidden',))
            else:
                self.tree.item(item, tags=())
        
        self.tree.tag_configure('hidden', background=DARK_BG)

    def start_scraping_thread(self):
        thread = threading.Thread(target=self.scrape_data)
        thread.start()

    def scrape_data(self):
        search_query = self.entry_search.get()
        try:
            count = int(self.entry_count.get())
        except ValueError:
            messagebox.showerror("Hata", "Lütfen geçerli bir sayı giriniz!")
            return

        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--window-size=1280,1024")
            driver = webdriver.Chrome(options=options)
        except Exception as e:
            messagebox.showerror("Hata", f"Chrome sürücüsü başlatılamadı: {str(e)}")
            return

        try:
            driver.get("https://www.google.com.tr/maps/")
            
            search_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "searchboxinput"))
            )
            
            search_box.send_keys(search_query)
            search_box.send_keys(Keys.ENTER)
            
            time.sleep(4)
            
            businesses_processed = 0
            last_height = driver.execute_script("return document.body.scrollHeight")
            
            while businesses_processed < count:
                try:
                    businesses = WebDriverWait(driver, 10).until(
                        EC.presence_of_all_elements_located((By.CLASS_NAME, "Nv2PK"))
                    )
                    
                    for business in businesses[businesses_processed:]:
                        if businesses_processed >= count:
                            break
                            
                        driver.execute_script("arguments[0].scrollIntoView(true);", business)
                        time.sleep(1)
                        business.click()
                        time.sleep(2)
                        
                        try:
                            business_name = driver.find_element(By.CLASS_NAME, "DUwDvf").text
                        except:
                            business_name = "Bilgi bulunamadı"
                        
                        try:
                            address = driver.find_element(By.CSS_SELECTOR, 
                                "button[data-item-id='address'] .Io6YTe").text
                        except:
                            address = "Bilgi bulunamadı"
                        
                        try:
                            phone_element = driver.find_element(By.CSS_SELECTOR, 
                                "button[data-item-id^='phone'] .Io6YTe")
                            phone_number = phone_element.text
                            phone_number = re.sub(r'\D', '', phone_number)
                            if phone_number.startswith('0'):
                                phone_number = phone_number[1:]
                            if len(phone_number) == 10:
                                phone_number = f'+90{phone_number}'
                            else:
                                phone_number = "Bilgi bulunamadı"
                        except:
                            phone_number = "Bilgi bulunamadı"
                        
                        self.tree.insert("", "end", values=(
                            business_name,
                            address,
                            phone_number,
                            "Hayır",
                            "Mesaj Gönder"
                        ))
                        
                        businesses_processed += 1
                        self.root.update()
                    
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2)
                    
                    new_height = driver.execute_script("return document.body.scrollHeight")
                    if new_height == last_height:
                        break
                    last_height = new_height
                    
                except Exception as e:
                    print(f"Veri çekme hatası: {str(e)}")
                    break
                    
        except Exception as e:
            messagebox.showerror("Hata", f"Veri çekme işlemi başarısız: {str(e)}")
            
        finally:
            try:
                driver.quit()
            except:
                pass
            messagebox.showinfo("Bilgi", f"Toplam {businesses_processed} işletme bilgisi çekildi.")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            data = []
            for item in self.tree.get_children():
                values = self.tree.item(item)["values"]
                data.append(values)
            
            df = pd.DataFrame(data, columns=["İşletme Adı", "Adres", "İletişim No", 
                                           "Mesaj Atıldı Mı?", "Mesaj Gönder"])
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Başarılı", "Veriler Excel dosyasına aktarıldı.")

    def on_tree_click(self, event):
        region = self.tree.identify('region', event.x, event.y)
        if region == 'cell':
            column = self.tree.identify_column(event.x)
            if column == '#5':  # 'Mesaj Gönder' sütunu
                selected_item = self.tree.identify_row(event.y)
                if selected_item:
                    item = self.tree.item(selected_item)
                    values = item["values"]
                    phone_number = values[2]
                    if phone_number != "Bilgi bulunamadı":
                        try:
                            webbrowser.open(f"https://wa.me/{phone_number}")
                            self.tree.set(selected_item, "Mesaj Atıldı Mı?", "Evet")
                        except Exception as e:
                            messagebox.showerror("Hata", f"Mesaj gönderilemedi: {str(e)}")

    def show_analytics(self):
    # Analiz penceresi oluşturma
        analytics_window = tk.Toplevel(self.root)
        analytics_window.title("Veri Analizi Raporu")
        analytics_window.geometry("800x600")
        analytics_window.configure(bg=DARK_BG)

        # Verileri toplama
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        
        df = pd.DataFrame(data, columns=["İşletme Adı", "Adres", "İletişim No", 
                                    "Mesaj Atıldı Mı?", "Mesaj Gönder"])

        # İstatistik frame'i
        stats_frame = tk.Frame(analytics_window, bg=DARK_BG, pady=10)
        stats_frame.pack(fill="x", padx=20)

        stats = {
            "Toplam İşletme": len(df),
            "Mesaj Gönderilen": len(df[df["Mesaj Atıldı Mı?"] == "Evet"]),
            "Mesaj Gönderilmeyen": len(df[df["Mesaj Atıldı Mı?"] == "Hayır"]),
            "Telefon Numarası Olan": len(df[df["İletişim No"] != "Bilgi bulunamadı"])
        }

        for idx, (key, value) in enumerate(stats.items()):
            tk.Label(stats_frame, 
                    text=f"{key}: {value}",
                    bg=DARK_BG,
                    fg=DARK_TEXT,
                    font=("Helvetica", 10, "bold")).grid(row=idx//2, 
                                                        column=idx%2, 
                                                        padx=10, 
                                                        pady=5,
                                                        sticky="w")

        # Grafik frame'i
        graph_frame = tk.Frame(analytics_window, bg=DARK_BG)
        graph_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Grafik stili
        plt.style.use('dark_background')
        fig = plt.Figure(figsize=(10, 6), dpi=100)
        fig.patch.set_facecolor(DARK_BG)

        # Mesaj durumu pasta grafiği
        ax1 = fig.add_subplot(121)
        message_stats = df["Mesaj Atıldı Mı?"].value_counts()
        ax1.pie(message_stats.values, 
                labels=message_stats.index,
                autopct='%1.1f%%',
                colors=[ACCENT_COLOR, '#e74c3c'])
        ax1.set_title("Mesaj Gönderim Durumu", color=DARK_TEXT)

        # Bölgesel dağılım grafiği
        ax2 = fig.add_subplot(122)
        df['Bölge'] = df['Adres'].apply(lambda x: x.split()[0] if x != "Bilgi bulunamadı" else "Bilinmiyor")
        region_stats = df['Bölge'].value_counts().head(5)
        bars = ax2.bar(region_stats.index, region_stats.values, color=ACCENT_COLOR)
        ax2.set_title("Bölgesel Dağılım (İlk 5)", color=DARK_TEXT)
        ax2.tick_params(colors=DARK_TEXT)
        plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45)

        # Grafikleri canvas'a yerleştir
        canvas = FigureCanvasTkAgg(fig, master=graph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # Excel'e aktarma butonu
    def export_analysis():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"analiz_raporu_{timestamp}.xlsx"
        
        with pd.ExcelWriter(filename) as writer:
            df.to_excel(writer, sheet_name='Tüm Veriler', index=False)
            
            stats_df = pd.DataFrame(list(stats.items()), 
                                columns=['Metrik', 'Değer'])
            stats_df.to_excel(writer, sheet_name='İstatistikler', index=False)
            
            region_stats.to_frame('İşletme Sayısı').to_excel(writer, 
                                                            sheet_name='Bölgesel Dağılım')

        messagebox.showinfo("Başarılı", f"Analiz raporu {filename} dosyasına kaydedildi.")

        export_btn = tk.Button(analytics_window,
                            text="Raporu Excel'e Aktar",
                            command=export_analysis,
                            bg=BUTTON_BG,
                            fg=DARK_TEXT,
                            font=("Helvetica", 10),
                            pady=5)
        export_btn.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = GoogleMapsApp(root)
    root.mainloop()