import tkinter as tk
from tkinter import ttk, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import threading
import pandas as pd

# Siyah beyaz stilinde tkinter UI oluşturma
class GoogleMapsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Haritalar Arama Botu")
        self.root.geometry("640x680")
        self.root.configure(bg="#FFFFFF")

        # Üst tarafta arama kutusu ve veri çekme seçenekleri
        self.frame_top = tk.Frame(self.root, bg="#FFFFFF", bd=10, relief=tk.FLAT)
        self.frame_top.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.2)

        self.label_search = tk.Label(self.frame_top, text="Aramak İstediğiniz Kelime:", bg="#FFFFFF", font=("Helvetica", 10))
        self.label_search.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        self.entry_search = tk.Entry(self.frame_top, font=("Helvetica", 10), bd=5, relief=tk.FLAT, fg="black", bg="#F8D8E8")
        self.entry_search.grid(row=0, column=1, padx=5, pady=5)
        
        self.label_count = tk.Label(self.frame_top, text="Çekilecek İşletme Sayısı:", bg="#FFFFFF", font=("Helvetica", 10))
        self.label_count.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        self.entry_count = tk.Entry(self.frame_top, font=("Helvetica", 10), bd=5, relief=tk.FLAT, fg="black", bg="#F8D8E8")
        self.entry_count.grid(row=1, column=1, padx=5, pady=5)

        self.button_start = tk.Button(self.frame_top, text="Verileri Çek", command=self.start_scraping_thread, font=("Helvetica", 10), bg="#FF0000", fg="#FFFFFF", relief=tk.RAISED, padx=5, pady=5)
        self.button_start.grid(row=0, column=2, rowspan=2, padx=10, pady=5)

        self.button_export = tk.Button(self.frame_top, text="Excel'e Aktar", command=self.export_to_excel, font=("Helvetica", 10), bg="#008000", fg="#FFFFFF", relief=tk.RAISED, padx=5, pady=5)
        self.button_export.grid(row=2, column=2, padx=10, pady=5)

        # Alt tarafta çekilen verileri göstermek için tablo
        self.frame_bottom = tk.Frame(self.root, bg="#FFFFFF", bd=10, relief=tk.FLAT)
        self.frame_bottom.place(relx=0.02, rely=0.3, relwidth=0.96, relheight=0.65)

        self.tree = ttk.Treeview(self.frame_bottom, columns=("İşletme Adı", "Adres", "İletişim No"), show='headings', height=15)
        self.tree.heading("İşletme Adı", text="İşletme Adı")
        self.tree.heading("Adres", text="Adres")
        self.tree.heading("İletişim No", text="İletişim No")
        self.tree.column("İşletme Adı", width=150)
        self.tree.column("Adres", width=250)
        self.tree.column("İletişim No", width=120)

        # Scrollbar ekleme
        self.scrollbar = ttk.Scrollbar(self.frame_bottom, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Stil ayarları
        style = ttk.Style()
        style.configure("Treeview", background="#F5F5F5", foreground="black", rowheight=25, fieldbackground="#F5F5F5")
        style.map("Treeview", background=[('selected', '#000000')], foreground=[('selected', '#FFFFFF')])

        self.tree.bind("<ButtonRelease-1>", self.on_tree_select)

    def start_scraping_thread(self):
        # Verileri çekmek için ayrı bir iş parçacığı oluşturma
        thread = threading.Thread(target=self.scrape_data)
        thread.start()

    def scrape_data(self):
        search_query = self.entry_search.get()
        try:
            count = int(self.entry_count.get())
        except ValueError:
            count = 15

        # Chrome WebDriver'ı başlat
        options = webdriver.ChromeOptions()
        options.add_argument("--window-size=1280,1024")
        driver = webdriver.Chrome(options=options)

        try:
            # Google Haritalar'ı aç
            driver.get("https://www.google.com.tr/maps/")

            # Arama kutusunu bul
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "searchboxinput")))
            search_box = driver.find_element(By.ID, "searchboxinput")

            # Kullanıcıdan alınan arama terimini arama kutusuna yaz
            search_box.send_keys(search_query)
            
            # Arama butonuna tıklayın veya Enter'a basarak arama yapın
            search_box.send_keys(Keys.ENTER)

            # Arama sonuçlarını beklemek için kısa bir süre uyutma
            time.sleep(4)

            # İlk 'count' kadar işletmeyi al ve her biri için bilgileri yazdır
            index = 0
            while index < count:
                businesses = driver.find_elements(By.CLASS_NAME, "Nv2PK")
                for business in businesses:
                    if index >= count:
                        break
                    try:
                        # İşletmeye tıkla
                        driver.execute_script("arguments[0].scrollIntoView(true);", business)
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(business)).click()
                        time.sleep(2)  # Bilgilerin yüklenmesi için bekleme süresi

                        # İşletme adı, adresi ve iletişim numarasını bul
                        try:
                            business_name = driver.find_element(By.CLASS_NAME, "DUwDvf").text
                        except:
                            business_name = "Bilgi bulunamadı"
                        
                        try:
                            address = driver.find_element(By.CSS_SELECTOR, "button[data-item-id='address'] .Io6YTe").text
                        except:
                            address = "Bilgi bulunamadı"
                        
                        try:
                            phone_number = driver.find_element(By.CSS_SELECTOR, "button[data-item-id^='phone'] .Io6YTe").text
                        except:
                            phone_number = "Bilgi bulunamadı"
                        
                        # Verileri tabloya ekle
                        self.tree.insert("", "end", values=(business_name, address, phone_number))

                        # Sayfa aşağı kaydırılır
                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(1)  # Kaydırma sonrası bekleme süresi

                        index += 1
                    except Exception as e:
                        print(f"Hata: {e}")

        finally:
            # İşlemler bittikten sonra tarayıcıyı kapat
            driver.quit()

    def export_to_excel(self):
        # Verileri Excel dosyasına aktarma işlemi
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            data = []
            for item in self.tree.get_children():
                values = self.tree.item(item, "values")
                data.append(values)
            
            df = pd.DataFrame(data, columns=["İşletme Adı", "Adres", "İletişim No"])
            df.to_excel(file_path, index=False)
            print("Veriler Excel dosyasına aktarıldı.")

    def on_tree_select(self, event):
        for item in self.tree.selection():
            self.tree.item(item, tags=["selected"])

if __name__ == "__main__":
    root = tk.Tk()
    app = GoogleMapsApp(root)
    root.mainloop()
