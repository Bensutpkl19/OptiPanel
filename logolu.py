import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import scrolledtext, messagebox, simpledialog
from PIL import Image, ImageTk
import itertools, io, os
from copy import deepcopy
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import Rectangle
from collections import defaultdict
import sys
import pandas as pd
from tkinter import filedialog, messagebox
import threading
import time
import random

# Rekürsiyon limitini artır (büyük ve karmaşık düzenler için gerekli olabilir)
sys.setrecursionlimit(2000)

# --- Kesim Payı Tanımı ---
BLADE_ALLOWANCE = 4.8  # 4.8mm bıçak payı
TRIMMING_MARGIN = 5.2 # Her kenardan 10mm traşlama payı

def load_materials_from_excel():
    """Excel dosyasından malzeme listesini okur ve bir sözlük listesi olarak döndürür."""
    try:
        file_path = filedialog.askopenfilename(
            title="Malzeme Listesi Excel Dosyasını Seçin",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls")]
        )

        if not file_path:
            messagebox.showinfo("Bilgi", "Herhangi bir dosya seçilmedi. Program sonlandırılıyor.")
            return []

        df = pd.read_excel(file_path)
        
        required_columns = ["kod", "isim", "eni", "boyu", "tur", "kalınlık"]
        if not all(col in df.columns for col in required_columns):
            messagebox.showerror("Hata", f"Excel dosyasında eksik sütunlar var. Gerekli sütunlar: {', '.join(required_columns)}")
            return []
        
        # Veri tiplerini dönüştürürken hataları ele al
        df['eni'] = pd.to_numeric(df['eni'], errors='coerce').fillna(0).astype(int)
        df['boyu'] = pd.to_numeric(df['boyu'], errors='coerce').fillna(0).astype(int)
        df['kalınlık'] = pd.to_numeric(df['kalınlık'], errors='coerce').fillna(0).astype(int)
        
        # 'bonus_oncelik' sütununu kontrol et ve yoksa varsayılan değer ata
        if 'bonus_oncelik' not in df.columns:
            df['bonus_oncelik'] = 999  # Varsayılan olarak düşük öncelik
        else:
            df['bonus_oncelik'] = pd.to_numeric(df['bonus_oncelik'], errors='coerce').fillna(999).astype(int)
            
        return df.to_dict('records')

    except Exception as e:
        messagebox.showerror("Hata", f"Excel dosyasını okurken bir hata oluştu: {e}")
        return []

# Uygulama başlamadan önce malzemeleri yükle
malzemeler = load_materials_from_excel()
if not malzemeler:
    sys.exit()

class SheetOptimizerGUI:
    def __init__(self, app):
        self.app = app
        self.app.title(" panel Kesim Optimizasyonu")
        self.app.geometry("1200x800")
        self.entries = {}
        self.malzemeler = malzemeler
        self.optimization_thread = None
        self.best_result = None
        self.max_bonus_parts_per_type = 0 # Varsayılan değer, kullanıcı tarafından ayarlanacak
        self.sheet_stocks_initial = {} # Başlangıç sac stoklarını tutacak
        self.available_sheet_dimensions = [] # Kullanıcının girdiği stokları olan sac boyutları

        self.status = ttk.Label(self.app, text="Hazır", anchor='w')
        self.status.pack(fill='x', side='bottom')

        self.main_frame = ttk.Frame(self.app, padding=10)
        self.main_frame.pack(fill='both', expand=True)

        self.left_frame = ttk.Frame(self.app, style='primary.TFrame')
        self.left_frame.pack(side='left', fill='y', padx=(0, 10))

        self.right_frame = ttk.Frame(self.app)
        self.right_frame.pack(side='right', fill='both', expand=True)

        self.logo_frame = ttk.Frame(self.left_frame)
        self.logo_frame.pack(fill='x', pady=(0, 10))

        self.entries_frame = ttk.Frame(self.left_frame)
        self.entries_frame.pack(fill='both', expand=True)

        self.bulk_frame = ttk.Frame(self.left_frame)
        self.bulk_frame.pack(fill='x', pady=(10, 0))

        self.load_logo(path="bellona-logo.png") # <--- Bu satırı böyle güncelleyin
        self.create_input_fields()
        self.create_bulk_input()
        self.create_output_tabs()

        self.optimize_button = ttk.Button(self.left_frame, text="Optimize Et", bootstyle="success", command=self.start_optimization)
        self.optimize_button.pack(pady=15)

    def load_logo(self, path="bellona-logo.png"):
        try:
            if os.path.exists(path):
                image = Image.open(path).resize((150, 80), Image.LANCZOS)
                self.logo = ImageTk.PhotoImage(image)
                label = ttk.Label(self.logo_frame, image=self.logo)
                label.pack(anchor='center', pady=(0, 10))
            else:
                label = ttk.Label(self.logo_frame, text="panel Kesim Optimizasyon", font=('Segoe UI', 16, 'bold'))
                label.pack(fill='x', pady=(0, 10))
        except Exception:
            ttk.Label(self.logo_frame, text=" panel Kesim Optimizasyon", font=('Segoe UI', 16, 'bold')).pack(fill='x', pady=(0, 10))

    def create_input_fields(self):
        label = ttk.Label(self.entries_frame, text="📋 Parça Listesi", font=('Segoe UI', 11, 'bold'))
        label.pack(pady=(0, 5))

        scroll_container = ttk.Frame(self.entries_frame)
        scroll_container.pack(fill='both', expand=True)

        self.input_canvas = tk.Canvas(scroll_container)
        self.input_scroll = ttk.Scrollbar(scroll_container, orient="vertical", command=self.input_canvas.yview)
        self.input_frame = ttk.Frame(self.input_canvas)

        self.input_frame.bind(
            "<Configure>",
            lambda e: self.input_canvas.configure(scrollregion=self.input_canvas.bbox("all"))
        )

        self.input_canvas.create_window((0, 0), window=self.input_frame, anchor="nw")
        self.input_canvas.configure(yscrollcommand=self.input_scroll.set)

        self.input_canvas.pack(side="left", fill="both", expand=True)
        self.input_scroll.pack(side="right", fill="y")

        for i, malzeme in enumerate(self.malzemeler):
            row_frame = ttk.Frame(self.input_frame)
            row_frame.pack(fill='x', pady=2, padx=5)

            label = ttk.Label(row_frame, text=f"{malzeme['kod']} {malzeme['isim']}", width=40, anchor='w')
            label.pack(side='left')

            entry = ttk.Entry(row_frame, width=10)
            entry.insert(0, "0")
            entry.pack(side='right')

            self.entries[malzeme['kod']] = entry

    def create_bulk_input(self):
        ttk.Label(self.bulk_frame, text="📦 Toplu Giriş (kod: adet)", font=('Segoe UI', 10, 'bold')).pack(pady=(0, 5))
        self.bulk_input = tk.Text(self.bulk_frame, height=7, width=50)
        self.bulk_input.pack(pady=(0, 5))

        def apply_bulk():
            for line in self.bulk_input.get("1.0", tk.END).strip().splitlines():
                if ":" in line:
                    code, qty = line.split(":")
                    if code.strip() in self.entries:
                        self.entries[code.strip()].delete(0, tk.END)
                        self.entries[code.strip()].insert(0, qty.strip())

        ttk.Button(self.bulk_frame, text="Uygula", command=apply_bulk, bootstyle="info-outline").pack(pady=(0, 10))

    def create_output_tabs(self):
        self.notebook = ttk.Notebook(self.right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.text_tab = ttk.Frame(self.notebook)
        self.output = scrolledtext.ScrolledText(self.text_tab, width=80, height=30)
        self.output.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.notebook.add(self.text_tab, text="Sonuçlar")

        self.vis_tab = ttk.Frame(self.notebook)
        self.vis_canvas = tk.Canvas(self.vis_tab)
        self.vis_scroll = ttk.Scrollbar(self.vis_tab, orient="vertical", command=self.vis_canvas.yview)
        self.vis_scroll_x = ttk.Scrollbar(self.vis_tab, orient="horizontal", command=self.vis_canvas.xview)

        self.vis_canvas.configure(yscrollcommand=self.vis_scroll.set, xscrollcommand=self.vis_scroll_x.set)
        self.vis_canvas.bind('<Configure>', lambda e: self.vis_canvas.configure(scrollregion=self.vis_canvas.bbox("all")))

        self.vis_scroll.pack(side="right", fill="y")
        self.vis_scroll_x.pack(side="bottom", fill="x")
        self.vis_canvas.pack(side="left", fill="both", expand=True)

        self.vis_frame = ttk.Frame(self.vis_canvas)
        self.vis_canvas.create_window((0, 0), window=self.vis_frame, anchor="nw")

        self.notebook.add(self.vis_tab, text="Yerleşim Görseli")

    def draw_layout(self, sheet_index, parts, material_label, original_sheet_dims, count=1):
        fig, ax = plt.subplots(figsize=(8, 6))
        
        if count > 1:
            fig.suptitle(f"Kesim Modeli #{sheet_index} - Yerleşim Planı ({count} Adet sunta İçin)", fontsize=12)
        else:
            fig.suptitle(f"Kesim Modeli #{sheet_index} - Yerleşim Planı", fontsize=12)

        ax.set_xlim(0, original_sheet_dims[0])
        ax.set_ylim(0, original_sheet_dims[1])
        ax.set_aspect('equal')
        ax.invert_yaxis()
        ax.grid(True, linestyle='--', alpha=0.7)
        ax.text(
            original_sheet_dims[0] / 2, -20,
            f"{material_label}", fontsize=14, weight='bold',
            ha='center', va='bottom', color='black')

        # Sacın dış kenarlarını (orijinal boyutları) çizer
        ax.add_patch(Rectangle((0, 0), original_sheet_dims[0], original_sheet_dims[1],
                                edgecolor='black', facecolor='none', lw=2, label="Ham sunta Boyutu"))

        # Kullanılabilir alanı temsil eden iç dikdörtgeni çizer
        usable_w = original_sheet_dims[0] - (2 * TRIMMING_MARGIN)
        usable_h = original_sheet_dims[1] - (2 * TRIMMING_MARGIN)
        
        # Eğer usable_w veya usable_h 0'dan küçükse çizme
        if usable_w > 0 and usable_h > 0:
            ax.add_patch(Rectangle((TRIMMING_MARGIN, TRIMMING_MARGIN), usable_w, usable_h,
                                    edgecolor='red', linestyle='--', facecolor='none', lw=1, 
                                    label=f"Kullanılabilir Alan ({usable_w}x{usable_h})"))
            ax.legend(loc='upper right', fontsize=8)

        colors = plt.cm.tab20.colors
        color_index = 0
        bonus_color_index = 0
        bonus_colors = plt.cm.Set2.colors # Fazla üretim için farklı bir renk paleti

        for part_info in parts:
            if part_info.get('type') == 'bonus':
                color =  'magenta'
                bonus_color_index += 1
            else:
                color = colors[color_index % len(colors)]
                color_index += 1

            # Parçaların konumuna traşlama payını ekle (görselleştirme için ofsetleme)
            x = part_info['x'] + TRIMMING_MARGIN
            y = part_info['y'] + TRIMMING_MARGIN
            width = part_info['eni_original'] # Nominal boyutları kullan
            height = part_info['boyu_original'] # Nominal boyutları kullan

            ax.add_patch(Rectangle(
                (x, y), width, height,
                edgecolor='black', facecolor=color, alpha=0.7, lw=2
            ))

            ax.text(
                x + width/2, y + height/2,
                f"{width}x{height}", # Sadece boyutları yaz
                ha='center', va='center', fontsize=6
            )

        canvas = FigureCanvasTkAgg(fig, master=self.vis_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        if not hasattr(self, 'figures'):
            self.figures = []
        self.figures.append(fig)
        plt.close(fig)
        return canvas

    def optimize_best_fit(self, parts_to_place, rect_w, rect_h, bonus_parts_allowed=True, max_bonus_per_type=0, current_bonus_counts=None):
        placed_parts = []
        unplaced_parts = deepcopy(parts_to_place)
        
        usable_rect_w = rect_w 
        usable_rect_h = rect_h 

        if usable_rect_w <= 0 or usable_rect_h <= 0:
            return [], unplaced_parts, [], current_bonus_counts if current_bonus_counts is not None else defaultdict(int)

        free_rects = [{'x': 0, 'y': 0, 'w': usable_rect_w, 'h': usable_rect_h}]
        
        if current_bonus_counts is None:
            current_bonus_counts = defaultdict(int)

        while unplaced_parts or (bonus_parts_allowed and any(v < max_bonus_per_type for v in current_bonus_counts.values())):
            best_fit = None
            best_score = float('inf')
            best_fit_index = -1
            best_part_index = -1
            is_bonus_placement = False

            # Önce talep edilen parçaları yerleştirmeye çalış
            for i, part in enumerate(unplaced_parts):
                if part['adet'] == 0:
                    continue

                orientations = [(part['eni'], part['boyu']), (part['boyu'], part['eni'])] if part.get('dondurulebilir', False) else [(part['eni'], part['boyu'])]

                for j, free_rect in enumerate(free_rects):
                    for current_w, current_h in orientations:
                        part_rotated = (current_w == part['boyu'])

                        part_nominal_w = current_w
                        part_nominal_h = current_h

                        if part_nominal_w <= free_rect['w'] and part_nominal_h <= free_rect['h']:
                            waste_area = (free_rect['w'] * free_rect['h']) - (part_nominal_w * part_nominal_h)
                            
                            score = waste_area
                            if score < best_score:
                                best_score = score
                                best_fit = {'part': part, 'rotated': part_rotated, 'w': part_nominal_w, 'h': part_nominal_h}
                                best_fit_index = j
                                best_part_index = i
                                is_bonus_placement = False
            
            # Eğer talep edilen parça yerleşemiyorsa ve bonus parçalara izin veriliyorsa
            if best_fit is None and bonus_parts_allowed:
                bonus_parts_for_current_type = []
                
                # Sadece mevcut malzemenin türüne ve kalınlığına ait bonus parçaları kullan
                if parts_to_place: 
                    current_group_type = parts_to_place[0]['malzeme']['tur']
                    current_group_kalinlik = parts_to_place[0]['malzeme']['kalınlık']
                    for m in self.malzemeler:
                        if m.get('tur') == current_group_type and m.get('kalınlık') == current_group_kalinlik:
                            bonus_parts_for_current_type.append(m)
                else: 
                     # Eğer hiç talep edilen parça yoksa, tüm bonus parçaları değerlendir
                     bonus_parts_for_current_type = self.malzemeler
                
                # Bonus parçaları önceliklerine göre sırala
                bonus_parts_for_current_type.sort(key=lambda m: m.get('bonus_oncelik', 999))

                for bonus_part_info in bonus_parts_for_current_type:
                    if self.max_bonus_parts_per_type > 0 and current_bonus_counts[bonus_part_info['kod']] >= self.max_bonus_parts_per_type:
                        continue 

                    orientations_b = [(bonus_part_info['eni'], bonus_part_info['boyu']), 
                                      (bonus_part_info['boyu'], bonus_part_info['eni'])] 
                    
                    for j, free_rect in enumerate(free_rects):
                        for current_w_b, current_h_b in orientations_b:
                            part_rotated_b = (current_w_b == bonus_part_info['boyu'])

                            part_nominal_w_b = current_w_b
                            part_nominal_h_b = current_h_b

                            if part_nominal_w_b <= free_rect['w'] and part_nominal_h_b <= free_rect['h']:
                                waste_area = (free_rect['w'] * free_rect['h']) - (part_nominal_w_b * part_nominal_h_b)
                                
                                bonus_priority_score = bonus_part_info.get('bonus_oncelik', 999) 
                                combined_score = (bonus_priority_score * 1_000_000_000) + waste_area 
                                
                                if combined_score < best_score: 
                                    best_score = combined_score
                                    best_fit = {'part': bonus_part_info, 'rotated': part_rotated_b, 'w': part_nominal_w_b, 'h': part_nominal_h_b}
                                    best_fit_index = j
                                    is_bonus_placement = True

            if best_fit:
                free_rect = free_rects.pop(best_fit_index)
                part_w, part_h = best_fit['w'], best_fit['h'] 

                # Malzeme bilgisi ekleme (önemli)
                part_material_info = best_fit['part'].get('malzeme', {'tur': 'Bilinmiyor', 'kalınlık': 'Bilinmiyor'}) # Eğer bonus parça ise kendi malzeme bilgisi yoksa default al

                if is_bonus_placement:
                    placed_parts.append({
                        'kod': best_fit['part']['kod'],
                        'isim': best_fit['part']['isim'],
                        'eni_original': best_fit['part']['eni'] if not best_fit['rotated'] else best_fit['part']['boyu'],
                        'boyu_original': best_fit['part']['boyu'] if not best_fit['rotated'] else best_fit['part']['eni'],
                        'x': free_rect['x'], 
                        'y': free_rect['y'], 
                        'adet': 1,
                        'rotated': best_fit['rotated'],
                        'type': 'bonus',
                        'malzeme': part_material_info # Malzeme bilgisini ekle
                    })
                    current_bonus_counts[best_fit['part']['kod']] += 1
                else:
                    part_to_place = unplaced_parts[best_part_index]
                    placed_parts.append({
                        'kod': part_to_place['kod'],
                        'isim': part_to_place['isim'],
                        'eni_original': part_to_place['eni'] if not best_fit['rotated'] else part_to_place['boyu'],
                        'boyu_original': part_to_place['boyu'] if not best_fit['rotated'] else part_to_place['eni'],
                        'x': free_rect['x'], 
                        'y': free_rect['y'], 
                        'adet': 1,
                        'rotated': best_fit['rotated'],
                        'malzeme': part_to_place['malzeme'] # Malzeme bilgisini ekle
                    })
                    part_to_place['adet'] -= 1
                
                # Kalan boşlukları hesaplarken bıçak payını düş
                remnantA1 = {'x': free_rect['x'] + part_w + BLADE_ALLOWANCE, 
                             'y': free_rect['y'], 
                             'w': free_rect['w'] - part_w - BLADE_ALLOWANCE, 
                             'h': free_rect['h']}
                remnantA2 = {'x': free_rect['x'], 
                             'y': free_rect['y'] + part_h + BLADE_ALLOWANCE, 
                             'w': part_w, 
                             'h': free_rect['h'] - part_h - BLADE_ALLOWANCE}
                areaA = (max(0, remnantA1['w']) * max(0, remnantA1['h'])) + (max(0, remnantA2['w']) * max(0, remnantA2['h']))

                remnantB1 = {'x': free_rect['x'], 
                             'y': free_rect['y'] + part_h + BLADE_ALLOWANCE, 
                             'w': free_rect['w'], 
                             'h': free_rect['h'] - part_h - BLADE_ALLOWANCE}
                remnantB2 = {'x': free_rect['x'] + part_w + BLADE_ALLOWANCE, 
                             'y': free_rect['y'], 
                             'w': free_rect['w'] - part_w - BLADE_ALLOWANCE, 
                             'h': part_h}
                areaB = (max(0, remnantB1['w']) * max(0, remnantB1['h'])) + (max(0, remnantB2['w']) * max(0, remnantB2['h']))
                
                if areaA > areaB:
                    remnants_to_add = [remnantA1, remnantA2]
                else:
                    remnants_to_add = [remnantB1, remnantB2]

                for remnant in remnants_to_add:
                    if remnant['w'] >= BLADE_ALLOWANCE and remnant['h'] >= BLADE_ALLOWANCE:
                        free_rects.append(remnant)

                free_rects.sort(key=lambda r: (r['w'] * r['h']), reverse=True)

                unplaced_parts = [p for p in unplaced_parts if p['adet'] > 0]
            else:
                break 

        return placed_parts, unplaced_parts, free_rects, current_bonus_counts

    def start_optimization(self):
        if self.optimization_thread and self.optimization_thread.is_alive():
            messagebox.showwarning("Uyarı", "Bir optimizasyon işlemi zaten devam ediyor.")
            return
        
        # Kullanıcıdan bonus parça limitini al
        try:
            self.max_bonus_parts_per_type = simpledialog.askinteger(
                "Bonus Parça Limiti", 
                "Boşlukları doldurmak için üretilebilecek her bir bonus parça tipi için maksimum adedi girin (0 sınırsız):", 
                parent=self.app, 
                minvalue=0, 
                initialvalue=0
            )
            if self.max_bonus_parts_per_type is None: # Kullanıcı iptal etti
                return
        except Exception as e:
            messagebox.showerror("Hata", f"Geçersiz giriş: {e}")
            return

        # Kullanıcıdan her bir sac boyutu için stok adedi al
        self.sheet_stocks_initial = {}
        self.available_sheet_dimensions = [] # Kullanılabilir sac boyutlarını ve stoklarını tutacak
        
        # Tanımlı sac boyutları
        sheet_options = [
            (2100, 3660), 
            (3660, 2100), 
            (2880, 2100), 
            (2100, 2880)  
        ]

        for w, h in sheet_options:
            stock = simpledialog.askinteger(
                "Sunta Stok Bilgisi",
                f"Lütfen {w}x{h} mm sunta levhadan kaç adet stoğunuz olduğunu girin:",
                parent=self.app,
                minvalue=0,
                initialvalue=0
            )
            if stock is None: # Kullanıcı iptal etti
                return
            if stock > 0: # Sadece stoğu olanları dikkate al
                self.sheet_stocks_initial[(w, h)] = stock
                self.available_sheet_dimensions.append((w, h))
        
        if not self.available_sheet_dimensions:
            messagebox.showerror("Hata", "Lütfen en az bir sunta boyutu için stok giriniz.")
            return

        self.status.config(text="Optimizasyon devam ediyor... Lütfen bekleyin.")
        self.optimize_button.config(state="disabled")
        self.output.delete('1.0', tk.END)
        for widget in self.vis_frame.winfo_children():
            widget.destroy()
        if hasattr(self, 'figures'):
            del self.figures

        self.optimization_thread = threading.Thread(target=self._run_all_strategies)
        self.optimization_thread.daemon = True
        self.optimization_thread.start()
        
        self.app.after(100, self.check_thread_status)

    def check_thread_status(self):
        if self.optimization_thread.is_alive():
            self.app.after(100, self.check_thread_status)
        else:
            self._display_results()
            self.status.config(text="Optimizasyon tamamlandı.")
            self.optimize_button.config(state="normal")
            
    def _run_all_strategies(self):
        parca_adetleri = {}
        for malzeme in self.malzemeler:
            try:
                adet = int(self.entries[malzeme['kod']].get())
                if adet > 0:
                    parca_adetleri[malzeme['kod']] = {
                        "adet": adet,
                        "malzeme": malzeme, 
                        "dondurulebilir": True 
                    }
            except ValueError:
                continue

        if not parca_adetleri:
            self.best_result = {"message": "Uyarı", "detail": "Lütfen en az bir parça için adet giriniz."}
            return

        total_requested_parts_count = sum(data['adet'] for data in parca_adetleri.values())
        
        global_best_plan_score = (float('inf'), float('inf')) # (Toplam Sac Sayısı, Toplam Atık Alanı)
        self.best_result = None # En iyi sonucu tutacak

        # Tüm parçaları tek bir liste halinde topla
        all_initial_parts = []
        for kod, data in parca_adetleri.items():
            for _ in range(data['adet']):
                all_initial_parts.append({
                    "kod": kod, "isim": data['malzeme']['isim'], 
                    "eni": data['malzeme']['eni'], "boyu": data['malzeme']['boyu'], 
                    "dondurulebilir": data['dondurulebilir'],
                    "malzeme": data['malzeme'] 
                })

        # Farklı sıralama anahtarları
        sort_keys = [
            lambda p: (p['eni'] * p['boyu']),      
            lambda p: max(p['eni'], p['boyu']),    
            lambda p: min(p['eni'], p['boyu']),    
            lambda p: p['eni'],                    
            lambda p: p['boyu'],                   
            lambda p: random.random()              
        ]
        
        # Her sac tipi ve her sıralama anahtarı kombinasyonu için bir deneme yap
        all_optimization_configs = []
        for sheet_dims in self.available_sheet_dimensions:
            for sort_key_func in sort_keys:
                all_optimization_configs.append((sheet_dims, sort_key_func))
        
        random.shuffle(all_optimization_configs) # Deneme sırasını karıştır

        for trial_index, (initial_sheet_dims, sort_key_func) in enumerate(all_optimization_configs):
            
            current_sheet_stocks = deepcopy(self.sheet_stocks_initial) # Her deneme için stokları sıfırla
            parts_remaining_for_trial = deepcopy(all_initial_parts) # Her deneme için parçaları sıfırla
            
            trial_levha_kesim_planlari = []
            trial_total_sheets_used = 0
            trial_placed_parts_count = 0
            trial_placed_parts_overall = defaultdict(int)
            trial_current_bonus_counts = defaultdict(int)

            # Sac tiplerini başlangıçtaki tercih edilen sac tipiyle sırala
            sheet_priority_order = [initial_sheet_dims] + [s for s in self.available_sheet_dimensions if s != initial_sheet_dims]
            
            current_sheet_index_in_priority = 0

            # Tüm parçalar yerleşene veya stoklar bitene kadar devam et
            while trial_placed_parts_count < total_requested_parts_count and current_sheet_index_in_priority < len(sheet_priority_order):
                current_sheet_dims = sheet_priority_order[current_sheet_index_in_priority]
                
                if current_sheet_stocks[current_sheet_dims] <= 0:
                    current_sheet_index_in_priority += 1
                    continue 

                original_sheet_w, original_sheet_h = current_sheet_dims
                usable_sheet_w = original_sheet_w - (2 * TRIMMING_MARGIN)
                usable_sheet_h = original_sheet_h - (2 * TRIMMING_MARGIN)

                if usable_sheet_w <= 0 or usable_sheet_h <= 0:
                    current_sheet_index_in_priority += 1
                    continue 

                # optimize_best_fit'e göndermeden önce parçaları sırala
                parts_to_optimize_on_sheet = []
                temp_dict = defaultdict(lambda: {'adet': 0, 'malzeme': None, 'dondurulebilir': False, 'eni':0, 'boyu':0, 'isim':''})
                for part_info in parts_remaining_for_trial:
                    key = part_info['kod']
                    if temp_dict[key]['malzeme'] is None:
                        temp_dict[key]['malzeme'] = part_info['malzeme']
                        temp_dict[key]['dondurulebilir'] = part_info['dondurulebilir']
                        temp_dict[key]['eni'] = part_info['eni']
                        temp_dict[key]['boyu'] = part_info['boyu']
                        temp_dict[key]['isim'] = part_info['isim']
                    temp_dict[key]['adet'] += 1

                for key, data in temp_dict.items():
                    parts_to_optimize_on_sheet.append({
                        "kod": key, "isim": data['isim'], "eni": data['eni'], 
                        "boyu": data['boyu'], "adet": data['adet'], 
                        "dondurulebilir": data['dondurulebilir'], "malzeme": data['malzeme']
                    })

                parts_to_optimize_on_sheet = sorted(parts_to_optimize_on_sheet, key=sort_key_func, reverse=True)
                
                placed_parts_on_sheet, remaining_after_placement_dict_list, _, trial_current_bonus_counts = self.optimize_best_fit(
                    parts_to_optimize_on_sheet, usable_sheet_w, usable_sheet_h,
                    bonus_parts_allowed=True,
                    max_bonus_per_type=self.max_bonus_parts_per_type,
                    current_bonus_counts=trial_current_bonus_counts
                )

                # Düzeltme: material_type_for_sheet ve material_thickness_for_sheet değişkenlerini 
                # her zaman başlatmak için bloğun başında tanımla.
                material_type_for_sheet = 'Bilinmiyor'
                material_thickness_for_sheet = 'Bilinmiyor'

                if placed_parts_on_sheet:
                    # İlk yerleşen parçanın malzeme bilgisini güvenli bir şekilde alalım
                    first_placed_part = placed_parts_on_sheet[0]
                    # 'malzeme' anahtarı yoksa veya sözlük değilse varsayılan değerleri kullan
                    if 'malzeme' in first_placed_part and isinstance(first_placed_part['malzeme'], dict):
                        material_type_for_sheet = first_placed_part['malzeme'].get('tur', 'Bilinmiyor')
                        material_thickness_for_sheet = first_placed_part['malzeme'].get('kalınlık', 'Bilinmiyor')

                if not placed_parts_on_sheet:
                    # Bu sac boyutuna hiçbir parça yerleşmediyse, bir sonraki sac boyutuna geç
                    current_sheet_index_in_priority += 1
                    continue
                
                # Sac kullanıldı, stoktan düş
                current_sheet_stocks[current_sheet_dims] -= 1
                trial_total_sheets_used += 1

                trial_levha_kesim_planlari.append({
                    "levha_no": trial_total_sheets_used, 
                    "levha_genisligi_orijinal": original_sheet_w, 
                    "levha_yuksekligi_orijinal": original_sheet_h, 
                    "tur": material_type_for_sheet, 
                    "kalinlik": material_thickness_for_sheet, 
                    "kesilen_parcalar": placed_parts_on_sheet,
                    "sheet_dimensions_original": (original_sheet_w, original_sheet_h)
                })

                # Yerleşen parçaları toplam placed_parts_count'a ve genel döküme ekle
                for p_info in placed_parts_on_sheet:
                    if p_info.get('type') != 'bonus': 
                        trial_placed_parts_count += p_info['adet']
                        trial_placed_parts_overall[p_info['kod']] += p_info['adet']
                
                # Kalan parçaları güncelle (remaining_after_placement_dict_list'ten tekil parçalar listesine dönüştür)
                parts_remaining_for_trial = []
                for p_rem_dict in remaining_after_placement_dict_list:
                    parts_remaining_for_trial.extend([
                        {
                            "kod": p_rem_dict['kod'], "isim": p_rem_dict['isim'], 
                            "eni": p_rem_dict['eni'], "boyu": p_rem_dict['boyu'], 
                            "dondurulebilir": p_rem_dict['dondurulebilir'],
                            "malzeme": p_rem_dict['malzeme'] 
                        } for _ in range(p_rem_dict['adet'])
                    ])
                
                # Eğer mevcut sacın stoğu bittiyse, bir sonraki sac tipine geç
                if current_sheet_stocks[current_sheet_dims] <= 0:
                    current_sheet_index_in_priority += 1
            
            # Bu deneme sonucunda elde edilen toplam atık alanını hesapla
            total_area_used_in_trial = sum(
                sum(p['eni_original'] * p['boyu_original'] for p in plan['kesilen_parcalar'])
                for plan in trial_levha_kesim_planlari
            )
            total_original_sheet_area_in_trial = sum(
                plan['levha_genisligi_orijinal'] * plan['levha_yuksekligi_orijinal']
                for plan in trial_levha_kesim_planlari
            )
            total_waste_in_trial = total_original_sheet_area_in_trial - total_area_used_in_trial

            # Bu denemenin sonucunu değerlendir (sac sayısı ve atık alanı)
            current_trial_score = (trial_total_sheets_used, total_waste_in_trial)

            # Eğer daha iyi bir sonuç bulunduysa kaydet
            if current_trial_score < global_best_plan_score:
                global_best_plan_score = current_trial_score
                self.best_result = {
                    "levha_kesim_planlari": deepcopy(trial_levha_kesim_planlari), # Kopyala!
                    "genel_sunta_sayisi": trial_total_sheets_used, 
                    "all_placed_parts_overall": deepcopy(trial_placed_parts_overall), # Kopyala!
                    "parca_adetleri": deepcopy(parca_adetleri), # Kopyala!
                    "strategy_name": f"Deneme {trial_index+1} (Başlangıç Sunta: {initial_sheet_dims[0]}x{initial_sheet_dims[1]}, Sıralama: {sort_key_func.__name__.split(' at ')[0]})",
                    "best_sheet_dims_used": initial_sheet_dims, # Bu denemede en başta denenen sac tipi
                    "final_sheet_stock_distribution": deepcopy(current_sheet_stocks) # Kopyala!
                }
            
            # Eğer tüm parçalar yerleştirildiyse, daha fazla denemeye gerek kalmayabilir.
            # Ancak yine de en iyi sonucu bulmak için tüm konfigürasyonları denemek daha garanti olur.
            # Performans kritikse bu 'if' bloğunun sonuna bir 'break' eklenebilir.
            # if trial_placed_parts_count >= total_requested_parts_count and trial_total_sheets_used == global_best_plan_score[0]:
            #     break

    def _display_results(self):
        if self.best_result is None:
            messagebox.showerror("Hata", "Optimizasyon işlemi tamamlanamadı veya bir sonuç üretilemedi. Lütfen parça ve sunta stoğu girdiğinizden emin olun.")
            return

        results = self.best_result
        
        if "message" in results:
            messagebox.showwarning(results["message"], results["detail"])
            return

        buffer = io.StringIO() 
        
        levha_kesim_planlari = results["levha_kesim_planlari"]
        genel_sunta_sayisi = results["genel_sunta_sayisi"]
        all_placed_parts_overall = results["all_placed_parts_overall"]
        parca_adetleri = results["parca_adetleri"]
        strategy_name = results["strategy_name"]
        final_sheet_stock_distribution = results["final_sheet_stock_distribution"] 

        print("\n\n=== GENEL SON ÖZET ===", file=buffer)
        print(f"EN İYİ SONUÇ ELDE EDİLDİ: {strategy_name}", file=buffer)
        
        # Kullanılan sac levha dağılımını göster
        used_sheet_counts = defaultdict(int)
        for plan in levha_kesim_planlari:
            used_sheet_counts[plan['sheet_dimensions_original']] += 1
        
        print("\n--- Kullanılan sunta Levha Dağılımı ---", file=buffer)
        if used_sheet_counts:
            for dims, count in used_sheet_counts.items():
                print(f"- {dims[0]}x{dims[1]} mm: {count} Adet Kullanıldı", file=buffer)
        else:
            print("Hiç sunta levha kullanılmadı.")
            
        print("\n--- Başlangıç ve Son Sunta Stoğu ---", file=buffer)
        for dims in self.sheet_stocks_initial.keys():
            initial_stock = self.sheet_stocks_initial[dims]
            remaining_stock = final_sheet_stock_distribution.get(dims, 0) 
            print(f"- {dims[0]}x{dims[1]} mm: Başlangıç Stoğu = {initial_stock}, Kalan Stok = {remaining_stock}", file=buffer)

        print(f"\nToplam kullanılan ham sunta sayısı: {genel_sunta_sayisi}", file=buffer)
        print(f"Sunta başına traşlama payı: Her kenardan {TRIMMING_MARGIN} mm", file=buffer)

        if levha_kesim_planlari:
            toplam_kullanilan_alan = sum(
                sum(p['eni_original'] * p['boyu_original'] for p in plan['kesilen_parcalar'])
                for plan in levha_kesim_planlari
            )
            total_original_sheet_area_used = sum(
                plan['levha_genisligi_orijinal'] * plan['levha_yuksekligi_orijinal']
                for plan in levha_kesim_planlari
            )
            toplam_trim = total_original_sheet_area_used - toplam_kullanilan_alan
            
            print(f"Toplam trim (traşlama payı dahil): {toplam_trim / 1000000:.3f} m²", file=buffer)
            if total_original_sheet_area_used > 0:
                print(f"Trim yüzdesi (traşlama payı dahil): {(toplam_trim / total_original_sheet_area_used) * 100:.2f}%", file=buffer)
            else:
                print("Trim yüzdesi hesaplanamadı (kullanıl sunta alanı 0).", file=buffer)
        else:
            print("Atık ve trim yüzdesi hesaplanamadı (hiç parça yerleşmedi).", file=buffer)

        print("\n\n=== GENEL ÜRETİM ÖZETİ (SON HALİ) ===", file=buffer)
        for p_malzeme in self.malzemeler:
            kod = p_malzeme['kod']
            isim = p_malzeme['isim']
            talep = parca_adetleri.get(kod, {}).get('adet', 0)
            toplam_uretilen = all_placed_parts_overall[kod]
            fazla = max(0, toplam_uretilen - talep)
            
            if talep > 0 or toplam_uretilen > 0:
                print(f"- {kod} {isim}: Talep = {talep}, Toplam Üretilen = {toplam_uretilen}, Fazla Üretim = {fazla}", file=buffer)

        grouped_layouts = defaultdict(list)
        for plan in levha_kesim_planlari:
            # Planın benzersiz bir "imzasını" oluştur
            # Malzeme bilgilerini de imzaya dahil ederek aynı boyutlardaki farklı malzemeleri ayırabiliriz
            # Ancak görselleştirmede bir sacda tek malzeme türü ve kalınlığı olduğu varsayımıyla devam edelim
            
            sorted_parts_for_signature = sorted(
                plan['kesilen_parcalar'], 
                key=lambda p: (p['x'], p['y'], p['eni_original'], p['boyu_original'], p.get('rotated', False), p.get('type', 'normal'))
            )
            
            layout_signature = tuple([
                (p['x'], p['y'], p['eni_original'], p['boyu_original'], p.get('rotated', False), p.get('type', 'normal')) 
                for p in sorted_parts_for_signature
            ])
            
            # Sac boyutunu, türünü ve kalınlığını da imzanın bir parçası yap
            full_signature = (plan['sheet_dimensions_original'], plan['tur'], plan['kalinlik'], layout_signature)
            
            grouped_layouts[full_signature].append(plan)

        print("\n\n=== KESİM MODELLERİ VE DETAYLI RAPOR ===", file=buffer)
        layout_model_index = 0
        for signature, plans in grouped_layouts.items():
            layout_model_index += 1
            sample_plan = plans[0] # Bu modeldeki bir örnek planı al
            
            model_count = len(plans) # Bu modelden kaç adet üretileceğini say
            original_sheet_dims = sample_plan['sheet_dimensions_original']
            
            print(f"\n--- Kesim Modeli #{layout_model_index} ---", file=buffer)
            print(f"Bu modelden {model_count} adet sunta kesilecektir.", file=buffer)
            print(f"Sunta Boyutu (Ham): {original_sheet_dims[0]}x{original_sheet_dims[1]} mm", file=buffer)
            print(f"sunta Türü: {sample_plan['tur']}, Kalınlık: {sample_plan['kalinlik']}mm", file=buffer) # Yeni eklenen satır
            print(f"sunta başına traşlama payı: Her kenardan {TRIMMING_MARGIN} mm", file=buffer)

            total_model_area_used = sum(p['eni_original'] * p['boyu_original'] for p in sample_plan['kesilen_parcalar'])
            total_model_sheet_area_original = original_sheet_dims[0] * original_sheet_dims[1]
            model_trim = total_model_sheet_area_original - total_model_area_used
            
            print(f"Bu modeldeki trim (traşlama payı dahil): {model_trim / 1000000:.3f} m²", file=buffer)
            if total_model_sheet_area_original > 0:
                print(f"Bu modeldeki trim yüzdesi (traşlama payı dahil): {(model_trim / total_model_sheet_area_original) * 100:.2f}%", file=buffer)
            else:
                print("Model trim yüzdesi hesaplanamadı (sunta alanı 0).", file=buffer)
            
            parts_in_model_summary = defaultdict(int)
            for p in sample_plan['kesilen_parcalar']:
                parts_in_model_summary[f"{p['kod']} ({p['eni_original']}x{p['boyu_original']})"] += 1

            print("\nModeldeki Parçalar (Adet - Boyut):", file=buffer)
            for part_desc, count in parts_in_model_summary.items():
                print(f"- {part_desc}: {count} adet", file=buffer)

            bonus_parts_in_model = [p for p in sample_plan['kesilen_parcalar'] if p.get('type') == 'bonus']
            if bonus_parts_in_model:
                print("\nModeldeki Bonus Parçalar:", file=buffer)
                bonus_summary = defaultdict(int)
                for p in bonus_parts_in_model:
                    bonus_summary[f"{p['kod']} {p['isim']} ({p['eni_original']}x{p['boyu_original']})"] += 1
                for bonus_desc, count in bonus_summary.items():
                    print(f"- {bonus_desc}: {count} adet", file=buffer)

            self.draw_layout(layout_model_index, sample_plan['kesilen_parcalar'], 
                             f"{sample_plan['tur']} {sample_plan['kalinlik']}mm", 
                             sample_plan['sheet_dimensions_original'], count=model_count)

        self.output.insert(tk.END, buffer.getvalue())
        buffer.close()

        if genel_sunta_sayisi > 0:
            self.notebook.select(self.vis_tab)
            
if __name__ == "__main__":
    app = ttk.Window(themename="flatly")
    try:
        SheetOptimizerGUI(app)
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Hata", f"Bir hata oluştu: {str(e)}")