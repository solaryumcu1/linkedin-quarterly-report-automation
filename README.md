[README.md](https://github.com/user-attachments/files/21315240/README.md)# linkedin-quarterly-report-automation


# linkedin-quarterly-report-automation 📊🤖


## 🚀 Features

- ✅ Automatically detects correct sheet and dominant year
- 📅 Quarterly comparison (e.g., Q1 vs Q2)
- 📈 Calculates Engagement Rate (ER) for each post
- 🎯 Classifies ER and Impression performance into human-readable labels[Uploading 

## 🧠 How It Works

1. Upload your LinkedIn Excel file
2. The script auto-detects the correct sheet and year
3. Quarterly engagement rates are calculated and compared
4. Custom labels and colors are applied in the Excel output
5. A summary chart is added at the bottom of the report
   

- 🎨 Applies custom Excel color formatting based on labels
- 📊 Generates bar charts for ER and Impression trends
- 🖥️ Converts to standalone `.exe` app for non-technical users

linkedin-quarterly-report-automation/
│
├── proje.py # Main Python script
        ├── 
            
            
            import tkinter as tk
            from tkinter import filedialog
            import pandas as pd
            import os
            import matplotlib.pyplot as plt
            from collections import defaultdict
            from openpyxl import load_workbook
            from openpyxl.styles import PatternFill
        
        Sütun eşleşme haritası ve normalize fonksiyonu
                  column_mapping = {
                      "date": ["date", "created date", "tarih", "Date"],
                      "engagement": [
                          "engagement", "Engagement", "engagement total", "etkileşim",
                          "engagement (organic + paid)", "engagement (paid + organic)",
                          "engagement\n(organic + paid)", "engagement\u202f", "engagement\xa0"
                      ],
                      "engagement_rate": ["engagement rate", "etkileşim oranı"],
                      "impressions": [
                          "impression", "impressions", "gösterim",
                          "impression\n", "impression (organic + paid)", "impression\n(organic + paid)"
                      ],
                      "title": ["post title", "context", "içerik", "title"],
                      "likes": ["like", "likes", "beğeni"],
                      "content_type": ["content type", "type", "içerik türü"],
                  }
                  
          
          def save_pie_chart(df, column_name, filename):
              import matplotlib.pyplot as plt
              if column_name in df.columns:
                  counts = df[column_name].value_counts().sort_index()
                  plt.figure(figsize=(6, 6))
                  plt.pie(counts, labels=counts.index, autopct='%1.1f%%', startangle=140)
                  plt.title(f"{column_name} Distribution")
                  plt.axis('equal')
                  plt.tight_layout()
                  plt.savefig(filename)
                  plt.close()
          
          
          def embed_chart_to_excel(excel_path, image_path, sheet_name="Charts", cell="B2"):
              from openpyxl import load_workbook
              from openpyxl.drawing.image import Image as ExcelImage
              wb = load_workbook(excel_path)
              if sheet_name in wb.sheetnames:
                  ws = wb[sheet_name]
              else:
                  ws = wb.create_sheet(sheet_name)
              img = ExcelImage(image_path)
              ws.add_image(img, cell)
              wb.save(excel_path)
          
          def normalize_columns(df, column_mapping=None):
              renamed = {}
              def clean(s):
                  if not isinstance(s, str):
                      return ""
                  return s.strip().lower().replace('\n', '').replace('\xa0', ' ').replace('  ', ' ')
              cleaned_cols = {clean(col): col for col in df.columns}
              if column_mapping:
                  for std_col, variants in column_mapping.items():
                      for variant in variants:
                          key = clean(variant)
                          if key in cleaned_cols:
                              renamed[cleaned_cols[key]] = std_col
                              break
              df = df.rename(columns=renamed)
              df.columns = [clean(col) for col in df.columns]
              return df
          
          def get_dominant_year_from_file(file_path):
              xl = pd.ExcelFile(file_path)
              all_years = []
              for sheet in xl.sheet_names:
                  for header_row in range(5):
                      try:
                          df = xl.parse(sheet, header=header_row)
                          # Eğer dict ise (çoklu sheet), her bir DataFrame için işle
                          if isinstance(df, dict):
                              for _df in df.values():
                                  _df.columns = [str(c).strip().lower() for c in _df.columns]
                                  if 'date' in _df.columns or 'created date' in _df.columns:
                                      date_col = 'date' if 'date' in _df.columns else 'created date'
                                      _df[date_col] = pd.to_datetime(_df[date_col], errors='coerce')
                                      years = _df[date_col].dt.year.dropna().tolist()
                                      all_years.extend(years)
                          else:
                              df.columns = [str(c).strip().lower() for c in df.columns]
                              if 'date' in df.columns or 'created date' in df.columns:
                                  date_col = 'date' if 'date' in df.columns else 'created date'
                                  df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                                  years = df[date_col].dt.year.dropna().tolist()
                                  all_years.extend(years)
                      except:
                          continue
              if not all_years:
                  raise ValueError(f" {file_path} içinde geçerli yıl bulunamadı.")
              year_series = pd.Series(all_years)
              return int(year_series.mode()[0])
          
          def load_valid_sheet(filepath):
              xls = pd.ExcelFile(filepath)
              best_df = None
              best_year_count = 0
              dominant_year = None
              for sheet in xls.sheet_names:
                  for header_row in range(10):
                      try:
                          df_try = pd.read_excel(filepath, sheet_name=sheet, header=header_row)
                          df_try = normalize_columns(df_try, column_mapping)
                          if 'date' not in df_try.columns:
                              continue
                          df_try['date'] = pd.to_datetime(df_try['date'], errors='coerce')
                          df_try = df_try.dropna(subset=['date'])
                          year_counts = df_try['date'].dt.year.value_counts()
                          if year_counts.empty:
                              continue
                          top_year = year_counts.idxmax()
                          top_count = year_counts.max()
                          if top_count > best_year_count:
                              best_df = df_try
                              best_year_count = top_count
                              dominant_year = top_year
                      except Exception:
                          continue
              if best_df is not None:
                  best_df['Year'] = dominant_year
                  return best_df
              else:
                  raise ValueError(f" {filepath} couldn't find any valid date.")
          
          
          
          
          def grafik_er_ve_impression(df):
              """
              ER Label ve Impression Label için pasta grafiklerini gösterir.
              """
              if "ER Label" in df.columns:
                  er_counts = df["ER Label"].value_counts().sort_index()
                  plt.figure(figsize=(6, 6))
                  plt.pie(er_counts, labels=er_counts.index, autopct='%1.1f%%', startangle=140)
                  plt.title("ER Label Distribution")
                  plt.axis('equal')
                  plt.tight_layout()
                  plt.show()
          
              if "Impression Label" in df.columns:
                  imp_counts = df["Impression Label"].value_counts().sort_index()
                  plt.figure(figsize=(6, 6))
                  plt.pie(imp_counts, labels=imp_counts.index, autopct='%1.1f%%', startangle=140)
                  plt.title("Impression Label Distribution")
                  plt.axis('equal')
                  plt.tight_layout()
                  plt.show()
          
          def renklendir_excel(dosya_adi):
              from openpyxl import load_workbook
              from openpyxl.styles import PatternFill
              wb = load_workbook(dosya_adi)
              if wb is None:
                  print("Workbook not found!")
                  return
              ws = wb.active
              if ws is None:
                  print("Worksheet not found!")
                  return
              er_renkleri = {
                  "Very Low Interest": "AC0000",
                  "Low Interest": "FF0000",
                  "Average": "FFF700",
                  "High Interest": "003EFF",
                  "Very High Interest": "04FF00"
              }
              impression_renkleri = {
                  "Very Narrow Distribution": "099A2E",
                  "Narrow Distribution": "FA8F8F",
                  "Mild Contraction": "BFA93E",
                  "Equal Distribution": "FFD500",
                  "Slight Increase": "00C5FF",
                  "Increased Distribution": "FF00CD",
                  "Wide Distribution": "CBA6F7",
                  "Very Wide Distribution": "00FF08"
              }
              try:
                  header_row = next(ws.iter_rows(min_row=1, max_row=1))
              except Exception:
                  print("Header row alınamadı!")
                  return
              headers = {cell.value: idx for idx, cell in enumerate(header_row, 1)}
              for row in ws.iter_rows(min_row=2):
                  if "ER Label" in headers:
                      etiket = row[headers["ER Label"] - 1].value
                      if isinstance(etiket, str) and etiket in er_renkleri:
                          renk = er_renkleri[etiket]
                          for hedef in ["ER Label", "ER Diff", "ER Comparison"]:
                              if hedef in headers:
                                  cell = row[headers[hedef] - 1]
                                  if cell is not None:
                                      cell.fill = PatternFill(start_color=renk, end_color=renk, fill_type="solid")
                  if "Impression Label" in headers:
                      etiket = row[headers["Impression Label"] - 1].value
                      if isinstance(etiket, str) and etiket in impression_renkleri:
                          renk = impression_renkleri[etiket]
                          for hedef in ["Impression Label", "Impression Diff", "Impression Comparison"]:
                              if hedef in headers:
                                  cell = row[headers[hedef] - 1]
                                  if cell is not None:
                                      cell.fill = PatternFill(start_color=renk, end_color=renk, fill_type="solid")
              yeni_dosya = dosya_adi.replace(".xlsx", "_renkli.xlsx")
              wb.save(yeni_dosya)
              print(f" Coloring completed: {yeni_dosya}")
          
          # --- GUI Sınıfı ---
          class ExcelComparerApp:
              def __init__(self, root):
                  self.root = root
                  self.root.title("Excel Yearly Comparison")
                  self.root.geometry("600x450")
                  self.root.configure(bg="#23272F")
          
                  self.file_1 = None
                  self.file_2 = None
          
                  # Başlık
                  self.lbl_title = tk.Label(root, text="Excel Yearly Comparison", font=("Arial", 18, "bold"), bg="#23272F", fg="#F5F5F5")
                  self.lbl_title.pack(pady=(20, 10))
          
                  # 1. Dosya
                  self.lbl_file1 = tk.Label(root, text="File 1: Not yet selected", font=("Arial", 13, "bold"), bg="#23272F", fg="#00FFCC", anchor="center")
                  self.lbl_file1.pack(pady=(20, 5))
                  self.btn_file1 = tk.Button(root, text="1. Select File", command=self.select_file1, width=35, font=("Arial", 11), bg="#444950", fg="#F5F5F5", activebackground="#FFD700")
                  self.btn_file1.pack()
          
                  # 2. Dosya
                  self.lbl_file2 = tk.Label(root, text="File 2: Not yet selected", font=("Arial", 13, "bold"), bg="#23272F", fg="#00FFCC", anchor="center")
                  self.lbl_file2.pack(pady=(20, 5))
                  self.btn_file2 = tk.Button(root, text="2. Select File", command=self.select_file2, width=35, font=("Arial", 11), bg="#444950", fg="#F5F5F5", activebackground="#FFD700")
                  self.btn_file2.pack()
          
                  # Karşılaştır Butonu
                  self.btn_compare = tk.Button(root, text="Compare and Save", command=self.compare_files, width=40, font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", activebackground="#FFD700")
                  self.btn_compare.pack(pady=25)
          
                  # Sonuç Kutusu
                  self.txt_result = tk.Text(root, height=8, width=70, font=("Consolas", 11), bg="#1A1C22", fg="#F5F5F5", borderwidth=2, relief="groove")
                  self.txt_result.pack(pady=(5, 10))
          
              def select_file1(self):
                  file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
                  if file_path:
                      self.file_1 = file_path
                      self.lbl_file1.config(text=f"File 1:: {os.path.basename(file_path)}")
          
              def select_file2(self):
                  file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
                  if file_path:
                      self.file_2 = file_path
                      self.lbl_file2.config(text=f"File 2:: {os.path.basename(file_path)}")
          
              def compare_files(self):
                  self.txt_result.delete(1.0, tk.END)
                  if not self.file_1 or not self.file_2:
                      self.txt_result.insert(tk.END, "Please select two files!\n")
                      return
                  try:
                      # Yıl eşlemesi
                      file_year_map = {}
                      for f in [self.file_1, self.file_2]:
                          year = get_dominant_year_from_file(f)
                          file_year_map[year] = f
          
                      current_year = max(file_year_map.keys())
                      reference_year = current_year - 1
          
                      df_current = load_valid_sheet(file_year_map[current_year])
                      df_current = normalize_columns(df_current, column_mapping)
                      df_current['date'] = pd.to_datetime(df_current['date'], errors='coerce')
                      df_current = df_current[df_current['date'].dt.year == current_year].copy()
          
                      df_previous = load_valid_sheet(file_year_map[reference_year])
                      df_previous = normalize_columns(df_previous, column_mapping)
                      df_previous['date'] = pd.to_datetime(df_previous['date'], errors='coerce')
                      df_previous = df_previous[df_previous['date'].dt.year == reference_year].copy()
          
                      #  TARİH, ÇEYREK VE AY (Quarter & Month hesaplamaları)
                      for df in [df_previous, df_current]:
                          df['date'] = pd.to_datetime(df['date'], errors='coerce')
                          df.dropna(subset=['date'], inplace=True)
                          df['Quarter'] = df['date'].dt.to_period('Q')
                          df['Month'] = df['date'].dt.to_period('M')
          
                      #  engagement_rate HESAPLAMA VEYA DÜZELTME
                      def ensure_engagement_rate(df, label):
                          if 'engagement' in df.columns and 'impressions' in df.columns:
                              df['engagement_rate'] = (df['engagement'] / df['impressions']) * 100
                              print(f" {label}: engagement_rate yeniden hesaplandı.")
                          elif 'engagement_rate' in df.columns:
                              mean_val = df['engagement_rate'].mean()
                              if mean_val < 1:  # oran formatında olabilir
                                  df['engagement_rate'] *= 100
                                  print(f" {label}: engagement_rate oran formatındaydı, % formatına çevrildi.")
                              else:
                                  print(f" {label}: engagement_rate zaten % formatında.")
                          else:
                              raise ValueError(f" {label}: Ne engagement ne de engagement_rate bulunamadı.")
          
                      # Her iki DataFrame'e uygula
                      ensure_engagement_rate(df_previous, "df_previous")
                      ensure_engagement_rate(df_current, "df_current")
          
                      # 1. Quarter’ı düzgün stringe çevir
                      df_current['Quarter_Label'] = df_current['date'].dt.to_period('Q').astype(str)
                      df_previous['Quarter_Label'] = df_previous['date'].dt.to_period('Q').astype(str)
          
                      # 2. Geçmiş yıl için çeyrek bazlı ortalamaları al
                      quarter_stats_prev = df_previous.groupby('Quarter_Label').agg({
                          'engagement_rate': 'mean',
                          'impressions': 'mean'
                      }).rename(columns={
                          'engagement_rate': 'Avg ER',
                          'impressions': 'Avg Impressions'
                      })
          
                      # 3. Ref Quarter etiketini üret (örneğin 2024Q1 → 2023Q1)
                      def get_ref_quarter_label(q_label):
                          try:
                              year, q = q_label.split('Q')
                              return f"{int(year) - 1}Q{q}"
                          except:
                              return None
          
                      df_current['Ref Quarter'] = df_current['Quarter_Label'].apply(get_ref_quarter_label)
          
                      # 4. Ref verileri eşleştir
                      df_current['Ref ER'] = df_current['Ref Quarter'].map(quarter_stats_prev['Avg ER'])
                      df_current['Ref Impressions'] = df_current['Ref Quarter'].map(quarter_stats_prev['Avg Impressions'])
          
                      # 5. Mutlak farkları hesapla
                      df_current['ER Diff'] = df_current['engagement_rate'] - df_current['Ref ER']
                      df_current['Impression Diff'] = df_current['impressions'] - df_current['Ref Impressions']
          
                      #  1. Farkları hesapla
                      df_current["ER Diff"] = df_current["engagement_rate"] - df_current["Ref ER"]
                      df_current["Impression Diff"] = df_current["impressions"] - df_current["Ref Impressions"]
          
                      #  2. Açıklamalı karşılaştırmaları oluştur
                      def er_aciklama(er, ref):
                          if pd.isna(er) or pd.isna(ref):
                              return "-"
                          if er > ref:
                              return f"{er:.2f} > {ref:.2f}"
                          elif er < ref:
                              return f"{er:.2f} < {ref:.2f}"
                          else:
                              return f"{er:.2f} = {ref:.2f}"
          
                      def imp_aciklama(imp, ref):
                          if pd.isna(imp) or pd.isna(ref):
                              return "-"
                          if imp > ref:
                              return f"{imp:.0f} > {ref:.0f}"
                          elif imp < ref:
                              return f"{imp:.0f} < {ref:.0f}"
                          else:
                              return f"{imp:.0f} = {ref:.0f}"
          
                      df_current["ER Comparison"] = df_current.apply(lambda row: er_aciklama(row["engagement_rate"], row["Ref ER"]), axis=1)
                      df_current["Impression Comparison"] = df_current.apply(lambda row: imp_aciklama(row["impressions"], row["Ref Impressions"]), axis=1)
          
                      #  3. Etiketleme 
              
                      def er_diff_etiket(diff):
                          if pd.isna(diff):
                              return "-"
                          elif diff >= 3:
                              return "Very High Interest"
                          elif diff >= 1:
                              return "High Interest"
                          elif diff > -1:
                              return "Average"
                          elif diff > -3:
                              return "Low Interest"
                          else:
                              return "Very Low Interest"
          
                      def imp_diff_etiket_v2(diff):
                          if pd.isna(diff):
                              return "-"
                          elif diff <= -2000:
                              return "Very Narrow Distribution"             
                          elif -1999 <= diff < -500:
                              return "Narrow Distribution"
                          elif -500 <= diff < 0:
                              return "Mild Contraction"
                          elif diff == 0:
                              return "Equal Distribution"
                          elif 0 < diff <= 500:
                              return "Slight Increase"
                          elif 500 < diff <= 1000:
                              return "Increased Distribution"
                          elif 1000 < diff <= 3000:
                              return "Wide Distribution"
                          else:
                              return "Very Wide Distribution"
          
                      df_current["ER Label"] = df_current["ER Diff"].apply(er_diff_etiket)
                      df_current["Impression Label"] = df_current["Impression Diff"].apply(imp_diff_etiket_v2)
          
                      df_current["ER Category (Detailed)"] = df_current["ER Diff"].apply(er_diff_etiket)
                      df_current["Impression Category (Detailed)"] = df_current["Impression Diff"].apply(imp_diff_etiket_v2)
          
                      desired_columns = [
                          "title", "date", "Quarter", "Month",
                          "impressions", "engagement_rate", "likes",
                          "Ref Quarter", "Ref ER", "Ref Impressions",
                          "ER Diff","ER Label","ER Comparison", "Impression Diff","Impression Label",
                          "Impression Comparison",
                             # Bu eski etiket varsa onu da en sona ekle
                      ]
          
                      # Yalnızca mevcut olanları sıralayarak al
                      final_columns = [col for col in desired_columns if col in df_current.columns]
                      df_current = df_current[final_columns]
          
                      from datetime import datetime
                      # Sonuç dosyalarını kaydet
                      output_dir = "output"
                      os.makedirs(output_dir, exist_ok=True)
                      output_path = os.path.join(output_dir, f"{current_year}_vs_{reference_year}_results.xlsx")
                      df_current.to_excel(output_path, index=False)
                      #  Grafik çizimlerini PNG olarak kaydet
                      save_pie_chart(df_current, "ER Label", "er_label_pie.png")
                      save_pie_chart(df_current, "Impression Label", "impression_label_pie.png")
          
                      #  Grafik görsellerini Excel'e göm
                      embed_chart_to_excel(output_path, "er_label_pie.png", sheet_name="Charts", cell="B2")
                      embed_chart_to_excel(output_path, "impression_label_pie.png", sheet_name="Charts", cell="L2")
                      # Renkli dosya
                      renklendir_excel(output_path)
                      renkli_path = output_path.replace(".xlsx", "_renkli.xlsx")
                      
                                      
                      
                      self.txt_result.insert(tk.END, f" Comparison completed!\n")
                      self.txt_result.insert(tk.END, f"Normal output file: {output_path}\n")
                      self.txt_result.insert(tk.END, f"Colorized output file: {renkli_path}\n")
                  except Exception as e:
                      self.txt_result.insert(tk.END, f"Error: {e}\n")
          
          if __name__ == "__main__":
              root = tk.Tk()
              app = ExcelComparerApp(root)
              root.mainloop()proje.py…]()



## 🚀 Features

- ✅ Automatically detects correct sheet and dominant year
- 📅 Quarterly comparison (e.g., Q1 vs Q2)
- 📈 Calculates Engagement Rate (ER) for each post
- 🎯 Classifies ER and Impression performance into human-readable labels
- 🎨 Applies custom Excel color formatting based on labels
- 📊 Generates bar charts for ER and Impression trends
- 🖥️ Converts to standalone `.exe` app for non-technical users



## 💻 Technologies Used

- Python 3.x
- pandas
- openpyxl
- matplotlib
- tkinter
- pyinstaller


## 💡 Why This Project?

This project was built during my internship to automate a repetitive manual analysis task.  
It evolved into a professional tool that can save time and increase consistency in performance reporting.

## 👨‍💻 Author

Mustafa Efe Kılıç  
[LinkedIn]([https://linkedin.com/in/Mustafa_Efe_Kılıç](https://www.linkedin.com/in/mustafa-efe-k%C4%B1l%C4%B1%C3%A7-25943925b?lipi=urn%3Ali%3Apage%3Ad_flagship3_profile_view_base_contact_details%3B9LohgOZJSEaCT1u5AJhr8g%3D%3D)) | [Medium](https://medium.com/@YOURUSERNAME)

---
