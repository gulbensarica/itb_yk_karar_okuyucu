import re
import pandas as pd
from pdfminer.high_level import extract_text
import os
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from pathlib import Path
import win32com.client as win32

def word2pdf(filedoc,temp):
    print( Path(filedoc))
    
    word = None
    doc = None
    if Path(filedoc).suffix in ['.doc', '.docx']:
        output = Path(temp+filedoc).with_suffix('.pdf')
        if output.exists():
            # pass
            return
        print("Convert WORD into PDF: "+filedoc)
        word = win32.DispatchEx('Word.Application')
        word.Visible = 0
        doc = word.Documents.Open(filedoc, False, False, True)
        # 'OutputFileName', 'ExportFormat', 'OpenAfterExport', 'OptimizeFor', 'Range',
        # 'From', 'To', 'Item', 'IncludeDocProps', 'KeepIRM', 'CreateBookmarks', 'DocStructureTags',
        # 'BitmapMissingFonts', 'UseISO19005_1', 'FixedFormatExtClassPtr'
        doc.ExportAsFixedFormat(
            filedoc.replace(".docx", '.pdf').replace('.doc','.pdf').replace('input','temp'),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1)
    if doc:
        doc.Close()
    if word:
        word.Quit()
 
def get_pdf_files(output_folder):
    pdf_files = []
    for file in os.listdir(output_folder):
        if file.endswith(".pdf"):
            pdf_files.append(os.path.join(output_folder, file))
    return pdf_files

def convert_to_text(pdf_path):
    pdf_text = extract_text(pdf_path)
    toplanti = re.split(r"\sSayfa\s+\d\s+\/\s+\d\s", pdf_text)
    toplanti = toplanti[0:len(toplanti)-1]
    x = r"Oy birliği ile karar verildi."
    text = []
    birlesmis_sayfa = ''

    for sayfa in toplanti:
        if x in sayfa:
            if birlesmis_sayfa:  
                text.append(birlesmis_sayfa + sayfa) 
                birlesmis_sayfa = ''
            else:
                text.append(sayfa) 
        else:
            birlesmis_sayfa += sayfa

    if birlesmis_sayfa:  
        text.append(birlesmis_sayfa)
    
    return text
"""
if __name__ == '__main__':
    folder_name = "Input"
    folder_path = os.path.join(os.getcwd(), folder_name)
    for doc_file in Path(folder_path).glob('*.doc?'):
        word2pdf(str(doc_file))
"""
def tarih(text):
    toplanti_tarihi =  r"Toplantı Tarihi:\s*?(\d{2}\/\d{2}\/\d{4})"
    tarihler = []
    tarih_eslesme = re.search(toplanti_tarihi, text) 
    if tarih_eslesme:
        toplanti_tarihi = tarih_eslesme.group(1)  
        #if toplanti_tarihi not in tarihler:
        tarihler.append(toplanti_tarihi)
    else:
        print("Tarih Bulunamadı.")
    if len(tarihler)>0:
        return tarihler[0]
    else:
        return ""

#pdf_files = get_pdf_files(folder_path)
#print(pdf_files)
def toplantı_no(text):
    toplanti_no =  r"Toplantı No:\s*(\d+)"
    no_eslesme = re.search(toplanti_no, text)  
    if no_eslesme:
        toplanti_no = no_eslesme.group(1)
    else: 
        print("Toplantı No Bulunamadı.")       
    return toplanti_no

def karar_no(text):
    kararno_pattern = r"Karar No:\s*(\d+)\s*"
    kararnolari = []
    karar_no_eslesme = re.search(kararno_pattern, text) 
    if karar_no_eslesme:
        kararno = karar_no_eslesme.group(1)
        kararnolari.append(kararno)
    else:
        print("Karar No bulunamadı.")
    if len(kararnolari)>0:
        return kararnolari[0]
    else:
        return "" 

def topic(text):
    konu = r"Konu\s+:\s+(.*?)\n"
    konular = []
    konu_eslesme = re.search(konu, text)
    if konu_eslesme:
        konu_adi= konu_eslesme.group(1)
        if konu_adi not in konular:
            konular.append(konu_adi)
    else:
        print("Konu Bulunamadı.")
    
    if len(konular)>0:
        return konular[0]
    else:
        return ""     

def kararlar(text):
    toplam_karar_sayisi = 0
    tum_kararlar= []
    karar = ''
    karar_sayisi = len(re.findall(r"Yönetim Kurulu Kararı:", text))
    toplam_karar_sayisi += karar_sayisi    
    if toplam_karar_sayisi == 1:
        karar_bul = re.search(r"Yönetim Kurulu Kararı:(.*?\s*Oy birliği ile karar verildi\.)", text, re.DOTALL)
        tum_kararlar.append(karar_bul.group(1))
    else:          
        regex = r'\s*İlgili Birim:.*?Yönetim Kurulu Kararı: .*? '
        regex2 = r'Oy birliği ile karar verildi\.(.*?Başkan.*?$)'
        temizlenmis_metin = re.sub(regex, '', text, flags=re.DOTALL)
        temizlenmis_metin2 = re.sub(regex2, 'Oy birliği ile karar verildi.', temizlenmis_metin, flags=re.DOTALL)
        tum_kararlar.append(temizlenmis_metin2)
    toplam_karar_sayisi = 0
    for her_parca in tum_kararlar:
        final_tum_kararlar=' '.join(tum_kararlar)
    return final_tum_kararlar
    
def yk_kararlari(pdf_path, all_data):
    toplanti = convert_to_text(pdf_path)
    for eleman in toplanti:           
        if len(toplanti) == 1:
            for item in toplanti:
                toplanti_tarihi = tarih(item)              
                toplanti_no = toplantı_no(item)
                karar_numarası = karar_no(item)
                konu = topic(item)
                ykkararlar = kararlar(item)

            data = {'Toplantı Tarihi': toplanti_tarihi,
                    'Toplantı No': toplanti_no,
                    'Karar No': karar_numarası,
                    'Konu': konu,
                    'Kararlar': ykkararlar}
            all_data.append(data)
            return all_data

        if len(toplanti) > 1:         
            for item in toplanti:          
                toplanti_tarihi = tarih(item)              
                toplanti_no = toplantı_no(item)
                karar_numarası = karar_no(item)
                konu = topic(item)
                ykkararlar = kararlar(item)

                data = {
                    'Toplantı Tarihi': toplanti_tarihi,
                    'Toplantı No': toplanti_no,
                    'Karar No': karar_numarası,
                    'Konu': konu,
                    'Kararlar': ykkararlar
                }
                all_data.append(data)
            return all_data


if __name__ == '__main__':

    input_folder = "input/"  # Word belgelerinin bulunduğu klasör
    temp_folder = "temp/"  # PDF dosyalarının kaydedileceği klasör 
    folder_path = os.path.join(os.getcwd(), input_folder)
    temp_path = os.path.join(os.getcwd(), temp_folder)

    for doc_file in Path(folder_path).glob('*.doc?'):
        word2pdf(str(doc_file),temp_path)

    # PDF dosyalarının yolu
    pdf_files = get_pdf_files(temp_folder)
# Tüm PDF dosyaları için yk_kararlari fonksiyonunu çalıştır
all_data = []
final_list=[]
for pdf_path in pdf_files:

    all_data = yk_kararlari(pdf_path,all_data)
    print(type(all_data))

# Önceden kaydedilmiş Excel dosyasını okuyun
excel_path = "YK Kararları.xlsx"
try:
    existing_df = pd.read_excel(excel_path)
except FileNotFoundError:
    existing_df = pd.DataFrame()

new_data_df = pd.DataFrame(all_data)
merged_df = pd.concat([existing_df, new_data_df])

merged_df['Toplantı No']=merged_df['Toplantı No'].astype(int)
merged_df['Karar No']=merged_df['Karar No'].astype(int)
merged_df=merged_df.drop_duplicates(subset=['Toplantı Tarihi', 'Toplantı No', 'Karar No'],keep='last')
merged_df.to_excel(excel_path, index=False)

print(merged_df)
try:
    os.remove('temp/')
except:
    pass




"""
# Sonuçları Excel dosyasına kaydet
df = pd.DataFrame(all_data)
print(df)
excel_path = "YK Kararları.xlsx"
df.to_excel(excel_path, index=False)
"""