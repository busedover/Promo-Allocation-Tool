
from datetime import datetime
import io
import calendar
import pandas as pd
import streamlit as st

def normalize_text(text):
  if pd.isna(text):
    return ''
  text = str(text).strip().upper()
  return (
      text.replace('İ', 'I')
      .replace('Ş', 'S')
      .replace('Ğ', 'G')
      .replace('Ü', 'U')
      .replace('Ö', 'O')
      .replace('Ç', 'C')
  )

def process_allocation(input_file):
  customer_mapping = {
      'AFILI': '40000719',
      'AMAZON': '40000809',
      'A101': '40000143',
      'BIM': '40000142',
      'CARREFOUR': '40000136',
      'CASH & CARRY': '40000590',
      'DOGUS': '40000156',
      'EVE': '40000570',
      'FILE': '40000566',
      'GRATIS': '40000146',
      'DISTRIBUTOR': '40000004',
      'LOCAL PERFUMERY': '40000572',
      'MIGROS': '40000133',
      'ROSSMANN': '40000148',
      'RKA': '40000594',
      'SOK': '40000135',
      'SALDOS': '40000720',
      'WATSONS': '40000145',
  }

  df = pd.read_excel(input_file)
  df.columns = df.columns.astype(str).str.strip()
  col_norm = {normalize_text(c): c for c in df.columns}

  # Kritik sütunları bul
  bar_customer_col = col_norm.get('BARCUSTOMER')
  barcode_col = col_norm.get('BARCODE')
  yorum_col = col_norm.get('YORUM') or col_norm.get('COMMENT')

  # Adet sütununu bul (Barcode, BarCustomer ve Yorum haricindeki ilk sayısal veya tarih sütunu)
  qty_col = None
  for c in df.columns:
    norm_c = normalize_text(c)
    if norm_c not in [
        'BARCUSTOMER',
        'BARCODE',
        'DESC',
        'PFL',
        'BRAND',
        'YORUM',
        'COMMENT',
    ]:
      qty_col = c
      break

  if not bar_customer_col or not barcode_col or not qty_col:
    st.error(
        '⚠ Dosyada BarCustomer, Barcode veya miktar sütunu bulunamadı! Sütun'
        ' adlarını kontrol edin.'
    )
    return pd.DataFrame()

  # 1. Yorum filtresi ("alokasyonda" olanları çıkar)
  if yorum_col:
    df['yorum_clean'] = df[yorum_col].astype(str).str.strip().str.lower()
    df = df[~df['yorum_clean'].str.contains('alokasyonda|allocation', na=False)]
    df = df.drop(columns=['yorum_clean'])

  # Boş veya 0 adetleri ele
  df = df[df[qty_col].notna()]
  df = df[df[qty_col] != 0]

  # Tarih hesaplamaları
  today = datetime.today()
  valid_from = today.strftime('%d.%m.%Y')
  last_day = calendar.monthrange(today.year, today.month)[1]
  valid_to = datetime(today.year, today.month, last_day).strftime('%d.%m.%Y')

  # Çıktı formatını oluştur
  output_df = pd.DataFrame()
  output_df['Sales Organization'] = ''
  output_df['Distribution Channel'] = ''
  output_df['Division'] = ''
  output_df['Plant'] = 'ZTR1'
  output_df['Storage Location'] = 'TL01'

  # Müşteri kodunu mapping'den al
  output_df['Customer'] = df[bar_customer_col].apply(
      lambda x: customer_mapping.get(normalize_text(x), str(x))
  )

  output_df['Material Number'] = df[barcode_col]
  output_df['Valid-From Date'] = valid_from
  output_df['Valid-To Date'] = valid_to
  output_df['Qty reserved'] = df[qty_col]
  output_df['Unit of measure'] = 'UN'
  output_df['X'] = 'X'

  return output_df

# Streamlit Arayüzü
st.set_page_config(page_title='Promo Alokasyon Tool', layout='centered')
st.title('🎯 Promo Alokasyon Tool')
st.write('Lütfen ham Excel dosyanızı yükleyin.')

uploaded_file = st.file_uploader('Excel Dosyası Seçin', type=['xlsx'])

if uploaded_file is not None:
  try:
    with st.spinner('Dosya işleniyor...'):
      result_df = process_allocation(uploaded_file)

    if not result_df.empty:
      st.success(f'✅ Dosya başarıyla işlendi! ({len(result_df)} satır)')

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False)
      processed_data = output.getvalue()

      st.download_button(
          label='📥 İşlenmiş Excel Dosyasını İndir',
          data=processed_data,
          file_name='promo_alokasyon_cikti.xlsx',
          mime=(
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          ),
      )
    else:
      st.warning('⚠ Dışarı aktarılacak uygun veri bulunamadı.')
  except Exception as e:
    st.error(f'⚠ Bir hata oluştu: {e}')
