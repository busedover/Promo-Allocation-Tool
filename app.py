
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

  # Yorum sütununu bul ve temizle
  yorum_col = None
  for c in df.columns:
    if 'yorum' in normalize_text(c) or 'comment' in normalize_text(c):
      yorum_col = c
      break

  if yorum_col:
    df['yorum_clean'] = df[yorum_col].astype(str).str.strip().str.lower()
    df = df[~df['yorum_clean'].str.contains('alokasyonda|allocation', na=False)]

    mask_promo = df['yorum_clean'].str.contains(
        'promo only|sadece promosyon', na=False
    )
    for cust_key in customer_mapping.keys():
      for c in df.columns:
        if normalize_text(c) == cust_key:
          df.loc[mask_promo & df[c].isna(), c] = 0

    df = df.drop(columns=['yorum_clean'])

  # Müşteri ve Barkod sütunlarını belirle
  existing_customers = [
      c for c in df.columns if normalize_text(c) in customer_mapping
  ]
  barcode_col = next(
      (c for c in df.columns if normalize_text(c) == 'BARCODE'), None
  )

  if not existing_customers or not barcode_col:
    st.error('⚠ Dosyada Barcode veya müşteri sütunları bulunamadı!')
    return pd.DataFrame()

  # Unpivot işlemi
  df_melted = df.melt(
      id_vars=[barcode_col],
      value_vars=existing_customers,
      var_name='Customer_Name',
      value_name='Qty_reserved',
  )

  # Adeti boş olan veya 0 olan satırları ele
  df_melted = df_melted[df_melted['Qty_reserved'].notna()]
  df_melted = df_melted[df_melted['Qty_reserved'] != 0]

  # Tarih hesaplamaları (Bugün ve Ayın Son Günü)
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
  output_df['Customer'] = df_melted['Customer_Name'].apply(
      lambda x: customer_mapping.get(normalize_text(x), '')
  )
  output_df['Material Number'] = df_melted[barcode_col]
  output_df['Valid-From Date'] = valid_from
  output_df['Valid-To Date'] = valid_to
  output_df['Qty reserved'] = df_melted['Qty_reserved']
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
