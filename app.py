
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
      'AFILI KOZMETIK': '40000719',
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
      'TRADITIONAL TRADE': '40000004',
      'LOCAL PERFUMERY': '40000572',
      'MIGROS': '40000133',
      'MION': '40000133',
      'MİON': '40000133',
      'ROSSMANN': '40000148',
      'RKA': '40000594',
      'SOK': '40000135',
      'SALDOS': '40000720',
      'WATSONS': '40000145',
  }

  df = pd.read_excel(input_file)
  df.columns = df.columns.astype(str).str.strip()
  col_norm = {normalize_text(c): c for c in df.columns}

  bar_customer_col = col_norm.get('BARCUSTOMER')
  barcode_col = col_norm.get('BARCODE')
  yorum_col = col_norm.get('YORUM') or col_norm.get('COMMENT')

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
        '⚠ Dosyada BarCustomer, Barcode veya miktar sütunu bulunamadı!'
    )
    return pd.DataFrame()

  if yorum_col:
    df['yorum_raw'] = df[yorum_col].astype(str)
    df['yorum_clean'] = df['yorum_raw'].str.strip().str.lower()
    df = df[~df['yorum_clean'].str.contains('alokasyonda|allocation', na=False)]
  else:
    df['yorum_raw'] = ''
    df['yorum_clean'] = ''

  processed_rows = []
  unique_barcodes = df[barcode_col].unique()

  for barcode in unique_barcodes:
    prod_df = df[df[barcode_col] == barcode]

    is_promo_only = prod_df['yorum_clean'].str.contains(
        'promo only|sadece promosyon', na=False
    ).any()
    is_migros_exc = prod_df['yorum_clean'].str.contains('migros exc|exc', na=False).any()

    if is_migros_exc:
      # Migros EXC kuralı: O müşteri quotasız adet alır, diğerleri quotalı 0 alır
      for _, row in prod_df.iterrows():
        processed_rows.append({
            'Customer_Name': row[bar_customer_col],
            'Barcode': row[barcode_col],
            'Qty': row[qty_col],
            'Comment': 'migros exc',  # quotasız işaretçisi
        })

      existing_customers = prod_df[bar_customer_col].apply(
          lambda x: normalize_text(x)
      ).tolist()

      for cust_name in customer_mapping.keys():
        if cust_name in ['MION', 'MİON']:
          continue
        found = any(
            normalize_text(cust_name) == normalize_text(ec)
            for ec in existing_customers
        )
        if not found:
          processed_rows.append({
              'Customer_Name': cust_name,
              'Barcode': barcode,
              'Qty': 0,
              'Comment': 'promo only',  # quotalı 0 işaretçisi
          })

    elif is_promo_only:
      existing_customers = prod_df[bar_customer_col].apply(
          lambda x: normalize_text(x)
      ).tolist()

      for _, row in prod_df.iterrows():
        processed_rows.append({
            'Customer_Name': row[bar_customer_col],
            'Barcode': row[barcode_col],
            'Qty': row[qty_col],
            'Comment': row['yorum_clean'],
        })

      for cust_name in customer_mapping.keys():
        if cust_name in ['MION', 'MİON']:
          continue
        found = any(
            normalize_text(cust_name) == normalize_text(ec)
            for ec in existing_customers
        )
        if not found:
          processed_rows.append({
              'Customer_Name': cust_name,
              'Barcode': barcode,
              'Qty': 0,
              'Comment': 'promo only',
          })
    else:
      for _, row in prod_df.iterrows():
        processed_rows.append({
            'Customer_Name': row[bar_customer_col],
            'Barcode': row[barcode_col],
            'Qty': row[qty_col],
            'Comment': row['yorum_clean'],
        })

  final_df = pd.DataFrame(processed_rows)

  final_df['Customer_Code'] = final_df['Customer_Name'].apply(
      lambda x: customer_mapping.get(normalize_text(x), str(x))
  )

  # Aynı ürün ve müşteri için adetleri topla
  final_df = (
      final_df.groupby(['Barcode', 'Customer_Code', 'Comment'], as_index=False)
      ['Qty']
      .sum()
  )

  today = datetime.today()
  valid_from = today.strftime('%d.%m.%Y')
  last_day = calendar.monthrange(today.year, today.month)[1]
  valid_to = datetime(today.year, today.month, last_day).strftime('%d.%m.%Y')

  output_df = pd.DataFrame()
  output_df['Sales Organization'] = ''
  output_df['Distribution Channel'] = ''
  output_df['Division'] = ''
  output_df['Plant'] = 'ZTR1'
  output_df['Storage Location'] = 'TL01'

  output_df['Customer'] = final_df['Customer_Code']
  output_df['Material Number'] = final_df['Barcode'].astype(str)
  output_df['Valid-From Date'] = valid_from
  output_df['Valid-To Date'] = valid_to
  output_df['Qty reserved'] = final_df['Qty']
  output_df['Unit of measure'] = 'UN'

  # X Sütunu Kuralı: Migros EXC olanlar quotasız (boş), Promo Only olanlar quotalı ('X')
  def determine_x(comment):
    c = str(comment).lower()
    if 'migros exc' in c or 'exc' in c:
      return ''
    elif 'promo only' in c or 'promosyon' in c:
      return 'X'
    return ''

  output_df['X'] = final_df['Comment'].apply(determine_x)

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
