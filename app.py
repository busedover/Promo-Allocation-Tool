
import io
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

  yorum_col = col_norm.get('YORUM') or col_norm.get('COMMENT')
  if yorum_col:
    df['yorum_clean'] = df[yorum_col].astype(str).str.strip().str.lower()
    df = df[~df['yorum_clean'].str.contains('alokasyonda|allocation', na=False)]

    mask_promo = df['yorum_clean'].str.contains(
        'promo only|sadece promosyon', na=False
    )
    for cust in customer_mapping.keys():
      actual_cust_col = next(
          (c for c in df.columns if normalize_text(c) == cust), None
      )
      if actual_cust_col:
        df.loc[mask_promo & df[actual_cust_col].isna(), actual_cust_col] = 0

    df = df.drop(columns=['yorum_clean'])

  existing_customers = [
      c for c in df.columns if normalize_text(c) in customer_mapping
  ]
  id_vars = [
      c
      for c in df.columns
      if normalize_text(c) in ['BARCODE', 'DESC', 'PFL', 'BRAND']
  ]

  df_melted = df.melt(
      id_vars=id_vars,
      value_vars=existing_customers,
      var_name='Customer_Name',
      value_name='Qty_reserved',
  )

  df_melted['Customer_Code'] = df_melted['Customer_Name'].apply(
      lambda x: customer_mapping.get(normalize_text(x), '')
  )

  output_df = pd.DataFrame()
  output_df['Sales Organization'] = ''
  output_df['Distribution Channel'] = ''
  output_df['Division'] = ''
  output_df['Plant'] = 'ZTR1'
  output_df['Storage Location'] = 'TL01'
  output_df['Customer'] = df_melted['Customer_Code']

  barcode_col = next(
      (c for c in df_melted.columns if normalize_text(c) in ['BARCODE', 'DESC']),
      df_melted.columns[0],
  )
  output_df['Material Number'] = df_melted[barcode_col]
  output_df['Valid-From Date'] = ''
  output_df['Valid-To Date'] = ''
  output_df['Qty reserved'] = df_melted['Qty_reserved']
  output_df['Unit of measure'] = 'UN'

  output_df['X'] = output_df['Qty reserved'].apply(
      lambda x: 'X' if pd.notna(x) and str(x).strip() not in ['', 'nan'] else ''
  )
  output_df = output_df[output_df['Qty reserved'].notna()]

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

    st.success('✅ Dosya başarıyla işlendi!')

    # Excel dosyasına dönüştür
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
  except Exception as e:
    st.error(f'⚠ Bir hata oluştu: {e}')
