import pandas as pd

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

def generate_promo_allocation_excel(
    input_file, output_file='promo_alokasyon_cikti.xlsx'
):
  # Müşteri Kodları Mapping Sözlüğü
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

  print(f'Excel dosyası okunuyor: {input_file}...')
  df = pd.read_excel(input_file)
  df.columns = df.columns.astype(str).str.strip()
  col_norm = {normalize_text(c): c for c in df.columns}

  # 1. Yorum filtresi ("alokasyonda" olanları çıkar)
  yorum_col = col_norm.get('YORUM') or col_norm.get('COMMENT')
  if yorum_col:
    df['yorum_clean'] = df[yorum_col].astype(str).str.strip().str.lower()
    df = df[~df['yorum_clean'].str.contains('alokasyonda|allocation', na=False)]

    # 2. Promo Only kuralı: Promo only olanlarda boş hücrelere 0 yaz
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

  # 3. Yatay tabloyu alt alta Excel formatına dönüştürme (Unpivot)
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

  # Excel Çıktı Şablonunu Oluştur
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

  # X sütunu kuralı: Qty doluysa 'X', boşsa boş bırak
  output_df['X'] = output_df['Qty reserved'].apply(
      lambda x: 'X' if pd.notna(x) and str(x).strip() not in ['', 'nan'] else ''
  )

  # Sadece Qty dolu olan satırları filtrele
  output_df = output_df[output_df['Qty reserved'].notna()]

  output_df.to_excel(output_file, index=False)
  print(f'✅ Alokasyon Excel dosyası başarıyla oluşturuldu: {output_file}')

if __name__ == '__main__':
  # Test için: generate_promo_allocation_excel("ham_veri.xlsx")
  pass
