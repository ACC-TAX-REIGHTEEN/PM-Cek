import os
import re
import pandas as pd

def normalize_text(text):
  if pd.isna(text):
    return ''
  clean_str = str(text).replace('\xa0', ' ').strip()
  return re.sub(r'\s+', ' ', clean_str)

def find_column_mapping(df_raw):
  col_mapping = {}
  top_rows = df_raw.iloc[:15]

  for c_idx in range(df_raw.shape[1]):
    col_cells = top_rows.iloc[:, c_idx].dropna().astype(str)
    combined_header = ' '.join(col_cells)
    norm = normalize_text(combined_header).lower()

    if 'tanggal' in norm and 'pajak' not in norm:
      col_mapping['Tanggal'] = c_idx
    elif 'tgl' in norm and 'pajak' in norm:
      col_mapping['Tgl. Pajak'] = c_idx
    elif 'pemasok' in norm:
      col_mapping['Nama Pemasok'] = c_idx
    elif 'referensi' in norm or 'ref' in norm:
      col_mapping['No. Referensi'] = c_idx
    elif 'faktur' in norm:
      col_mapping['No. Faktur Pajak'] = c_idx
    elif 'pajak' in norm and 'jumlah' in norm:
      col_mapping['Jumlah Pajak'] = c_idx

  return col_mapping

def clean_cell_value(val):
  if pd.isna(val):
    return ''
  s = normalize_text(val)
  if s.lower() == 'nan':
    return ''
  return s

def parse_and_split_pm(
    input_filepath='PM.xls', output_filepath='PM_temp.xlsx'
):
  if not os.path.exists(input_filepath):
    if os.path.exists('PM.xlsx'):
      input_filepath = 'PM.xlsx'
    else:
      raise FileNotFoundError(
          f"File '{input_filepath}' tidak ditemukan! Pastikan file berada di"
          ' folder yang sama.'
      )

  print(f"Membaca file input: '{input_filepath}'...")
  df_raw = pd.read_excel(input_filepath, header=None)

  col_map = find_column_mapping(df_raw)
  print('Kolom terdeteksi:', col_map)

  pdm_data = []
  ppn_data = []
  jasa_data = []

  current_section = None

  for idx, row in df_raw.iterrows():
    row_text = ' '.join(row.dropna().astype(str))
    norm_row = normalize_text(row_text)

    if 'PDM PPh Ps 23' in norm_row or '15.04-00' in norm_row:
      current_section = 'PDM'
      continue
    elif 'PPn Masukan (15.01)' in norm_row or (
        'PPn Masukan' in norm_row and '15.01' in norm_row
    ):
      current_section = 'PPN'
      continue
    elif 'Total dari P : PPN' in norm_row:
      current_section = 'AFTER_PPN'
      continue
    elif 'Transaksi Lainnya' in norm_row:
      if current_section in ['AFTER_PPN', 'PPN']:
        current_section = 'JASA'
        continue
      elif current_section is None:
        current_section = 'PDM'
        continue

    if any(
        keyword in norm_row
        for keyword in [
            'Total dari',
            'Saldo Awal',
            'Saldo Akhir',
            'Nama Pemasok',
            'Jumlah Pajak',
        ]
    ):
      continue

    c_tgl = col_map.get('Tanggal')
    c_tgl_pajak = col_map.get('Tgl. Pajak')
    c_pemasok = col_map.get('Nama Pemasok')
    c_ref = col_map.get('No. Referensi')
    c_faktur = col_map.get('No. Faktur Pajak')
    c_jml = col_map.get('Jumlah Pajak')

    val_tgl = (
        clean_cell_value(row[c_tgl])
        if c_tgl is not None and c_tgl < len(row)
        else ''
    )
    val_tgl_pajak = (
        clean_cell_value(row[c_tgl_pajak])
        if c_tgl_pajak is not None and c_tgl_pajak < len(row)
        else ''
    )
    val_pemasok = (
        clean_cell_value(row[c_pemasok])
        if c_pemasok is not None and c_pemasok < len(row)
        else ''
    )
    val_ref = (
        clean_cell_value(row[c_ref])
        if c_ref is not None and c_ref < len(row)
        else ''
    )
    val_faktur = (
        clean_cell_value(row[c_faktur])
        if c_faktur is not None and c_faktur < len(row)
        else ''
    )
    val_jml = (
        clean_cell_value(row[c_jml])
        if c_jml is not None and c_jml < len(row)
        else ''
    )

    if not val_tgl and not val_ref:
      continue

    if 'tanggal' in val_tgl.lower() or 'tanggal' in val_tgl_pajak.lower():
      continue

    item = {
        'Tanggal': val_tgl,
        'Tgl. Pajak': val_tgl_pajak,
        'Nama Pemasok': val_pemasok,
        'No. Referensi': val_ref,
        'No. Faktur Pajak': val_faktur,
        'Jumlah Pajak': val_jml,
    }

    if current_section == 'PDM':
      pdm_data.append(item)
    elif current_section == 'PPN':
      ppn_data.append(item)
    elif current_section == 'JASA':
      jasa_data.append(item)

  target_columns = [
      'Tanggal',
      'Tgl. Pajak',
      'Nama Pemasok',
      'No. Referensi',
      'No. Faktur Pajak',
      'Jumlah Pajak',
  ]

  df_pdm = pd.DataFrame(pdm_data, columns=target_columns)
  df_ppn = pd.DataFrame(ppn_data, columns=target_columns)
  df_jasa = pd.DataFrame(jasa_data, columns=target_columns)

  with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
    df_pdm.to_excel(writer, sheet_name='PDM PPh Ps 23', index=False)
    df_ppn.to_excel(writer, sheet_name='PPn Masukan (15.01)', index=False)
    df_jasa.to_excel(writer, sheet_name='Jasa', index=False)

  print('\nProses pembersihan dan pemisahan selesai!')
  print(f"File hasil telah disimpan di: '{output_filepath}'")
  print(f" - Sheet 'PDM PPh Ps 23'       : {len(df_pdm)} baris")
  print(f" - Sheet 'PPn Masukan (15.01)'  : {len(df_ppn)} baris")
  print(f" - Sheet 'Jasa'                 : {len(df_jasa)} baris")

if __name__ == '__main__':
  parse_and_split_pm()