import datetime
import math
import pandas as pd
import numpy as np

INPUT_FILENAME = 'input.xlsx'
OUTPUT_FILENAME = 'sampling.xlsx'
NO_URUT_FILENAME = 'no_urut.csv'

def load_excel_from_local():
  return pd.read_excel(INPUT_FILENAME, sheet_name=0), pd.read_excel(INPUT_FILENAME, sheet_name=1)

def load_csv_from_local():
  return pd.read_csv(NO_URUT_FILENAME)

def eomonth(d, months=0):
  months = int(months)
  y, m = divmod(d.month + months + 1, 12)
  if m == 0:
    y -= 1
    m = 12
  return datetime.datetime(d.year + y, m, 1) - datetime.timedelta(days = 1)

def get_params_from_df(params_df):
  return {
    'TanggalMulai': params_df.iat[0,1],
    'BulanNaikGaji': params_df.iat[2,1],
    'PersenKenaikanGaji': params_df.iat[1,1],
    'PersenSetoran': params_df.iat[4,1],
    'PersenTingkatBunga': params_df.iat[3,1],
    'PersenBiaya': params_df.iat[5,1]
  }

def get_naik_gaji(d, month):
  return 1 if d.month == month else 0

def get_kenaikan_gaji(naik_gaji, prsn_kenaikan_gaji):
  return prsn_kenaikan_gaji if naik_gaji == 1 else 0

def hitung_iuran(gaji, prsn_setoran):
  return gaji * prsn_setoran

def hitung_imbal_hasil(tingkat_bunga, saldo):
  return (tingkat_bunga / 12) * saldo

def hitung_biaya(prsn_biaya, saldo):
  return -1 * (prsn_biaya * saldo)

def hitung_saldo_akhir(saldo_awal, iuran, imbal_hasil, biaya):
  return saldo_awal + iuran + imbal_hasil + biaya

def get_usia_awal(tgl_mulai, tgl_lahir):
  ttgl_mulai=(12*tgl_mulai.year + tgl_mulai.month)
  ttgl_lahir=(12*tgl_lahir.year + tgl_lahir.month)
  return (ttgl_mulai-ttgl_lahir)/12

def update_data(dt, params):
  dt['Iuran'] = hitung_iuran(dt['Gaji'], params['PersenSetoran'])
  dt['ImbalHasil'] = hitung_imbal_hasil(params['PersenTingkatBunga'], dt['SaldoAwal'])
  dt['Biaya'] = hitung_biaya(params['PersenBiaya'], dt['SaldoAwal'])
  dt['SaldoAkhir'] = hitung_saldo_akhir(dt['SaldoAwal'], dt['Iuran'], dt['ImbalHasil'], dt['Biaya'])

def init_data_awal(input_df, params):
  dt = {
    'No':1,
    'Tanggal': params['TanggalMulai'],
    'Usia': get_usia_awal(params['TanggalMulai'], input_df['TglLahir']),
    'SaldoAwal': input_df['SaldoSekarang']
  }
  dt['Gaji'] = input_df['GajiSebulan']
  dt['NaikGaji'] = get_naik_gaji(dt['Tanggal'], params['BulanNaikGaji'])
  dt['KenaikanGaji'] = get_kenaikan_gaji(dt['NaikGaji'], params['PersenKenaikanGaji'])
  update_data(dt, params)
  return [dt]

def get_process_df(input_df, params):
  arr_dt = init_data_awal(input_df, params)
  usia_awal = arr_dt[0]['Usia']
  i = arr_dt[0]['No']
  usia = usia_awal
  while(usia < input_df['UsiaAkhir']):
    last_el = arr_dt[i-1]
    usia = usia_awal + (i) / 12
    i += 1
    dt = {     
      'No': i,     
      'Tanggal': eomonth(arr_dt[0]['Tanggal'], (i-1)),
      'Usia': usia,
      'SaldoAwal': last_el['SaldoAkhir']
    }
    dt['NaikGaji'] = get_naik_gaji(dt['Tanggal'], params['BulanNaikGaji'])
    dt['KenaikanGaji'] = get_kenaikan_gaji(dt['NaikGaji'], params['PersenKenaikanGaji'])
    dt['Gaji'] = last_el['Gaji'] + (last_el['Gaji'] * dt['KenaikanGaji'])
    update_data(dt, params)
    arr_dt.append(dt)
  df = pd.DataFrame.from_records(arr_dt)
  return df


if __name__ == '__main__':
  inputdf, paramsdf = load_excel_from_local()
  nourutdf = load_csv_from_local()
  params = get_params_from_df(paramsdf)
  with pd.ExcelWriter(OUTPUT_FILENAME) as writer:
    for index, dt in nourutdf.iterrows():
      row = inputdf.iloc[index]
      df = get_process_df(row, params)
      df.to_excel(writer, sheet_name=row['Nomor ID'], index=False)
  