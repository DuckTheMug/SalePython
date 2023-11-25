import pandas as pd
import os

#a
with pd.ExcelFile(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'DuLieuThucHanh2_V1.xlsx')) as reader:
    df_san_pham = pd.read_excel(reader, 'San Pham')
    df_nhan_vien = pd.read_excel(reader, 'Nhan Vien')
    df_hoa_don = pd.read_excel(reader, 'Hoa Don')
    df_thong_tin = pd.read_excel(reader, 'Thong Tin Hoa Don')

print('San Pham: ')
print(df_san_pham)
print('Nhan Vien: ')
print(df_nhan_vien)
print('Hoa Don: ')
print(df_hoa_don)
print('Thong Tin Hoa Don: ')
print(df_thong_tin)

#b
df = df_thong_tin.groupby(df_thong_tin['ID San Pham'], as_index=False, dropna=False)[['So Luong']].sum()
df_ban_hang = pd.merge(left=df_san_pham[['Ten', 'ID San Pham']], right=df, how='left', on='ID San Pham').fillna(0)
print('Ban Hang: ')
print(df_ban_hang)
print('Mat hang ban chay nhat: ')
print(df_ban_hang[df_ban_hang['So Luong'] == df_ban_hang['So Luong'].max()])

#c
df = pd.merge(left=df, right=df_san_pham[['ID San Pham', 'Gia']], on='ID San Pham')
print('Doanh thu: ')
print((df['So Luong'] * df['Gia']).sum())

#d
df_san_pham['So Luong'] = df_san_pham['So Luong'] - df_ban_hang['So Luong']
print('San pham sau khi update so luong: ')
print(df_san_pham)

#e
df_thong_tin = df_thong_tin.groupby(by=['ID Hoa Don', 'ID San Pham'], as_index=False).sum()
print(df_thong_tin)

#f

### TODO ###

#g
with pd.ExcelWriter(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'DuLieuThucHanh2_V1_output.xlsx')) as writer:
    df_san_pham.to_excel(excel_writer=writer, sheet_name='San Pham', index=False)
    df_nhan_vien.to_excel(excel_writer=writer, sheet_name='Nhan Vien', index=False)
    df_hoa_don.to_excel(excel_writer=writer, sheet_name='Hoa Don', index=False)
    df_thong_tin.to_excel(excel_writer=writer, sheet_name='Thong Tin Hoa Don', index=False)
    df_ban_hang.to_excel(excel_writer=writer, sheet_name='San Pham Da Ban', index=False)
    