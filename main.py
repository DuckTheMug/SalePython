import pandas as pd
import os
import matplotlib.pyplot as plt

#a
with pd.ExcelFile(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'DuLieuThucHanh2_V1.xlsx')) as reader:
    df_san_pham = pd.read_excel(reader, 'San Pham')
    df_nhan_vien = pd.read_excel(reader, 'Nhan Vien')
    df_hoa_don = pd.read_excel(reader, 'Hoa Don')
    df_thong_tin = pd.read_excel(reader, 'Thong Tin Hoa Don')
os.system('cls' if os.name == 'nt' else 'clear')
print('San Pham: ')
print(df_san_pham)
os.system('pause' if os.name == 'nt' else "/bin/bash -c 'read -s -n 1 -p \"Press any key to continue...\"'")
os.system('cls' if os.name == 'nt' else 'clear')
print('Nhan Vien: ')
print(df_nhan_vien)
os.system('pause' if os.name == 'nt' else "/bin/bash -c 'read -s -n 1 -p \"Press any key to continue...\"'")
os.system('cls' if os.name == 'nt' else 'clear')
print('Hoa Don: ')
print(df_hoa_don)
os.system('pause' if os.name == 'nt' else "/bin/bash -c 'read -s -n 1 -p \"Press any key to continue...\"'")
os.system('cls' if os.name == 'nt' else 'clear')
print('Thong Tin Hoa Don: ')
print(df_thong_tin)
os.system('pause' if os.name == 'nt' else "/bin/bash -c 'read -s -n 1 -p \"Press any key to continue...\"'")

#b
df = df_thong_tin.groupby(by=df_thong_tin['ID San Pham'], as_index=False, dropna=False)[['So Luong']].sum()
df_ban_hang = pd.merge(left=df_san_pham[['Ten', 'ID San Pham']], right=df, how='left', on='ID San Pham').fillna(0)
os.system('cls' if os.name == 'nt' else 'clear')
print('Ban Hang: ')
print(df_ban_hang)
os.system('pause' if os.name == 'nt' else "/bin/bash -c 'read -s -n 1 -p \"Press any key to continue...\"'")
os.system('cls' if os.name == 'nt' else 'clear')
print('Mat hang ban chay nhat: ')
print(df_ban_hang[df_ban_hang['So Luong'] == df_ban_hang['So Luong'].max()])
os.system('pause' if os.name == 'nt' else "/bin/bash -c 'read -s -n 1 -p \"Press any key to continue...\"'")

#c
df = pd.merge(left=df, right=df_san_pham[['ID San Pham', 'Gia']], on='ID San Pham')
doanhthu = (df['So Luong'] * df['Gia']).sum()
os.system('cls' if os.name == 'nt' else 'clear')
print(f'Doanh thu: {doanhthu}')
os.system('pause' if os.name == 'nt' else "/bin/bash -c 'read -s -n 1 -p \"Press any key to continue...\"'")

#d
df_san_pham['So Luong'] = df_san_pham['So Luong'] - df_ban_hang['So Luong']
os.system('cls' if os.name == 'nt' else 'clear')
print('San pham sau khi update so luong: ')
print(df_san_pham)
os.system('pause' if os.name == 'nt' else "/bin/bash -c 'read -s -n 1 -p \"Press any key to continue...\"'")

#e
df_thong_tin = df_thong_tin.groupby(by=['ID Hoa Don', 'ID San Pham'], as_index=False).sum()
os.system('cls' if os.name == 'nt' else 'clear')
print('Thong tin hoa don sau khi update trung lap: ')
print(df_thong_tin)
os.system('pause' if os.name == 'nt' else "/bin/bash -c 'read -s -n 1 -p \"Press any key to continue...\"'")

#g
with pd.ExcelWriter(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'DuLieuThucHanh2_V1_output.xlsx')) as writer:
    df_san_pham.to_excel(excel_writer=writer, sheet_name='San Pham', index=False)
    df_nhan_vien.to_excel(excel_writer=writer, sheet_name='Nhan Vien', index=False)
    df_hoa_don.to_excel(excel_writer=writer, sheet_name='Hoa Don', index=False)
    df_thong_tin.to_excel(excel_writer=writer, sheet_name='Thong Tin Hoa Don', index=False)
    df_ban_hang.to_excel(excel_writer=writer, sheet_name='San Pham Da Ban', index=False)
    
#f
df_ban_hang[['Ten', 'So Luong']].set_index('Ten').sort_values(by='So Luong', ascending=False).plot(kind='bar', title = 'Bieu do cho san pham da ban', figsize=(20, 15)).set_xlabel('Ten',fontsize = 8)
plt.savefig('Figure.svg')
plt.show()