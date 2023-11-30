from vn_fullname_generator import generator
import shortuuid
import numpy as np
import random
import pandas as pd
import os

lstReID = []  # ID Hoa Don
lstIDEm = []  # ID Nhan Vien
lstIDIt = []  # ID San Pham
lstNEm = []  # Ten Nhan Vien
lstNIt = []  # Ten San Pham
lstIIt = []  # So Luong San Pham
lstSIt = []  # So Luong Hoa Don
lstP = []  # Gia
re = []  # List hoa don trung gian
reInfo = []  # List thong tin hoa don trung gian

for i in range(50):
    lstReID.append(shortuuid.ShortUUID().random(8))
    lstIDEm.append(shortuuid.ShortUUID().random(8))
    lstIDIt.append(shortuuid.ShortUUID().random(8))
    lstNEm.append(generator.generate())
    lstNIt.append('Sản phẩm {num}'.format(num=str(np.random.randint(500))))
    lstIIt.append(np.random.randint(500, 1000))
    lstSIt.append(np.random.randint(10, 50))
    lstP.append(np.random.randint(5, 10) * 10000)

for i in lstReID:
    re.append((i, random.choice(lstIDEm)))

for i in lstReID:
    reInfo.append((i, random.choice(lstIDIt)))

### SPECIAL SEEDING ###
for i in range(np.random.randint(2,5)):
    seed = np.random.randint(0, 50)
    for j in range(np.random.randint(1,4)):
        reInfo[np.random.randint(0,50)] = reInfo[seed]
     

dfSP = pd.DataFrame({'ID San Pham': lstIDIt,
                     'Ten': lstNIt,
                     'So Luong': lstIIt,
                     'Gia': lstP})

dfNV = pd.DataFrame({'ID Nhan Vien': lstIDEm,
                     'Ten': lstNEm})


dfHD = pd.DataFrame(re, columns=['ID Hoa Don', 'ID Nhan Vien'])
dfTTHD = pd.concat(
    [pd.DataFrame(reInfo, columns=['ID Hoa Don', 'ID San Pham']), pd.DataFrame(lstSIt, columns=['So Luong'])], axis=1)

with pd.ExcelWriter(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'DuLieuThucHanh2_V1.xlsx')) as writer:
    dfSP.to_excel(excel_writer=writer, sheet_name='San Pham', index=False)
    dfNV.to_excel(excel_writer=writer, sheet_name='Nhan Vien', index=False)
    dfHD.to_excel(excel_writer=writer, sheet_name='Hoa Don', index=False)
    dfTTHD.to_excel(excel_writer=writer, sheet_name='Thong Tin Hoa Don', index=False)
