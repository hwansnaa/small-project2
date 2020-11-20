import pytesseract
from PIL import Image
import numpy as np
import time
import pandas as pd
import timeit
from tqdm import trange
import math
# tesseract 경로
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'


def isNumber(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def Have(s):
    try:
        Image.open(s)
        return True
    except PermissionError:
        return False
    except FileNotFoundError:
        return False

def isnan(value):
    try:
        return math.isnan(float(value))
    except:
        return False

# 원하는 파일의 리스트 이름이 적힌 Excel의 경로를 읽어온다.
df = pd.read_excel(input())


start = timeit.default_timer()
test = df.shape[0]
for testnumber in trange(test):
    #이미 검수가 완료된 경우에 대한 예외처리
    if (isnan(df['correct'][testnumber]) == False):
        continue
    #파일이 없는 경우에 대한 에러 검출
    if ((Have(df['location'][testnumber]+'\\'+df['name'][testnumber]+'.png') == False )) :
        #df['correct'][testnumber] = 'X'
        continue
    _text = pytesseract.image_to_string(Image.open(df['location'][testnumber]+'\\'+df['name'][testnumber]+'.png'),lang = 'kor')
    _textsplit = _text.split()
    #testcase = ['파일','폴더']
    #추후에 정확도를 높일 때
    param2 = _textsplit.count('파일')
    ARRAY1 = np.arange(param2)
    for i in range(len(ARRAY1)):
        if (i>=1):
            ARRAY1[i] = _textsplittemp.index('파일')
            _textsplittemp = _textsplittemp[ARRAY1[i]+1:]
            ARRAY1[i] += ARRAY1[i-1]+1
        else:
            ARRAY1[i] = _textsplit.index('파일')
            _textsplittemp = _textsplit[ARRAY1[i]+1:]

    param3 = _textsplit.count('일')
    ARRAY2 = np.arange(param3)
    for i in range(len(ARRAY2)):
        if (i>=1):
            ARRAY2[i] = _textsplittemp.index('일')
            _textsplittemp = _textsplittemp[ARRAY2[i]+1:]
            ARRAY2[i] += ARRAY2[i-1]+1
        else:
            ARRAY2[i] = _textsplit.index('일')
            _textsplittemp = _textsplit[ARRAY2[i]+1:]


    VALUE1 = []
    for i in range(len(ARRAY1)):
        _textsplit[ARRAY1[i]+1] = _textsplit[ARRAY1[i]+1].replace(',','')
        if isNumber(_textsplit[ARRAY1[i]+1])==True :
            VALUE1.append(int(_textsplit[ARRAY1[i]+1]))

    for i in range(len(ARRAY2)):
        _textsplit[ARRAY2[i]+1] = _textsplit[ARRAY2[i]+1].replace(',','')
        if isNumber(_textsplit[ARRAY2[i]+1])==True :
            VALUE1.append(int(_textsplit[ARRAY2[i]+1]))

    if VALUE1.count(df['count'][testnumber]) == 2:
        df['correct'][testnumber] = 'o'
    # else:
    #     df['correct'][testnumber] = 'X'
df.to_excel('./output/Result.xlsx')
stop = timeit.default_timer()
print(stop - start)


# # 예외처리
# 1. 이미 검수가 완료 된 경우
# 2. 경로에 캡쳐파일이 없는 경우
# 3. 읽어온 파일에 해당 컬럼에 기입된 숫자가 2개가 아닌 경우
#  - 캡쳐본이 한개
#  - 캡쳐본은 두개이지만, 서로의 숫자가 다르거나, 엑셀에 기입된 숫자와 매치되지 않는 경우
#  - 정해진 규격에 맞지 않는 경우
