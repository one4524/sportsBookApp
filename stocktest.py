import numpy as np
import pandas as pd
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt

import win32com.client

instCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')

Connect = instCpCybos.IsConnect

if (Connect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 객체생성

instCpStockCode = win32com.client.Dispatch('CpUtil.CpStockCode')  # 주식 코드 가져오기

# 종목 코드 가져오기: CpCodeMgr 클래스
instCpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
codeList = instCpCodeMgr.GetStockListByMarket(1)  # 코스피(거래소)
industryCodeList = instCpCodeMgr.GetIndustryList()  # 업종 코드 리스트

# CpSysDib 모듈의 MarketEye 클래스
instMarketEye = win32com.client.Dispatch('CpSysDib.MarketEye')

#########################################################################################################

df = pd.read_csv("D:\stockdata\electronic.csv", encoding='euc-kr')  # csv 파일 가져오기

df.fillna(0, inplace=True)  # 결손값 0으로 치환 - 비어있거나 n/a값이 모두 0

stockcodelist = []
stocknamelist = []

data = np.array(df)

start = 1
end = 651
# 651은 전기전자업 csv파일의 열의 길이, 921은 서비스업, 1291은 금융업

# csv파일에서 값이 문자열인 항목에서 쉼표를 제거하고 float형으로 바꾸는 반복문
for i in range(6, 26):  # 6 ~ 25 행이 실제 값들이 있는 행
    for j in range(start, end):  # 650은 이 csv파일의 열의 길이
        z = data[i, j]
        if type(z) == str:
            float_z = z.replace(',', '')
            data[i, j] = float(float_z)



# csv파일에서 코드와 이름 가져오기
for i in range(start, end, 10):  # 1부터 값이 시작하고 651까지 있음, 각 종목은 10개의 변수를 가짐
    stockcodelist.append(data[0, i])  # 리스트에 포함되어 있는 종목 코드 리스트
    stocknamelist.append(data[1, i])  # 리스트에 포함되어 있는 종목 이름 리스트

for i in range(start, end, 10):
    if data[25, i + 1] == 0:            # 2개 변수 이상이 비어있는 종목은 패스
        if data[25, i + 2] == 0 | data[25, i + 3]:
            continue

    print(data[1, i])
    # 대신증권 API를 통해 해당 종목의 정보를 가져오기
    instMarketEye.SetInputValue(0, (94, 90, 75, 89, 70, 77))
    # 94 - 이자보상비율, 90 - 영업이익증가율, 75 - 부채비율, 89 - BPS, 70 - EPS, 77 - ROE
    instMarketEye.SetInputValue(1, data[0, i])

    instMarketEye.BlockRequest()     # 서버에 정보 요청

    interest = instMarketEye.GetDataValue(0, 0)   # 이자보상비율

    profit_growth = instMarketEye.GetDataValue(1, 0)    # 영업이익증가율
    debt = instMarketEye.GetDataValue(2, 0)     # 부채비율
    bps = instMarketEye.GetDataValue(3, 0)      # BPS
    eps = instMarketEye.GetDataValue(4, 0)      # EPS
    roe = instMarketEye.GetDataValue(5, 0)      # ROE



    # 데이터 결손인 부분을 현재의 재무제표 상의 값으로 채워넣기
    if data[25, i + 7] == 0:
        data[25, i+7] = interest
    if data[25, i + 6] == 0:
        data[25, i+6] = roe
    if data[25, i+5] == 0:
        data[25, i+5] = profit_growth
    if data[25, i+3] == 0:
        data[25, i+3] = debt
    if data[25, i+2] == 0:
        data[25, i+2] = bps
    if data[25, i+1] == 0:
        data[25, i+1] = eps


    # 만약 중간에 자료가 비어있으면 해당 분기의 다음 분기의 값으로 채운다.
    for a in range(start, end):
        for b in range(24, 5, -1):
            if data[b, a] == 0:
                data[b, a] = data[b+1, a]


##########################################################################
Xdf = pd.DataFrame()
ydf = pd.DataFrame()


for i in range(start, end, 10):
    df = pd.DataFrame()

    # x 변수와 y 변수 가져오기
    if data[25, i + 4] == 0:
        continue
    else:
        X = data[6:26, i + 1:i + 10]
        df = pd.DataFrame(X)
        Xdf = pd.concat([Xdf, df], ignore_index=True)



        y = data[6:26, i]
        df = pd.DataFrame(y)
        ydf = pd.concat([ydf, df], ignore_index=True)

X_sum = np.array(Xdf)
Y_sum = np.array(ydf)
X_mean = np.mean(X_sum, axis=0)
print('변수들의 평균', X_mean)
Y_mean = np.mean(Y_sum, axis=0)
print('가격의 평균', Y_mean)


for i in range(0, len(Y_sum)):

    if Y_sum[i, 0] > Y_mean * 50:
        Y_sum[i, 0] /= 50
        X_sum[i, 0] /= 45
        X_sum[i, 1] /= 45

    elif Y_sum[i, 0] > Y_mean * 40:
        Y_sum[i, 0] /= 40
        X_sum[i, 0] /= 35
        X_sum[i, 1] /= 35

    elif Y_sum[i, 0] > Y_mean*30:
        Y_sum[i, 0] /= 30
        X_sum[i, 0] /= 23
        X_sum[i, 1] /= 23

    elif Y_sum[i, 0] > Y_mean*20:
        Y_sum[i, 0] /= 20
        X_sum[i, 0] /= 15
        X_sum[i, 1] /= 15

    elif Y_sum[i, 0] > Y_mean*10:
        Y_sum[i, 0] /= 5
        X_sum[i, 0] /= 3
        X_sum[i, 1] /= 3
    elif Y_sum[i, 0] > Y_mean*5:
        Y_sum[i, 0] /= 2
        X_sum[i, 0] /= 2
        X_sum[i, 1] /= 2


X_mean = np.mean(X_sum, axis=0)
print('2변수들의 평균', X_mean)
Y_mean = np.mean(Y_sum, axis=0)
print('2가격의 평균', Y_mean)
Xdf = pd.DataFrame(X_sum)
ydf = pd.DataFrame(Y_sum)

# 이상치 제거 알고리즘
##########################################################################################################
# 출처 : https://mkjjo.github.io/python/2019/01/10/outlier.html
"""
Isolation Forest

다차원 데이터셋에서 효율적으로 작동하는 아웃라이어 제거 방법이다. Isolation Forest는 랜덤하게 선택된 Feature의 MinMax값을 랜덤하게 분리한 관측치들로 구성된다.

재귀 분할은 트리 구조로 나타낼 수 있으므로 샘플을 분리하는 데 필요한 분할 수는 루트 노드에서 종결 노드까지의 경로 길이와 동일하다.

이러한 무작위 Forest에 대해 평균된이 경로 길이는 정규성과 결정의 척도가 된다.

이상치에 대한 무작위 분할을 그 경로가 현저하게 짧아진다. 따라서 특정 샘플에 대해 더 짧은 경로 길이를 생성할 때 아웃라이어일 가능성이 높다.
"""
"""
# Isolation Forest 방법을 사용하기 위해, 변수로 선언을 해 준다.
clf = IsolationForest(max_samples=1000, random_state=1)

# fit 함수를 이용하여, 데이터셋을 학습시킨다. race_for_out은 dataframe의 이름이다.
clf.fit(Xdf)

# predict 함수를 이용하여, outlier를 판별해 준다. 0과 1로 이루어진 Series형태의 데이터가 나온다.
X_pred_outliers = clf.predict(Xdf)


# 원래의 dataframe에 붙이기. 데이터가 0인 것이 outlier이기 때문에, 0인 것을 제거하면 outlier가 제거된  dataframe을 얻을 수 있다.
out = pd.DataFrame(X_pred_outliers)
out = out.rename(columns={0: "out"})
Xdf = pd.concat([Xdf, out], 1)


clf.fit(ydf)

# predict 함수를 이용하여, outlier를 판별해 준다. 0과 1로 이루어진 Series형태의 데이터가 나온다.
y_pred_outliers = clf.predict(ydf)


# 원래의 dataframe에 붙이기. 데이터가 0인 것이 outlier이기 때문에, 0인 것을 제거하면 outlier가 제거된  dataframe을 얻을 수 있다.
out = pd.DataFrame(y_pred_outliers)
out = out.rename(columns={0: "out"})
ydf = pd.concat([Xdf, out], 1)
"""

print(Xdf.shape)


#####################################################################################################

# 다중 선형 회귀 분석

my_stock = [[5006, 42406, 37, 73, 15, 13, 61, 0.9, 6]]



X_train, X_test, y_train, y_test = train_test_split(Xdf, ydf, train_size=0.8, test_size=0.2)

scaler_x = StandardScaler().fit(X_train)
standardized_X = scaler_x.transform(X_train)
standardized_X_test = scaler_x.transform(X_test)

scaler_y = StandardScaler().fit(y_train)
standardized_y = scaler_y.transform(y_train)
standardized_y_test = scaler_y.transform(y_test)



#mlr = LinearRegression()
#mlr.fit(standardized_X, standardized_y)

#y_predict = mlr.predict(standardized_X_test)

#print(mlr.score(standardized_X_test, standardized_y_test))

mlr = LinearRegression()
mlr.fit(X_train, y_train)

y_predict = mlr.predict(X_test)
y_pre = mlr.predict(my_stock)
print("예측 가격 = ", y_pre)

print(mlr.score(X_test, y_test))
#plt.scatter(y_test, y_predict, alpha=0.5)



#plt.scatter(standardized_y_test, y_predict, alpha=0.5)
plt.xlabel("Actual Price")
plt.ylabel("Predicted Price")
plt.title("MULTIPLE LINEAR REGRESSION")
plt.show()


