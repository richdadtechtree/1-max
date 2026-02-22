import pandas as pd

df = pd.read_csv('202511 인구 및 연령.csv', encoding='cp949', skiprows=1)
print('전체 컬럼 수:', len(df.columns))
print('\n컬럼 목록:')
for i, col in enumerate(df.columns):
    print(f'{i}: {col}')

print('\n첫 3행 데이터:')
print(df.head(3))
