import pandas as pd

# 다른 방법으로 읽기 시도
df = pd.read_csv('202511 인구 및 연령.csv', encoding='cp949')
print('skiprows 없이 전체 컬럼 수:', len(df.columns))
print('\n컬럼 목록:')
for i, col in enumerate(df.columns):
    print(f'{i}: "{col}"')

print('\n처음 5행:')
print(df.head(5))

print('\n\n=== skiprows=1로 읽기 ===')
df2 = pd.read_csv('202511 인구 및 연령.csv', encoding='cp949', skiprows=1)
print('컬럼 수:', len(df2.columns))
print('컬럼:', df2.columns.tolist())
