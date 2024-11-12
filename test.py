import pandas as pd

# 假设 df_quote_programmatrix 是你的 DataFrame

# 将列名标准化为小写并去除空格
standardized_columns = {col: col.replace(" ", "").lower() for col in df_quote_programmatrix.columns}
df_quote_programmatrix.rename(columns=standardized_columns, inplace=True)

# 将要查找的列名标准化
search_column = 'AV Level 2'.replace(" ", "").lower()

# 查找不为空的行的索引
av_index = df_quote_programmatrix[df_quote_programmatrix.loc[:, search_column].isnull() == False].index.tolist()
