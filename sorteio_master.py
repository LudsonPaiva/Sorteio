import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

arquivo_excel_1 = 'base_sorteio.xlsx'

df_nomes = pd.read_excel(arquivo_excel_1, sheet_name='Funcionário')
df_produtos = pd.read_excel(arquivo_excel_1, sheet_name='Produtos')
# df_nomes.at[all(), "A"] = df_produtos.at[all(), "F"]


print(df_nomes, df_produtos)