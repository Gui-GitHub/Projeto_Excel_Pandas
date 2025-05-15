import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import os

# Caminho do arquivo original e final
caminho_pasta = r'X:\arrumar_excel'
arquivo_entrada = os.path.join(caminho_pasta, 'sc_xls_20250513150234_686_grid_db_angariacao_excel.xlsx')
arquivo_saida = os.path.join(caminho_pasta, 'funcionarios_dependentes_formatado.xlsx')

# Carrega o Excel original
df = pd.read_excel(arquivo_entrada)

# Lista com linhas formatadas
dados_formatados = []

for _, row in df.iterrows():
    dados_formatados.append({
        'Nome': row['Nome Funcionario'],
        'Cpf Funcionario (repetido)': row['Cpf Funcionario'],
        'Plano Escolhido': row['Id Plano Escolhido'],
        'Sexo': '',
        'Cpf Dep': '',
        'Data Nascimento': '',
        'Nome Mae': '',
        'Parentesco': '',
        'Data Cadastro Update': row['Data Cadastro Update']
    })
    
    for i in range(1, 7):
        if pd.notna(row.get(f'Nome Dep {i}')):
            data_nasc = row.get(f'Data Nascimento Dep {i}', '')
            if pd.notna(data_nasc):
                data_nasc = pd.to_datetime(data_nasc).strftime('%d/%m/%Y')
            else:
                data_nasc = ''
            dados_formatados.append({
                'Nome': row[f'Nome Dep {i}'],
                'Cpf Funcionario (repetido)': row['Cpf Funcionario'],
                'Plano Escolhido': row['Id Plano Escolhido'],
                'Sexo': row.get(f'Sexo Dep {i}', ''),
                'Cpf Dep': row.get(f'Cpf Dep {i}', ''),
                'Data Nascimento': data_nasc,
                'Nome Mae': row.get(f'Nome Mae Dep {i}', ''),
                'Parentesco': row.get(f'Parentesco Dep {i}', ''),
                'Data Cadastro Update': row['Data Cadastro Update']
            })

# Cria DataFrame final
df_final = pd.DataFrame(dados_formatados)

# Salva o DataFrame como Excel
df_final.to_excel(arquivo_saida, index=False)

# Aplica formatações com openpyxl
wb = load_workbook(arquivo_saida)
ws = wb.active

# Estilos
cabecalho_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
cabecalho_font = Font(color="FFFFFF", bold=True)
alinhamento_centro = Alignment(horizontal="center", vertical="center")

# Aplica estilos no cabeçalho
for col in range(1, ws.max_column + 1):
    celula = ws.cell(row=1, column=col)
    celula.fill = cabecalho_fill
    celula.font = cabecalho_font
    celula.alignment = alinhamento_centro

# Aplica alinhamento nas células do restante
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        cell.alignment = alinhamento_centro

# Salva com estilo
wb.save(arquivo_saida)

print(f"Arquivo gerado com sucesso em:\n{arquivo_saida}")