import pandas as pd
import re

# Lista de domínios pessoais comuns
dominios_pessoais = ['gmail', 'hotmail', 'outlook']

# Função para validar e classificar e-mail
def email_valido(email):
    if not isinstance(email, str):
        return False
    email = email.strip().lower()
    padrao = r'^[a-z0-9._%-]+@[a-z0-9.-]+\.[a-z]{2,}$'
    return re.match(padrao, email) is not None

def tipo_email(email):
    if not isinstance(email, str):
        return 'Inválido'
    email = email.lower()
    if any(d in email for d in dominios_pessoais):
        return 'Pessoal'
    return 'Corporativo'

# Função para ajustar e validar celular
def ajustar_celular(celular):
    celular_str = ''.join(filter(str.isdigit, str(celular)))
    if len(celular_str) == 9 and celular_str.startswith('9') and celular_str[1] in '6789':
        return celular_str, True
    elif len(celular_str) == 8 and celular_str[0] in '6789':
        return '9' + celular_str, True
    return celular_str, False

# Carregar a planilha
arquivo = 'dados_credito.xlsx'
planilha = pd.ExcelFile(arquivo, engine='openpyxl')
abas = ['DADOS_PF', 'DADOS_PJ']
resultado = {}

for aba in abas:
    df = planilha.parse(aba)
    ajustes = []

    for i, row in df.iterrows():
        email = row.get('EMAIL_PESSOA')
        celular = row.get('Celular')

        email_ok = email_valido(email)
        celular_ajustado, celular_ok = ajustar_celular(celular)

        df.at[i, 'Celular'] = celular_ajustado
        ajustes.append('X' if not email_ok or not celular_ok else '')

    df.insert(loc=len(df.columns), column='AD', value=ajustes)
    df.loc[0, 'AD'] = 'AJUSTE'
    resultado[aba] = df

# Salvar a nova planilha
with pd.ExcelWriter('planilha_credito_ajustado.xlsx', engine='openpyxl') as writer:
    for aba, df in resultado.items():
        df.to_excel(writer, sheet_name=aba, index=False)
