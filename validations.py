import pandas as pd
import re

dominios_pessoais = ['gmail.com', 'hotmail.com', 'outlook.com', 'yahoo.com', 'uol.com']

def email_valido(email):
    if not isinstance(email, str):
        return False
    email = email.strip().lower()
    padrao = r'^[a-z0-9._%-]+@[a-z0-9.-]+\.[a-z]{2,}$'
    if not re.match(padrao, email):
        return False
    dominio = email.split('@')[-1]
    return dominio in dominios_pessoais

def ajustar_celular(celular):
    celular_str = ''.join(filter(str.isdigit, str(celular)))
    if len(celular_str) == 9 and celular_str.startswith('9') and celular_str[1] in '6789':
        return celular_str, True
    elif len(celular_str) == 8 and celular_str[0] in '6789':
        return '9' + celular_str, True
    return celular_str, False

arquivo = 'dados_credito.xlsx'
planilha = pd.ExcelFile(arquivo, engine='openpyxl')
abas = ['DADOS_PF', 'DADOS_PJ']
resultado = {}

for aba in abas:
    df = planilha.parse(aba)
    ajustes = []
    motivo_email = []
    motivo_celular = []

    for i, row in df.iterrows():
        email = row.get('EMAIL_PESSOA')
        celular = row.get('Celular')

        email_ok = email_valido(email)
        celular_ajustado, celular_ok = ajustar_celular(celular)

        df.at[i, 'Celular'] = celular_ajustado
        ajustes.append('X' if not email_ok or not celular_ok else '')
        motivo_email.append('Email válido' if email_ok else 'Email inválido ou domínio não permitido')
        motivo_celular.append('Celular válido' if celular_ok else 'Celular inválido')

    df['AD'] = ajustes
    df['AE'] = motivo_email
    df['AF'] = motivo_celular
    df.at[0, 'AD'] = 'AJUSTE'
    resultado[aba] = df

with pd.ExcelWriter('planilha_credito_ajustado.xlsx', engine='openpyxl') as writer:
    for aba, df in resultado.items():
        df.to_excel(writer, sheet_name=aba, index=False)
