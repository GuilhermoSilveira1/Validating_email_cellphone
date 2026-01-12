
import pandas as pd
import re
import os

dominios_pessoais = [
    'gmail.com', 'gmail.com.br', 'hotmail.com.br', 'hotmail.com',
    'outlook.com', 'outlook.com.br', 'yahoo.com', 'yahoo.com.br', 'uol.com.br'
]

def email_valido(email: str) -> bool:
    if not isinstance(email, str):
        return False
    email = email.strip().lower()
    padrao = r'^[a-z0-9._%-]+@[a-z0-9.-]+\.[a-z]{2,}$'
    if not re.match(padrao, email):
        return False
    dominio = email.split('@')[-1]
    return dominio in dominios_pessoais

def apenas_digitos(valor) -> str:
    """
    Converte valor para string de dígitos, corrigindo casos de float vindo do Excel (ex.: 98765432.0).
    Nunca adiciona dígitos que não existiam.
    """
    # Corrige floats inteiros (ex.: 98765432.0 -> '98765432', sem o '0' final)
    if isinstance(valor, float):
        if valor.is_integer():
            valor = str(int(valor))
        else:
            valor = str(valor)  # mantém formato textual, vamos extrair dígitos abaixo
    s = ''.join(ch for ch in str(valor) if ch.isdigit())
    return s

def remover_prefixo_operadora(s: str) -> str:
    """
    Remove prefixos de tronco/operadora:
    - '0' simples
    - '0XX' (Ex.: 015, 021, 041 etc)
    Só remove se fizer sentido (comprimento suficiente).
    """
    if not s:
        return s
    if s.startswith('0'):
        # Se houver ao menos 3 dígitos, remove '0' + 2 dígitos de operadora
        if len(s) >= 3 and s[1:3].isdigit():
            return s[3:]
        # Senão, remove apenas zeros à esquerda
        return s.lstrip('0')
    return s

def normalizar_telefone_br(valor, exigir_celular=True):
    """
    Normaliza/valida telefone BR.
    - Remove caracteres não numéricos.
    - Remove prefixos de tronco/operadora (0, 0XX).
    - Preserva/remover DDI 55: mantemos internamente sem '+' para padronizar (opcional no retorno).
    - Se tiver 10 ou 11 dígitos após DDI/prefixos, assume que os 2 primeiros são DDD.
    - Regras:
        * Celular: local 9 dígitos e começa com '9' (segundo dígito usualmente 6–9).
          Se local 8 dígitos e começa com 6–9, adiciona '9'.
        * Fixo: local 8 dígitos e começa com 2–5.
    Parâmetro `exigir_celular`: se True, números fixos são considerados inválidos.

    Retorna: (numero_normalizado_digitos, valido_bool, tipo_str, motivo_str, formato_nacional, formato_e164)
    """
    s = apenas_digitos(valor)
    if not s:
        return '', False, 'desconhecido', 'vazio', '', ''

    # Remove prefixos de operadora/tronco
    s = remover_prefixo_operadora(s)

    # Trata DDI
    ddi = ''
    if s.startswith('55'):
        ddi = '55'
        s = s[2:]

    # Detecta DDD (2 dígitos) para comprimentos comuns
    ddd = ''
    local = s
    if len(s) in (10, 11):  # DDD presente
        ddd = s[:2]
        local = s[2:]
    elif len(s) in (8, 9):   # sem DDD
        local = s
    else:
        motivo = f'comprimento inválido ({len(s)} dígitos)'
        numero = f'{ddi}{ddd}{local}'
        return numero, False, 'desconhecido', motivo, '', ''

    tipo = 'desconhecido'
    valido = False
    motivo = ''

    # Normalização de celular/fixo
    if len(local) == 9:
        if local[0] == '9' and local[1] in '6789':
            tipo, valido = 'celular', True
        elif local[0] == '9':
            # Ainda é celular; segundo dígito fora de 6–9 pode existir em casos raros.
            tipo, valido = 'celular', True
        else:
            motivo = '9 dígitos mas não inicia com 9 (provável celular incorreto)'
    elif len(local) == 8:
        if local[0] in '6789':
            # Provável celular antigo sem o 9 — adiciona
            local = '9' + local
            tipo, valido = 'celular', True
        elif local[0] in '2345':
            # Fixo
            tipo = 'fixo'
            valido = not exigir_celular  # válido só se não exigirmos celular
            if not valido:
                motivo = 'fixo informado, mas é exigido celular'
        else:
            motivo = '8 dígitos inválidos para BR'
    else:
        motivo = 'tamanho de local inválido'

    # Monta número normalizado só com dígitos
    numero = f'{ddi}{ddd}{local}'

    # Formatos amigáveis
    formato_nacional = ''
    if ddd and len(local) in (8, 9):
        if len(local) == 9:
            formato_nacional = f'({ddd}) {local[0]}{local[1:5]}-{local[5:]}'
        else:
            formato_nacional = f'({ddd}) {local[:4]}-{local[4:]}'
    elif not ddd and len(local) in (8, 9):
        if len(local) == 9:
            formato_nacional = f'{local[0]}{local[1:5]}-{local[5:]}'
        else:
            formato_nacional = f'{local[:4]}-{local[4:]}'

    formato_e164 = ''
    if ddi == '55':
        # E.164 com '+' só se tivermos DDD
        if ddd:
            formato_e164 = f'+55{ddd}{local}'
        else:
            formato_e164 = f'+55{local}'

    if not valido and not motivo:
        motivo = 'número inválido pelas regras de validação'

    return numero, valido, tipo, motivo, formato_nacional, formato_e164

# ---------------------------
# Uso no seu fluxo
# ---------------------------

# arquivo = input("Digite o caminho completo do arquivo Excel (.xlsx): ").strip()
arquivo = "dados_credito.xlsx"
diretorio_saida = input("Digite o caminho da pasta onde deseja salvar o arquivo gerado: ").strip()

if not os.path.isdir(diretorio_saida):
    print("Diretório inválido. Verifique o caminho e tente novamente.")
    raise SystemExit()

try:
    planilha = pd.ExcelFile(arquivo, engine='openpyxl')
except FileNotFoundError:
    print("Arquivo não encontrado. Verifique o caminho e tente novamente.")
    raise SystemExit()

abas = ['DADOS_PF', 'DADOS_PJ']
resultado = {}

for aba in abas:
    # Força tipos como string para evitar floats no telefone
    df = planilha.parse(
        aba,
        dtype={'Celular': str, 'EMAIL_PESSOA': str}
    )

    ajustes = []
    motivo_email = []
    motivo_celular = []

    for i, row in df.iterrows():
        email = row.get('EMAIL_PESSOA')
        celular = row.get('Celular')

        email_ok = email_valido(email)

        numero_norm, cel_ok, tipo_num, motivo_num, formato_nac, formato_e164 = normalizar_telefone_br(
            celular, exigir_celular=True  # mude para False se aceitar fixo
        )

        # Grava o número normalizado (só dígitos). Se preferir, troque para formato_nac.
        df.at[i, 'Celular'] = numero_norm

        ajustes.append('X' if not email_ok or not cel_ok else '')
        motivo_email.append('Email válido' if email_ok else 'Email inválido ou domínio não permitido')
        if cel_ok:
            rotulo = f'Celular válido ({tipo_num})'
        else:
            rotulo = f'Celular inválido: {motivo_num or "regra não atendida"}'
        motivo_celular.append(rotulo)

    df['AD'] = ajustes
    df['AE'] = motivo_email
    df['AF'] = motivo_celular
    df.at[0, 'AD'] = 'AJUSTE'
    resultado[aba] = df

caminho_saida = os.path.join(diretorio_saida, 'planilha_credito_validada.xlsx')

with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
    for aba, df in resultado.items():
        df.to_excel(writer, sheet_name=aba, index=False)

print(f"✅ Arquivo gerado com sucesso em:\n{caminho_saida}")
