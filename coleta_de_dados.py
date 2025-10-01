import pdfplumber
import json 
import os 
import re
import sys 
import glob
import datetime

if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(__file__)

config = os.path.join(script_dir, "Config")
entrada = os.path.join(script_dir, "Entrada")
pdfs = glob.glob(os.path.join(entrada, '*.pdf'))
saida = os.path.join(script_dir, "Saida")

for pasta in [entrada, saida, config]:
    if not os.path.exists(pasta):
        os.makedirs(pasta)
        print(f"Diretório {pasta} criado.")

if not pdfs:
    print("Nenhum arquivo PDF encontrado na pasta de entrada.")
    sys.exit(1)

def parece_nome(palavra):
    return re.fullmatch(r'[A-Z][a-zà-ÿ]+', palavra) or re.fullmatch(r'[A-ZÀ-Ý]{2,}', palavra)

def extracao_de_dados(page):
    nome = None
    cpf = None
    data_de_pagamento = None

    dados = {
        '1114/10-I-PM': None,
        'ALE I': None,
        'ALE II': None,
        'ALE III': None,
        'Salario_Base': None,
    }

    print(f'Processando páginas {page.page_number}, por favor aguarde')
    text = page.extract_text()
    if not text:
        print(f"Não foi possível extrair texto da página {page.page_number}.")
        return nome, cpf, data_de_pagamento, dados
    
    # Extrair Dados
    lines = text.split('\n')

    for i, line in enumerate(lines):
        linha = line.strip()

        if data_de_pagamento is None and linha.startswith("Data Pagto"):
            if i + 1 < len(lines):
                data_pagina = lines[i + 1].strip()
                match = re.search(r'(\d{2}/\d{2}/\d{4})', data_pagina)
                if match:
                    dia, mes, ano = match.group().split('/')
                    data_de_pagamento = f"01/{mes}/{ano}"


        elif data_de_pagamento is None and "Data Pagamento" in linha:
            if i + 1 < len(lines):
                data_pagina = lines[i + 1].strip()
                match = re.search(r'(\d{2}/\d{2}/\d{4})', data_pagina)
                if match:
                    dia, mes, ano = match.group().split('/')
                    data_de_pagamento = f"01/{mes}/{ano}"

        if nome is None and linha.startswith("Nome"):
            if i + 2 < len(lines):
                linha_nome = lines[i + 2].strip()
                palavras = linha_nome.split()
                nome_valido = []
                for palavra in palavras:
                    if palavra.upper() in {"ATIVO", "INATIVO", "APOSENTADO"}:
                        break
                    if re.match(r'^[A-Za-zÀ-ÿ]+$', palavra):
                        nome_valido.append(palavra)
                    else:
                        break
                nome = ' '.join(nome_valido)

        elif nome is None and linha.startswith("NOME"):
            if i + 1 < len(lines):
                linha_nome = lines[i + 1].strip()
                if linha_nome:
                    palavras = linha_nome.split()
                    nome_valido = []
                    for palavra in palavras:
                        if re.match(r'^[A-Za-zÀ-ÿ\-]+$', palavra):
                            nome_valido.append(palavra)
                        else:
                            break
                    if nome_valido:
                        nome = ' '.join(nome_valido)


        if cpf is None and ("CPF" in linha.upper() or "C.P.F" in linha.upper()):
            match = re.search(r'\d{3}\.\d{3}\.\d{3}-\d{2}', linha)
            if match:
                cpf = match.group()
            elif i + 1 < len(lines):
                cpf_match = re.search(r'\d{3}\.\d{3}\.\d{3}-\d{2}', lines[i + 1])
                if cpf_match:
                    cpf = cpf_match.group()
        
        if 'tempo_de_servico' not in dados:
            codigo = next((c for c in ['009.001', '009001'] if c in line and 'ADICIONAL TEMPO DE SERVICO' in line), None)
            if codigo:
                partes = line.split()
                try:
                    dados['tempo_de_servico'] = partes[6]
                    dados['adicional_de_tempo_de_servico_valor'] = partes[-1]
                except IndexError:
                    pass

        
        if dados.get('Salario_Base') is None and ('001.001' in linha or '001001' in linha) and 'SALARIO BASE (PADRAO)' in linha:
            partes = linha.split()
            dados['Salario_Base'] = partes[-1]

        if ('012.075' in linha or '012075' in linha) and '1114/10-I-PM' in linha:
            partes = linha.split()
            dados['ALE I'] = partes[-1]
            

        if ('012.047' in linha or '012047' in linha) and 'PM-NIVEL I' in linha:
            partes = linha.split()
            dados['ALE I'] = partes[-1]

        if ('012.048' in linha or '012048' in linha) and 'PM-NIVEL II' in linha:
            partes = linha.split()
            dados['ALE II'] = partes[-1]
        
        if ('012.076' in linha or '012076' in linha) and '1114/10-II-PM' in linha:
            partes = linha.split()
            dados['ALE II'] = partes[-1]
            
        if ('012.049' in linha or '012049' in linha) and 'PM-NIVEL III' in linha:
            partes = linha.split()
            dados['ALE III'] = partes[-1]
    
        if ('012.077' in linha or '012077' in linha) and '1114/10-III-PM' in linha:
            partes = linha.split()
            dados['ALE III'] = partes[-1]
            
    return data_de_pagamento, nome, cpf, dados

#Processamento dos arquivos PDF
def processo_1():
    for pdf_path in pdfs:
        nome_arquivo = os.path.basename(pdf_path).replace('.pdf', '')

        with pdfplumber.open(pdf_path) as pdf:
            valores_lista = []

            primeira_pagina = pdf.pages[0]
            data, nome, cpf, _ = extracao_de_dados(primeira_pagina)

            json_path = os.path.join(config, f'{nome}.json')

            for idx, page in enumerate(pdf.pages):
                data_pagina, nome, cpf, dados_pdf = extracao_de_dados(page)

                bloco_valores = {
                    'indice_original': idx,
                    'Data da Competência': data_pagina,
                    **dados_pdf
                }
                valores_lista.append(bloco_valores)

        valores_validos = [v for v in valores_lista if v['Data da Competência']]
        valores_ordenados = sorted(valores_validos, key=lambda x: (datetime.datetime.strptime(x['Data da Competência'], "%d/%m/%Y"), x['indice_original']))


        dados_finais = {
            'Informações Pessoais': {
                'Data da Competência': valores_ordenados[0]['Data da Competência'] if valores_ordenados else None,
                'Nome': nome,
                'CPF': cpf,
            },
            'Valores das Gratificações': valores_ordenados
        }

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(dados_finais, f, ensure_ascii=False, indent=4)


def parse_valor(valor_str):
    return float(valor_str.replace('.', '').replace(',', '.')) if valor_str else 0.0

def calcular_reajustes(valores_gratificacoes, ale_50_valor):
    reajustes = []
    salario_anterior = None

    for i, atual in enumerate(valores_gratificacoes):
        salario_atual = parse_valor(atual.get("Salario_Base"))

        if salario_anterior is not None:
            ale_anterior = (
                parse_valor(valores_gratificacoes[i-1].get("1114/10-I-PM")) or
                parse_valor(valores_gratificacoes[i-1].get("ALE I")) or
                parse_valor(valores_gratificacoes[i-1].get("ALE II")) or
                parse_valor(valores_gratificacoes[i-1].get("1114/10-II-PM")) or
                parse_valor(valores_gratificacoes[i-1].get("ALE III")) or
                parse_valor(valores_gratificacoes[i-1].get("1114/10-III-PM"))
            )
            ale_atual = (
                parse_valor(atual.get("1114/10-I-PM")) or
                parse_valor(atual.get("ALE I")) or
                parse_valor(atual.get("ALE II")) or
                parse_valor(atual.get("1114/10-II-PM")) or
                parse_valor(atual.get("ALE III")) or
                parse_valor(atual.get("1114/10-III-PM"))
            )

            # Exceção: somar metade do ALE anterior se sumiu
            salario_corrigido_anterior = salario_anterior
            if ale_anterior is None:
                salario_corrigido_anterior += ale_50_valor

            elif salario_corrigido_anterior != 0:
                reajuste = round(((salario_atual + ale_50_valor) / (salario_corrigido_anterior) - 1) * 100, 2)
                
                # Adicionar apenas se o reajuste for positivo e relevante
                if reajuste > 0:
                    reajustes.append({
                        "de": valores_gratificacoes[i-1]["Data da Competência"],
                        "Salário Base Anteior": f"R$ {salario_corrigido_anterior:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                        "Salário Base Anterior + ALE 50%": f"R$ {(salario_corrigido_anterior + ale_50_valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                        "para": atual["Data da Competência"],
                        "Salário Base com Reajuste": f"R$ {salario_atual:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                        "Salário Base com Reajuste + ALE 50%": f"R$ {(salario_atual + ale_50_valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                        "reajuste_percentual": f"{reajuste:.2f}%",
                    })

        salario_anterior = salario_atual

    return reajustes

   #Processamento dos arquivos JSON
def processo_2():
    for arquivo_json in glob.glob(os.path.join(config, '*.json')):
        if os.path.basename(arquivo_json) == 'selic.json':
            continue
        
        with open(arquivo_json, 'r', encoding='utf-8') as f:
            dados_finais = json.load(f)

        valores_ordenados = dados_finais['Valores das Gratificações']

        valores_invertidos = list(reversed(valores_ordenados))

        ale_registros = []

        for registro in valores_invertidos:
            for ale_chave in ['1114/10-I-PM', 'ALE I', 'ALE II', 'ALE III']:
                valor_ale = registro.get(ale_chave)
                if valor_ale:
                    ale_registros.append((registro['Data da Competência'], valor_ale))
                    break
            if len(ale_registros) == 5:
                break

        maior_ale_valor = 0
        maior_ale_data = None

        for data, valor_ale in ale_registros:
            valor_numerico = float(valor_ale.replace('.', '').replace(',', '.'))
            if valor_numerico > maior_ale_valor:
                maior_ale_valor = valor_numerico
                maior_ale_data = data

        ale_50_valor = round(maior_ale_valor / 2, 2)
        ale_50_formatado = f"{ale_50_valor:.2f}".replace('.', ',')

        resultados_para_calculo = {
            'Valores Para fins de cálculo': {
                'ALE 50%': ale_50_formatado
            }
        }

        dados_finais['Valores Para fins de cálculo'] = resultados_para_calculo['Valores Para fins de cálculo']
        dados_finais['Valores Para fins de cálculo']['Reajustes Percentuais'] = calcular_reajustes(valores_ordenados, ale_50_valor)

        with open(arquivo_json, "w", encoding="utf-8") as f:
            json.dump(dados_finais, f, ensure_ascii=False, indent=4)

        if maior_ale_data:
            print(f"Maior ALE encontrado na data {maior_ale_data} com valor de R$ {maior_ale_valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        else:
            print("Nenhum valor de ALE encontrado.")

if __name__ == "__main__":
    processo_1()
    processo_2()