# Calculadora de Cobrança ALE

Este repositório contém uma automação em Python para calcular valores de cobrança de Adicional de Local de Exercício (ALE) a partir de arquivos PDF, processar os dados, calcular juros e gerar planilhas Excel com os resultados.

## Visão Geral

A automação é dividida em etapas, orquestradas pelo script `iniciar_processo.py`:

1.  **Coleta de Dados**: Extrai informações relevantes (nome, CPF, datas, valores de ALE e salário base) de arquivos PDF de entrada.
2.  **Processamento de Dados**: Calcula reajustes percentuais e determina o valor de 50% do maior ALE encontrado.
3.  **Preenchimento de Planilha**: Popula um modelo de planilha Excel com os dados extraídos e calculados.
4.  **Cálculo de Juros**: Utiliza uma planilha de juros externa (`juros_ale.xlsm`) e a taxa SELIC para calcular os juros sobre os valores.
5.  **Limpeza**: Remove arquivos temporários gerados durante o processo.

## Estrutura do Projeto

```
ale_cobranca/
├── Config/             # Contém arquivos de configuração e modelos (JSON, XLSX, XLSM)
│   ├── juros_ale.xlsm  # Planilha para cálculo de juros
│   ├── modelo_ale.xlsx # Modelo da planilha de saída
│   └── selic.json      # Armazena a taxa SELIC
├── Entrada/            # Pasta para colocar os arquivos PDF de entrada
├── Saida/              # Pasta onde as planilhas Excel resultantes serão salvas
├── coleta_de_dados.py  # Script para extração de dados de PDFs e processamento inicial
├── iniciar_processo.py # Orquestrador principal da automação
├── juros.py            # Script para cálculo de juros e preenchimento final da planilha
├── limpar_armazenamento.py # Script para limpar arquivos temporários
└── planilha.py         # Script para preencher a planilha modelo com os dados
```

## Como Usar

### Pré-requisitos

-   Python 3.x
-   Bibliotecas Python: `pdfplumber`, `openpyxl`, `xlwings`, `python-dateutil`

### Instalação

1.  Clone este repositório:
    ```bash
    git clone https://github.com/paulo-moretti/Calculadora-Cobranca-ALE.git
    cd Calculadora-Cobranca-ALE
    ```
2.  Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```
    (Você precisará criar o arquivo `requirements.txt` com as dependências listadas acima.)

### Execução

1.  Coloque os arquivos PDF que contêm os dados de ALE na pasta `Entrada/`.
2.  Certifique-se de que os arquivos `juros_ale.xlsm`, `modelo_ale.xlsx` e `selic.json` estejam na pasta `Config/`.
3.  Execute o script principal:
    ```bash
    python iniciar_processo.py
    ```

As planilhas Excel resultantes serão salvas na pasta `Saida/`.

## Contribuição

Sinta-se à vontade para contribuir com melhorias, correções de bugs ou novas funcionalidades. Abra uma *issue* ou envie um *pull request*.

## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes. (Você precisará criar este arquivo se desejar incluir uma licença.)

---

© 2024 Paulo Eduardo Moretti. Todos os direitos reservados.
