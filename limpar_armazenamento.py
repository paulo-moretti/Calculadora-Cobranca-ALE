import os

script_dir = os.path.dirname(__file__)
Config = os.path.join(script_dir, 'Config')

arquivos_preservar = {'juros_ale.xlsm', 'modelo_ale.xlsx', 'selic.json'}

arquivos_na_pasta = os.listdir(Config)

arquivos_apagados = 0

for arquivo in arquivos_na_pasta:
    if arquivo not in arquivos_preservar:
        caminho_arquivo = os.path.join(Config, arquivo)
        try:
            if os.path.isfile(caminho_arquivo):
                os.remove(caminho_arquivo)
                arquivos_apagados += 1
                print(f"[✓] Arquivo removido: {arquivo}")
        except Exception as e:
            print(f"[✗] Erro ao tentar remover {arquivo}: {e}")

if arquivos_apagados == 0:
    print("Nenhum arquivo para limpar na pasta Config.")
else:
    print(f"[✓] Limpeza concluída. {arquivos_apagados} arquivo(s) removido(s).")