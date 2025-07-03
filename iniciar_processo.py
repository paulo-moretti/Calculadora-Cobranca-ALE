import subprocess
import os

script_dir = os.path.dirname(__file__)

limpar_armazenamento_script = os.path.join(script_dir, 'limpar_armazenamento.py')
coleta_script = os.path.join(script_dir, 'coleta_de_dados.py')
planilha_script = os.path.join(script_dir, 'planilha.py')
limpar_armazenamento_script = os.path.join(script_dir, 'limpar_armazenamento.py')
calcular_juros_script = os.path.join(script_dir, 'juros.py')

print("Iniciando coleta de dados...")
subprocess.run(["python", coleta_script], check=True)

print("Preenchendo a(s) planilha(s)...")
subprocess.run(["python", planilha_script], check=True)

print("Calculando Juros e Finalizando...")
subprocess.run(["python", calcular_juros_script], check=True)

print("Iniciando limpeza de armazenamento...")
subprocess.run(["python", limpar_armazenamento_script], check=True)

print("[✓] Processo completo!")