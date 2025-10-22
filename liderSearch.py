from pathlib import Path

#file_path = Path(r"C:\Users\GARCIACUNHABERNARDO\OneDrive - TK Elevator\Documents\Email-automação-novo\Email-Automate-System\Ativos com liderança 20102025 limpo.xlsx")

import pandas as pd
import json

def excel_to_json(file_path):
    # Ler o arquivo Excel
    df = pd.read_excel(file_path)
    
    # Criar lista para armazenar os dados
    dados_json = []
    
    # Iterar sobre as linhas do DataFrame
    for index, row in df.iterrows():
        nome_funcionario = row['Nome']
        matricula_funcionario = row["Cadastro"]
        lider = row['Lider Imediato']
        matricula_lider = row["Matr. Líder"]
        id_lider = row["8ID Líder"]
        centro_custo = row["CCusto"]
        
        # Verificar se o líder não está vazio
        if pd.notna(lider) and lider != '':
            dados_json.append({
                "Nome Funcionario": nome_funcionario,
                "Matricula": matricula_funcionario,
                "LIDER": lider,
                "Matricula Lider": matricula_lider,
                "8ID Lider": id_lider,
                "Centro de Custo": centro_custo
            })
    
    # Converter para JSON
    json_output = json.dumps(dados_json, ensure_ascii=False, indent=2)
    
    return json_output

# Exemplo de uso
if __name__ == "__main__":
    # Substitua pelo caminho do seu arquivo
    arquivo_excel = Path(r"C:\Users\GARCIACUNHABERNARDO\OneDrive - TK Elevator\Documents\Email-automação-novo\Email-Automate-System\Ativos com liderança 20102025 limpo.xlsx")
    
    try:
        json_resultado = excel_to_json(arquivo_excel)
        print(json_resultado)
        
        # Salvar em arquivo JSON
        with open("lideranca.json", "w", encoding="utf-8") as f:
            f.write(json_resultado)
        print("\nArquivo 'lideranca.json' salvo com sucesso!")
        
    except FileNotFoundError:
        print(f"Erro: Arquivo '{arquivo_excel}' não encontrado.")
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")