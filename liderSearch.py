from pathlib import Path
import pandas as pd
import json

def excel_to_json(file_path):
    # Ler o arquivo Excel
    df = pd.read_excel(file_path)
    
    # Criar dicionário para armazenar os dados
    dict_json = {}
    
    # Iterar sobre as linhas do DataFrame
    for index, row in df.iterrows():
        nome_funcionario = row['Nome']
        matricula_funcionario = row["Cadastro"]
        id_funcionario = row["ID HRIS"]
        cargo_funcionario = row['Cargo']
        filial_funcionario= row['Filial']
        lider = row['Lider Imediato']
        matricula_lider = row["Matr. Líder"]
        id_lider = row["8ID Líder"]
        centro_custo = row["CCusto"]
        
        # Verificar se o líder não está vazio
        if pd.notna(lider) and lider != '':
            if matricula_funcionario not in dict_json:
                dict_json[matricula_funcionario] = {
                    "Nome Funcionario": nome_funcionario,
                    "Matricula": matricula_funcionario,
                    "8ID Funcionario": id_funcionario,
                    "Cargo": cargo_funcionario,
                    "Filial Funcionario": filial_funcionario,
                    "Lider": lider,
                    "Matricula Lider": matricula_lider,
                    "8ID Lider": id_lider,
                    "Centro de Custo": centro_custo
                }

    # Converter para JSON
    json_output = json.dumps(dict_json, ensure_ascii=False, indent=2)
    
    return json_output

# Exemplo de uso
if __name__ == "__main__":
    # Substitua pelo caminho do seu arquivo
    arquivo_excel = Path(r".\arquivos\Ativos_16-03-2026.xlsx")
    
    try:
        json_resultado = excel_to_json(arquivo_excel)
        print(json_resultado)
        
        # Salvar em arquivo JSON
        with open("lideranca_dict.json", "w", encoding="utf-8") as f:
            f.write(json_resultado)
        print("\nArquivo 'lideranca_dict.json' salvo com sucesso!")
        
    except FileNotFoundError:
        print(f"Erro: Arquivo '{arquivo_excel}' não encontrado.")
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")