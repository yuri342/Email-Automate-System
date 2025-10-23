import pdfplumber
import re
from pathlib import Path as path
import json

pdf = path(r"C:\Users\GARCIACUNHABERNARDO\OneDrive - TK Elevator\Documents\Email-automação-novo\Email-Automate-System\relatorio 5025, 5026, 5033, 5058, 5027, 5065.pdf")
#json_data = path(r"ModeloEmail\dados.json")


def filtrar_linhas_pdf(linhas):
    """
    Filtra linhas de PDF removendo cabeçalhos, rodapés e linhas indesejadas
    
    Args:
        linhas: Lista de linhas extraídas do PDF
        
    Returns:
        Lista filtrada apenas com linhas de dados relevantes
    """
    linhas_filtradas = []
    
    # Padrões para identificar linhas indesejadas
    padroes_excluir = [
        r'^Totais do Colaborador',  # Rodapé de totais
        r'^Total:',                 # Linhas de total
        r'^\d+/\d+/\d+',           # Datas completas (rodapé)
        r'^Página \d+ de \d+',     # Numeração de página
        r'^Relatório:',            # Cabeçalhos
        r'^Período:',              # Cabeçalhos
        r'^Data:',                 # Cabeçalhos
        r'^Empresa:',              # Cabeçalhos
        r'^Matrícula',             # Cabeçalhos de coluna
        r'^Nome',                  # Cabeçalhos de coluna
        r'^Data\s+Dia',           # Cabeçalhos de coluna
        r'^Marcações',            # Cabeçalhos de coluna
        r'^Situação',             # Cabeçalhos de coluna
        r'^Horas',                # Cabeçalhos de coluna
        r'^-+',                   # Linhas separadoras
        r'^\s*$',                 # Linhas vazias
    ]
    
    for linha in linhas:
        linha = linha.strip()
        
        # Pular linhas vazias
        if not linha:
            continue
            
        # Verificar se a linha corresponde a algum padrão de exclusão
        excluir = False
        for padrao in padroes_excluir:
            if re.match(padrao, linha, re.IGNORECASE):
                excluir = True
                break
        
        # ✅ CORREÇÃO: Verificar se contém "Total" (case insensitive)
        if not excluir and re.search(r'total', linha, re.IGNORECASE):
            # Mas manter se for um ID (não queremos excluir linhas com IDs)
            if not re.match(r'\b\d{8}\b', linha.split()[0] if linha.split() else ''):
                excluir = True
        
        # ✅ CORREÇÃO: Verificar se é rodapé numérico (contém apenas números e símbolos)
        if not excluir and re.match(r'^[\d\s:/-]+$', linha):
            excluir = True
        
        if not excluir:
            linhas_filtradas.append(linha)
    
    return linhas_filtradas



def nao_e_hora(item):
    padrao_hora = r'^\d{3}:\d{2}$' # Exemplo: 123:45
    return not re.match(padrao_hora, str(item))

def criar_funcionario(matricula, nome, escala, turma, horario, horarioId):
    return {
        "id": matricula,
        "nome": nome,
        "escala": escala,
        "turma": turma,
        "horarioId": horarioId,
        "horario": horario,
        "dias_trabalho": []
    }

def adicionar_dia_trabalho(funcionario, data, dia_semana, marcacoes):
    dia = {
        "data": data,
        "dia_semana": dia_semana,
        "marcacoes": marcacoes,
        "situacoes": []
    }
    funcionario["dias_trabalho"].append(dia)
    return dia

def adicionar_situacao(dia_trabalho, codigo, descricao, horas):
    situacao = {
        "codigo": codigo,
        "descricao": descricao,
        "horas": horas
    }
    dia_trabalho["situacoes"].append(situacao)
    return situacao

#pdf - Reader

with pdfplumber.open(pdf) as pdf:
    padraoid = r"\b\d{8}\b"
    padraoData = r"\b\d{2}/\d{2}\b"
    padraohora3 = r"\d{1,3}:\d{2}"
    padraohora2 = r"\d{2}:\d{2}"
    funcionarios = []
    funcionario_atual = None
    pagina = 0
    ids = set()  # Conjunto para rastrear IDs únicos
    global linhas_filtradas

    for numero_pagina, pagina in enumerate(pdf.pages):
        print(f"\n--- Página {numero_pagina + 1} ---")
        texto_completo = pagina.extract_text()
        
        # Divide o texto em linhas e ignora as primeiras N
        linhas = texto_completo.split('\n')
        linhas_filtradas = linhas[6:-1]  # Ignora as 6 primeiras linhas

        # Processo de pegar funcionario dados.
        for linha in linhas_filtradas:
            if re.match(padraoid, linha) :

            

                id = linha.split()[0]  # Primeiro elemento
                nome = ' '.join(linha.split()[1:])  # Restante como nome

                # ✅ ADICIONE ESTA VERIFICAÇÃO
                if id in ids:
                    print(f"⚠️  Funcionário {id} já existe - pulando")
                    continue  # Pula para a próxima linha
                
                try:
                    linha2 = linhas_filtradas[linhas_filtradas.index(linha)+1]
                    escala = linha2.split()[0]  # Primeiro elemento da próxima linha
                    turma = linha2.split()[1]  # Segundo elemento da próxima linha
                    horarioId = linha2.split()[2]  # Terceiro elemento da próxima linha
                    horario = ' '.join(linha2.split()[3:])  # Restante como horário
                except:
                    continue


                print("Dados do Funcionário:\n")
                print(horarioId)
                print(horario)
                print(turma)
                print(escala)
                print(f"Matrícula: {id}")
                print(f"Nome: {nome}\n")
                print("-------------------------");
                funcionario_atual = criar_funcionario(id, nome, escala, turma, horario, horarioId)
                if funcionario_atual:
                    funcionarios.append(funcionario_atual)
                    # ✅ ADICIONE ESTA LINHA - Marcar ID como adicionado
                    ids.add(id)



            elif re.match(padraoData, linha):
                data = linha.split()[0]  # Primeiro elemento
                dia_semana = linha.split()[1]  # Segundo elemento
                
                marcacoes = ' '.join(linha.split()[2:])  # Restante como marcações
                marcacoesOld = re.findall(padraohora3, marcacoes)
                marcacoes = list(filter(nao_e_hora, marcacoesOld))
                marcacoesF = ' '.join(marcacoes)
                linhaSemM = re.sub(marcacoesF, '', linha).strip()

                situacaoUnit = linhaSemM.split()[2:]  # Restante como marcações
                situacaoUnit[1:-1] = [' '.join(situacaoUnit[1:-1])]            

                situacao = []
                situacao.append(situacaoUnit)
                
                adicionar_dia_trabalho(funcionario_atual, data, dia_semana, marcacoes)

                if linhas_filtradas.index(linha) < len(linhas_filtradas)-1:
                    proxLinha = linhas_filtradas[linhas_filtradas.index(linha)+1]   
                    if proxLinha is not None and not re.match(padraoData, proxLinha) and not re.match(padraoid, proxLinha) and not "Total" in proxLinha and not '0390' in proxLinha:
                        proxlinhaf = proxLinha.split()
                        proxlinhaf[1:-1] = [' '.join(proxlinhaf[1:-1])] 
                        situacao.append(proxlinhaf)
                else:
                    proxLinha = None

                for sit in situacao:
                    adicionar_situacao(funcionario_atual["dias_trabalho"][-1], sit[0], sit[1], sit[2])
                print(f"Data: {data} | Dia: {dia_semana} | Marcacoes: {marcacoes} | Situacao: {situacao} \n")


import json
import time
mest = funcionarios[0]["dias_trabalho"][0]["data"].split("/")[1]
dados_json = {
    "mes":  funcionarios[0]["dias_trabalho"][0]["data"].split("/")[1], # Pegando o mês da primeira data do primeiro funcionário
    "total_leituras": len(funcionarios),
    "Empregados": funcionarios, 
    "ultima_atualizacao": time.time()
}

with open("projetos"+str(mest)+".json", "w", encoding="utf-8") as arquivo:
    json.dump(dados_json, arquivo, ensure_ascii=False, indent=2)