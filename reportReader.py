import pdfplumber
import re
from pathlib import Path as path
import json

pdf_path = path(r"./arquivos/abono 11-09_10-10.PDF")
 
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
 
BLOCO = 100  # ajuste

padraoid = r"\b\d{8,9}\b"
padraoData = r"\b\d{2}/\d{2}\b"
padraohora3 = r"\d{1,3}:\d{2}"
padraohora2 = r"\d{2}:\d{2}"
funcionarios = []
funcionario_atual = None
ids = set()  # Conjunto para rastrear IDs únicos

data = ''
dataAnterior = ''

id = 0
id_anterior = None
escala = 0
turma = 0 
horarioId = 0 
horario = 0 
situacao = []
total_colab = False

with pdfplumber.open(pdf_path) as pdf:
    total = len(pdf.pages)

for inicio in range(0, total, BLOCO):
    with pdfplumber.open(pdf_path) as pdf:
        for i in range(inicio, min(inicio + BLOCO, total)):
            #texto = pdf.pages[i].extract_text().split('\n')[6:-1]
            
            for linha in pdf.pages[i].extract_text().split('\n')[5:-1]: # Ignora as 6 primeiras linhas
                print(linha)
                if re.match(padraoid, linha) :
                    id_anterior = id
                    id = linha.split()[0]  # Primeiro elemento
                    nome = ' '.join(linha.split()[1:])  # Restante como nome

                    # B: Testa se encontrou novo funcionário, o que significa que o anterior terminou; então adiciona as situações.
                    # B: Dá para depois testar se isso aqui chega a ser utilizado, mas vou deixar por enquanto por questão de segurança.
                    if not(id in ids) and ids:
                        for sit in situacao:
                            adicionar_situacao(funcionario_atual["dias_trabalho"][-1], sit[0], sit[1], sit[2])
                        
                        escala = 0  
                        turma = 0 
                        horarioId = 0
                        horario = 0
                        total_colab = False

    
                    # ✅ ADICIONE ESTA VERIFICAÇÃO
                    if id in ids:
                        continue  # Pula para a próxima linha
    
                    funcionario_atual = criar_funcionario(id, nome, escala, turma, horario, horarioId)
                    if funcionario_atual:
                        funcionarios.append(funcionario_atual)
                        # ✅ ADICIONE ESTA LINHA - Marcar ID como adicionado
                        ids.add(id)
    
                # B: Caso encontre o final do pdf, adiciona as situações do último funcionário.
                if ((linha.split()[0]).casefold() == "Total".casefold()) and ((linha.split()[1]).casefold() == "Geral:".casefold()):
                    for sit in situacao:    
                        adicionar_situacao(funcionario_atual["dias_trabalho"][-1], sit[0], sit[1], sit[2])
                

                # B: Ignora quando chega no total do colaborador.
                if ((linha.split()[0]).casefold() == "Total".casefold()) and ((linha.split()[1]).casefold() == "Colaborador:".casefold()):
                    total_colab = True
                    data = ''
                    dataAnterior = ''
                
                
                if total_colab and ids:
                    continue

                elif re.match(padraoData, linha):
                    # B: Pega os casos de mudança de página em que reaparece o mesmo dia com situações que não couberam
                    # na página anterior.
                    if (linha.split()[0] == dataAnterior) and (data != '') and (id_anterior == id):
                        linhaSit = ' '.join(linha.split()[2:])
                        situacaoUnit = linhaSit.split()
                        situacaoUnit[1:-1] = [' '.join(situacaoUnit[1:-1])]
                        situacao.append(situacaoUnit)
                        
                    else: 
                        try:
                            for sit in situacao:
                                adicionar_situacao(funcionario_atual["dias_trabalho"][-1], sit[0], sit[1], sit[2])
                        except:           
                            pass
                
                        data = linha.split()[0]  # Primeiro elemento
                        dia_semana = linha.split()[1]  # Segundo elemento
                        marcacoes = ' '.join(linha.split()[2:])  # Restante como marcações
                        marcacoesOld = re.findall(padraohora3, marcacoes)
                        marcacoes = list(filter(nao_e_hora, marcacoesOld))
                        marcacoesF = ' '.join(marcacoes)
                        linhaSemM = re.sub(marcacoesF, '', linha).strip()

                        situacaoUnit = linhaSemM.split()[2:]  # Restante como situação.
                        situacaoUnit[1:-1] = [' '.join(situacaoUnit[1:-1])]            

                        situacao = []
            
                        situacao.append(situacaoUnit)
                        
                        adicionar_dia_trabalho(funcionario_atual, data, dia_semana, marcacoes)
                    
                    
                # B: Pega as linhas que só tem a situação, sem a data escrita e marcações.
                elif len(linha.split()[0]) == 3:
                    situacaoUnit = linha.split()
                    situacaoUnit[1:-1] = [' '.join(situacaoUnit[1:-1])]
                    situacao.append(situacaoUnit)

                
                # B: Adiciona escala, turma, etc quando encontra.
                if (escala == 0) and (":" in linha.split()[-1]) and (len(linha.split()[0]) == 4):
                    escala = linha.split()[0]  # Primeiro elemento da próxima linha
                    turma = linha.split()[1]  # Segundo elemento da próxima linha
                    horarioId = linha.split()[2]  # Terceiro elemento da próxima linha
                    horario = ' '.join(linha.split()[3:])  # Restante como horário

                    funcionario_atual["escala"] = escala
                    funcionario_atual["turma"] = turma
                    funcionario_atual["horarioId"] = horarioId
                    funcionario_atual["horario"] = horario

                dataAnterior = data

import json
import time
mest = funcionarios[0]["dias_trabalho"][0]["data"].split("/")[1]
dados_json = {
    "mes":  funcionarios[0]["dias_trabalho"][0]["data"].split("/")[1], # Pegando o mês da primeira data do primeiro funcionário
    "total_leituras": len(funcionarios),
    "Empregados": funcionarios,
    "ultima_atualizacao": time.time()
}
 
with open("teste_leitura.json", "w", encoding="utf-8") as arquivo:
    json.dump(dados_json, arquivo, ensure_ascii=False, indent=2)
 
