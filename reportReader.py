import pdfplumber
import re
from pathlib import Path as path
import json

pdf = path(r"relatorio 5004 a 5033 11-12_16-12.PDF")
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
    padraoid = r"\b\d{8,9}\b"
    padraoData = r"\b\d{2}/\d{2}\b"
    padraohora3 = r"\d{1,3}:\d{2}"
    padraohora2 = r"\d{2}:\d{2}"
    funcionarios = []
    funcionario_atual = None
    pagina = 0
    ids = set()  # Conjunto para rastrear IDs únicos
    global linhas_filtradas
   
    data = ''
    dataAnterior = ''

    id = 0
    escala = 0  
    turma = 0 
    horarioId = 0 
    horario = 0 
    situacao = []
    total_colab = False
    
    for numero_pagina, pagina in enumerate(pdf.pages):
        texto_completo = pagina.extract_text()
       
        # Divide o texto em linhas e ignora as primeiras N
        linhas = texto_completo.split('\n')
        linhas_filtradas = linhas[6:-1]  # Ignora as 6 primeiras linhas
 
        # Processo de pegar funcionario dados.
        for linha in linhas_filtradas:
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
 
with open("reporte_dezembro.json", "w", encoding="utf-8") as arquivo:
    json.dump(dados_json, arquivo, ensure_ascii=False, indent=2)
 
