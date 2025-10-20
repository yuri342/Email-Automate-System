import win32com.client
import time
import pathlib
from emailModel import construir_email_body_multiplos_funcionarios
import json
import re
import pandas as pd
import os
from datetime import datetime
from datetime import datetime, timedelta
from itertools import permutations

#Variaveis Globais
global total_Extras
global datas_Extras_nAut
global datas_interjor
global nome
global horario
global ops
global cna
global funcionariosEnviados
global lideresEnviados

def montar_funcionario(lider, nome_colaborador, HorasPendentes="", Fechamento_folha="", 
                      data_inicio="", data_final="", ultimo_Ponto="", 
                      primeiro_ponto_outro="", interjornadas=None, 
                      horas_extras_nao_autorizadas=None, ops=None):
    """
    Monta a estrutura de dados para um funcionário incluindo o líder
    
    Args:
        lider: Nome do líder/gestor
        nome_colaborador: Nome completo do colaborador
        HorasPendentes: Horas pendentes para compensação (formato "HH:MM")
        Fechamento_folha: Data de fechamento da folha (formato "DD/MM/AAAA")
        data_inicio: Data inicial para análise
        data_final: Data final para análise  
        ultimo_Ponto: Último ponto do dia anterior
        primeiro_ponto_outro: Primeiro ponto do dia seguinte
        interjornadas: Lista de tuplas/dicionários com dados de interjornada
        horas_extras_nao_autorizadas: Lista de tuplas/dicionários com HE não autorizadas
        ops: Lista de opções [1, 2, 3] para mostrar seções específicas
        
    Returns:
        dict: Estrutura completa do funcionário com líder
    """
    
    if interjornadas is None:
        interjornadas = []
    if horas_extras_nao_autorizadas is None:
        horas_extras_nao_autorizadas = []
    if ops is None:
        ops = []
    
    return {
        'lider': lider,
        'nome_colaborador': nome_colaborador,
        'HorasPendentes': HorasPendentes,
        'Fechamento_folha': Fechamento_folha,
        'data_inicio': data_inicio,
        'data_final': data_final,
        'ultimo_Ponto': ultimo_Ponto,
        'primeiro_ponto_outro': primeiro_ponto_outro,
        'interjornadas': interjornadas,
        'horas_extras_nao_autorizadas': horas_extras_nao_autorizadas,
        'ops': ops
    }

def buscar_email_na_gal(nome):
    """
    Busca email de uma pessoa na Global Address List (GAL) do Outlook,
    tentando todas as permutações do nome caso a busca direta falhe.
    
    Args:
        nome (str): Nome da pessoa para buscar
        
    Returns:
        str: Email encontrado ou None se não encontrar
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Divide o nome em palavras
        palavras = nome.split()
        # Gera todas as permutações das palavras
        todas_permutacoes = [' '.join(p) for p in permutations(palavras)]

        for nomes_permutados in todas_permutacoes:
            recipient = namespace.CreateRecipient(nomes_permutados)
            recipient.Resolve()
            
            if recipient.Resolved:
                if recipient.AddressEntry.Type == "EX":
                    # Usuário Exchange - pega o email SMTP
                    exchange_user = recipient.AddressEntry.GetExchangeUser()
                    if exchange_user:
                        email = exchange_user.PrimarySmtpAddress
                        print(f"✅ Email encontrado na GAL ({nomes_permutados}): {email}")
                        return email
                else:
                    # Outro tipo de entrada
                    print(f"✅ Email encontrado ({nomes_permutados}): {email}")
                    return email

        print(f"❌ Nenhum e-mail encontrado na GAL para o nome '{nome}'")
        return None

    except Exception as e:
        print(f"❌ Erro ao buscar email para '{nome}': {str(e)}")
        return None

def buscar_multiplos_emails(nomes):
    """
    Busca emails para múltiplos nomes
    
    Args:
        nomes: Lista de nomes ou string separada por vírgula
        
    Returns:
        dict: {nome: email} com os emails encontrados
    """
    if isinstance(nomes, str):
        nomes = [nome.strip() for nome in nomes.split(',')]
    
    resultados = {}
    for nome in nomes:
        if nome:
            email = buscar_email_na_gal(nome)
            if email:
                resultados[nome] = email
    
    return resultados


def criar_planilha_empregado_lider(array_empregados, array_lideres, nome_arquivo=None):
    """
    Cria uma planilha Excel com duas colunas: Empregado e Lider
    
    Args:
        array_empregados: Lista com nomes dos empregados
        array_lideres: Lista com nomes dos líderes (deve ter mesmo tamanho que array_empregados)
        nome_arquivo: Nome do arquivo Excel (opcional)
    
    Returns:
        str: Caminho do arquivo salvo
    """
    
    # Verifica se os arrays têm o mesmo tamanho
    if len(array_empregados) != len(array_lideres):
        raise ValueError("Os arrays de empregados e líderes devem ter o mesmo tamanho")
    
    # Cria o DataFrame
    df = pd.DataFrame({
        'Empregado': array_empregados,
        'Lider': array_lideres
    })
    
    # Define o nome do arquivo se não foi fornecido
    if nome_arquivo is None:
        data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"relacao_empregado_lider_{data_atual}.xlsx"
    elif not nome_arquivo.endswith('.xlsx'):
        nome_arquivo += '.xlsx'
    
    # Salva a planilha
    df.to_excel(nome_arquivo, index=False, engine='openpyxl')
    
    print(f"✅ Planilha criada com sucesso: {nome_arquivo}")
    print(f"📊 Total de registros: {len(df)}")
    
    return nome_arquivo

def adicionar_registro_planilha(empregado, lider, nome_arquivo="relacao_empregado_lider.xlsx"):
    """
    Adiciona um novo registro à planilha existente ou cria uma nova
    
    Args:
        empregado: Nome do empregado
        lider: Nome do líder
        nome_arquivo: Nome do arquivo Excel
    """
    try:
        # Tenta carregar a planilha existente
        df = pd.read_excel(nome_arquivo, engine='openpyxl')
        
        # Cria um novo DataFrame com o registro a ser adicionado
        novo_registro = pd.DataFrame({
            'Empregado': [empregado],
            'Lider': [lider]
        })
        
        # Concatena com o DataFrame existente
        df = pd.concat([df, novo_registro], ignore_index=True)
        
    except FileNotFoundError:
        # Se o arquivo não existe, cria um novo
        df = pd.DataFrame({
            'Empregado': [empregado],
            'Lider': [lider]
        })
    
    # Salva a planilha
    df.to_excel(nome_arquivo, index=False, engine='openpyxl')
    print(f"✅ Registro adicionado: {empregado} -> {lider}")


def print_relatorio_dinamico(total_Extras, datas_Extras, datas_interjor, nome, horario, ops):
    print("=" * 60)
    print("📊 RELATÓRIO DE ANÁLISE DE PONTO")
    print("=" * 60)
    print(f"👤 Colaborador: {nome}")
    print(f"🕒 Horário: {horario}")
    print(f"📈 Total de Horas Extras: {total_Extras:.2f}h")
    
    if datas_Extras:
        print(f"📅 Datas com Horas Extras: {(datas_Extras)}")
    else:
        print("📅 Datas com Horas Extras: Nenhuma")
    
    if datas_interjor:
        print(f"⚠️  Datas com Interjornada: {datas_interjor}" +"\n")
    else:
        print("⚠️  Datas com Interjornada: Nenhuma" +"\n")
    
    if ops:
        print(f"🔧 Itens do Relatório: {', '.join(map(str, ops))}")
    else:
        print("🔧 Itens do Relatório: Nenhum")
    print("=" * 60)


def enviar_email_outlook(destinatario, assunto, corpo, cc=None, anexo=None, 
                         enviar_automatico=True, formato_html=True):
    """
    Envia e-mail via Outlook - MODO MANUAL SE NÃO RESOLVER
    
    Args:
        destinatario (str ou list): E-mail(s) OU nome(s) da lista corporativa
        assunto (str): Assunto do e-mail
        corpo (str): Corpo do e-mail
        cc (str ou list, optional): E-mail(s) OU nome(s) em cópia
        anexo (str ou list, optional): Caminho(s) do(s) arquivo(s) para anexar
        enviar_automatico (bool): Se True envia automaticamente, senão exibe para revisão
        formato_html (bool): Se True usa HTML, senão texto plano
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        
        # Configurações básicas
        mail.Subject = assunto
        
        if formato_html:
            mail.HTMLBody = corpo
        else:
            mail.Body = corpo
        
        # Função para adicionar destinatários
        def adicionar_destinatarios(emails, tipo="To"):
            if isinstance(emails, str):
                emails = [email.strip() for email in emails.split(',')]
            
            for email in emails:
                if email:
                    recipient = mail.Recipients.Add(email)
                    recipient.Type = tipo
        
        # Processa destinatários
        adicionar_destinatarios(destinatario, 1)  # 1 = To
        if cc:
            adicionar_destinatarios(cc, 2)  # 2 = CC
        
        # Resolve destinatários
        time.sleep(2)
        mail.Recipients.ResolveAll()
        
        # Verifica se todos os destinatários foram resolvidos
        destinatarios_nao_resolvidos = []
        for recipient in mail.Recipients:
            if not recipient.Resolved:
                destinatarios_nao_resolvidos.append(recipient.Name)
                print(f"❌ Destinatário não resolvido: '{recipient.Name}'")
        
        # ✅ SE HÁ DESTINATÁRIOS NÃO RESOLVIDOS, VAI PARA MODO MANUAL
        if destinatarios_nao_resolvidos:
            print(f"\n⚠️ {len(destinatarios_nao_resolvidos)} destinatário(s) não encontrado(s):")
            for nome in destinatarios_nao_resolvidos:
                print(f"   - {nome}")
            
            print("\n📝 Abrindo modo manual para edição...")
            print("   Você pode:")
            print("   1. Corrigir os destinatários diretamente no Outlook")
            print("   2. Enviar manualmente quando estiver pronto")
            print("   3. Fechar a janela para cancelar")
            
            # ✅ SEMPRE ABRE PARA EDIÇÃO MANUAL QUANDO HÁ ERROS
            mail.Display()
            print("✅ E-mail aberto para edição manual")
            return True
        
        # Adiciona anexos
        if anexo:
            if isinstance(anexo, str):
                anexo = [anexo]
            
            for arquivo in anexo:
                if os.path.exists(arquivo):
                    mail.Attachments.Add(arquivo)
                else:
                    print(f"⚠️ Aviso: Arquivo não encontrado: {arquivo}")
        
        # Envia ou exibe (só chega aqui se TODOS os destinatários foram resolvidos)
        if enviar_automatico:
            mail.Send()
            print(f"✅ E-mail enviado com sucesso!")
            return True
        else:
            mail.Display()
            print(f"✉️ E-mail aberto para envio manual")
            return True
            
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {str(e)}")
        return False


from datetime import datetime, timedelta
def calcular_intervalo_datetime(horario1, horario2):
    # Cria objetos datetime para o mesmo dia
    data_base = datetime.now().date()
    dt1 = datetime.combine(data_base, datetime.strptime(horario1, '%H:%M').time())
    dt2 = datetime.combine(data_base, datetime.strptime(horario2, '%H:%M').time())
    
    # Se o segundo horário for menor, assume que é do dia seguinte
    if dt2 < dt1:
        dt2 += timedelta(days=1)
    
    diferenca = dt2 - dt1
    total_minutos = int(diferenca.total_seconds() // 60)
    horas = total_minutos // 60
    minutos = total_minutos % 60
    
    return {
        'horas': horas,
        'minutos': minutos,
        'total_minutos': total_minutos,
        'formato_string': f"{horas:02d}:{minutos:02d}",
        'virada_dia': dt2.day > dt1.day
    }


def buscar_gerente_viaAtivo(nome, ativopath):
    with open(ativopath, 'r', encoding='utf-8') as arquivo:
        ativos1 = json.load(arquivo)
        for funcionario in ativos1:
            if funcionario["Nome Funcionario"] == nome:
                return funcionario["LIDER"]
    return None
        
    

def diferenca_dias(data1, data2):
    """
    Calcula diferença em dias entre duas datas no formato 'DD/MM'
    Considera que ambas são do mesmo ano
    """
    try:
        # Adicionar o ano atual para converter para datetime
        ano_atual = datetime.now().year
        data1_completa = datetime.strptime(f"{data1}/{ano_atual}", "%d/%m/%Y")
        data2_completa = datetime.strptime(f"{data2}/{ano_atual}", "%d/%m/%Y")
        
        diferenca = abs((data2_completa - data1_completa).days)
        return diferenca
    except ValueError as e:
        print(f"Erro ao converter datas: {e}")
        return None


def horas_para_minutos(horario):
    horas, minutos = map(int, horario.split(':'))
    return horas * 60 + minutos


def subtrair_horarios(horario1, horario2):
    h1, m1 = map(int, horario1.split(':'))
    h2, m2 = map(int, horario2.split(':'))
    
    # Converter tudo para minutos
    total_minutos1 = h1 * 60 + m1
    total_minutos2 = h2 * 60 + m2
    
    # Subtrair
    diferenca_minutos = total_minutos1 - total_minutos2
    
    # Converter de volta para horas e minutos
    horas = diferenca_minutos // 60
    minutos = diferenca_minutos % 60
    
    return horas, minutos

funcionariosEnviados = []
lideresEnviados = []
funcionarios = []
jsonArq = pathlib.Path(r"projetos09.json")


criar_planilha_empregado_lider(funcionariosEnviados, lideresEnviados, "aracaju-Enviados")
with open(jsonArq, 'r', encoding='utf-8') as arquivo:
    empregados = json.load(arquivo)
    for empregado in empregados["Empregados"]:
        total_Extras = 0
        datas_Extras_nAut = []
        datas_interjor = []
        nome = ""
        horario = ""
        ops = []
        cna = False
        

        nome = empregado["nome"]
        horario = empregado["horario"]
        for dia in empregado["dias_trabalho"]:
            dias = empregado["dias_trabalho"]
            index_atual = empregado["dias_trabalho"].index(dia)
            marcas = dia["marcacoes"]

            if index_atual < len(empregado["dias_trabalho"]) - 1:
              dia_seguinte = empregado["dias_trabalho"][index_atual + 1]
              data_seguinte = dia_seguinte["data"]
              if marcas and dia_seguinte["marcacoes"]:
                ultima_hoje = marcas[-1]
                primeira_amanha = dia_seguinte["marcacoes"][0]
                calc = calcular_intervalo_datetime(ultima_hoje, primeira_amanha)
                if 0 < calc["total_minutos"] <= 660 and diferenca_dias(dia_seguinte["data"], dia["data"]) >= 1: #11 horas em minutos
                    calc = calcular_intervalo_datetime(ultima_hoje, primeira_amanha)
                    datas_interjor.append([dia["data"], dia["dia_semana"], dia["marcacoes"], dia_seguinte["data"],dia_seguinte["dia_semana"], dia_seguinte["marcacoes"], calc["formato_string"]])
            

            for sit in dia["situacoes"]:

                if 601 <= int(sit["codigo"]) <= 688 :
                    # print(sit)
                    horasIntinmin = horas_para_minutos(sit["horas"])
                    horasInt = horasIntinmin / 60
                    total_Extras += horasInt
                    if dia["data"] not in datas_Extras_nAut:
                        datas_Extras_nAut.append([dia["data"], dia["dia_semana"], dia["marcacoes"], sit["codigo"], sit["horas"], sit["descricao"]]) 
                    cna = True
                elif 301 <= int(sit["codigo"]) <= 336:
                    horasIntinmin = horas_para_minutos(sit["horas"])
                    horasInt = horasIntinmin / 60
                    total_Extras += horasInt
                elif 698 == int(sit["codigo"]):
                    continue
                elif 351 <= int(sit["codigo"]) <= 390:
                    if sit["horas"] == "DSR":
                        if dia["data"] not in datas_Extras_nAut:
                            datas_Extras_nAut.append([dia["data"], dia["dia_semana"], dia["marcacoes"], sit["codigo"], sit["horas"], sit["descricao"]])
                            cna = False
                    else:
                        horasIntinmin = horas_para_minutos(sit["horas"])
                        horasInt = horasIntinmin / 60
                        total_Extras += horasInt
                        if dia["data"] not in datas_Extras_nAut:
                            datas_Extras_nAut.append([dia["data"], dia["dia_semana"], dia["marcacoes"], sit["codigo"], sit["horas"], sit["descricao"]])
                            cna = True
                elif sit["codigo"] == "DSR":
                    continue

        if total_Extras >= 9: 
          ops.append(1)

        if cna == True :
            ops.append(2)

        if datas_interjor and len(datas_interjor) > 0:
            ops.append(3)

        print_relatorio_dinamico(total_Extras, datas_Extras_nAut, datas_interjor, nome, horario, ops)
# Gerar email apenas se houver irregularidades
        if len(ops) > 0:
            print(ops)
            from datetime import date, timedelta
            hoje = "10/10/2025"
            amanha = date.today() + timedelta(days=1)

            # bodye = construir_email_body(
            #     nome_colaborador=nome,
            #     periodo=f"11/{empregados["mes"]}/2025 á {hoje}",  # Usar mês do JSON
            #     HorasPendentes=str(total_Extras),
            #     Fechamento_folha="10/10/2025",
            #     data_inicio=datas_interjor[0][0]if 3 in ops else "",
            #     data_final=datas_interjor[-1][-4]if 3 in ops else "",
            #     ultimo_Ponto=datas_interjor[0][2][-1]if 3 in ops else "",
            #     primeiro_ponto_outro=datas_interjor[-1][5][0]if 3 in ops else "",
            #     interjornadas=datas_interjor if 3 in ops else None,
            #     horas_extras_nao_autorizadas=datas_Extras_nAut,
            #     ops=ops
            # )

            ativo = pathlib.Path(r"lideranca.json")
            lider = buscar_gerente_viaAtivo(nome, ativo)
            liderpart = lider.split()
            liderlimpo = " ".join(lider.split())
            email = buscar_email_na_gal(lider)

            funcionarios.append(montar_funcionario(
                lider=lider,
                nome_colaborador=nome,
                HorasPendentes=str(total_Extras),
                Fechamento_folha="10/11/2025",
                data_inicio=datas_interjor[0][0]if 3 in ops else "",
                data_final=datas_interjor[-1][-4]if 3 in ops else "",
                ultimo_Ponto=datas_interjor[0][2][-1]if 3 in ops else "",
                primeiro_ponto_outro=datas_interjor[-1][5][0]if 3 in ops else "",
                interjornadas=datas_interjor if 3 in ops else None,
                horas_extras_nao_autorizadas=datas_Extras_nAut,
                ops=ops
            ))

for func in funcionarios:
    lider = func['lider']
    print("Processando email para liderança - " + lider +"\n"+"-"*30)
    bodye = construir_email_body_multiplos_funcionarios(periodo=f"11/{empregados["mes"]}/2025 á {hoje}",funcionarios=funcionarios)
                        
    try:
        # Primeiro tenta enviar usando o nome
        enviar_email_outlook(
            destinatario=[liderlimpo],
            assunto="Relatório de Horas Extras",
            corpo=bodye,
            cc=["maicon.borba@tkelevator.com", "bernardo.cunha@tkelevator.com"],
            enviar_automatico=False
        )
        sucesso = True
    except Exception as e:
        print(f"Falha ao enviar para {lider} pelo nome: {e}")
        sucesso = False
    # Se falhar e tiver e-mail, tenta novamente usando o e-mail
    if not sucesso and email:
        try:
            enviar_email_outlook(
                destinatario=[email],
                assunto="Relatório de Horas Extras",
                corpo=bodye,
                cc=["maicon.borba@tkelevator.com", "bernardo.cunha@tkelevator.com"],
                enviar_automatico=True
            )
            sucesso = True
            print(f"E-mail enviado usando endereço direto: {email}")
        except Exception as e2:
            print(f"Falha ao enviar para {email}: {e2}")
    # Se algum envio funcionou, registra

    if sucesso == True :
        funcionariosEnviados.append(nome)
        lideresEnviados.append(lider)
        adicionar_registro_planilha(nome, lider, "13a17-outubro-Enviados.xlsx")
    



      




