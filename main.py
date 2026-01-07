import pathlib
from emailModel import construir_email_body_multiplos_funcionarios
import pandas as pd
import json
from datetime import datetime, timedelta
from arquivos.services import *
from premailer import transform
import sys

aceito = False
teste = False

while aceito != True:
    try:
        start = int(input("Você iniciou o main do Email-Automate-System. Digite uma opção: \n - 1: Rodar o script \n - 2: Rodar em modo teste \n - 3: Cancelar \n Digite a opção: "))
        
        if (start >= 4) or (start <= 0):
            raise Exception

        aceito = True
    except:
        print("--- Opção não aceita, tente novamente ---\n")

if start == 3:
    print("Cancelando...")
    sys.exit(0)
elif start == 2:
    print("Rodando em modo teste.")
    teste = True
elif start == 1:
    ("Rodando script.")

# -------------------------------------------------------------------------------

ativo_path = pathlib.Path(r".\arquivos\lideranca_dict.json")
with open(ativo_path, 'r', encoding='utf-8') as arquivo:
        ativos_dict = json.load(arquivo)

email_path = pathlib.Path(r".\arquivos\emails.xlsx")
df_email = pd.read_excel(email_path, sheet_name="Planilha1")

funcionarios = []
lista_problemas = []

dict_lider = dict()

func_manual = set()
lider_manual = set()

class Funcionario:
    def __init__(self, nome, matricula, cargo, lider_id, lider_matricula, lider_nome, lider_cargo, horas_extras, datas_sem_descanso, datas_interjornada, ops):
        self.nome = nome
        self.matricula = matricula
        self.cargo = cargo
        self.lider_id = lider_id
        self.lider_matricula = lider_matricula
        self.lider_nome = lider_nome
        self.lider_cargo = lider_cargo
        self.horas_extras = horas_extras
        self.datas_sem_descanso = datas_sem_descanso
        self.datas_interjornada = datas_interjornada
        self.ops = ops
        

#----------------------------------------------------------------------------------------

jsonArq = pathlib.Path(r"report_janeiro.json")
with open(jsonArq, 'r', encoding='utf-8') as arquivo:
    empregados = json.load(arquivo)
    # Itera sobre cada empregado do arquivo "projetosX.json"
    for empregado in empregados["Empregados"]:
        total_Extras = 0
        datas_interjor = []
        datas_sem_descanso = []
        dias_sequencia = 0
        nome = ""
        horario = ""
        ops = []
       
        matricula = empregado["id"]
        nome = empregado["nome"]
        horario = empregado["horario"]
        cargo = buscar_cargo_viaAtivo(matricula, ativos_dict)

        # Testa se consegue encontrar líder.
        try:
            lider_id, lider_matricula, lider_nome  = buscar_lideranca(matricula, ativos_dict)
            lider_cargo = buscar_cargo_viaAtivo(lider_matricula, ativos_dict)
        except:
            print("Problema ao processar lider de funcionário: ", nome, "/", matricula)
            lista_problemas.append({"Matricula": matricula, "Nome": nome})
            continue

        # Ignora quem não for técnico.
        try:
            if (len(cargo.split()) < 2) or (cargo.split()[0] != "TECNICO" and cargo.split()[1] != "SERVICOS"):
                continue
        except: pass

        # Itera sobre cada dia de trabalho.
        for dia in empregado["dias_trabalho"]:
            dias = empregado["dias_trabalho"]
            index_atual = empregado["dias_trabalho"].index(dia)
            marcas = dia["marcacoes"]
            trabalhado_hoje = False

            if index_atual < len(empregado["dias_trabalho"]) - 1:
              dia_seguinte = empregado["dias_trabalho"][index_atual + 1]
              data_seguinte = dia_seguinte["data"]
              
              if marcas and dia_seguinte["marcacoes"]:
                ultima_hoje = marcas[-1]
                primeira_amanha = dia_seguinte["marcacoes"][0]
                interjornada = diferenca_horas(data1=dia["data"], hora1=ultima_hoje, data2=dia_seguinte["data"], hora2=primeira_amanha)
                # Filtra a interjornada. Ignora quem tiver menos de 1 hora (provavelmente é um dos casos
                # de bater o ponto 23:59 de um dia e depois 00:01 do outro.)
                if 1 < interjornada < 11:
                    datas_interjor.append([dia["data"], dia["dia_semana"], dia["marcacoes"], dia_seguinte["data"],dia_seguinte["dia_semana"], dia_seguinte["marcacoes"], f"{int(interjornada):02d}:{int((interjornada - int(interjornada)) * 60):02d}"])                    
                    
                    # Coloca no manual quem tiver menos de 5 horas de interjornada para evitar erros
                    # com casos especiais que geram problema no cálculo.
                    if interjornada < 6: 
                        if (matricula not in func_manual) and matricula:
                            func_manual.add(matricula)

            # Testa situação para somar hora extra.
            for sit in dia["situacoes"]:

                if sit["codigo"] in set_horas_extras:
                    horasIntinmin = horas_para_minutos(sit["horas"])
                    horasInt = horasIntinmin / 60
                    total_Extras += horasInt 

                elif sit["codigo"] in {"698", "699"}:
                    horasIntinmin = horas_para_minutos(sit["horas"])
                    horasInt = horasIntinmin / 60
                    total_Extras -= horasInt
                
                if sit["codigo"] in set_descanso: trabalhado_hoje = True
        

            # Testa se o funcionário fez descanso semanal.
            data_atual = datetime.strptime(f"{dia["data"]}/2025", "%d/%m/%Y")
            
            try:
                if data_atual - data_anterior == timedelta(days=1) and trabalhado_hoje and trabalhado_ontem:
                    dias_sequencia += 1
                                  
                if ((data_atual - data_anterior) != timedelta(days=1)) or (index_atual >= len(empregado["dias_trabalho"]) - 1) or not trabalhado_hoje:
                    if (dias_sequencia + 1) > 6:
                        data_final = data_atual if (index_atual >= len(empregado["dias_trabalho"]) - 1) else data_anterior
                        
                        # Faz append de uma tupla com o dia de início da sequência, dia final e total de dias seguidos.
                        datas_sem_descanso.append(((data_final-timedelta(days=(dias_sequencia))).strftime("%d/%m/%Y"), data_final.strftime("%d/%m/%Y"), dias_sequencia + 1))
                    dias_sequencia = 0
            except:
                pass

            data_anterior = data_atual
            trabalhado_ontem = trabalhado_hoje


        # Alterar número mínimo dependendo de quantas semanas faz desde o início do ponto.
        if total_Extras >= 5:
          ops.append(1)

        if datas_sem_descanso:
            ops.append(2)

        if datas_interjor and len(datas_interjor) > 0:
            ops.append(3)
        

        # Gerar email apenas se houver irregularidades
        if ops:

            funcionario = Funcionario(
                nome=nome,
                matricula=matricula,
                cargo=cargo,
                lider_id=lider_id,
                lider_matricula=lider_matricula,
                lider_nome=lider_nome,
                lider_cargo=lider_cargo,
                horas_extras=str(total_Extras),
                datas_sem_descanso=datas_sem_descanso,
                datas_interjornada=datas_interjor,
                ops=ops)
            
            if (lider_id not in dict_lider) and (lider_id):
                dict_lider[lider_id] = []

            try:
                dict_lider[lider_id].append(funcionario)
                
                if matricula in func_manual:
                    lider_manual.add(lider_id)
            except:
                print("Problema ao processar funcionário: ", funcionario)
            
# Cria planilha com funcionários que não conseguiram ser processados.
nome_arquivo_problemas = 'problemas_processamento_' + datetime.now().strftime('%d-%m-%Y %H-%M-%S') + '.xlsx'

try:
    df_problemas = pd.DataFrame(lista_problemas)
    df_problemas.to_excel(nome_arquivo_problemas, index=False, engine='openpyxl')
except Exception as Exception_Planilha:
    print("Problema com a planilha de erros de processamento: ", Exception_Planilha)

# Itera sobre cada liderança para enviar o email.
enviados = []
email_gerente = None

trava = 0

for lider_id, funcionarios_deste_lider in dict_lider.items():
    
    if trava>2 and teste: sys.exit(0)
    trava += 1

    email_gerente = None
    try:        
        lider_matricula = funcionarios_deste_lider[0].lider_matricula
        lider_cargo_gerente = buscar_cargo_viaAtivo(lider_matricula, ativos_dict)
       
        while lider_cargo_gerente.split()[0].casefold() != "GERENTE".casefold():
            #print(lider_cargo_gerente)
            id_superior, superior_matricula, superior_nome = buscar_lideranca(lider_matricula, ativos_dict)

            id_lider = id_superior
            lider_matricula = superior_matricula
            lider_cargo_gerente = buscar_cargo_viaAtivo(lider_matricula, ativos_dict)
        
        filtro = df_email.loc[df_email['ID'] == int(id_lider), 'Email']
   
        if not filtro.empty:
            email_gerente = filtro.iloc[0]
        else:
            email_gerente = None
            
    except Exception as e: 
        print("Erro gerente: ", e)
        pass

    
    # Busca o e-mail do lider
    filtro = df_email.loc[df_email['ID'] == int(lider_id), 'Email']
   
    if not filtro.empty:
        email_lider = filtro.iloc[0]
    else:
        email_lider = None

    # Construir corpo do email apenas com os funcionários desta liderança.
    bodye = construir_email_body_multiplos_funcionarios(
        periodo=f"11/12 A 04/01",
        funcionarios=funcionarios_deste_lider
    )
    
    # Coloca o CSS no formato inline para ser melhor lido pelo Outlook.
    bodye_inline = transform(bodye, disable_validation=True)

    cc = ["maicon.borba@tkelevator.com", "fernanda.barboza@tkelevator.com", "yuri.souza@tkelevator.com"]

    if email_gerente:
        cc.append(email_gerente)


    sucesso = False
    try:
        enviar_email_outlook(
            destinatario=email_lider,
            assunto="Relatório de Horas Extras",
            corpo=bodye_inline,
            cc=cc,
            enviar_automatico=True if lider_id not in lider_manual and not teste else False
        )
        sucesso = True
        
    except Exception as e:
        print(f"❌ Falha ao enviar para {funcionarios_deste_lider[0].lider_nome} pelo nome: {e}")
        sucesso = False
   
    # Se houver problema, tenta novamente procurando o email na GAL.
    if not sucesso:
        email_gal = buscar_email_na_gal(funcionarios_deste_lider[0].lider_nome)
        if email_gal:
            try:
                enviar_email_outlook(
                    destinatario=email_gal,
                    assunto="Relatório de Horas Extras",
                    corpo=bodye_inline,
                    cc=cc,
                    enviar_automatico=True if lider_id not in lider_manual and not teste else False
                )
                
                sucesso = True
                print(f"✅ E-mail enviado com sucesso para: {email_gal}")
                
            except Exception as e2:
                print(f"❌ Falha ao enviar para {email_gal}: {e2}")

    
    # Registra os funcionários cujo email teve sucesso no envio.
    if sucesso:
        try:
            for func in funcionarios_deste_lider:
                nome_funcionario = func.nome
                enviados.append({"Matricula": func.matricula, "Empregado": nome_funcionario, "Lider": func.lider_nome, "Lider Matricula": func.lider_matricula})
        except:
            print(f"****Problema em adicionar lider na lista: Líder: {funcionarios_deste_lider[0].lider_nome}****")

# Adicionar registros na planilha final de enviados.
nome_arquivo_enviados = "enviados_" + datetime.now().strftime('%d-%m-%Y %H-%M-%S') + '.xlsx'
try:
    df_enviados = pd.DataFrame(enviados)
except Exception as Exception_Enviados:
    print("Erro na planilha enviados:", Exception_Enviados)

df_enviados.to_excel(nome_arquivo_enviados, index=False, engine='openpyxl')
