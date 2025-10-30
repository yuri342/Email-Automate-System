import pathlib
from emailModel import construir_email_body_multiplos_funcionarios
import json
import re
from datetime import datetime
from datetime import datetime, timedelta
from services import *

funcionariosEnviados = []
lideresEnviados = []
funcionarios = []
jsonArq = pathlib.Path(r"projetos10.json")

func_manual = set()
lider_manual = set()

with open(jsonArq, 'r', encoding='utf-8') as arquivo:
    empregados = json.load(arquivo)
    # B: Itera sobre cada empregado do arquivo "projetosX.json"
    for empregado in empregados["Empregados"]:
        total_Extras = 0
        datas_interjor = []
        datas_sem_descanso = []
        dias_sequencia = 0
        nome = ""
        horario = ""
        ops = []
       

        nome = empregado["nome"]
        horario = empregado["horario"]

        # Ignora quem não for técnico.
        ativo = pathlib.Path(r"lideranca.json")
        cargo = buscar_cargo_viaAtivo(nome, ativo)

        try:
            if cargo.split()[0] != "TECNICO":
                continue
        except:
            pass

        # B: Itera sobre cada dia de trabalho.
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
                # B: Filtra a interjornada. Ignora quem tiver menos de 1 hora (provavelmente é um dos casos
                # de bater o ponto 23:59 de um dia e depois 00:01 do outro.
                if 1 < interjornada < 11:
                    datas_interjor.append([dia["data"], dia["dia_semana"], dia["marcacoes"], dia_seguinte["data"],dia_seguinte["dia_semana"], dia_seguinte["marcacoes"], f"{int(interjornada):02d}:{int((interjornada - int(interjornada)) * 60):02d}"])                    
                    
                    # B: Coloca no manual quem tiver menos de 5 horas de interjornada para evitar erros
                    # casos especiais com casos especiais que geram problema no cálculo.
                    if interjornada < 6: 
                        if (nome not in func_manual) and nome:
                            func_manual.add(nome)

            # B: Testa situação para somar hora extra.
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
        

            # B: Testa se o funcionário fez descanso semanal.
            data_atual = datetime.strptime(f"{dia["data"]}/2025", "%d/%m/%Y")
            
            try:
                if data_atual - data_anterior == timedelta(days=1) and trabalhado_hoje and trabalhado_ontem:
                    dias_sequencia += 1
                                  
                if ((data_atual - data_anterior) != timedelta(days=1)) or (index_atual >= len(empregado["dias_trabalho"]) - 1) or not trabalhado_hoje:
                    if (dias_sequencia + 1) > 6:
                        data_final = data_atual if (index_atual >= len(empregado["dias_trabalho"]) - 1) else data_anterior
                        
                        datas_sem_descanso.append(((data_final-timedelta(days=(dias_sequencia))).strftime("%d/%m/%Y"), data_final.strftime("%d/%m/%Y"), dias_sequencia + 1))
                    dias_sequencia = 0
            except:
                pass

            
            data_anterior = data_atual
            trabalhado_ontem = trabalhado_hoje


        # B: Alterar número mínimo dependendo de quantas semanas faz desde o início do ponto.
        if total_Extras >= 10: 
          ops.append(1)

        if datas_sem_descanso:
            ops.append(2)

        if datas_interjor and len(datas_interjor) > 0:
            ops.append(3)
        

        #print_relatorio_dinamico(total_Extras, datas_Extras_nAut, datas_interjor, nome, horario, ops)
# Gerar email apenas se houver irregularidades
        if len(ops) > 0:
            #print(ops)
            from datetime import date, timedelta
            hoje = "03/11/2025"
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

            
            lider = buscar_gerente_viaAtivo(nome, ativo)

            funcionarios.append(montar_funcionario(
                lider=lider,
                nome_colaborador=nome,
                cargo_colaborador=cargo,
                HorasPendentes=str(total_Extras),
                Fechamento_folha="10/11/2025",
                data_inicio=datas_interjor[0][0]if 3 in ops else "",
                data_final=datas_interjor[-1][-4]if 3 in ops else "",
                ultimo_Ponto=datas_interjor[0][2][-1]if 3 in ops else "",
                primeiro_ponto_outro=datas_interjor[-1][5][0]if 3 in ops else "",
                interjornadas=datas_interjor if 3 in ops else None,
                datas_sem_descanso = datas_sem_descanso if 2 in ops else None,
                ops=ops
            ))

# Primeiro, agrupa os funcionários por liderança
funcionarios_por_lider = {}
for func in funcionarios:
    lider = func['lider']
    if lider not in funcionarios_por_lider:
        funcionarios_por_lider[lider] = []
    funcionarios_por_lider[lider].append(func)

    # B: Encontra o lider que tem um funcionário marcado para manual.
    func_nome = func["nome_colaborador"]
    if func_nome in func_manual:
        if (lider not in lider_manual) and lider:
            lider_manual.add(lider)


# Agora processa cada liderança separadamente
for lider, funcionarios_deste_lider in funcionarios_por_lider.items():
    print(f"\n{'='*50}")
    print(f"{funcionarios_deste_lider}")
    print(f"\n{'='*50}")
    print(f"Processando email para liderança: {lider}")
    print(f"Total de funcionários: {len(funcionarios_deste_lider)}")
    print(f"{'='*50}")
    
    # Limpar o nome do líder para busca de email
    if lider is not None and lider != "":
        liderlimpo = " ".join(lider.split()[1:]) + ", " + lider.split()[0]
    else:
        liderlimpo = "Líder Não Informado"

    # Buscar email do líder
    email = buscar_email_na_gal(lider)
    
    # Construir corpo do email apenas com os funcionários desta liderança
    bodye = construir_email_body_multiplos_funcionarios(
        periodo=f"11/10 A 25/10",
        funcionarios=funcionarios_deste_lider
    )
    
    sucesso = False
    
    try:
        # Primeiro tenta enviar usando o nome
        print(f"📧 Tentando enviar para: {lider}")
        enviar_email_outlook(
            destinatario=lider,
            assunto="Relatório de Horas Extras",
            corpo=bodye,
            cc=["maicon.borba@tkelevator.com", "yuri.souza@tkelevator.com"],
            #enviar_automatico=True if lider not in lider_manual else False
            enviar_automatico=True if lider not in lider_manual else False
        )
        sucesso = True
        print(f"✅ Email aberto para envio manual: {lider}")
        
    except Exception as e:
        print(f"❌ Falha ao enviar para {lider} pelo nome: {e}")
        sucesso = False
    
    # Se falhar e tiver e-mail, tenta novamente usando o e-mail
    if not sucesso and email:
        try:
            print(f"📧 Tentando enviar para email: {email}")
            enviar_email_outlook(
                destinatario=email,
                assunto="Relatório de Horas Extras",
                corpo=bodye,
                cc=["maicon.borba@tkelevator.com", "yuri.souza@tkelevator.com"],
                enviar_automatico=True if lider not in lider_manual else False
            )
            sucesso = True
            print(f"✅ E-mail enviado com sucesso para: {email}")
            
        except Exception as e2:
            print(f"❌ Falha ao enviar para {email}: {e2}")
            sucesso = False
    
    # Se algum envio funcionou, registra na planilha
    if sucesso:
        # Registrar todos os funcionários desta liderança
        for func in funcionarios_deste_lider:
            nome_funcionario = func['nome_colaborador']
            funcionariosEnviados.append(nome_funcionario)
            lideresEnviados.append(lider)
            adicionar_registro_planilha(nome_funcionario, lider, "teste+.xlsx")
        
        print(f"📊 Registrados {len(funcionarios_deste_lider)} funcionários da liderança {lider}")
    else:
        print(f"⚠️  Nenhum email enviado para liderança: {lider}")


print(f"\n{'='*50}")
print("RESUMO DO PROCESSAMENTO:")
print(f"Total de lideranças processadas: {len(funcionarios_por_lider)}")
print(f"Total de funcionários enviados: {len(funcionariosEnviados)}")
print(f"{'='*50}")
    



      




