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

def gerar_funcionario_html(funcionario: dict, periodo: str) -> str:
    """
    Gera o HTML para um único funcionário
    """
    ops = funcionario.get('ops', [])
    interjornadas = funcionario.get('interjornadas', [])
    horas_extras_nao_autorizadas = funcionario.get('horas_extras_nao_autorizadas', [])
    
    # ===== SEÇÃO 1: COMPENSAÇÃO BANCO DE HORAS (OP 1) =====
    secao_compensacao = ""
    if 1 in ops:
        HorasPendentes = funcionario.get('HorasPendentes', '')
        HorasPendentestrimed = HorasPendentes[:4] if HorasPendentes else ""
        Fechamento_folha = funcionario.get('Fechamento_folha', '')
        
        secao_compensacao = f"""
        <div class="situacao-box situacao-interjornada">
            <h3 style="color: #000000; margin-top: 0;">⏰ Compensação Banco de Horas Extras</h3>
            
            <div class="data-item">
                <span class="data-label">Horas excedentes para compensação:</span>
                <span class="data-value">{HorasPendentestrimed}</span>
                <span class="status status-interjornada">Banco de Horas</span>
            </div>
            
            <div class="data-item">
                <span class="data-label">Data para fechamento da folha:</span>
                <span class="data-value">{Fechamento_folha}</span>
                <span class="status status-interjornada"></span>
            </div>
            
            <div class="data-item">
                <span class="data-label">Ordem:</span>
                <span class="data-value">Solicitamos informar quais serão os planos de compensação para esses acumulados.</span>
            </div>
        </div>
        """
    
    # ===== SEÇÃO 2: INTERJORNADA (OP 3) =====
    secao_interjornada = ""
    if 3 in ops:
        tabela_interjornadas = ""
        for inter in interjornadas:
            data_inicio = inter[0]
            marcacoes_inicio = inter[2]
            data_fim = inter[3]
            marcacoes_fim = inter[5]
            diferenca = inter[-1]
            tabela_interjornadas += f"""
            <div class="data-item">
                <span class="data-label">{data_inicio} ({marcacoes_inicio})</span>
                <span class="data-value">→ {data_fim} ({marcacoes_fim})</span>
                <span class="status status-outro">{diferenca}h</span>
            </div>
            """
        
        secao_interjornada = f"""
        <div class="situacao-box situacao-interjornada">
            <h3 style="color: #005b96; margin-top: 0;">🚩 Interjornada</h3>
            
            <div class="data-item">
                <span class="data-label">Período Analisado:</span>
                <span class="data-value">{periodo}</span>
            </div>
            
            <div class="data-item">
                <span class="data-label">Mínimo Legal:</span>
                <span class="data-value">11 horas</span>
                <span class="status status-outro">Não Respeitado</span>
            </div>

            <div class="situacao-box situacao-outro">
                <p style="text-align: justify;">Verifica-se que não está sendo observado o intervalo mínimo de 11 (onze) horas consecutivas entre o término de uma jornada e o início da seguinte, conforme dispõe o artigo 66 da CLT. Essa prática configura descumprimento da legislação trabalhista, podendo resultar em autuações e passivos para a empresa</p>
            </div>

            <div style="margin-top: 15px;">
                <h4 style="margin-bottom: 10px;">Ocorrências Identificadas:</h4>
                {tabela_interjornadas if tabela_interjornadas else "<p>Nenhuma ocorrência de interjornada identificada.</p>"}
            </div>
        </div>
        """
    
    # ===== SEÇÃO 3: HORAS EXTRAS NÃO AUTORIZADAS (OP 2) =====
    secao_horas_extras = ""
    if 2 in ops:
        lista_horas_extras = ""
        for he in horas_extras_nao_autorizadas:
            he_data = he[0]
            he_sit_cod = he[-2]
            he_hora = he[-3]
            he_sit = he[-1]
            lista_horas_extras += f"""
            <div class="data-item">
                <span class="data-label">{he_data}</span>
                <span class="data-value">{he_sit}</span>
                <span class="data-value">{he_sit_cod}</span>
                <span class="status status-outro-2">{he_hora}</span>
            </div>
            """
        
        Fechamento_folha = funcionario.get('Fechamento_folha', '')
        secao_horas_extras = f"""
        <div class="situacao-box situacao-interjornada">
            <h3 style="color: #ff3300; margin-top: 0;">⚠️ Horas Extras Não Autorizadas</h3>
            
            <div class="data-item">
                <span class="data-label">Prazo para Compensação:</span>
                <span class="data-value">{Fechamento_folha}</span>
                <span class="status status-outro">Até fim da folha</span>
            </div>

            <div class="situacao-box situacao-interjornada">
                <p style="text-align: justify;">
                  As horas extras não estão sendo justificadas com as respectivas OS's (Ordem de Serviço). Solicito que o respectivo empregado justifique suas horas extras.  
                </p>
                <div style="margin-top: 15px;">
                    <h4 style="margin-bottom: 10px;">Detalhamento:</h4>
                    {lista_horas_extras if lista_horas_extras else "<p>Nenhuma hora extra não autorizada identificada.</p>"}
                </div>
            </div>
        </div>
        """
    
    # HTML do funcionário
    nome_colaborador = funcionario.get('nome_colaborador', '')
    
    return f"""
    <div class="Employebox texto-controlado">
        <span>
            <h3>{nome_colaborador}</h3>
            <h5 style="margin-block-start: 0; margin-block-end: 0;">Scroll horizontal. ➡️</h5>
            <hr>
        </span>
        {secao_compensacao}
        {secao_interjornada}
        {secao_horas_extras}
    </div>
    """

def construir_email_body_multiplos_funcionarios(periodo: str, funcionarios: list) -> str:
    """
    Constrói o corpo do email HTML para múltiplos funcionários
    
    Args:
        periodo: Período de análise
        funcionarios: Lista de dicionários com dados de cada funcionário
            Exemplo: [
                {
                    'nome_colaborador': 'João Silva',
                    'HorasPendentes': '08:30',
                    'Fechamento_folha': '05/02/2024',
                    'data_inicio': '15/01/2024',
                    'data_final': '16/01/2024',
                    'ultimo_Ponto': '18:30',
                    'primeiro_ponto_outro': '07:00',
                    'interjornadas': [...],
                    'horas_extras_nao_autorizadas': [...],
                    'ops': [1, 2, 3]
                },
                {
                    'nome_colaborador': 'Maria Santos',
                    'HorasPendentes': '05:15',
                    'Fechamento_folha': '05/02/2024',
                    'ops': [1, 3]
                },
                ...
            ]
    """
    
    # Gerar HTML para cada funcionário
    funcionarios_html = ""
    for func in funcionarios:
        funcionario_html = gerar_funcionario_html(func, periodo)
        funcionarios_html += funcionario_html
    
    # HTML completo
    html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {{
            font-family: Arial, sans-serif;
            color: #222;
            line-height: 1.5;
            font-size: 14px;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 100%;
            margin: 0 auto;
            background: white;
            padding: 25px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        .header {{
            border-bottom: 2px solid #005b96;
            padding-bottom: 15px;
            margin-bottom: 20px;
        }}
        .situacao-box {{
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            padding: 15px;
            margin: 15px 0;
            background: #f9f9f9;
        }}
        .situacao-interjornada {{
            border-left: 4px solid #7b00ff;
        }}
        .situacao-outro {{
            border-left: 4px solid #ff4800;
        }}
        .data-item {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 10px;
            margin: 5px 0;
            background: white;
            border-radius: 3px;
            border: 1px solid #eaeaea;
        }}
        .data-label {{
            font-weight: bold;
            color: #2c3e50;
        }}
        .data-value {{
            color: #34495e;
        }}
        .status {{
            padding: 4px 8px;
            border-radius: 3px;
            font-size: 12px;
            font-weight: bold;
        }}
        .status-interjornada {{
            background: #d5edda;
            color: #155724;
        }}
        .status-outro {{
            background: #f8d7da;
            color: #721c24;
        }}
        .status-outro-2 {{
            background: #a375fa;
            color: #000000;
        }}
        .observacoes {{
            margin-top: 20px;
            padding: 15px;
            background: #fff3cd;
            border-radius: 4px;
            border-left: 4px solid #ffc107;
        }}
        
        /* CONTAINER HORIZONTAL COM SCROLL */
        .overflow-box {{
            display: flex;
            flex-direction: row;
            gap: 20px;
            padding: 20px;
            overflow-x: auto;
            align-items: flex-start;
            scroll-behavior: smooth;
        }}
        
        .Employebox {{
            flex: 0 0 auto;
            width: 400px;
            min-width: 400px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        
        .texto-controlado {{
            word-wrap: break-word;
            overflow-wrap: break-word;
            max-width: 100%;
        }}

        .social-links {{
            margin-top: 15px;
            font-size: 12px;
        }}
        .social-links a {{
            color: #005b96;
            text-decoration: none;
            margin: 0 5px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Cabeçalho Corporativo -->
        <div class="header">
            <h2 style="color: #000000; margin: 0 0 5px 0;">Controle de Horas Extras - TK Elevator</h2>
            <p style="color: #666; margin: 0; font-size: 13px;">Relatório de Situações e Datas - Apuração</p>
        </div>

        <!-- Saudação -->
        <p>Prezada Liderança,</p>
        
        <p>
            Em continuidade às análises de ponto dos colaboradores, informo que, neste mês,
            iniciamos uma avaliação mais criteriosa, com o objetivo de apoiar as lideranças
            no acompanhamento da jornada de seus empregados. Verificamos, no período de <b>{periodo}</b>, a ocorrência
            das seguintes irregularidades dos seguintes empregados:
        </p>

        <!-- CONTAINER COM TODOS OS FUNCIONÁRIOS -->
        <div class="overflow-box">
            {funcionarios_html}
        </div>

        <!-- Observações -->
        <div class="observacoes" style="background-color: #e6e6e6;">
            <h3 style="margin-top: 0;">📝 Observações:</h3>
            <p>• Casos especiais devem ser compensados conforme acordo ou suas devidas rotinas e OS's</p>
            <p>• Próxima verificação no final do fechamento da proxima folha</p>
            <p>• Está apuração sera verificada para auditar das ordems requisitadas</p>
        </div>

        <!-- Rodapé Corporativo TKE -->
        <div class="signature">
            <table>
                <tr>
                    <td style="padding: 4px 0; vertical-align: top;">
                        <b style="font-size:14px; color:#000;">Yuri Bertola de Souza</b><br>
                        <span style="color:#444;">Planejamento e Projetos HR</span><br>
                        <span style="color:#444;">Latin América</span>
                    </td>
                </tr>
                <tr>
                    <td style="padding: 8px 0;">
                        <span style="color:#000;"><b>T</b> +55 51 2129.7638</span><br>
                        <span>TK Elevator | R Santa Maria 1000 | CEP 92500-000 | Guaíba - RS | Brasil | 
                            <a href="https://www.tkelevator.com" style="color:#005b96; text-decoration:none;">www.tkelevator.com</a>
                        </span>
                    </td>
                </tr>
            </table>
            
            <div class="social-links">
                <a href="https://www.facebook.com/TKE.Brasil/">Facebook</a> |
                <a href="https://www.instagram.com/tke.brasil/">Instagram</a> |
                <a href="https://x.com/TKE_BR">Twitter</a> |
                <a href="https://www.linkedin.com/company/tke-global/">LinkedIn</a> |
                <a href="https://blog.br.tkelevator.com/">Blog</a>
            </div>
        </div>
    </div>
</body>
</html>
"""
    return html
