import pandas as pd
import win32com.client as win32
import re
from datetime import date
import matplotlib.pyplot as plt
import base64

#Definindo e-mail de destino:
email = 'dl_ctio_coreoperations_mobilevoiceplatforms@timbrasil.com.br'

#Endereço do export do NETFLOW 
path_r041 = r'\\internal.timbrasil.com.br\FileServer\TBR\Network\Management&Control\ProjectMonitoring\PMOnline-Exports\Bases\Mobile Voice Platforms\[R041] Relatório de Anexos da OC  - CORE.xlsx'
path_r001 = r'\\internal.timbrasil.com.br\FileServer\TBR\Network\Management&Control\ProjectMonitoring\PMOnline-Exports\Bases\Mobile Voice Platforms\[R001] Ordens Forwarded  e  Oper. Core Mobile Voice.xlsx'

#faz a leitura da base de histórico para incrementar grafico de linhas
path_baseLineChart = r'\\internal.timbrasil.com.br\FileServer\TBR\Network\NetworkAssurance\CoreVASNetworkAssurance\Coor_CVNA_NSS_THD\CORE_NGN\Estrutura Rede -  Italtel\Netflow\dataLineChart.xlsx'
dfLineChart = pd.read_excel(path_baseLineChart) 

# Verificar se a planilha esta atualizada com a data de hoje. 
sheet2 = pd.read_excel(path_r001, sheet_name='Data Atualização')

# Verificar a data na terceira linha da segunda coluna
cell_value = sheet2.iloc[1, 1]  # índices começam em zero
current_date = date.today()  

# Extrair a data (sem a hora) da célula e da data atual
cell_date = cell_value.date()

print('++++++++++++++++++++++++++')
print('Data da BASE do NETFLOW:')
print(cell_date)
print('++++++++++++++++++++++++++')
print('Data de hoje:')
print(current_date)

if cell_date == current_date:
    df = pd.read_excel(path_r001)
    df1 = pd.read_excel(path_r041)

    #Tratando as OCs na fila de Oper. MOBILE VOICE:   
    #filtrar apenas as linhas que contenham a string 'Oper. Core Mobile Voice' na coluna 'Atribuição'redefinindo o indice para 0
    df_mobilevoice = df[df['Atribuição'].str.contains('Oper. Core Mobile Voice')].reset_index(drop=True)
     
    # Seleciona as colunas desejadas do export R001 e cria um novo DataFrame
    r001 = df_mobilevoice.loc[:, ['Regional',
                         'Order ID',
                         'Ordem Complexa',
                         'Data Inicio Ordem',
                         'Tipo Ordem Complexa',
                         'Nome Proprietário',
                         'Nome Atividade',
                         'Elemento ID']]
    
    # Seleciona as colunas desejadas do export R041 e cria um novo DataFrame  
    r041 = df1.loc[:, ['Ordem Complexa',
                      'Caminho',
                      'Detalhamento do Projeto']]
                      
    #vizualizando todas as colunas e linhas da base de dados
    pd.set_option('display.max_columns',None)
    pd.set_option('display.max_rows',None)
    
    # Criação de uma nova coluna na r001 na posicao 6 para receber os dados da coluna "caminho_2" da r041
    r001.insert(loc=6, column='Detalhamento do Projeto', value= '')
    r001['anexos'] = ''
    
    # Formata a coluna de data no formato brasileiro
    r001['Data Inicio Ordem'] = r001['Data Inicio Ordem'].dt.strftime('%d/%m/%Y')
    
    # Renomear a coluna "Caminho" da r041 para evitar conflito com a df
    r041.rename(columns={'Caminho': 'Caminho_2'}, inplace=True)
    r041.rename(columns={'Detalhamento do Projeto': 'Detalhamento do Projeto_2'}, inplace=True)
    
    # Renomear a coluna "Tipo Ordem Complexa" para Origem
    r001.rename(columns={'Tipo Ordem Complexa': 'Origem'}, inplace=True)

    # definindo a expressão regular para a string "TIM Network V.8.X CORE CSP - IRL", onde "X" pode ser um dígito de 0 a 9
    #regex = re.compile('^TIM\s+Network\s+V\.8\.[0-9]\s+CORE\s+CSP\s+-\s+IRL$')
    regex = re.compile('^TIM\s+Network\s+V\.([89]\.[0-9]|9\.0)\s+CORE\s+CSP\s+-\s+IRL$')

    # percorrendo cada valor da coluna 'Origem' e altera para: ITX, Acesso ou outros
    for i in range(len(r001)):
        valor = r001.loc[i, 'Origem']
        if regex.match(valor):
            r001.loc[i, 'Origem'] = 'ITX'
        
        elif valor == 'TIM Network V.9.0 Acesso + MW Remanejamento':
            r001.loc[i, 'Origem'] = 'Acesso'

        elif valor == 'TIM Network V 7.0 Packet Core + FAM':
            r001.loc[i, 'Origem'] = 'Engenharia'
        
        else:
            r001.loc[i, 'Origem'] = 'outros'
    
    # Comparação cruzada das duas planilhas usando o método merge
    df_merged = pd.merge(r001, r041, on='Ordem Complexa', how='inner')
    
    # Loop para atualizar a nova coluna na nova_df com os valores da coluna "Caminho_2" da r041
    for i in range(len(df_merged)):
        ordem_complexa = df_merged['Ordem Complexa'][i]
        Caminho_2 = df_merged['Caminho_2'][i]
        Detalhamento_do_Projeto_2 = df_merged['Detalhamento do Projeto_2'][i]
        r001.loc[r001['Ordem Complexa'] == ordem_complexa, 'anexos'] = Caminho_2
        r001.loc[r001['Ordem Complexa'] == ordem_complexa, 'Detalhamento do Projeto'] = Detalhamento_do_Projeto_2
    
    #Identifica as linhas onde a coluna 'nome proprietário' está em branco '-'
    nouser = r001['Nome Proprietário'].str.contains('-')
    nouser=(r001[nouser])
    print('#################################################################')
    print('Novas Ordens:')
    print(nouser[["Ordem Complexa", "Data Inicio Ordem"]])
    
    #criar lista de Nome de Atividade para cada grupo
    Nome_Atividade_ngnI = ['6.5.13 Executa configuração CORE NGN Italtel',
                          '6.5.13.2 Executa configuração CORE NGN Italtel',
                          '6.6.4 Executa testes e elabora DT(CORE NGN Italtel – Sigla NGN-I )',
                           '6.5.17 Executa configuração CORE SBC',
                           '6.5.17.2 Executa configuração CORE SBC',
                           '2.2 Analisa demanda (Oper CORE NGN Italtel)',
                           '2.2 Analisa demanda (Oper MV) 3',
                           '2.4 Executa demanda e anexa evidências (Oper CORE NGN Italtel)']
    
    Nome_Atividade_ngnH = ['6.5.12 Executa configuração CORE NGN Huawei',
                           '6.5.12.2 Executa configuração CORE NGN Huawei',
                           '6.6.3  Executa testes e elabora DT(CORE NGN Huawei)',
                           '2.2 Analisa demanda (Oper CORE NGN Huawei)']
    
    Nome_Atividade_coreCS = ['6.5.10 Executa configuração CORE CS', 
                            '6.5.10.2 Executa configuração CORE CS', 
                            '6.2.33 Altera labels no Core', 
                            '6.6.1  Executa testes e elabora DT (CORE CS – Sigla CS)',
                            '2.2 Analisa demanda (Oper Core CS)',
                            '2.18 Executa projeto Core CS',
                            '2.18 Executa projeto Core CS (NT56)',
                            '2.4 Executa demanda e anexa evidências (Oper Core CS)']
    
    Nome_Atividade_aceitacao = ['7.3.1.2 Realiza aceitação lógica e ou check de integrações (Core Mobile Voice)',
                                '7.3.1.1 Checa integração (CMV)',
                                '7.1.7 Verifica necessidade de execução de PQR (CMV)']
    
    #definindo os  proprietarios wllctel para contagem de OCs desta empresa.
    wllctel = ['Dalmo Silas de Castilho',
              'Leandro Monte',
              'Romulo Quaresma Nobrega',
              'Kenison Silva',
              'Marco Antonio Matos',
              'Eduardo Marcelo Chaves da Silveira']

    #Defindo os dados para filtrar no campo "Regional" para CORE CS
    TRJ = ['TRJ','rj']
    TLE = ['TLE']
    TSL = ['TSL']
    TNO = ['TNO']
    TNE = ['TNE']
    TCO = ['TCO']
    TSP = ['TSP']
        
    #pegar linhas que correspondem as OCs de NGN Italtel
    ngnI = nouser[nouser['Nome Atividade'].isin(Nome_Atividade_ngnI)]
    ngnI = ngnI.reset_index(drop=True)
    ngnI.index = pd.RangeIndex(start=1, stop=len(ngnI)+1, step=1)
    wllNgnI = r001[r001['Nome Atividade'].isin(Nome_Atividade_ngnI) & (r001['Nome Proprietário'].isin(wllctel))]
    allNgnI = r001[r001['Nome Atividade'].isin(Nome_Atividade_ngnI)]
    cNoNgnI = ngnI.shape[0]  #contagem de OCs NGN I sem atribuição
    callNgnI = allNgnI.shape[0] # contagem de todas as OCs para NGN I
    cwllNgnI = wllNgnI.shape[0] # contagem de OCs NGN I atribuidas para WLLCTEL
    cTimNgnI = callNgnI - cwllNgnI - cNoNgnI
    
    #pegar linhas que correspondem as OCs de NGN huawei
    ngnH = nouser[nouser['Nome Atividade'].isin(Nome_Atividade_ngnH)].reset_index(drop=True)
    ngnH.index = pd.RangeIndex(start=1, stop=len(ngnH)+1, step=1)
    wllNgnH = r001[r001['Nome Atividade'].isin(Nome_Atividade_ngnH) & (r001['Nome Proprietário'].isin(wllctel))]
    allNgnH = r001[r001['Nome Atividade'].isin(Nome_Atividade_ngnH)]
    cNoNgnH = ngnH.shape[0]
    callNgnH = allNgnH.shape[0]
    cwllNgnH = wllNgnH.shape[0]
    cTimNgnH = callNgnH - cwllNgnH - cNoNgnH
    
    #pegar linhas que correspondem as OCs de CORE CS TRJO
    core_trj = nouser[(nouser['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (nouser['Regional'].isin(TRJ))].reset_index(drop=True)
    core_trj.index = pd.RangeIndex(start=1, stop=len(core_trj)+1, step=1)
    wllTrj = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TRJ)) & (r001['Nome Proprietário'].isin(wllctel))]
    allTrj = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TRJ))]
    cNoTrj =  core_trj.shape[0]                 
    callTrj = allTrj.shape[0]
    cwllTrj = wllTrj.shape[0]
    cTimTrj = callTrj - cwllTrj - cNoTrj

    #pegar linhas que correspondem as OCs de CORE CS TSP
    core_tsp = nouser[(nouser['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (nouser['Regional'].isin(TSP))].reset_index(drop=True)
    core_tsp.index = pd.RangeIndex(start=1, stop=len(core_tsp)+1, step=1)
    wllTsp = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TSP)) & (r001['Nome Proprietário'].isin(wllctel))]
    allTsp = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TSP))]
    cNoTsp =  core_tsp.shape[0]                 
    callTsp = allTsp.shape[0]
    cwllTsp = wllTsp.shape[0]
    cTimTsp = callTsp - cwllTsp - cNoTsp

    #pegar linhas que correspondem as OCs de CORE CS TCO
    core_tco = nouser[(nouser['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (nouser['Regional'].isin(TCO))].reset_index(drop=True)
    core_tco.index = pd.RangeIndex(start=1, stop=len(core_tco)+1, step=1)
    wllTco = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TCO)) & (r001['Nome Proprietário'].isin(wllctel))]
    allTco = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TCO))]
    cNoTco =  core_tco.shape[0]                 
    callTco = allTco.shape[0]
    cwllTco = wllTco.shape[0]
    cTimTco = callTco - cwllTco - cNoTco

    #pegar linhas que correspondem as OCs de CORE CS TLE
    core_tle = nouser[(nouser['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (nouser['Regional'].isin(TLE))].reset_index(drop=True)
    core_tle.index = pd.RangeIndex(start=1, stop=len(core_tle)+1, step=1)
    wllTle = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TLE)) & (r001['Nome Proprietário'].isin(wllctel))]
    allTle = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TLE))]
    cNoTle =  core_tle.shape[0]                 
    callTle = allTle.shape[0]
    cwllTle = wllTle.shape[0]
    cTimTle = callTle - cwllTle - cNoTle

    #pegar linhas que correspondem as OCs de CORE CS TNE
    core_tne = nouser[(nouser['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (nouser['Regional'].isin(TNE))].reset_index(drop=True)
    core_tne.index = pd.RangeIndex(start=1, stop=len(core_tne)+1, step=1)
    wllTne = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TNE)) & (r001['Nome Proprietário'].isin(wllctel))]
    allTne = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TNE))]
    cNoTne =  core_tne.shape[0]                 
    callTne = allTne.shape[0]
    cwllTne = wllTne.shape[0]
    cTimTne = callTne - cwllTne - cNoTne

    #pegar linhas que correspondem as OCs de CORE CS TNO
    core_tno = nouser[(nouser['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (nouser['Regional'].isin(TNO))].reset_index(drop=True)
    core_tno.index = pd.RangeIndex(start=1, stop=len(core_tno)+1, step=1)
    wllTno = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TNO)) & (r001['Nome Proprietário'].isin(wllctel))]
    allTno = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TNO))]
    cNoTno =  core_tno.shape[0]                 
    callTno = allTno.shape[0]
    cwllTno = wllTno.shape[0]
    cTimTno = callTno - cwllTno - cNoTno

    #pegar linhas que correspondem as OCs de CORE CS TSL
    core_tsl = nouser[(nouser['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (nouser['Regional'].isin(TSL))].reset_index(drop=True)
    core_tsl.index = pd.RangeIndex(start=1, stop=len(core_tsl)+1, step=1)
    wllTsl = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TSL)) & (r001['Nome Proprietário'].isin(wllctel))]
    allTsl = r001[(r001['Nome Atividade'].isin(Nome_Atividade_coreCS)) & (r001['Regional'].isin(TSL))]
    cNoTsl =  core_tsl.shape[0]                 
    callTsl = allTsl.shape[0]
    cwllTsl = wllTsl.shape[0]
    cTimTsl = callTsl - cwllTsl - cNoTsl

    #pegar linhas que correspondem as OCs de Aceitacao
    aceitacao = nouser[(nouser['Nome Atividade'].isin(Nome_Atividade_aceitacao))]
    wllAc = r001[r001['Nome Atividade'].isin(Nome_Atividade_aceitacao) & (r001['Nome Proprietário'].isin(wllctel))]
    allAc = r001[r001['Nome Atividade'].isin(Nome_Atividade_aceitacao)]
    cNoAc = aceitacao.shape[0]
    callAc = allAc.shape[0]
    cwllAc = wllAc.shape[0]
    cTimAc = callAc - cwllAc - cNoAc

    # Excluir do DataFrame DemaisOCs as linhas que contenham 'Elemento ID' presente em NGN 
    DemaisOCs = nouser[~nouser['Nome Atividade'].isin(Nome_Atividade_ngnI + Nome_Atividade_ngnH + Nome_Atividade_coreCS + Nome_Atividade_aceitacao)].reset_index(drop=True)
    #DemaisOCs.index = pd.RangeIndex(start=1, stop=len(core_tsl)+1, step=1)
    wllDemais = r001[~r001['Nome Atividade'].isin(Nome_Atividade_ngnI + Nome_Atividade_ngnH + Nome_Atividade_coreCS + Nome_Atividade_aceitacao) & (r001['Nome Proprietário'].isin(wllctel))]
    allDemais = r001[~r001['Nome Atividade'].isin(Nome_Atividade_ngnI + Nome_Atividade_ngnH + Nome_Atividade_coreCS + Nome_Atividade_aceitacao)]
    cNoDemais = DemaisOCs.shape[0]
    callDemais = allDemais.shape[0]
    cwllDemais = wllDemais.shape[0]
    cTimDemais = callDemais - cwllDemais - cNoDemais

    # conta o número de linhas usando a propriedade shape para determinar o numero de OCs na fila
    totalocs = r001.shape[0]
    totalWll = cwllTsp + cwllTrj + cwllTsl + cwllTco + cwllTle + cwllTno + cwllTne + cwllNgnH + cwllNgnI + cwllAc + cwllDemais
    totalTim = cTimTsp + cTimTrj + cTimTsl + cTimTco + cTimTle + cTimTno + cTimTne + cTimNgnH + cTimNgnI + cTimAc + cTimDemais
    totalNo = cNoTsp + cNoTrj + cNoTsl + cNoTco + cNoTle + cNoTno + cNoTne + cNoNgnH + cNoNgnI + cNoAc + cNoDemais
    
    ###PLOTAGEM DE GRAFICO DE BARRAS
    # Tamanho da figura para o gráfico de barras
    fig_bar, ax_bar = plt.subplots(figsize=(6, 3))

    categories = ['TRJ', 'TSP', 'TSL', 'TCO','TLE', 'TNE', 'TNO', 'NGN I', 'NGN H']
    values1 = [(cTimTrj), (cTimTsp), (cTimTsl), (cTimTco), (cTimTle), (cTimTne), (cTimTno), (cTimNgnI), (cTimNgnH)]
    values2 = [(cwllTrj), (cwllTsp), (cwllTsl), (cwllTco), (cwllTle), (cwllTne), (cwllTno), (cwllNgnI), (cwllNgnH)]
    values3 = [(cNoTrj), (cNoTsp), (cNoTsl), (cNoTco), (cNoTle), (cNoTne), (cNoTno), (cNoNgnI), (cNoNgnH)]

    bar1 = plt.bar(categories, values1, label=(f'TIM: {totalTim}'), color='blue')
    bar2 = plt.bar(categories, values2, bottom=values1, label=(f'WLLCTEL: {totalWll}'), color='orange')
    bar3 = plt.bar(categories, values3, bottom=[i + j for i, j in zip(values1, values2)], label=(f'SEM ATRIBUICÃO: {totalNo} '), color='gray')

    #plt.xlabel('Regionais')
    plt.ylabel('Qtde de Ocs')
    plt.title(f'Total de OCs em Execução: {totalocs}')
    plt.legend()
    
    # Salvar o gráfico como imagem
    bar_chart_image_path = r'\\internal.timbrasil.com.br\FileServer\TBR\Network\NetworkAssurance\CoreVASNetworkAssurance\Coor_CVNA_NSS_THD\CORE_NGN\Estrutura Rede -  Italtel\Netflow\bar_chart_image.png'

    plt.savefig(bar_chart_image_path)
    plt.close()

    # Converter imagem em base64
    with open(bar_chart_image_path, "rb") as image_bar_chart:
        image_bar_chart_data = base64.b64encode(image_bar_chart.read()).decode()

    #PLOTAGEM DE GRAFICO DE LINHAS 
    # Criação do DataFrame com as colunas Dia, TotalOCs, TimOcs e WllOcs
    new_row_line_chart = {'Dia': current_date, 'TotalOcs': totalocs, 'TimOcs': (totalocs - totalWll), 'WllOcs': totalWll}
    dfLineChart.loc[len(dfLineChart)] = new_row_line_chart

    # Remover linhas mais antigas caso ultrapasse 30 linhas
    if len(dfLineChart) > 30:
        dfLineChart = dfLineChart.drop([0], axis=0).reset_index(drop=True)

    #print(dfLineChart)
    dfLineChart.to_excel(path_baseLineChart, index=False)

    # Converter a coluna 'Dia' para o formato de string
    dfLineChart['Dia'] = pd.to_datetime(dfLineChart['Dia'])  # Se a coluna 'Dia' não estiver no formato datetime

    # # Plotar o gráfico de linhas
    plt.figure(figsize=(6, 3))  # Define o tamanho da figura
    plt.plot(dfLineChart['Dia'].dt.strftime('%d/%m'), dfLineChart['TotalOcs'], label='Total', color='black')#, marker='o')
    plt.plot(dfLineChart['Dia'].dt.strftime('%d/%m'), dfLineChart['TimOcs'], label='Tim', color='blue')#, marker='x')
    plt.plot(dfLineChart['Dia'].dt.strftime('%d/%m'), dfLineChart['WllOcs'], label='Wllctel', color='orange')

    # Configurar os rótulos dos eixos e o título
    plt.xlabel('Dia/Mês', fontsize=8)
    plt.ylabel('Qtde de OCs')
    plt.title('Volumetria - Últimos 30 dias')

    # Adicionar legenda
    plt.legend()

    # Rotacionar os rótulos do eixo x para melhor visualização
    plt.xticks(rotation=45,fontsize=6)

    # Salvar o gráfico como imagem
    line_chart_image_path = r'\\internal.timbrasil.com.br\FileServer\TBR\Network\NetworkAssurance\CoreVASNetworkAssurance\Coor_CVNA_NSS_THD\CORE_NGN\Estrutura Rede -  Italtel\Netflow\line_chart_image.png'

    plt.savefig(line_chart_image_path)
    plt.close()

    # Converter imagem em base64
    with open(line_chart_image_path, "rb") as image_line_chart:
        image_line_chart_data = base64.b64encode(image_line_chart.read()).decode()

    #Salva as ordens de Oper. Mobile Voice em excel para enviar por e-mail.
    filename = r'\\internal.timbrasil.com.br\FileServer\TBR\Network\NetworkAssurance\CoreVASNetworkAssurance\Coor_CVNA_NSS_THD\CORE_NGN\All_OCs_Mobile_Voice.xlsx'
    r001.to_excel(filename, index=False)
    
    # Inicializa o objeto de envio de e-mail do Outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    
    # Define os destinatários, assunto e corpo do e-mail
    #mail.To = ('jolenski@timbrasil.com.br; acamacho@timbrasil.com.br')
    #mail.To = ('dl_ctio_coreoperations_mobilevoiceplatforms@timbrasil.com.br')
    mail.To = (email)
    mail.Subject = 'OCs - NETFLOW - FILA --> Oper. Mobile Voice - BETA'
    mail.HTMLBody = f'''
    
    <p>Prezados(as), bom dia!</p> 
        
    <img src="data:image/png;base64,{image_bar_chart_data}" alt="Gráfico de Barras">
    <img src="data:image/png;base64,{image_line_chart_data}" alt="Gráfico de Barras">

    <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Segue o relatório contendo as novas OCs:</p>

    '''
    
    if len(ngnI) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">NGN ITALTEL</p>'
        mail.HTMLBody += ngnI.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">NGN ITALTEL.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'
    
     
    if len(ngnH) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">NGN HUAWEI</p>'
        mail.HTMLBody += ngnH.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">NGN HUAWEI.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'
    
    if len(core_trj) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - RJO</p>'
        mail.HTMLBody += core_trj.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - RJO.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'
    
    if len(core_tno) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TNO</p>'
        mail.HTMLBody += core_tno.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TNO.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'
    
    if len(core_tco) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TCO</p>'
        mail.HTMLBody += core_tco.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TCO.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'
    
    
    if len(core_tle) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TLE</p>'
        mail.HTMLBody += core_tle.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TLE.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'
    
    
    if len(core_tsp) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TSP</p>'
        mail.HTMLBody += core_tsp.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TSP.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'
    
    if len(core_tne) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TNE</p>'
        mail.HTMLBody += core_tne.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TNE.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'
    
    if len(core_tsl) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TSL</p>'
        mail.HTMLBody += core_tsl.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">CORE CS - TSL.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'

    if len(aceitacao) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">ACEITAÇÕES</p>'
        mail.HTMLBody += aceitacao.to_html()
    else:
        mail.HTMLBody += '<p style="font-size: 16pt">ACEITAÇÕES.</p>'
        mail.HTMLBody += '<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Não há OCs sem atribuição.</p>'
    
    if len(DemaisOCs) >= 1:
        mail.HTMLBody += '<p style="font-size: 16pt">DEMAIS OCs</p>'
        mail.HTMLBody += DemaisOCs.to_html()
    
    # Adiciona um anexo (opcional)
    attachment = filename
    mail.Attachments.Add(attachment)
    
    # Envia o e-mail
    mail.Send()

else:
    print('DB do Netflow não esta com a data de hoje, verificar.')

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = (email)
    mail.Subject = 'OCs - NETFLOW - FILA --> Oper. Mobile Voice - BETA'
    mail.HTMLBody = f'''
    
    <p>Prezados(as), bom dia!</p> 
    
    <p>RELATÓRIO NÃO DISPONÍVEL: PMOnline desatualizado</b></p> 
   
    '''
    mail.Send()