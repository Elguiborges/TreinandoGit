import pandas as pd
import win32com.client as win32
from IPython.display import display, HTML
import paramiko
from io import BytesIO
from datetime import datetime

#   Configurações do servidor SSH 
ssh_host = "10.216.127.32"  # endereço do servidor SSH
ssh_port = 22  # Porta SSH padrão
ssh_user = "microstrategy"  # nome de usuário SSH
ssh_password = "Micro!2022"  

#   Criando a conexão SSH
ssh_client = paramiko.SSHClient()
ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())  # Aceitar automaticamente chaves de host desconhecidas
ssh_client.load_system_host_keys()
ssh_client.connect(ssh_host, port=ssh_port, username=ssh_user, password=ssh_password)
#   Pasta onde estão os arquivos Excel no servidor SSH
remote_xlsx_folder = "/opt/codes/ftp/microstrategy/" 

#   Usar SFTP para listar os arquivos Excel na pasta remota
sftp = ssh_client.open_sftp()
remote_files = sftp.listdir(remote_xlsx_folder)

#   Identificar o arquivo mais recente com base na data de modificação
latest_file = None
latest_mod_time = None

for remote_file in remote_files:
    if "Movimento" in remote_file:
        remote_file_path = remote_xlsx_folder + remote_file
        file_attributes = sftp.stat(remote_file_path)
        mod_time = file_attributes.st_mtime
        if latest_mod_time is None or mod_time > latest_mod_time:
            latest_mod_time = mod_time
            latest_file = remote_file

dataarquivo = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d')
dataatual = datetime.today().strftime('%Y-%m-%d')

#  Verificar se um arquivo foi encontrado na pasta remota
if latest_file is not None:
    #   Cria um buffer de bytes temporário para armazenar o arquivo Xlsx
    with BytesIO() as xlsx_buffer:
        # Baixa o arquivo mais recente para o buffer de bytes
        sftp.getfo(remote_xlsx_folder + latest_file, xlsx_buffer)
        xlsx_buffer.seek(0)  # Volta para o início do buffer

        # Lê o conteúdo do buffer de bytes como um DataFrame do Pandas
        df = pd.read_excel(xlsx_buffer, header=2)
        print ("-"*100)
        print(f"Lido o arquivo mais recente: {dataarquivo}")
else:
    print("Nenhum arquivo Excel encontrado na pasta remota.")

#   Fechar a conexão SSH
sftp.close()
ssh_client.close()

#   Selecionando quais colunas usar
colunas = df.loc[:, ['Data Movimento WO/NO',
                         'ID Ticket WO/NO',
                         'Status Ticket WO/NO',
                         'Login Usuario Executor WO/NO',
                         'Unnamed: 26',
                         'Operadora NE WO/NO', 
                         'Descrição WO/NO',
                         'Tipo Ne',
                         'Regra Atribuida WO/NO']]

#   Removendo repetições e exibindo por ordem alfabéica de Operadora
colunas = colunas.rename(columns={colunas.columns[4]: 'Nome Usuario WO/NO'})
colunas.drop_duplicates(keep='first',inplace=True)
colunas.sort_values(by=['Operadora NE WO/NO','Data Movimento WO/NO'], inplace=True, ascending=[True, True])
colunas = colunas.loc[(colunas['Regra Atribuida WO/NO'] == 'ANALYSER-CORE CONFIG CS&NGN')]
colunas.drop(columns=['Regra Atribuida WO/NO'], inplace=True)
colunas['Data Movimento WO/NO'] = colunas['Data Movimento WO/NO'].dt.strftime('%d-%m-%Y')

#   Declarando dataframe das WO não atribuídas
df_na = colunas.loc[(colunas['Status Ticket WO/NO'] == 'ASSIGNED') &  (colunas['Login Usuario Executor WO/NO'] == 'NAO INFORMADO')]
df_na.reset_index(drop=True, inplace=True)
df_na.index = pd.RangeIndex(start=1, stop=len(df_na)+1, step=1)
num_na = len(df_na)

#   Declarando dataframe com o restante das WO pendentes
df_restante = colunas.loc[(colunas['Login Usuario Executor WO/NO'] != 'NAO INFORMADO') & (colunas['Status Ticket WO/NO'] == 'ASSIGNED') | (colunas['Status Ticket WO/NO'] == 'PENDING') | (colunas['Status Ticket WO/NO'] == 'WORKING')]
df_restante.reset_index(drop=True, inplace=True)
df_restante.index = pd.RangeIndex(start=1, stop=len(df_restante)+1, step=1)
num_restante = len(df_restante)

#   Convertendo os DataFrames em HTML
df_na_html = df_na.to_html()
df_na_html = df_na_html.replace('<table', '<table style="font-size: 13px;"')
df_restante_html = df_restante.to_html()
df_restante_html = df_restante_html.replace('<table', '<table style="font-size: 13px;"')

#   Enviando o dataframe por e-mail
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'dl_ctio_coreoperations_mobilevoiceplatforms@timbrasil.com.br'
mail.Subject = 'WO/NO - xTTS - Regra --> ANALYSER-CORE CONFIG CS&NGN'

if dataarquivo == dataatual:
    mail.HTMLBody = f'''
    <p>Prezados(as), bom dia!</p>

    <p>Total de WO's/NO's em execução: {num_restante+num_na}</p>

    <p>Segue o relatório contendo as novas WO's/NO's:</p>

    {df_na_html}

    <p>Segue também o relatório com as demais WO's/NO's em tratamento:</p>

    {df_restante_html}

    <p>Agradecemos desde já!</p>

    <p>Att.,</p>
    <p>Guilherme Borges</p>
    '''
    mail.Send()
    print ("-"*100)
    print("Email Enviado.")

else:
    print ("-"*100)
    print(f"A data do arquivo é diferente da data atual ({dataatual}). O que deseja fazer?")
    resposta = input("1- Enviar mesmo assim \n2- Enviar aviso de indisponibilidade \n3- Cancelar envio \n")
    if resposta == "1":
        mail.HTMLBody = f'''
        <p>Prezados(as), bom dia!</p>

        <p>Total de WO's/NO's em execução: {num_restante+num_na}</p>

        <p>Segue o relatório contendo as novas WO's/NO's:</p>

        {df_na_html}

        <p>Segue também o relatório com as demais WO's/NO's em tratamento:</p>

        {df_restante_html}

        <p>Agradecemos desde já!</p>

        <p>Att.,</p>
        <p>Guilherme Borges</p>
        '''
        mail.Send()
        print ("-"*100)
        print("Email Enviado.")

    elif resposta == "2":
        mail.HTMLBody = f'''
        <p>Prezados(as), bom dia!</p>

        <p>O relatório não está disponível ou não está atualizado com a data de hoje em nossa base de dados.</p>
        <p>Pedimos desculpas pela ausência e ficamos à disposição em caso de dúvidas.</p>

        <p>Att.,</p>
        <p>Guilherme Borges</p>
        '''
        mail.Send()
        print("Email Enviado.")

    else:
        print ("-"*100)
        print("Envio cancelado.")