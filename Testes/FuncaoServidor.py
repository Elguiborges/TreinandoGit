import paramiko
import pandas as pd
from io import BytesIO

def acessar_df(ssh_host,ssh_port,ssh_user,ssh_password,remote_xlsx_folder):

#   Criando a conexão SSH
    ssh_client = paramiko.SSHClient()
    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())  # Aceitar automaticamente chaves de host desconhecidas
    ssh_client.load_system_host_keys()
    ssh_client.connect(ssh_host, port=ssh_port, username=ssh_user, password=ssh_password)

#   Usar SFTP para listar os arquivos Excel na pasta remota
    sftp = ssh_client.open_sftp()
    remote_files = sftp.listdir(remote_xlsx_folder)

#   Identificar o arquivo mais recente com base na data de modificação
    latest_file = None
    latest_mod_time = None

    for remote_file in remote_files:
        remote_file_path = remote_xlsx_folder + remote_file
        file_attributes = sftp.stat(remote_file_path)
        mod_time = file_attributes.st_mtime
        if latest_mod_time is None or mod_time > latest_mod_time:
            latest_mod_time = mod_time
            latest_file = remote_file

# Verificar se um arquivo foi encontrado na pasta remota
    if latest_file is not None:
    # Cria um buffer de bytes temporário para armazenar o arquivo XLSX
        with BytesIO() as xlsx_buffer:
        # Baixa o arquivo mais recente para o buffer de bytes
            sftp.getfo(remote_xlsx_folder + latest_file, xlsx_buffer)
            xlsx_buffer.seek(0)  # Volta para o início do buffer

        # Lê o conteúdo do buffer de bytes como um DataFrame do Pandas
            df = pd.read_excel(xlsx_buffer, header = 2)
            print(f"Lido o arquivo mais recente: {latest_file}")
    else:
        print("Nenhum arquivo Excel encontrado na pasta remota.")

# Fechar a conexão SSH
    sftp.close()
    ssh_client.close()

    return df