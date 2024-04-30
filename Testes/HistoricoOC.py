import pandas as pd
from datetime import date

#   Declarando o diretório e atribuindo a um DataFrame no pandas
path_hist = r'\\internal.timbrasil.com.br\FileServer\TBR\Network\NetworkAssurance\CoreVASNetworkAssurance\Coor_CVNA_NSS_THD\CORE_NGN\Historico_OCs.xlsx'
df_hist = pd.read_excel(path_hist)
path_ocs = r'\\internal.timbrasil.com.br\FileServer\TBR\Network\NetworkAssurance\CoreVASNetworkAssurance\Coor_CVNA_NSS_THD\CORE_NGN\All_OCs_Mobile_Voice.xlsx'
df_ocs = pd.read_excel(path_ocs)

#   Declarando as variáveis Dia e Total de OC's
current_date = date.today()
totalocs = df_ocs.shape[0]

#   Inserindo a linha na tabela de histórico
novaLinha = {'Dia': current_date, 'TotalOcs': totalocs}
df_hist.loc[len(df_hist)] = novaLinha
df_hist['Dia'] = pd.to_datetime(df_hist['Dia']) #Formatando em datetime

#   Condição para manter apenas os últimos 30 dias
if len(df_hist) > 30:
    df_hist = df_hist.drop([0], axis=0).reset_index(drop=True)

#   Transformando o DataFrame em Excel
df_hist.to_excel(path_hist, index=False)