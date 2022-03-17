import datetime

import pandas as pd

from src.utils import ajusta_data
from src.dict_config import iconsConfianca, iconsEngajamento, iconsImpacto, iconsStatus, iconHome

""""As funções desse módulo precisam ser personalizada conforme o caso de uso"""


def load_data(data):
    # Carrega aba Entregas
    df = pd.read_excel(data, sheet_name='Entregas')

    # Carrega aba Ações
    marcos = pd.read_excel(data, sheet_name='Acoes')
    marcos = calc_progresso(marcos)

    df['hyperlink'] = df['Pilar'].apply(lambda x: iconHome.get(x, ""))

    df = atualiza_marcos(df, marcos)

    df = calc_progresso(df)

    df['Desvio'] = df['Progresso (%)'] - df['Progresso Esperado']
    df['Desvio'].fillna(0, inplace=True)
    df.loc[df['Status'] == "Concluído"]['Desvio'] = 0

    df = atualiza_status(df)

    # Converte datetimes para string
    for col in df:
        if df[col].dtype == 'datetime64[ns]':
            df[col] = df[col].apply(ajusta_data).fillna("")

    # Converte valor percentual para string
    df['Progresso (%)'] = df['Progresso (%)'].fillna(0).apply(lambda x: str(int(x)) + "%")

    # Busca imagens
    df['Grau de Confiança'] = df['Grau de Confiança'].apply(lambda x: iconsConfianca.get(x, ""))
    df['Engajamento Unidade'] = df['Engajamento Unidade'].apply(lambda x: iconsEngajamento.get(x, ""))
    df['Impacto'] = df['Impacto'].apply(lambda x: iconsImpacto.get(x, ""))
    df['IconeStatus'] = df['Status'].apply(lambda x: iconsStatus.get(x, ""))
    df['IconeStatusMarco'] = df['Status Marco Entrega'].apply(lambda x: iconsStatus.get(x, ""))

    for i in range(1, 7):
        df[f'Status Marco {i}'] = df[f'Status Marco {i}'].apply(lambda x: iconsStatus.get(x, ""))

    df.sort_values(by=['Pilar', 'Nº'], ascending=[True, True], inplace=True)

    return df


def atualiza_marcos(df, marcos):
    marcos.sort_values(by=['Data fim previsto', 'Número ação'], ascending=[True, True], inplace=True)

    for i in range(1, 7):
        df[f'Status Marco {i}'] = ''
        df[f'Marco {i}'] = ''
        df[f'Prazo Marco {i}'] = ''

    for i, entrega in df.iterrows():
        if marcos.loc[marcos['Nº Entrega'] == entrega['Nº']].shape[0] <= 6:
            nrmarco = 1
            for x, marco in marcos.loc[marcos['Nº Entrega'] == entrega['Nº']].iterrows():
                df.at[i, {f'Status Marco {nrmarco}'}] = marco['Status']
                df.at[i, {f'Marco {nrmarco}'}] = marco['Nome ação']
                df.at[i, {f'Prazo Marco {nrmarco}'}] = ajusta_data(marco['Data fim previsto'])
                nrmarco += 1
        else:
            nrmarco = 1
            for x, marco in marcos.loc[(marcos['Nº Entrega'] == entrega['Nº']) & (marcos['Status'] != 'Concluído')][
                            0:5].iterrows():
                df.at[i, {f'Status Marco {nrmarco}'}] = marco['Status']
                df.at[i, {f'Marco {nrmarco}'}] = marco['Nome ação']
                df.at[i, {f'Prazo Marco {nrmarco}'}] = ajusta_data(marco['Data fim previsto'])
                nrmarco += 1

    return df


def calc_progresso(df):
    df['Progresso Esperado'] = 0
    for i, row in df.iterrows():
        if row.Status == "Concluído":
            df.at[i, 'Progresso Esperado'] = 1
        if row.Status == "A Iniciar":
            df.at[i, 'Progresso Esperado'] = 0
        else:
            try:
                df.at[i, 'Progresso Esperado'] = min(
                    ((datetime.date.today() - row['Data início previsto'].date()) / row['Prazo previsto']).days / row[
                        'Prazo previsto'], 1)
            except:
                df.at[i, 'Progresso Esperado'] = 0
    return df


def atualiza_status(df):
    for i, row in df.iterrows():
        if row.Status == "Desenvolvimento":
            if row.Desvio > 0.2:
                df.at[i, 'Status'] = 'Atraso'
    return df
