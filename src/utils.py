import copy
import datetime
import pandas as pd


def change_pic(slide, old_picture, pic_file):
    if pic_file == "":
        old_picture._element.getparent().remove(old_picture._element)
        return None
    # Save the position and shape of old picture
    x, y, cx, cy = old_picture.left, old_picture.top, old_picture.width, old_picture.height

    # Add the new picture using old's picture shape and size
    slide.add_picture(pic_file, x, y, cx, cy)

    # Remove the old picture
    old_picture._element.getparent().remove(old_picture._element)
    return None


def ajusta_data(date):
    try:
        return date.date().strftime("%d/%m/%Y")
    except ValueError:
        return ''
    except AttributeError:
        return ''


# Em teste - Instavel
def duplicate_slide(pres, index):
    source = pres.slides[index]
    try:
        blank_slide_layout = pres.slide_layouts[6]
    except:
        blank_slide_layout = pres.slide_layouts[len(pres.slide_layouts) - 1]
    dest = pres.slides.add_slide(blank_slide_layout)

    for shp in source.shapes:
        el = shp.element
        newel = copy.deepcopy(el)
        dest.shapes._spTree.insert_element_before(newel, 'p:extLst')

    return dest


iconsConfianca = {
    "Alto": 'media/confiancaAlto.png',
    "Médio": 'media/confiancaMedio.png',
    "Baixo": 'media/confiancaBaixo.png',
}

iconsEngajamento = {
    "Alto": 'media/engajamentoAlto.png',
    "Médio": 'media/engajamentoMedio.png',
    "Baixo": 'media/engajamentoBaixo.png',
}

iconsImpacto = {
    "Alto": 'media/impactoAlto.png',
    "Médio": 'media/impactoMedio.png',
    "Baixo": 'media/impactoBaixo.png',
}

iconsStatus = {
    "Concluído": 'media/statusConcluido.png',
    "Desenvolvimento": 'media/statusDesenvolvimento.png',
    "Em desenvolvimento": 'media/statusDesenvolvimento.png',
    "Atraso": 'media/statusAtraso.png',
    "Suspenso": 'media/statusSuspenso.png',
    "A Iniciar": 'media/statusIniciar.png',
    "A iniciar": 'media/statusIniciar.png'
}

iconHome = {
    "Governança": 7,
    "Soluções Analíticas": 46,
    "Ambientes e Ferramentas Analíticas": 126,
    "Ambiente e Ferramentas": 126,
    "ECOA": 140
}


def atualiza_status(df):
    for i, row in df.iterrows():
        if row.Status == "Desenvolvimento":
            if row.Desvio > 0.2:
                df.at[i, 'Status'] = 'Atraso'
    return df


def calc_progresso(df):
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


def atualiza_marcos(df, marcos):
    marcos.sort_values(by=['Data fim previsto', 'Número ação'], ascending=[True, True], inplace=True)

    df['Status Marco 1'] = ''
    df['Status Marco 2'] = ''
    df['Status Marco 3'] = ''
    df['Status Marco 4'] = ''
    df['Status Marco 5'] = ''
    df['Status Marco 6'] = ''

    df['Marco 1'] = ''
    df['Marco 2'] = ''
    df['Marco 3'] = ''
    df['Marco 4'] = ''
    df['Marco 5'] = ''
    df['Marco 6'] = ''

    df['Prazo Marco 1'] = ''
    df['Prazo Marco 2'] = ''
    df['Prazo Marco 3'] = ''
    df['Prazo Marco 4'] = ''
    df['Prazo Marco 5'] = ''
    df['Prazo Marco 6'] = ''

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


def carrega_marcos():
    marcos = pd.read_excel(
        r"PATH_TO_FILE.xlsx")

    marcos['Progresso Esperado'] = 0

    marcos = calc_progresso(marcos)

    return marcos


def trata_base(df):

    df['hyperlink'] = df['Pilar'].apply(lambda x: iconHome.get(x, ""))

    marcos = carrega_marcos()

    df = atualiza_marcos(df, marcos)

    df['Progresso Esperado'] = 0

    df = calc_progresso(df)

    df['Desvio'] = df['Progresso (%)'] - df['Progresso Esperado']
    df['Desvio'].fillna(0, inplace=True)
    df.loc[df['Status'] == "Concluído"]['Desvio'] = 0

    df = atualiza_status(df)

    df['Data início previsto'] = df['Data início previsto'].apply(ajusta_data).fillna("")
    df['Data fim previsto'] = df['Data fim previsto'].apply(ajusta_data).fillna("")
    df['Data Fim'] = df['Data Fim'].apply(ajusta_data).fillna("")
    # df['Prazo Marco Entrega'] = df['Prazo Marco Entrega'].apply(ajusta_data).fillna("")
    df['Progresso (%)'] = df['Progresso (%)'].fillna(0).apply(lambda x: str(int(x)) + "%")
    df['Grau de Confiança'] = df['Grau de Confiança'].apply(lambda x: iconsConfianca.get(x, ""))
    df['Engajamento Unidade'] = df['Engajamento Unidade'].apply(lambda x: iconsEngajamento.get(x, ""))
    df['Impacto'] = df['Impacto'].apply(lambda x: iconsImpacto.get(x, ""))
    df['IconeStatus'] = df['Status'].apply(lambda x: iconsStatus.get(x, ""))
    df['IconeStatusMarco'] = df['Status Marco Entrega'].apply(lambda x: iconsStatus.get(x, ""))

    df['Status Marco 1'] = df['Status Marco 1'].apply(lambda x: iconsStatus.get(x, ""))
    df['Status Marco 2'] = df['Status Marco 2'].apply(lambda x: iconsStatus.get(x, ""))
    df['Status Marco 3'] = df['Status Marco 3'].apply(lambda x: iconsStatus.get(x, ""))
    df['Status Marco 4'] = df['Status Marco 4'].apply(lambda x: iconsStatus.get(x, ""))
    df['Status Marco 5'] = df['Status Marco 5'].apply(lambda x: iconsStatus.get(x, ""))
    df['Status Marco 6'] = df['Status Marco 6'].apply(lambda x: iconsStatus.get(x, ""))

    df.sort_values(by=['Pilar', 'Nº'], ascending=[True, True], inplace=True)

    return df


def clear_text(shape):
    """" Clear all texts from a shape """
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for run in p.runs:
                run.text = ""
    return shape


def change_text(shape, new_text):
    """" Insert text in a shape """
    if shape.has_text_frame:
        shape.text_frame.paragraphs[0].runs[0].text = new_text
    return shape
