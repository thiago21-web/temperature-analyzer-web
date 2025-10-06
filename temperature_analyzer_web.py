
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from io import BytesIO
import os
import numpy as np
import re
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Configuração da página
st.set_page_config(page_title="Analisador de Temperatura e Energia", layout="wide")

# Funções de lógica adaptadas
def get_button_colors(modo):
    return "#00FF00" if modo == "SITRAD" else "#FFA500" if modo == "DATALOGGER" else "#FFFF00"

def get_marker_color(filtro=False):
    return "orange" if filtro else "red"

def atualizar_menu_tipo_valor(modo, df):
    if modo == "DATALOGGER":
        return ["Temperatura", "Umidade"]
    elif modo == "SITRAD":
        return ["Temperatura"] + [col for col in df.columns if col.startswith("Dados_Extra")]
    else:  # ENERGIA
        return ["Potência"] + [col for col in df.columns if col.startswith("Potência_Trafo")]

@st.cache_data
def analisar_arquivo(uploaded_file, modo, aba=None):
    dados = []
    if modo == "SITRAD":
        excel = pd.ExcelFile(uploaded_file)
        abas = excel.sheet_names if aba is None else [aba]
        for sheet in abas:
            df = excel.parse(sheet)
            columns = ["DataHora", "Temperatura"] + [f"Dados_Extra{i+1}" for i in range(len(df.columns)-2)]
            df = df.iloc[:, :len(columns)]
            df.columns = columns
            df["Aba"] = sheet
            df["DataHora"] = pd.to_datetime(df["DataHora"], errors="coerce")
            df["Temperatura"] = pd.to_numeric(df["Temperatura"], errors="coerce")
            for col in columns[2:-1]:
                df[col] = pd.to_numeric(df[col], errors="coerce")
            dados.append(df)
    elif modo == "DATALOGGER":
        excel = pd.ExcelFile(uploaded_file)
        abas = excel.sheet_names[1:2] if aba is None else [aba]
        for sheet in abas:
            df = excel.parse(sheet)
            df = df.iloc[:, 1:4]
            df.columns = ["DataHora", "Temperatura", "Umidade"]
            df["Aba"] = sheet
            df["DataHora"] = pd.to_datetime(df["DataHora"], errors="coerce")
            df["Temperatura"] = pd.to_numeric(df["Temperatura"], errors="coerce")
            df["Umidade"] = pd.to_numeric(df["Umidade"], errors="coerce")
            dados.append(df)
    else:  # ENERGIA
        df = pd.read_csv(uploaded_file)
        if len(df.columns) < 2:
            raise ValueError("O arquivo CSV deve conter pelo menos 2 colunas (Data/Hora e Potência).")
        columns = ["DataHora", "Potência"]
        if len(df.columns) > 2:
            columns.extend([f"Potência_Trafo{i+2}" for i in range(min(3, len(df.columns)-2))])
        df = df.iloc[:, :len(columns)]
        df.columns = columns
        df["Aba"] = "ENERGIA"
        if pd.api.types.is_numeric_dtype(df["DataHora"]):
            df["DataHora"] = pd.to_datetime(df["DataHora"], unit="ms", errors="coerce").dt.tz_localize("UTC").dt.tz_convert("America/Sao_Paulo")
        else:
            df["DataHora"] = pd.to_datetime(df["DataHora"], errors="coerce")
        if df["DataHora"].isna().all():
            raise ValueError("Formato de data/hora inválido.")
        df["DataHora_str"] = df["DataHora"].dt.strftime("%Y/%m/%d %H:%M:%S")
        df["Potência"] = pd.to_numeric(df["Potência"], errors="coerce")
        for col in columns[2:]:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        dados.append(df)
    return pd.concat(dados, ignore_index=True)

def gerar_grafico(df, modo, filtro_ativo, filtro_valor_min, filtro_valor_max, filtro_valor_tipo, escala_y_min, escala_y_max, title="Gráfico de Dados", pontos_marcados=None, pontos_filtrados=None):
    fig = make_subplots(rows=1, cols=1, shared_xaxes=True, specs=[[ {"secondary_y": True} if modo == "DATALOGGER" else {"secondary_y": False}]])
    main_column = "Potência" if modo == "ENERGIA" else "Temperatura"
    default_color = "navy"
    highlight_color = "yellow"
    unit = "W" if modo == "ENERGIA" else "°C"
    x = df["DataHora"]
    y = df[main_column]
    fig.add_trace(go.Scatter(x=x, y=y, mode="lines+markers", name=main_column, line=dict(color=default_color)), secondary_y=False)
    if filtro_ativo == "valor" and filtro_valor_min is not None and filtro_valor_max is not None and filtro_valor_tipo == main_column:
        mask = (y >= filtro_valor_min) & (y <= filtro_valor_max) & (~pd.isna(y))
        fig.add_trace(go.Scatter(x=x[mask], y=y[mask], mode="lines+markers", name=f"{main_column} (Filtro)", line=dict(color=highlight_color)), secondary_y=False)
    if modo == "DATALOGGER" and "Umidade" in df.columns:
        y_umidade = df["Umidade"]
        fig.add_trace(go.Scatter(x=x, y=y_umidade, mode="lines+markers", name="Umidade", line=dict(color="lightblue")), secondary_y=True)
        if filtro_ativo == "valor" and filtro_valor_min is not None and filtro_valor_max is not None and filtro_valor_tipo == "Umidade":
            mask = (y_umidade >= filtro_valor_min) & (y_umidade <= filtro_valor_max) & (~pd.isna(y_umidade))
            fig.add_trace(go.Scatter(x=x[mask], y=y_umidade[mask], mode="lines+markers", name="Umidade (Filtro)", line=dict(color=highlight_color)), secondary_y=True)
        fig.update_yaxes(title_text="Umidade (%)", secondary_y=True)
    if modo == "SITRAD" or modo == "ENERGIA":
        colors = ["darkgreen", "purple", "deeppink"]
        extra_columns = [c for c in df.columns if c.startswith("Dados_Extra") or c.startswith("Potência_Trafo")]
        for i, col in enumerate(extra_columns):
            y_extra = df[col]
            fig.add_trace(go.Scatter(x=x, y=y_extra, mode="lines+markers", name=col, line=dict(color=colors[i % len(colors)])), secondary_y=False)
            if filtro_ativo == "valor" and filtro_valor_min is not None and filtro_valor_max is not None and filtro_valor_tipo == col:
                mask = (y_extra >= filtro_valor_min) & (y_extra <= filtro_valor_max) & (~pd.isna(y_extra))
                fig.add_trace(go.Scatter(x=x[mask], y=y_extra[mask], mode="lines+markers", name=f"{col} (Filtro)", line=dict(color=highlight_color)), secondary_y=False)
    if pontos_filtrados:
        for _, x, y, tipo in pontos_filtrados:
            fig.add_trace(go.Scatter(x=[x], y=[y], mode="markers+text", name=f"{tipo} (Filtro)", marker=dict(color="orange", size=10, symbol="x"), text=[f"{x.strftime('%Y/%m/%d %H:%M:%S')}\n{y:.1f}"], textposition="top center"), secondary_y=(tipo == "Umidade"))
    if pontos_marcados:
        for _, _, x, y, tipo in pontos_marcados:
            fig.add_trace(go.Scatter(x=[x], y=[y], mode="markers+text", name=f"{tipo} (Marcado)", marker=dict(color="red", size=10, symbol="x"), text=[f"{x.strftime('%Y/%m/%d %H:%M:%S')}\n{y:.1f}"], textposition="top center"), secondary_y=(tipo == "Umidade"))
    fig.update_layout(title=title, xaxis_title="Data e Hora", yaxis_title=f"{main_column} ({unit})", legend_orientation="h", height=600)
    if escala_y_min is not None and escala_y_max is not None:
        fig.update_yaxes(range=[escala_y_min, escala_y_max], secondary_y=False)
    return fig

def mostrar_estatisticas(df, modo, pontos_filtrados=None, pontos_marcados=None):
    st.subheader("Estatísticas")
    main_column = "Potência" if modo == "ENERGIA" else "Temperatura"
    total = len(df)
    validos = len(df.dropna(subset=[main_column]))
    invalidos = total - validos
    st.write(f"Total: {total} | Válidos: {validos} | Inválidos: {invalidos}")
    unit = "W" if modo == "ENERGIA" else "°C"
    max_val = df[main_column].max()
    min_val = df[main_column].min()
    mean_val = df[main_column].mean()
    st.write(f"{main_column} - Máxima: {max_val:.2f}{unit} | Mínima: {min_val:.2f}{unit} | Média: {mean_val:.2f}{unit}")
    if modo == "DATALOGGER" and "Umidade" in df.columns:
        umidade_max = df["Umidade"].max()
        umidade_min = df["Umidade"].min()
        umidade_media = df["Umidade"].mean()
        st.write(f"Umidade - Máxima: {umidade_max:.2f}% | Mínima: {umidade_min:.2f}% | Média: {umidade_media:.2f}%")
    if modo == "SITRAD" or modo == "ENERGIA":
        extra_columns = [c for c in df.columns if c.startswith("Dados_Extra") or c.startswith("Potência_Trafo")]
        for col in extra_columns:
            max_val = df[col].max()
            min_val = df[col].min()
            mean_val = df[col].mean()
            unit_extra = "W" if col.startswith("Potência_Trafo") else "°C"
            st.write(f"{col} - Máxima: {max_val:.2f}{unit_extra} | Mínima: {min_val:.2f}{unit_extra} | Média: {mean_val:.2f}{unit_extra}")
    inicio = df["DataHora"].min().strftime("%Y/%m/%d %H:%M:%S") if not pd.isnull(df["DataHora"].min()) else "N/A"
    fim = df["DataHora"].max().strftime("%Y/%m/%d %H:%M:%S") if not pd.isnull(df["DataHora"].max()) else "N/A"
    st.write(f"Intervalo: {inicio} até {fim}")
    if pontos_filtrados:
        st.write("Pontos Filtrados:")
        for _, x, y, tipo in pontos_filtrados:
            unit_p = "%" if tipo == "Umidade" else ("W" if "Potência" in tipo else "°C")
            st.write(f"{x.strftime('%Y/%m/%d %H:%M:%S')} - {y:.1f}{unit_p} ({tipo})")
    if pontos_marcados:
        st.write("Pontos Marcados Manualmente:")
        for _, _, x, y, tipo in pontos_marcados:
            unit_p = "%" if tipo == "Umidade" else ("W" if "Potência" in tipo else "°C")
            st.write(f"{x.strftime('%Y/%m/%d %H:%M:%S')} - {y:.1f}{unit_p} ({tipo})")

def main():
    st.title("Analisador de Temperatura e Energia")

    # Inicialização do estado da sessão
    if 'dados_consolidados' not in st.session_state:
        st.session_state['dados_consolidados'] = pd.DataFrame()
    if 'dados_filtrados' not in st.session_state:
        st.session_state['dados_filtrados'] = pd.DataFrame()
    if 'pontos_marcados' not in st.session_state:
        st.session_state['pontos_marcados'] = []
    if 'pontos_filtrados' not in st.session_state:
        st.session_state['pontos_filtrados'] = []
    if 'faixas' not in st.session_state:
        st.session_state['faixas'] = []
    if 'filtro_ativo' not in st.session_state:
        st.session_state['filtro_ativo'] = None
    if 'filtro_valor_min' not in st.session_state:
        st.session_state['filtro_valor_min'] = None
    if 'filtro_valor_max' not in st.session_state:
        st.session_state['filtro_valor_max'] = None
    if 'filtro_valor_tipo' not in st.session_state:
        st.session_state['filtro_valor_tipo'] = None
    if 'escala_y_min' not in st.session_state:
        st.session_state['escala_y_min'] = None
    if 'escala_y_max' not in st.session_state:
        st.session_state['escala_y_max'] = None

    # Seleção de modo
    modo = st.selectbox("Selecione o Modo", ["SITRAD", "DATALOGGER", "ENERGIA"], key="modo")
    if st.button("Resetar Tudo"):
        st.session_state['dados_consolidados'] = pd.DataFrame()
        st.session_state['dados_filtrados'] = pd.DataFrame()
        st.session_state['pontos_marcados'] = []
        st.session_state['pontos_filtrados'] = []
        st.session_state['faixas'] = []
        st.session_state['filtro_ativo'] = None
        st.session_state['filtro_valor_min'] = None
        st.session_state['filtro_valor_max'] = None
        st.session_state['filtro_valor_tipo'] = None
        st.session_state['escala_y_min'] = None
        st.session_state['escala_y_max'] = None
        st.rerun()

    # Upload de arquivo
    file_types = ["xlsx", "xls"] if modo != "ENERGIA" else ["csv"]
    uploaded_file = st.file_uploader("Selecione o Arquivo", type=file_types)
    if uploaded_file:
        try:
            if modo in ["SITRAD", "DATALOGGER"]:
                excel = pd.ExcelFile(uploaded_file)
                aba = st.selectbox("Selecione a Aba", excel.sheet_names)
                st.session_state['dados_consolidados'] = analisar_arquivo(uploaded_file, modo, aba)
            else:
                st.session_state['dados_consolidados'] = analisar_arquivo(uploaded_file, modo)
            st.session_state['dados_filtrados'] = st.session_state['dados_consolidados']
            st.success("Arquivo analisado com sucesso!")
        except Exception as e:
            st.error(str(e))
            return

    if not st.session_state['dados_consolidados'].empty:
        df = st.session_state['dados_filtrados'] if not st.session_state['dados_filtrados'].empty else st.session_state['dados_consolidados']
        main_column = "Potência" if modo == "ENERGIA" else "Temperatura"
        columns = atualizar_menu_tipo_valor(modo, df)

        # Filtros
        with st.expander("Configurações de Filtros e Escala"):
            filtro_hora = st.text_input("Filtro de Hora (HH:MM)", "")
            if st.button("Aplicar Filtro de Hora", key="filtro_hora"):
                if filtro_hora:
                    try:
                        hora, minuto = map(int, filtro_hora.split(":"))
                        if hora > 23 or minuto > 59:
                            st.error("Hora inválida.")
                        else:
                            mask = (df["DataHora"].dt.hour == hora) & (df["DataHora"].dt.minute == minuto)
                            st.session_state['dados_filtrados'] = df[mask]
                            st.session_state['filtro_ativo'] = "hora"
                            st.session_state['pontos_filtrados'] = []
                            for _, row in st.session_state['dados_filtrados'].iterrows():
                                x = row["DataHora"]
                                y = row[main_column]
                                st.session_state['pontos_filtrados'].append((None, x, y, main_column))
                                if modo == "DATALOGGER" and "Umidade" in row:
                                    y_umidade = row["Umidade"]
                                    if not pd.isna(y_umidade):
                                        st.session_state['pontos_filtrados'].append((None, x, y_umidade, "Umidade"))
                                if modo in ["SITRAD", "ENERGIA"]:
                                    for col in [c for c in df.columns if c.startswith("Dados_Extra") or c.startswith("Potência_Trafo")]:
                                        y_extra = row[col]
                                        if not pd.isna(y_extra):
                                            st.session_state['pontos_filtrados'].append((None, x, y_extra, col))
                            st.success("Filtro de hora aplicado.")
                    except:
                        st.error("Formato inválido. Use HH:MM.")

            filtro_valor_tipo = st.selectbox("Tipo de Valor para Filtro", columns)
            filtro_valor_min = st.number_input("Valor Mínimo", value=0.0, step=0.1)
            filtro_valor_max = st.number_input("Valor Máximo", value=100.0, step=0.1)
            if st.button("Aplicar Filtro de Valor", key="filtro_valor"):
                if filtro_valor_min > filtro_valor_max:
                    st.error("Mínimo deve ser menor que máximo.")
                else:
                    mask = (df[filtro_valor_tipo] >= filtro_valor_min) & (df[filtro_valor_tipo] <= filtro_valor_max) & (~df[filtro_valor_tipo].isna())
                    st.session_state['dados_filtrados'] = df[mask]
                    st.session_state['filtro_ativo'] = "valor"
                    st.session_state['filtro_valor_min'] = filtro_valor_min
                    st.session_state['filtro_valor_max'] = filtro_valor_max
                    st.session_state['filtro_valor_tipo'] = filtro_valor_tipo
                    st.session_state['pontos_filtrados'] = [(None, row["DataHora"], row[filtro_valor_tipo], filtro_valor_tipo) for _, row in st.session_state['dados_filtrados'].iterrows()]
                    st.success("Filtro de valor aplicado.")

            escala_y_min = st.number_input("Escala Y Mínima", value=0.0, step=0.1)
            escala_y_max = st.number_input("Escala Y Máxima", value=100.0, step=0.1)
            if st.button("Aplicar Escala Y", key="escala_y"):
                if escala_y_min >= escala_y_max:
                    st.error("Mínimo deve ser menor que máximo.")
                else:
                    st.session_state['escala_y_min'] = escala_y_min
                    st.session_state['escala_y_max'] = escala_y_max
                    st.success("Escala Y aplicada.")

        # Marcar pontos manualmente
        with st.expander("Marcar Ponto Manualmente"):
            data_ponto = st.text_input("Data/Hora do Ponto (YYYY/MM/DD HH:MM:SS)")
            valor_ponto = st.number_input("Valor do Ponto", value=0.0, step=0.1)
            tipo_ponto = st.selectbox("Tipo do Ponto", columns)
            if st.button("Marcar Ponto"):
                try:
                    x = pd.to_datetime(data_ponto)
                    st.session_state['pontos_marcados'].append((None, None, x, valor_ponto, tipo_ponto))
                    st.success("Ponto marcado!")
                except:
                    st.error("Formato de data inválido.")

        # Gráfico
        title = "Gráfico de Dados" if st.session_state['dados_filtrados'].empty else f"Gráfico da Faixa: {df['DataHora'].min().strftime('%Y/%m/%d %H:%M:%S')} a {df['DataHora'].max().strftime('%Y/%m/%d %H:%M:%S')}"
        fig = gerar_grafico(df, modo, st.session_state['filtro_ativo'], st.session_state['filtro_valor_min'], st.session_state['filtro_valor_max'], st.session_state['filtro_valor_tipo'], st.session_state['escala_y_min'], st.session_state['escala_y_max'], title, st.session_state['pontos_marcados'], st.session_state['pontos_filtrados'])
        st.plotly_chart(fig, use_container_width=True)

        # Estatísticas
        mostrar_estatisticas(df, modo, st.session_state['pontos_filtrados'], st.session_state['pontos_marcados'])

        # Múltiplas Faixas
        with st.expander("Gerenciar Múltiplas Faixas"):
            if st.button("Adicionar Faixa"):
                st.session_state['faixas'].append({"inicio": "", "fim": "", "nome": f"faixa_{len(st.session_state['faixas'])+1}"})
            for i, faixa in enumerate(st.session_state['faixas']):
                col1, col2, col3, col4 = st.columns([3, 3, 3, 1])
                with col1:
                    faixa["inicio"] = st.text_input(f"Início Faixa {i+1}", faixa["inicio"], key=f"inicio_{i}")
                with col2:
                    faixa["fim"] = st.text_input(f"Fim Faixa {i+1}", faixa["fim"], key=f"fim_{i}")
                with col3:
                    faixa["nome"] = st.text_input(f"Nome Faixa {i+1}", faixa["nome"], key=f"nome_{i}")
                with col4:
                    if st.button("Remover", key=f"remover_{i}"):
                        st.session_state['faixas'].pop(i)
                        st.rerun()
            if st.button("Exportar Faixas"):
                for faixa in st.session_state['faixas']:
                    try:
                        inicio = pd.to_datetime(faixa["inicio"], format="%Y/%m/%d %H:%M:%S")
                        fim = pd.to_datetime(faixa["fim"], format="%Y/%m/%d %H:%M:%S")
                        filtrado = df[(df["DataHora"] >= inicio) & (df["DataHora"] <= fim)]
                        if filtrado.empty:
                            st.warning(f"Nenhum dado para a faixa {faixa['nome']}.")
                            continue
                        buf = BytesIO()
                        filtrado.to_excel(buf, index=False)
                        st.download_button(f"Baixar {faixa['nome']}.xlsx", buf.getvalue(), f"{faixa['nome']}.xlsx")
                    except:
                        st.error(f"Erro na faixa {faixa['nome']}: Formato de data inválido.")

        # Exportar Consolidado
        buf = BytesIO()
        df_export = df.copy()
        if "DataHora_str" in df_export.columns:
            df_export = df_export.drop(columns=["DataHora"]).rename(columns={"DataHora_str": "DataHora"})
        else:
            df_export["DataHora"] = df_export["DataHora"].dt.strftime("%Y/%m/%d %H:%M:%S")
        df_export.to_excel(buf, index=False)
        st.download_button("Exportar Excel Consolidado", buf.getvalue(), "dados_consolidados.xlsx")

        # Exportar Gráfico
        buf = BytesIO()
        fig.write_image(buf, format="png")
        st.download_button("Exportar Gráfico PNG", buf.getvalue(), "grafico.png")

if __name__ == "__main__":
    main()