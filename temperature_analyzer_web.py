import streamlit as st
import pandas as pd
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

# CSS avançado para um visual mais próximo do CustomTkinter
st.markdown("""
<style>
    .main {
        background-color: #f5f7fa;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }
    .stButton>button {
        border-radius: 8px;
        font-size: 16px;
        font-family: Roboto, sans-serif;
        padding: 10px 20px;
        margin: 5px;
        transition: all 0.3s ease;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.3);
    }
    .stButton>button.sitrad {
        background: linear-gradient(to right, #00cc00, #00ff00);
        color: white;
    }
    .stButton>button.datalogger {
        background: linear-gradient(to right, #ff9900, #ffcc00);
        color: white;
    }
    .stButton>button.energia {
        background: linear-gradient(to right, #cccc00, #ffff00);
        color: white;
    }
    .stButton>button.reset {
        background: linear-gradient(to right, #cc0000, #ff3333);
        color: white;
    }
    .stButton>button.action {
        background: linear-gradient(to right, #007bff, #00ccff);
        color: white;
    }
    .stTextInput>div>input, .stNumberInput>div>input, .stSelectbox>div>select {
        border-radius: 5px;
        border: 1px solid #ced4da;
        padding: 8px;
        font-family: Roboto, sans-serif;
    }
    .section-container {
        background-color: white;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
    }
    h1, h2, h3 {
        color: #2c3e50;
        font-family: Roboto, sans-serif;
        margin-bottom: 10px;
    }
    .stPlotlyChart {
        border: 1px solid #e9ecef;
        border-radius: 8px;
        padding: 10px;
        background-color: white;
    }
    hr {
        border: 0;
        height: 1px;
        background: #e9ecef;
        margin: 15px 0;
    }
</style>
""", unsafe_allow_html=True)

# Funções de lógica
def get_button_colors(modo):
    return "sitrad" if modo == "SITRAD" else "datalogger" if modo == "DATALOGGER" else "energia"

def get_marker_color(filtro=False):
    return "orange" if filtro else "red"

def atualizar_menu_tipo_valor(modo, df):
    if modo == "DATALOGGER":
        return ["Temperatura", "Umidade"]
    elif modo == "SITRAD":
        return ["Temperatura"] + [col for col in df.columns if col.startswith("Dados_Extra") and pd.api.types.is_numeric_dtype(df[col])]
    else:  # ENERGIA
        return ["Potência"] + [col for col in df.columns if col.startswith("Potência_Trafo") and pd.api.types.is_numeric_dtype(df[col])]

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
        extra_columns = [c for c in df.columns if (c.startswith("Dados_Extra") or c.startswith("Potência_Trafo")) and pd.api.types.is_numeric_dtype(df[c])]
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
    fig.update_layout(
        title=dict(text=title, x=0.5, xanchor="center", font=dict(size=20)),
        xaxis_title="Data e Hora",
        yaxis_title=f"{main_column} ({unit})",
        legend_orientation="h",
        legend=dict(x=0.5, xanchor="center", y=-0.1),
        height=600,
        margin=dict(l=50, r=50, t=80, b=80),
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="Roboto, sans-serif", size=14)
    )
    if escala_y_min is not None and escala_y_max is not None:
        fig.update_yaxes(range=[escala_y_min, escala_y_max], secondary_y=False)
    return fig

def mostrar_estatisticas(df, modo, pontos_filtrados=None, pontos_marcados=None):
    with st.container():
        st.markdown("### Estatísticas")
        main_column = "Potência" if modo == "ENERGIA" else "Temperatura"
        total = len(df)
        validos = len(df.dropna(subset=[main_column]))
        invalidos = total - validos
        st.write(f"**Total:** {total} | **Válidos:** {validos} | **Inválidos:** {invalidos}")
        unit = "W" if modo == "ENERGIA" else "°C"
        if not df[main_column].isna().all() and pd.api.types.is_numeric_dtype(df[main_column]):
            max_val = df[main_column].max()
            min_val = df[main_column].min()
            mean_val = df[main_column].mean()
            st.write(f"**{main_column}** - Máxima: {max_val:.2f}{unit} | Mínima: {min_val:.2f}{unit} | Média: {mean_val:.2f}{unit}")
        else:
            st.write(f"**{main_column}** - Dados insuficientes para estatísticas")
        if modo == "DATALOGGER" and "Umidade" in df.columns and not df["Umidade"].isna().all() and pd.api.types.is_numeric_dtype(df["Umidade"]):
            umidade_max = df["Umidade"].max()
            umidade_min = df["Umidade"].min()
            umidade_media = df["Umidade"].mean()
            st.write(f"**Umidade** - Máxima: {umidade_max:.2f}% | Mínima: {umidade_min:.2f}% | Média: {umidade_media:.2f}%")
        elif modo == "DATALOGGER":
            st.write("**Umidade** - Dados insuficientes para estatísticas")
        if modo == "SITRAD" or modo == "ENERGIA":
            extra_columns = [c for c in df.columns if (c.startswith("Dados_Extra") or c.startswith("Potência_Trafo")) and pd.api.types.is_numeric_dtype(df[c])]
            for col in extra_columns:
                if not df[col].isna().all():
                    max_val = df[col].max()
                    min_val = df[col].min()
                    mean_val = df[col].mean()
                    unit_extra = "W" if col.startswith("Potência_Trafo") else "°C"
                    st.write(f"**{col}** - Máxima: {max_val:.2f}{unit_extra} | Mínima: {min_val:.2f}{unit_extra} | Média: {mean_val:.2f}{unit_extra}")
                else:
                    st.write(f"**{col}** - Dados insuficientes para estatísticas")
        inicio = df["DataHora"].min().strftime("%Y/%m/%d %H:%M:%S") if not pd.isnull(df["DataHora"].min()) else "N/A"
        fim = df["DataHora"].max().strftime("%Y/%m/%d %H:%M:%S") if not pd.isnull(df["DataHora"].max()) else "N/A"
        st.write(f"**Intervalo:** {inicio} até {fim}")
        if pontos_filtrados:
            st.write("**Pontos Filtrados:**")
            for _, x, y, tipo in pontos_filtrados:
                unit_p = "%" if tipo == "Umidade" else ("W" if "Potência" in tipo else "°C")
                st.write(f"{x.strftime('%Y/%m/%d %H:%M:%S')} - {y:.1f}{unit_p} ({tipo})")
        if pontos_marcados:
            st.write("**Pontos Marcados Manualmente:**")
            for _, _, x, y, tipo in pontos_marcados:
                unit_p = "%" if tipo == "Umidade" else ("W" if "Potência" in tipo else "°C")
                st.write(f"{x.strftime('%Y/%m/%d %H:%M:%S')} - {y:.1f}{unit_p} ({tipo})")

def main():
    st.markdown('<div class="main">', unsafe_allow_html=True)
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

    # Layout em colunas para imitar Tkinter
    col_left, col_right = st.columns([1, 3])

    with col_left:
        # Seleção de modo
        with st.container():
            st.markdown("### Modo de Operação")
            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button("SITRAD", key="modo_sitrad", help="Modo SITRAD"):
                    st.session_state['modo'] = "SITRAD"
            with col2:
                if st.button("DATALOGGER", key="modo_datalogger", help="Modo DATALOGGER"):
                    st.session_state['modo'] = "DATALOGGER"
            with col3:
                if st.button("ENERGIA", key="modo_energia", help="Modo ENERGIA"):
                    st.session_state['modo'] = "ENERGIA"
            if st.button("Resetar Tudo", key="reset", help="Limpar todos os dados e filtros"):
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
        with st.container():
            st.markdown("### Carregar Arquivo")
            modo = st.session_state.get('modo', 'SITRAD')
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

        # Filtros
        with st.container():
            st.markdown("### Filtros e Escala")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                st.write("**Filtro de Hora**")
                filtro_hora = st.text_input("Hora (HH:MM)", "")
                if st.button("Aplicar Filtro de Hora", key="filtro_hora", type="primary"):
                    if filtro_hora:
                        try:
                            hora, minuto = map(int, filtro_hora.split(":"))
                            if hora > 23 or minuto > 59:
                                st.error("Hora inválida.")
                            else:
                                df = st.session_state['dados_filtrados'] if not st.session_state['dados_filtrados'].empty else st.session_state['dados_consolidados']
                                mask = (df["DataHora"].dt.hour == hora) & (df["DataHora"].dt.minute == minuto)
                                st.session_state['dados_filtrados'] = df[mask]
                                st.session_state['filtro_ativo'] = "hora"
                                st.session_state['pontos_filtrados'] = []
                                main_column = "Potência" if modo == "ENERGIA" else "Temperatura"
                                for _, row in st.session_state['dados_filtrados'].iterrows():
                                    x = row["DataHora"]
                                    y = row[main_column]
                                    st.session_state['pontos_filtrados'].append((None, x, y, main_column))
                                    if modo == "DATALOGGER" and "Umidade" in row:
                                        y_umidade = row["Umidade"]
                                        if not pd.isna(y_umidade):
                                            st.session_state['pontos_filtrados'].append((None, x, y_umidade, "Umidade"))
                                    if modo in ["SITRAD", "ENERGIA"]:
                                        for col in [c for c in df.columns if (c.startswith("Dados_Extra") or c.startswith("Potência_Trafo")) and pd.api.types.is_numeric_dtype(df[c])]:
                                            y_extra = row[col]
                                            if not pd.isna(y_extra):
                                                st.session_state['pontos_filtrados'].append((None, x, y_extra, col))
                                st.success("Filtro de hora aplicado.")
                        except:
                            st.error("Formato inválido. Use HH:MM.")
            with col_f2:
                st.write("**Filtro de Valor**")
                df = st.session_state['dados_filtrados'] if not st.session_state['dados_filtrados'].empty else st.session_state['dados_consolidados']
                columns = atualizar_menu_tipo_valor(modo, df)
                filtro_valor_tipo = st.selectbox("Tipo de Valor", columns)
                filtro_valor_min = st.number_input("Valor Mínimo", value=0.0, step=0.1)
                filtro_valor_max = st.number_input("Valor Máximo", value=100.0, step=0.1)
                if st.button("Aplicar Filtro de Valor", key="filtro_valor", type="primary"):
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
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                st.write("**Escala Y**")
                escala_y_min = st.number_input("Mínima", value=0.0, step=0.1)
            with col_s2:
                escala_y_max = st.number_input("Máxima", value=100.0, step=0.1)
            if st.button("Aplicar Escala Y", key="escala_y", type="primary"):
                if escala_y_min >= escala_y_max:
                    st.error("Mínimo deve ser menor que máximo.")
                else:
                    st.session_state['escala_y_min'] = escala_y_min
                    st.session_state['escala_y_max'] = escala_y_max
                    st.success("Escala Y aplicada.")

        # Marcar pontos manualmente
        with st.container():
            st.markdown("### Marcar Ponto Manualmente")
            col_p1, col_p2, col_p3 = st.columns(3)
            with col_p1:
                data_ponto = st.text_input("Data/Hora (YYYY/MM/DD HH:MM:SS)")
            with col_p2:
                valor_ponto = st.number_input("Valor", value=0.0, step=0.1)
            with col_p3:
                tipo_ponto = st.selectbox("Tipo", columns)
            if st.button("Marcar Ponto", key="marcar_ponto", type="primary"):
                try:
                    x = pd.to_datetime(data_ponto)
                    st.session_state['pontos_marcados'].append((None, None, x, valor_ponto, tipo_ponto))
                    st.success("Ponto marcado!")
                except:
                    st.error("Formato de data inválido.")

    with col_right:
        if not st.session_state['dados_consolidados'].empty:
            df = st.session_state['dados_filtrados'] if not st.session_state['dados_filtrados'].empty else st.session_state['dados_consolidados']
            # Gráfico
            with st.container():
                st.markdown("### Gráfico")
                title = "Gráfico de Dados" if st.session_state['dados_filtrados'].empty else f"Gráfico da Faixa: {df['DataHora'].min().strftime('%Y/%m/%d %H:%M:%S')} a {df['DataHora'].max().strftime('%Y/%m/%d %H:%M:%S')}"
                fig = gerar_grafico(df, modo, st.session_state['filtro_ativo'], st.session_state['filtro_valor_min'], st.session_state['filtro_valor_max'], st.session_state['filtro_valor_tipo'], st.session_state['escala_y_min'], st.session_state['escala_y_max'], title, st.session_state['pontos_marcados'], st.session_state['pontos_filtrados'])
                st.plotly_chart(fig, use_container_width=True)

            # Verificar exportação de gráfico
            try:
                buf = BytesIO()
                fig.write_image(buf, format="png")
                st.session_state['grafico_exportavel'] = True
            except Exception as e:
                st.session_state['grafico_exportavel'] = False
                st.warning("Exportação de gráfico PNG não disponível devido a limitações do ambiente.")

            # Estatísticas
            mostrar_estatisticas(df, modo, st.session_state['pontos_filtrados'], st.session_state['pontos_marcados'])

    # Múltiplas Faixas e Exportações
    with st.container():
        st.markdown("### Gerenciar Múltiplas Faixas")
        if st.button("Adicionar Faixa", type="primary"):
            st.session_state['faixas'].append({"inicio": "", "fim": "", "nome": f"faixa_{len(st.session_state['faixas'])+1}"})
        for i, faixa in enumerate(st.session_state['faixas']):
            col_f1, col_f2, col_f3, col_f4 = st.columns([3, 3, 3, 1])
            with col_f1:
                faixa["inicio"] = st.text_input(f"Início Faixa {i+1}", faixa["inicio"], key=f"inicio_{i}")
            with col_f2:
                faixa["fim"] = st.text_input(f"Fim Faixa {i+1}", faixa["fim"], key=f"fim_{i}")
            with col_f3:
                faixa["nome"] = st.text_input(f"Nome Faixa {i+1}", faixa["nome"], key=f"nome_{i}")
            with col_f4:
                if st.button("Remover", key=f"remover_{i}", type="primary"):
                    st.session_state['faixas'].pop(i)
                    st.rerun()
        if st.button("Exportar Faixas", type="primary"):
            for faixa in st.session_state['faixas']:
                try:
                    inicio = pd.to_datetime(faixa["inicio"], format="%Y/%m/%d %H:%M:%S")
                    fim = pd.to_datetime(faixa["fim"], format="%Y/%m/%d %H:%M:%S")
                    filtrado = df[(df["DataHora"] >= inicio) & (df["DataHora"] <= fim)]
                    if filtrado.empty:
                        st.warning(f"Nenhum dado para a faixa {faixa['nome']}.")
                        continue
                    buf = BytesIO()
                    df_export = filtrado.copy()
                    if "DataHora_str" in df_export.columns:
                        df_export = df_export.drop(columns=["DataHora"]).rename(columns={"DataHora_str": "DataHora"})
                    else:
                        df_export["DataHora"] = df_export["DataHora"].dt.strftime("%Y/%m/%d %H:%M:%S")
                    df_export.to_excel(buf, index=False)
                    st.download_button(f"Baixar {faixa['nome']}.xlsx", buf.getvalue(), f"{faixa['nome']}.xlsx", type="primary")
                except:
                    st.error(f"Erro na faixa {faixa['nome']}: Formato de data inválido.")

    with st.container():
        st.markdown("### Exportações")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            buf = BytesIO()
            df_export = df.copy() if not st.session_state['dados_consolidados'].empty else pd.DataFrame()
            if not df_export.empty:
                if "DataHora_str" in df_export.columns:
                    df_export = df_export.drop(columns=["DataHora"]).rename(columns={"DataHora_str": "DataHora"})
                else:
                    df_export["DataHora"] = df_export["DataHora"].dt.strftime("%Y/%m/%d %H:%M:%S")
                df_export.to_excel(buf, index=False)
                st.download_button("Exportar Excel Consolidado", buf.getvalue(), "dados_consolidados.xlsx", type="primary")
        with col_e2:
            if st.session_state.get('grafico_exportavel', False):
                buf = BytesIO()
                fig.write_image(buf, format="png")
                st.download_button("Exportar Gráfico PNG", buf.getvalue(), "grafico.png", type="primary")
            else:
                st.write("Exportação de PNG indisponível.")

    st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
