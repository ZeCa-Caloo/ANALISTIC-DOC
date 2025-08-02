import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
from io import BytesIO
import chardet
import pytz

st.set_page_config(page_title="Dashboard Inteligente", layout="wide")

st.title("Dashboard Inteligente - HTML, XLSX e CSV")

# =============================
# Funções auxiliares
# =============================
def ler_arquivo(uploaded_file):
    ext = uploaded_file.name.split('.')[-1].lower()
    if ext == "xlsx":
        return pd.read_excel(uploaded_file)
    elif ext == "html":
        try:
            tables = pd.read_html(uploaded_file)
            return tables[0]
        except Exception as e:
            st.error(f"Erro ao ler HTML: {e}")
            return None
    elif ext == "csv":
        try:
            raw_data = uploaded_file.read()
            result = chardet.detect(raw_data)
            encoding_detectado = result["encoding"]
            st.info(f"Arquivo CSV detectado com encoding: **{encoding_detectado}**")

            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=None, engine="python", encoding=encoding_detectado)
        except Exception as e:
            st.error(f"Erro ao ler CSV: {e}")
            return None
    else:
        return None

def to_excel(df):
    df_copy = df.copy()
    for col in df_copy.columns:
        if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
            if getattr(df_copy[col].dt, "tz", None) is not None:
                df_copy[col] = df_copy[col].dt.tz_localize(None)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_copy.to_excel(writer, index=False)
    return output.getvalue()

def to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

def to_json(df):
    return df.to_json(orient="records", force_ascii=False).encode('utf-8')

def gerar_insights(df):
    insights = []
    insights.append(f"O conjunto de dados possui {df.shape[0]} linhas e {df.shape[1]} colunas.")
    insights.append(f"As colunas disponíveis são: {', '.join(df.columns)}.")
    num_cols = df.select_dtypes(include='number').columns
    if len(num_cols) > 0:
        insights.append(f"Foram encontradas {len(num_cols)} colunas numéricas.")
        for col in num_cols:
            media = df[col].mean()
            insights.append(f"A média da coluna '{col}' é {media:.2f}.")
    else:
        insights.append("Não há colunas numéricas para calcular estatísticas básicas.")
    cat_cols = df.select_dtypes(exclude='number').columns
    if len(cat_cols) > 0:
        insights.append(f"Foram encontradas {len(cat_cols)} colunas categóricas.")
        for col in cat_cols:
            valor_mais_freq = df[col].mode()[0] if not df[col].mode().empty else "Nenhum"
            insights.append(f"Na coluna '{col}', o valor mais frequente é '{valor_mais_freq}'.")
    return "\n".join(insights)

def converter_datas_para_timestamp(df):
    df_convertido = df.copy()
    for col in df_convertido.columns:
        if pd.api.types.is_datetime64_any_dtype(df_convertido[col]):
            df_convertido[col] = df_convertido[col].view('int64') / 1e9
    return df_convertido

def detectar_colunas_datetime(df):
    """Detecta colunas datetime e converte para UTC-3."""
    fuso = pytz.timezone("America/Sao_Paulo")
    for col in df.columns:
        if df[col].dtype == object:
            try:
                temp = pd.to_datetime(df[col], errors='raise', utc=True)
                temp = temp.dt.tz_convert(fuso)
                df[col] = temp
            except Exception:
                pass
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            if getattr(df[col].dt, "tz", None) is not None:
                df[col] = df[col].dt.tz_convert(fuso)
    return df

def formatar_datas_para_exibicao(df):
    """Cria uma cópia para exibir datas formatadas dd/mm/aaaa hh:mm:ss"""
    df_exibir = df.copy()
    for col in df_exibir.columns:
        if pd.api.types.is_datetime64_any_dtype(df_exibir[col]):
            df_exibir[col] = df_exibir[col].dt.strftime("%d/%m/%Y %H:%M:%S")
    return df_exibir

# =============================
# Upload múltiplo
# =============================
uploaded_files = st.file_uploader(
    "Selecione arquivos HTML, XLSX ou CSV (múltiplos arquivos permitidos)",
    type=["html", "xlsx", "csv"], accept_multiple_files=True)

if uploaded_files:
    dfs = []
    for file in uploaded_files:
        df_temp = ler_arquivo(file)
        if df_temp is not None:
            dfs.append(df_temp)

    if dfs:
        df = pd.concat(dfs, ignore_index=True)

        aba1, aba2, aba3, aba4, aba5, aba6, aba7 = st.tabs([
            "📄 Dados",
            "🔍 Filtros",
            "📊 Gráficos",
            "📈 Estatísticas",
            "📊 Dashboard Automático",
            "🤖 Insights Automáticos",
            "⬇️ Download"
        ])

        with aba1:
            st.subheader("Visualização dos Dados Combinados")
            df = detectar_colunas_datetime(df)
            st.dataframe(formatar_datas_para_exibicao(df))

        with aba2:
            st.subheader("Filtrar Dados")
            colunas = st.multiselect("Selecione colunas para filtrar", df.columns)
            df_filtrado = df.copy()
            for col in colunas:
                valores = df[col].dropna().unique().tolist()
                selecao = st.multiselect(f"Valores para {col}", valores)
                if selecao:
                    df_filtrado = df_filtrado[df_filtrado[col].isin(selecao)]
            st.dataframe(formatar_datas_para_exibicao(df_filtrado))

        with aba3:
            st.subheader("Visualização de Gráficos")
            if not df_filtrado.empty:
                colunas_num = df_filtrado.select_dtypes(include="number").columns
                colunas_cat = df_filtrado.select_dtypes(exclude="number").columns
                tipo_grafico = st.selectbox(
                    "Selecione o tipo de gráfico",
                    ["Histograma", "Barras", "Linha", "Pizza"]
                )
                modo_grafico = st.radio(
                    "Selecione a biblioteca para visualização",
                    ["Plotly (Interativo)", "Matplotlib (Estático)"]
                )
                if tipo_grafico in ["Histograma", "Linha"] and len(colunas_num) > 0:
                    colunas_escolhidas = st.multiselect("Selecione colunas numéricas", colunas_num)
                elif tipo_grafico in ["Barras", "Pizza"] and len(colunas_cat) > 0:
                    colunas_escolhidas = st.multiselect("Selecione colunas categóricas", colunas_cat)
                else:
                    colunas_escolhidas = []
                if colunas_escolhidas:
                    if modo_grafico == "Plotly (Interativo)":
                        if tipo_grafico == "Histograma":
                            for col in colunas_escolhidas:
                                fig = px.histogram(df_filtrado, x=col, nbins=10, title=f"Histograma - {col}")
                                st.plotly_chart(fig, use_container_width=True)
                        elif tipo_grafico == "Barras":
                            for col in colunas_escolhidas:
                                contagem = df_filtrado[col].value_counts().reset_index()
                                contagem.columns = [col, "Contagem"]
                                fig = px.bar(contagem, x=col, y="Contagem", title=f"Barras - {col}")
                                st.plotly_chart(fig, use_container_width=True)
                        elif tipo_grafico == "Linha":
                            fig = px.line(df_filtrado[colunas_escolhidas])
                            fig.update_layout(title="Gráfico de Linha (múltiplas colunas)")
                            st.plotly_chart(fig, use_container_width=True)
                        elif tipo_grafico == "Pizza":
                            for col in colunas_escolhidas:
                                contagem = df_filtrado[col].value_counts().reset_index()
                                contagem.columns = [col, "Contagem"]
                                fig = px.pie(contagem, names=col, values="Contagem", title=f"Pizza - {col}")
                                st.plotly_chart(fig, use_container_width=True)
                    else:
                        for col in colunas_escolhidas:
                            fig, ax = plt.subplots()
                            if tipo_grafico == "Histograma":
                                df_filtrado[col].plot(kind="hist", bins=10, rwidth=0.8, ax=ax)
                                ax.set_title(f"Histograma - {col}")
                            elif tipo_grafico == "Barras":
                                df_filtrado[col].value_counts().plot(kind="bar", ax=ax)
                                ax.set_title(f"Barras - {col}")
                            elif tipo_grafico == "Linha":
                                df_filtrado[col].plot(kind="line", ax=ax)
                                ax.set_title(f"Linha - {col}")
                            elif tipo_grafico == "Pizza":
                                df_filtrado[col].value_counts().plot(kind="pie", autopct='%1.1f%%', ax=ax)
                                ax.set_ylabel('')
                                ax.set_title(f"Pizza - {col}")
                            st.pyplot(fig)
                else:
                    st.info("Selecione pelo menos uma coluna para gerar o gráfico.")

        with aba4:
            st.subheader("Estatísticas Descritivas")
            if not df_filtrado.empty:
                df_filtrado = detectar_colunas_datetime(df_filtrado)
                st.write("**Estatísticas das colunas numéricas:**")
                st.dataframe(df_filtrado.describe())
                st.write("**Contagem de valores por coluna:**")
                st.dataframe(df_filtrado.count())
                usar_timestamp = st.checkbox(
                    "Converter colunas de datas (datetime) em valores numéricos (timestamp) para incluir na correlação"
                )
                df_corr = df_filtrado.copy()
                if usar_timestamp:
                    df_corr = converter_datas_para_timestamp(df_corr)
                colunas_numericas = df_corr.select_dtypes(include="number")
                if colunas_numericas.shape[1] > 1:
                    st.write("**Matriz de Correlação:**")
                    st.dataframe(colunas_numericas.corr())
                else:
                    st.info("Não há colunas numéricas suficientes para calcular a correlação.")
            else:
                st.info("Nenhum dado disponível para gerar estatísticas.")

        with aba5:
            st.subheader("Dashboard Automático")
            if not df_filtrado.empty:
                colunas_num = df_filtrado.select_dtypes(include="number").columns
                colunas_cat = df_filtrado.select_dtypes(exclude="number").columns
                if len(colunas_num) > 0:
                    col = colunas_num[0]
                    st.write(f"Histograma automático para {col}")
                    st.plotly_chart(px.histogram(df_filtrado, x=col), use_container_width=True)
                if len(colunas_cat) > 0:
                    col = colunas_cat[0]
                    st.write(f"Barras automáticas para {col}")
                    contagem = df_filtrado[col].value_counts().reset_index()
                    contagem.columns = [col, "Contagem"]
                    st.plotly_chart(px.bar(contagem, x=col, y="Contagem"), use_container_width=True)
                if len(colunas_num) >= 2:
                    st.write("Gráfico de linha automático para duas primeiras colunas numéricas")
                    st.plotly_chart(px.line(df_filtrado[colunas_num[:2]]), use_container_width=True)
            else:
                st.info("Carregue dados e aplique filtros para gerar gráficos automáticos.")

        with aba6:
            st.subheader("Insights Automáticos")
            if not df_filtrado.empty:
                insights = gerar_insights(df_filtrado)
                st.text_area("Resumo gerado automaticamente:", insights, height=300)
            else:
                st.info("Nenhum dado para gerar insights.")

        with aba7:
            st.subheader("Exportar Dados Filtrados")
            excel_bytes = to_excel(df_filtrado)
            csv_bytes = to_csv(df_filtrado)
            json_bytes = to_json(df_filtrado)
            st.download_button(
                "Baixar em Excel",
                data=excel_bytes,
                file_name="dados_filtrados.xlsx",
                mime="application/vnd.ms-excel")
            st.download_button(
                "Baixar em CSV",
                data=csv_bytes,
                file_name="dados_filtrados.csv",
                mime="text/csv")
            st.download_button(
                "Baixar em JSON",
                data=json_bytes,
                file_name="dados_filtrados.json",
                mime="application/json")
    else:
        st.warning("Nenhum dado válido encontrado.")
