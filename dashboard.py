import streamlit as st
import pandas as pd
import plotly.express as px
from docx import Document
from datetime import timedelta
from fpdf import FPDF
import matplotlib.pyplot as plt
import tempfile
import os

st.set_page_config(page_title="📊 Dashboard Interativo com Datas", layout="wide")
st.title("📊 Dashboard de Execução Diária")

upload_file = st.file_uploader(
    "📥 Envie seu arquivo (CSV, Excel, TXT, DOCX):",
    type=["csv", "xlsx", "xlsm", "txt", "docx"]
)

def salvar_grafico_como_imagem(fig, filename):
    fig.write_image(filename)

def exportar_pdf(graficos_salvos, nome_arquivo="relatorio.pdf"):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    for path in graficos_salvos:
        pdf.add_page()
        pdf.image(path, x=10, y=20, w=180)
    pdf.output(nome_arquivo)
    return nome_arquivo

if upload_file is not None:
    file_type = upload_file.name.split(".")[-1].lower()
    try:
        if file_type in ["csv", "txt"]:
            df = pd.read_csv(upload_file, encoding="utf-8", sep=None, engine="python")
        elif file_type in ["xlsx", "xlsm"]:
            aba = st.text_input("📄 Nome da aba do Excel:", value="Sheet1")
            df = pd.read_excel(upload_file, sheet_name=aba)
        elif file_type == "docx":
            doc = Document(upload_file)
            text = [p.text for p in doc.paragraphs if p.text.strip()]
            df = pd.DataFrame(text, columns=["Conteúdo"])
        else:
            st.error("Tipo de arquivo não suportado.")
            st.stop()

        st.subheader("👀 Prévia dos Dados")
        st.dataframe(df.head(), use_container_width=True)

        if not df.empty and df.select_dtypes(include='object').shape[1] > 0:
            st.subheader("📅 Configurar Coluna de Data")
            date_column = st.selectbox("Selecione a coluna que representa datas:", df.columns)

            try:
                df[date_column] = pd.to_datetime(df[date_column], dayfirst=True, errors="coerce")
                df = df.dropna(subset=[date_column])
                df["Data Formatada"] = df[date_column].dt.strftime("%d/%m/%Y")
                st.success("Coluna de data convertida com sucesso!")
            except Exception as e:
                st.error(f"Erro ao converter coluna para data: {e}")
                st.stop()

            min_date = df[date_column].min()
            max_date = df[date_column].max()

            if pd.isna(min_date) or pd.isna(max_date):
                st.error("As datas são inválidas ou não foram corretamente reconhecidas.")
                st.stop()

            start_date, end_date = st.date_input(
                "🗓️ Filtrar por período:",
                value=(min_date.date(), max_date.date()),
                min_value=min_date.date(),
                max_value=max_date.date()
            )

            # Botões de período
            col_button1, col_button2, col_button3 = st.columns([1, 1, 6])
            days_range = (end_date - start_date).days

            if col_button1.button("⬅️ Período Anterior"):
                start_date -= timedelta(days=days_range + 1)
                end_date -= timedelta(days=days_range + 1)
            if col_button2.button("➡️ Próximo Período"):
                start_date += timedelta(days=days_range + 1)
                end_date += timedelta(days=days_range + 1)

            mask = (df[date_column].dt.date >= start_date) & (df[date_column].dt.date <= end_date)
            filtered_df = df.loc[mask]

            st.markdown("### 📄 Dados Filtrados")
            st.dataframe(filtered_df, use_container_width=True)

            # Conversão segura
            for col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", ".").str.replace("%", ""), errors="coerce")
                except:
                    pass

            numeric_columns = df.select_dtypes(include='number').columns.tolist()
            category_columns = df.select_dtypes(include='object').columns.tolist()

            if filtered_df.empty:
                st.warning("Nenhum dado encontrado no intervalo selecionado.")
            else:
                st.markdown("---")
                st.subheader("📈 Gráficos Interativos")

                col1, col2 = st.columns(2)
                imagens_salvas = []

                with col1:
                    st.markdown("#### 📉 Gráfico de Linha")
                    if numeric_columns:
                        line_y = st.selectbox("Métrica para linha:", numeric_columns, key="line")
                        fig_line = px.line(filtered_df, x="Data Formatada", y=line_y, title=f"Linha - {line_y}")
                        fig_line.update_traces(
                            line=dict(color="#9D68B2"),
                            mode="lines+markers+text",
                            text=filtered_df[line_y].round(2),
                            textposition="top center"
                        )
                        st.plotly_chart(fig_line, use_container_width=True)

                        temp_file_line = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                        salvar_grafico_como_imagem(fig_line, temp_file_line.name)
                        imagens_salvas.append(temp_file_line.name)
                    else:
                        st.info("Nenhuma coluna numérica disponível.")

                with col2:
                    st.markdown("#### 📊 Gráfico de Barras")
                    if numeric_columns:
                        bar_y = st.selectbox("Métrica para barras:", numeric_columns, key="bar")
                        fig_bar = px.bar(
                            filtered_df,
                            x="Data Formatada",
                            y=bar_y,
                            title=f"Barras - {bar_y}",
                            text=filtered_df[bar_y].round(2)
                        )
                        fig_bar.update_traces(marker_color="#00163C", textposition="outside")
                        st.plotly_chart(fig_bar, use_container_width=True)

                        temp_file_bar = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                        salvar_grafico_como_imagem(fig_bar, temp_file_bar.name)
                        imagens_salvas.append(temp_file_bar.name)
                    else:
                        st.info("Nenhuma coluna numérica disponível.")

                st.markdown("#### 🥧 Gráfico de Pizza")
                if category_columns:
                    pie_col = st.selectbox("Coluna categórica para pizza:", category_columns)
                    pie_data = filtered_df[pie_col].value_counts().reset_index()
                    pie_data.columns = [pie_col, "Contagem"]
                    fig_pie = px.pie(pie_data, names=pie_col, values="Contagem", title=f"Distribuição de {pie_col}")
                    st.plotly_chart(fig_pie, use_container_width=True)

                    temp_file_pie = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    salvar_grafico_como_imagem(fig_pie, temp_file_pie.name)
                    imagens_salvas.append(temp_file_pie.name)
                else:
                    st.info("Nenhuma coluna categórica para gráfico de pizza.")

                st.markdown("---")
                if st.button("📄 Exportar PDF com Gráficos"):
                    output_path = os.path.join(tempfile.gettempdir(), "relatorio.pdf")
                    exportar_pdf(imagens_salvas, output_path)
                    with open(output_path, "rb") as f:
                        st.download_button("📥 Baixar PDF", f, file_name="relatorio.pdf")
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
else:
    st.info("Envie um arquivo para começar.")
