def main():
    st.title("Automação Trello")
    st.write("Faça upload do arquivo Excel para processar")

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx'])

    if uploaded_file is not None:
        if st.button("Processar Arquivo"):
            try:
                trello_data, faltas_atualizadas, colunas_exportadas = processar_planilha(uploaded_file)
                
                st.success("Arquivo processado com sucesso!")
                st.write("Colunas exportadas:", ", ".join(colunas_exportadas))

                # Criar buffer para os arquivos Excel
                trello_buffer = BytesIO()
                faltas_buffer = BytesIO()
                
                # Salvar os DataFrames nos buffers
                with pd.ExcelWriter(trello_buffer, engine='openpyxl') as writer:
                    trello_data.to_excel(writer, index=False)
                
                with pd.ExcelWriter(faltas_buffer, engine='openpyxl') as writer:
                    faltas_atualizadas.to_excel(writer, index=False)

                # Preparar os buffers para download
                trello_buffer.seek(0)
                faltas_buffer.seek(0)

                st.download_button(
                    label="Baixar arquivo Trello formatado",
                    data=trello_buffer,
                    file_name=f"Trello_Formatado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.download_button(
                    label="Baixar planilha atualizada",
                    data=faltas_buffer,
                    file_name=f"Faltas_Atualizadas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Erro ao processar arquivo: {str(e)}")
