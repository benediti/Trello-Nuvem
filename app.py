import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

def processar_planilha(uploaded_file):
    # Lê o arquivo Excel
    df = pd.read_excel(uploaded_file)
    
    # Padroniza nomes das colunas
    df.columns = [col.strip().upper() for col in df.columns]
    
    # Adiciona coluna de verificação se não existir
    if 'ID VERIFICACAO' not in df.columns:
        df['ID VERIFICACAO'] = ''
    
    # Processa apenas linhas não processadas
    registros = []
    data_atual = datetime.now().strftime('%Y-%m-%d')
    colunas_exportadas = set()

    for index, row in df[df['ID VERIFICACAO'] != 'PROCESSADO'].iterrows():
        descricao = (
            f"Matrícula: {row.get('MATRÍCULA', '')}\n"
            f"Localização: {row.get('LOCALIZAÇÃO', '')}\n"
            f"Dia: {row.get('DIA', '')}\n"
        )

        # Verifica batidas
        batidas = [row.get(col, '').strip() if pd.notna(row.get(col, '')) else '' 
                  for col in ['BATIDAS', 'ENTRADA 1', 'SAÍDA 1', 'ENTRADA 2', 'SAÍDA 2']]
        
        if all(not batida or batida == '00:00' for batida in batidas):
            registros.append({
                'list': 'SEM BATIDA',
                'Card Name': row.get('NOME', 'Sem Nome'),
                'desc': descricao,
                'checklist': 'Sem registros de batida',
                'Data': data_atual
            })
            colunas_exportadas.add('SEM BATIDA')

        # Verifica outros campos
        campos = {
            'ATRASO': 'ATRASO',
            'FALTA': 'FALTA',
            'BANCO DE HORAS': 'BANCO DE HORAS',
            'HORA EXTRA 50% (N.A.)': 'HORA EXTRA 50%',
            'HORA EXTRA 100% (N.A.)': 'HORA EXTRA 100%',
            'DSR DESCONTADO': 'DSR DESCONTADO',
            'ADICIONAL NOTURNO': 'ADICIONAL NOTURNO',
            'EXPEDIENTE': 'EXPEDIENTE'
        }

        for campo, lista in campos.items():
            valor = str(row.get(campo, '')).strip() if pd.notna(row.get(campo, '')) else ''
            if valor and valor != '00:00':
                registros.append({
                    'list': lista,
                    'Card Name': row.get('NOME', 'Sem Nome'),
                    'desc': descricao,
                    'checklist': valor,
                    'Data': data_atual
                })
                colunas_exportadas.add(lista)

        # Atualiza o status de processamento
        df.loc[index, 'ID VERIFICACAO'] = 'PROCESSADO'

    return pd.DataFrame(registros), df, sorted(list(colunas_exportadas))

def main():
    st.title("Automação Trello")
    st.write("Faça upload do arquivo Excel para processar")

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx'])

    if uploaded_file is not None:
        if st.button("Processar Arquivo"):
            try:
                # Processa a planilha
                trello_data, faltas_atualizadas, colunas_exportadas = processar_planilha(uploaded_file)
                
                st.success("Arquivo processado com sucesso!")
                st.write("Colunas exportadas:", ", ".join(colunas_exportadas))

                # Prepara os arquivos Excel para download
                def to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    return output.getvalue()

                # Cria os botões de download
                st.download_button(
                    label="Baixar arquivo Trello formatado",
                    data=to_excel(trello_data),
                    file_name=f"Trello_Formatado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.download_button(
                    label="Baixar planilha atualizada",
                    data=to_excel(faltas_atualizadas),
                    file_name=f"Faltas_Atualizadas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Erro ao processar arquivo: {str(e)}")

if __name__ == "__main__":
    main()
