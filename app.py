import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

def verificar_colunas_obrigatorias(df):
    colunas_necessarias = {
        'NOME', 'MATRÍCULA', 'LOCALIZAÇÃO', 'DIA', 'BATIDAS', 
        'ENTRADA 1', 'SAÍDA 1', 'ENTRADA 2', 'SAÍDA 2', 
        'ATRASO', 'FALTA', 'BANCO DE HORAS', 
        'HORA EXTRA 50% (N.A.)', 'HORA EXTRA 100% (N.A.)', 
        'DSR DESCONTADO', 'ADICIONAL NOTURNO', 'EXPEDIENTE'
    }
    
    colunas_atuais = set(df.columns)
    colunas_faltantes = colunas_necessarias - colunas_atuais
    
    if colunas_faltantes:
        raise ValueError(f"Colunas obrigatórias faltando: {', '.join(colunas_faltantes)}")
    
    return True

def processar_planilha(uploaded_file):
    # Lê o arquivo Excel
    df = pd.read_excel(uploaded_file)
    
    # Padroniza nomes das colunas
    df.columns = [col.strip().upper() for col in df.columns]
    
    # Verifica se todas as colunas necessárias existem
    verificar_colunas_obrigatorias(df)
    
    # Adiciona coluna de verificação se não existir
    if 'ID VERIFICACAO' not in df.columns:
        df['ID VERIFICACAO'] = ''
    
    # Processa apenas linhas não processadas
    registros = []
    data_atual = datetime.now().strftime('%Y-%m-%d')
    colunas_exportadas = set()

    for index, row in df[df['ID VERIFICACAO'] != 'PROCESSADO'].iterrows():
        descricao = (
            f"Matrícula: {row['MATRÍCULA']}\n"
            f"Localização: {row['LOCALIZAÇÃO']}\n"
            f"Dia: {row['DIA']}\n"
        )

        # Verifica batidas - agora usando as colunas exatas
        batidas = [
            str(row['BATIDAS']).strip() if pd.notna(row['BATIDAS']) else '',
            str(row['ENTRADA 1']).strip() if pd.notna(row['ENTRADA 1']) else '',
            str(row['SAÍDA 1']).strip() if pd.notna(row['SAÍDA 1']) else '',
            str(row['ENTRADA 2']).strip() if pd.notna(row['ENTRADA 2']) else '',
            str(row['SAÍDA 2']).strip() if pd.notna(row['SAÍDA 2']) else ''
        ]
        
        if all(not batida or batida == '00:00' for batida in batidas):
            registros.append({
                'list': 'SEM BATIDA',
                'Card Name': row['NOME'],
                'desc': descricao,
                'checklist': 'Sem registros de batida',
                'Data': data_atual
            })
            colunas_exportadas.add('SEM BATIDA')

        # Mapeamento exato dos campos para verificação
        campos_verificacao = {
            'ATRASO': 'ATRASO',
            'FALTA': 'FALTA',
            'BANCO DE HORAS': 'BANCO DE HORAS',
            'HORA EXTRA 50% (N.A.)': 'HORA EXTRA 50%',
            'HORA EXTRA 100% (N.A.)': 'HORA EXTRA 100%',
            'DSR DESCONTADO': 'DSR DESCONTADO',
            'ADICIONAL NOTURNO': 'ADICIONAL NOTURNO',
            'EXPEDIENTE': 'EXPEDIENTE'
        }

        for coluna, lista in campos_verificacao.items():
            valor = str(row[coluna]).strip() if pd.notna(row[coluna]) else ''
            if valor and valor != '00:00':
                registros.append({
                    'list': lista,
                    'Card Name': row['NOME'],
                    'desc': descricao,
                    'checklist': valor,
                    'Data': data_atual
                })
                colunas_exportadas.add(lista)

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
                    output.seek(0)
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

            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {str(e)}")

if __name__ == "__main__":
    main()
