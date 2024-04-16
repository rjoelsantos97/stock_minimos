import streamlit as st
import pandas as pd
import io

# Função para processar o arquivo Excel
def processar_arquivo(arquivo_excel):
    folhas = ["Stock Feira", "Stock Frielas", "Stock Coimbra", "Stock Lousada", "Stock Sintra", "Stock Albergaria", "Stock Braga", "Stock Porto", "Stock Seixal"]
    resultados = []
    todas_refs = []

    # Lê e processa cada folha para coletar todas as referências e seus ABCs
    for folha in folhas:
        dados = pd.read_excel(arquivo_excel, sheet_name=folha)
        todas_refs.append(dados[['Ref', 'ABC']])

    # Concatena todas as referências para verificar a condição ABC = 'A'
    todas_refs = pd.concat(todas_refs)
    refs_abc_a = todas_refs.groupby('Ref').filter(lambda x: all(x['ABC'] == 'A'))

    # Processa cada folha considerando apenas as refs com ABC = 'A' em todos os armazéns
    for folha in folhas:
        dados = pd.read_excel(arquivo_excel, sheet_name=folha)
        dados = dados[dados['Ref'].isin(refs_abc_a['Ref'])]
        if folha == "Stock Feira":
            dados_filtrados = dados[dados['Stock_Min'] > 0]
            if not dados_filtrados.empty:
                dados_filtrados['Quantidade abaixo stock minimo'] = dados_filtrados['Stock_Min'] - dados_filtrados['Stock_Atual']
                filtrados = dados_filtrados[dados_filtrados['Quantidade abaixo stock minimo'] <= 0]
                if not filtrados.empty:
                    resultado_folha = filtrados[['Ref', 'Quantidade abaixo stock minimo', 'ABC']]
                    resultado_folha['Armazém'] = folha.split()[-1]
                    resultado_folha = resultado_folha[['Armazém', 'Ref', 'Quantidade abaixo stock minimo', 'ABC']]
                    resultados.append(resultado_folha)
        else:
            if not dados.empty:
                resultado_folha = dados[['Ref', 'ABC']]
                resultado_folha['Armazém'] = folha.split()[-1]
                resultado_folha['Quantidade abaixo stock minimo'] = 'N/A'  # Para folhas que não são 'Stock Feira'
                resultado_folha = resultado_folha[['Armazém', 'Ref', 'Quantidade abaixo stock minimo', 'ABC']]
                resultados.append(resultado_folha)
    
    if resultados:
        return pd.concat(resultados)
    else:
        return pd.DataFrame()

# Streamlit app layout
st.title("Análise de Stock Mínimo - Super A's")

# Upload do arquivo
uploaded_file = st.file_uploader("Carregue o ficheiro Excel aqui:", type=['xlsx'])

# Botão para executar a análise
if st.button('Executar Análise'):
    if uploaded_file is not None:
        with st.spinner('A executar análise, por favor aguarde. Pode demorar alguns minutos...'):
            # Processa o arquivo carregado
            df_resultado = processar_arquivo(uploaded_file)
            
            if not df_resultado.empty:
                # Mostra o DataFrame resultante
                st.success('Análise concluída!')
                st.dataframe(df_resultado)

                # Transforma o DataFrame em um arquivo Excel para download
                towrite = io.BytesIO()
                df_resultado.to_excel(towrite, index=False, engine='openpyxl')  # Usa o engine 'openpyxl'
                towrite.seek(0)  # Volta ao início do stream

                # Link para download do resultado
                st.download_button(label="Download ficheiro Excel processado", data=towrite, file_name='resultado_stock_minimo.xlsx', mime="application/vnd.ms-excel")
            else:
                st.error("Nenhum resultado encontrado para mostrar.")
    else:
        st.error("Por favor, carregue um ficheiro para análise.")

# Footer
footer_html = "<div style='background-color: #f1f1f1; color: #707070; font-size: 16px; padding: 10px; text-align: center; border-top: 1px solid #e0e0e0;'>Desenvolvido por NAPS Parts & Solutions</div>"
st.markdown(footer_html, unsafe_allow_html=True)
