import streamlit as st
import pandas as pd
import io

# Função para processar o arquivo Excel com armazéns selecionados
def processar_arquivo(arquivo_excel, folhas_selecionadas):
    resultados = []
    todas_refs = []

    # Lê e processa cada folha selecionada para coletar todas as referências e seus ABCs
    for folha in folhas_selecionadas:
        dados = pd.read_excel(arquivo_excel, sheet_name=folha)
        todas_refs.append(dados[['Ref', 'ABC']])
        
    # Concatena todas as referências para verificar a condição ABC = 'A'
    todas_refs = pd.concat(todas_refs)
    refs_abc_a = todas_refs.groupby('Ref').filter(lambda x: all(x['ABC'] == 'A'))

    # Processa cada folha considerando apenas as refs com ABC = 'A' em todos os armazéns
    for folha in folhas_selecionadas:
        dados = pd.read_excel(arquivo_excel, sheet_name=folha)
        dados = dados[dados['Ref'].isin(refs_abc_a['Ref'])]
        if folha == "Stock Feira":
            dados_filtrados = dados[dados['Stock_Min'] > 0]
            if not dados_filtrados.empty:
                dados_filtrados['Quantidade abaixo stock minimo'] = dados_filtrados['Stock_Min'] - dados_filtrados['Stock_Atual']
                filtrados = dados_filtrados[dados_filtrados['Quantidade abaixo stock minimo'] <= 0]
                if not filtrados.empty:
                    filtrados['Total Pendentes'] = dados_filtrados.groupby('Ref')['Pendentes'].transform('sum')
                    resultado_folha = filtrados[['Armazém', 'Ref', 'Quantidade abaixo stock minimo', 'ABC', 'Marca', 'Familia', 'LinhaProduto', 'Total Pendentes']]
                    resultados.append(resultado_folha)
        else:
            if not dados.empty:
                dados['Total Pendentes'] = dados.groupby('Ref')['Pendentes'].transform('sum')
                resultado_folha = dados[['Armazém', 'Ref', 'ABC', 'Marca', 'Familia', 'LinhaProduto', 'Total Pendentes']]
                resultado_folha['Quantidade abaixo stock minimo'] = 'N/A'  # Para folhas que não são 'Stock Feira'
                resultados.append(resultado_folha)

    if resultados:
        resultado_final = pd.concat(resultados)
        return resultado_final
    else:
        return pd.DataFrame()

# Lista de folhas disponíveis para seleção
opcoes_folhas = ["Stock Feira", "Stock Frielas", "Stock Coimbra", "Stock Lousada", "Stock Sintra", "Stock Albergaria", "Stock Braga", "Stock Porto", "Stock Seixal"]

# Streamlit app layout
st.title("Análise de Stock Mínimo - Super A's ")

# Permitir ao usuário selecionar folhas
folhas_selecionadas = st.multiselect("Selecione os armazéns para análise:", opcoes_folhas, default=opcoes_folhas)

# Upload do arquivo
uploaded_file = st.file_uploader("Carregue o ficheiro Excel aqui:", type=['xlsx'])

# Botão para executar a análise
if st.button('Executar Análise'):
    if uploaded_file is not None and folhas_selecionadas:
        with st.spinner('A executar análise, por favor aguarde. Pode demorar alguns minutos...'):
            df_resultado = processar_arquivo(uploaded_file, folhas_selecionadas)
            if not df_resultado.empty:
                # Mostra o DataFrame resultante
                st.success('Análise concluída!')
                st.dataframe(df_resultado)

                # Transforma o DataFrame em um arquivo Excel para download
                towrite = io.BytesIO()
                df_resultado.to_excel(towrite, index=False, engine='openpyxl')
                towrite.seek(0)  # Volta ao início do stream

                # Link para download do resultado
                st.download_button(label="Baixar arquivo Excel processado", data=towrite, file_name='resultado_stock_minimo.xlsx', mime="application/vnd.ms-excel")
            else:
                st.error("Nenhum resultado encontrado para mostrar.")
    else:
        st.error("Por favor, carregue um arquivo e selecione pelo menos um armazém para análise.")

# Footer
footer_html = "<div style='background-color: #f1f1f1; color: #707070; font-size: 16px; padding: 10px; text-align: center; border-top: 1px solid #e0e0e0;'>Desenvolvido por NAPS Parts & Solutions</div>"
st.markdown(footer_html, unsafe_allow_html=True)
