import streamlit as st
import pandas as pd
import io

# Função para processar o arquivo Excel
def processar_arquivo(arquivo_excel):
    folhas = ["Stock Feira"]
    resultados = []

    # Lê e processa cada folha
    for folha in folhas:
        dados = pd.read_excel(arquivo_excel, sheet_name=folha)
        dados_filtrados = dados[dados['Stock_Min'] > 0]
        if not dados_filtrados.empty:
            dados_filtrados['Quantidade abaixo stock minimo'] = dados_filtrados['Stock_Min'] - dados_filtrados['Stock_Atual']
            filtrados = dados_filtrados[dados_filtrados['Quantidade abaixo stock minimo'] <= 0]
            if not filtrados.empty:
                resultado_folha = filtrados[['Ref', 'Quantidade abaixo stock minimo', 'ABC','Marca','Familia','LinhaProduto']]
                resultado_folha['Armazém'] = folha.split()[-1]
                resultado_folha = resultado_folha[['Armazém', 'Ref', 'Quantidade abaixo stock minimo', 'ABC','Marca','Familia','LinhaProduto']]
                resultados.append(resultado_folha)
    
    if resultados:
        resultado_final = pd.concat(resultados)
        return resultado_final
    else:
        return pd.DataFrame()

# Streamlit app layout
st.title('Análise de Stock Minimo')

# Upload do arquivo
uploaded_file = st.file_uploader("Carregue o arquivo Excel aqui:", type=['xlsx'])

if uploaded_file is not None:
    # Processa o arquivo carregado
    df_resultado = processar_arquivo(uploaded_file)
    
    if not df_resultado.empty:
        # Mostra o DataFrame resultante
        st.write("Resultados:")
        st.dataframe(df_resultado)

        # Transforma o DataFrame em um arquivo Excel para download
        towrite = io.BytesIO()
        df_resultado.to_excel(towrite, index=False, engine='openpyxl')  # Usa o engine 'openpyxl'
        towrite.seek(0)  # Volta ao início do stream

        # Link para download do resultado
        st.download_button(label="Baixar arquivo Excel processado", data=towrite, file_name='resultado_stock_minimo.xlsx', mime="application/vnd.ms-excel")
    else:
        st.write("Nenhum resultado encontrado para mostrar.")
