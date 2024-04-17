import streamlit as st
import pandas as pd
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import smtplib
import traceback

# Configurações de email SMTP
def send_email(receiver_email, file_stream):
    sender_email = "napsparts@sapo.pt"  # Seu email
    sender_password = "Naps2022#?"  # Sua senha

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Análise de Stock Mínimo"

    # Corpo da mensagem
    body = "Encontre em anexo a análise de stock mínimo."
    message.attach(MIMEText(body, "plain"))

    # Anexando o arquivo Excel
    file_stream.seek(0)
    part = MIMEApplication(file_stream.read(), Name='resultado_stock_minimo.xlsx')
    part['Content-Disposition'] = 'attachment; filename="resultado_stock_minimo.xlsx"'
    message.attach(part)

    # Conectando ao servidor e enviando o email
    try:
        server = smtplib.SMTP('smtp.sapo.pt', 587)  # Servidor SMTP
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        server.quit()
        return "E-mail enviado com sucesso!"
    except Exception as e:
        return f"Erro ao enviar email: {e}\n{traceback.format_exc()}"

# Função para processar o arquivo Excel com armazéns selecionados
def processar_arquivo(arquivo_excel, folhas_selecionadas):
    resultados = []
    todas_refs = []

    for folha in folhas_selecionadas:
        dados = pd.read_excel(arquivo_excel, sheet_name=folha)
        todas_refs.append(dados[['Ref', 'ABC']])
        
    todas_refs = pd.concat(todas_refs)
    refs_abc_a = todas_refs.groupby('Ref').filter(lambda x: all(x['ABC'] == 'A'))

    for folha in folhas_selecionadas:
        dados = pd.read_excel(arquivo_excel, sheet_name=folha)
        dados = dados[dados['Ref'].isin(refs_abc_a['Ref'])]
        if folha == "Stock Feira":
            dados_filtrados = dados[dados['Stock_Min'] > 0]
            if not dados_filtrados.empty:
                dados_filtrados['Quantidade abaixo stock minimo'] = dados_filtrados['Stock_Min'] - dados_filtrados['Stock_Atual']
                dados_filtrados['Total Pendentes'] = dados_filtrados.groupby('Ref')['Pendentes'].transform('sum')
                dados_filtrados['Armazém'] = folha.split()[-1]
                resultado_folha = dados_filtrados[['Armazém', 'Ref', 'Quantidade abaixo stock minimo', 'ABC', 'Marca', 'Familia', 'LinhaProduto', 'Total Pendentes']]
                resultados.append(resultado_folha)
        else:
            if not dados.empty:
                dados['Total Pendentes'] = dados.groupby('Ref')['Pendentes'].transform('sum')
                dados['Armazém'] = folha.split()[-1]
                dados['Quantidade abaixo stock minimo'] = 'N/A'
                resultado_folha = dados[['Armazém', 'Ref', 'ABC', 'Marca', 'Familia', 'LinhaProduto', 'Total Pendentes']]
                resultados.append(resultado_folha)

    return pd.concat(resultados) if resultados else pd.DataFrame()

# Layout do app Streamlit
st.title("Análise de Stock Mínimo - Super A's ")

opcoes_folhas = ["Stock Feira", "Stock Frielas", "Stock Coimbra", "Stock Lousada", "Stock Sintra", "Stock Albergaria", "Stock Braga", "Stock Porto", "Stock Seixal"]
folhas_selecionadas = st.multiselect("Selecione os armazéns para análise:", opcoes_folhas, default=opcoes_folhas)

uploaded_file = st.file_uploader("Carregue o ficheiro Excel aqui:", type=['xlsx'])
if uploaded_file:
    st.session_state.uploaded_file = uploaded_file

if st.button('Executar Análise') and 'uploaded_file' in st.session_state and folhas_selecionadas:
    df_resultado = processar_arquivo(st.session_state.uploaded_file, folhas_selecionadas)
    if not df_resultado.empty:
        st.success('Análise concluída!')
        st.dataframe(df_resultado)
        towrite = io.BytesIO()
        df_resultado.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.session_state.towrite = towrite  # Salvar towrite em session_state
        st.download_button("Baixar arquivo Excel processado", towrite, "resultado_stock_minimo.xlsx", "application/vnd.ms-excel")

receiver_email = st.text_input("Digite o e-mail para enviar a análise:")
if st.button("Enviar Análise por E-mail") and 'towrite' in st.session_state and receiver_email:
    with st.spinner('Enviando e-mail, por favor aguarde...'):
        send_result = send_email(receiver_email, st.session_state.towrite)
    st.success(send_result)
