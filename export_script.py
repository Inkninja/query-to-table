import os
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine
import psycopg2
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Configurações do banco de dados
DB_HOST = os.environ.get('DB_HOST')
DB_PORT = str(os.environ.get('DB_PORT'))
DB_NAME = os.environ.get('DB_NAME')
DB_USER = os.environ.get('DB_USER')
DB_PASS = os.environ.get('DB_PASS')

# Configurações do email
EMAIL_USER = os.environ.get('EMAIL_USER')
EMAIL_PASS = os.environ.get('EMAIL_PASS')
EMAIL_TO = os.environ.get('EMAIL_TO')

print("=== Iniciando execução do script ===")

# Sua consulta SQL
query = """
SELECT *
FROM sua_tabela
WHERE data_criacao >= date_trunc('month', current_date - interval '1' month)
AND data_criacao < date_trunc('month', current_date);
"""

try:
    print("\n1. Conectando ao banco de dados...")
    conn_string = f"postgresql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
    engine = create_engine(conn_string)
    
    print("\n2. Executando query...")
    df = pd.read_sql_query(query, engine)
    
    print("\n3. Verificando resultados...")
    print(f"Número de linhas retornadas: {len(df)}")
    print(f"Colunas: {', '.join(df.columns)}")
    
    # Gerar nome do arquivo com data
    data_atual = datetime.now().strftime('%Y_%m')
    nome_arquivo = f'relatorio_{data_atual}.xlsx'
    caminho_completo = os.path.join(os.getcwd(), nome_arquivo)
    
    print(f"\n4. Salvando arquivo Excel em: {caminho_completo}")
    df.to_excel(nome_arquivo, index=False)
    
    # Verificar se o arquivo foi criado
    if os.path.exists(caminho_completo):
        tamanho = os.path.getsize(caminho_completo)
        print(f"✓ Arquivo criado com sucesso! Tamanho: {tamanho/1024:.2f} KB")
        
        print("\n5. Enviando email...")
        # Criar a mensagem
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = EMAIL_TO
        msg['Subject'] = f'Relatório Mensal - {data_atual}'
        
        # Corpo do email
        body = f"""
        Olá,

        Segue em anexo o relatório mensal gerado em {datetime.now().strftime('%d/%m/%Y')}.

        Atenciosamente,
        Sistema de Relatórios
        """
        msg.attach(MIMEText(body, 'plain'))
        
        # Anexar o arquivo
        with open(nome_arquivo, 'rb') as f:
            part = MIMEApplication(f.read(), Name=nome_arquivo)
            part['Content-Disposition'] = f'attachment; filename="{nome_arquivo}"'
            msg.attach(part)
        
        # Conectar ao servidor SMTP do Outlook
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        
        # Enviar email
        server.send_message(msg)
        server.quit()
        print("✓ Email enviado com sucesso!")

except Exception as e:
    print(f"\n❌ Erro durante a execução: {str(e)}")
    raise

print("\n=== Fim da execução ===")
