import os
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine
import psycopg2

# Configurações do banco de dados
DB_HOST = os.environ.get('DB_HOST')
DB_PORT = str(os.environ.get('DB_PORT'))  # Garantindo que é string
DB_NAME = os.environ.get('DB_NAME')
DB_USER = os.environ.get('DB_USER')
DB_PASS = os.environ.get('DB_PASS')

# Imprime informações de debug (ocultando a senha)
print(f"Tentando conectar com:")
print(f"Host: {DB_HOST}")
print(f"Port: {DB_PORT}")
print(f"Database: {DB_NAME}")
print(f"User: {DB_USER}")

# Sua consulta SQL
query = """
SELECT *
FROM sua_tabela
WHERE data_criacao >= date_trunc('month', current_date - interval '1' month)
AND data_criacao < date_trunc('month', current_date);
"""

try:
    # Primeiro tenta conectar com psycopg2 diretamente para testar
    print("Testando conexão com psycopg2...")
    conn = psycopg2.connect(
        host=DB_HOST,
        port=DB_PORT,
        database=DB_NAME,
        user=DB_USER,
        password=DB_PASS
    )
    print("Conexão psycopg2 bem sucedida!")
    conn.close()
    
    # Se funcionou, tenta com SQLAlchemy
    print("Conectando com SQLAlchemy...")
    conn_string = f"postgresql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
    engine = create_engine(conn_string)
    
    # Testa a conexão
    with engine.connect() as connection:
        print("Conexão SQLAlchemy bem sucedida!")
        
    # Executar a consulta e carregar em um DataFrame
    print("Executando query...")
    df = pd.read_sql_query(query, engine)
    
    # Gerar nome do arquivo com data
    data_atual = datetime.now().strftime('%Y_%m')
    nome_arquivo = f'relatorio_{data_atual}.xlsx'
    
    # Salvar como Excel
    df.to_excel(nome_arquivo, index=False)
    print(f'Arquivo {nome_arquivo} gerado com sucesso!')

except Exception as e:
    print(f"Erro detalhado: {str(e)}")
    print("\nVerifique se:")
    print("1. O endereço do host está correto")
    print("2. A porta está correta")
    print("3. O banco está aceitando conexões externas")
    print("4. As credenciais estão corretas")
    print("5. O firewall permite conexões na porta especificada")
    raise
