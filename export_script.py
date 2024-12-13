import os
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine
import psycopg2

# Configurações do banco de dados
DB_HOST = os.environ.get('DB_HOST')
DB_PORT = str(os.environ.get('DB_PORT'))
DB_NAME = os.environ.get('DB_NAME')
DB_USER = os.environ.get('DB_USER')
DB_PASS = os.environ.get('DB_PASS')

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
        
        # Listar arquivos no diretório
        print("\n5. Arquivos no diretório:")
        for arquivo in os.listdir():
            print(f"- {arquivo}")
    else:
        print("✗ Erro: Arquivo não foi criado!")

except Exception as e:
    print(f"\n❌ Erro durante a execução: {str(e)}")
    raise

print("\n=== Fim da execução ===")
