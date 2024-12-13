# Arquivo: export_script.py
import os
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine

# Configurações do banco de dados
DB_HOST = os.environ.get('DB_HOST')
DB_NAME = os.environ.get('DB_NAME')
DB_USER = os.environ.get('DB_USER')
DB_PASS = os.environ.get('DB_PASS')

# Sua consulta SQL
query = """
SELECT *
FROM sua_tabela
WHERE data_criacao >= date_trunc('month', current_date - interval '1' month)
AND data_criacao < date_trunc('month', current_date);
"""

# Conectar ao banco de dados
conn_string = f"postgresql://{DB_USER}:{DB_PASS}@{DB_HOST}/{DB_NAME}"
engine = create_engine(conn_string)

# Executar a consulta e carregar em um DataFrame
df = pd.read_sql_query(query, engine)

# Gerar nome do arquivo com data
data_atual = datetime.now().strftime('%Y_%m')
nome_arquivo = f'relatorio_{data_atual}.xlsx'

# Salvar como Excel
df.to_excel(nome_arquivo, index=False)

print(f'Arquivo {nome_arquivo} gerado com sucesso!')
