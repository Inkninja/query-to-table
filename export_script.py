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
SELECT
		cce.idcontacorrenteembarque,
    pr.nrprocesso,
    cce.dtlancamento AS "Data Lancamento",
    cce.dtpagamento AS "Data Pagamento",
    cce.dtvencimento AS "Data Vencimento",
    pe.appessoa AS "Cliente",
    usco.nmusuario,
    COALESCE(pr.nrconhecimento, pr.nrconhecmaster) AS "Conhecimento",
    pec.nmpessoa AS "Fornecedor",
    it.nmitemdespesa,
    CASE 
        WHEN cce.tpprocedencia = 'S' THEN cce.vritem * -1
        ELSE cce.vritem
    END AS "Valor Despesa",
    cce.nrdocumento,
    --CASE 
        --WHEN cce.dtpagamento IS NULL THEN 'A PAGAR'
        --ELSE 'PAGO'
    --END AS "Status",
    pgi.observacao,
		pe.cnpj as "CNPJ CLIENTE",
		pec.cnpj as "CNPJ FORNECEDOR",
		cp.dtcompetencia AS "Data Emissão",
		puo.nmpessoa AS "Unidade Operacional",
		puf.nmpessoa AS "Unidade Faturamento"
		
FROM processo pr
LEFT JOIN pessoa pe ON pe.idpessoa = pr.idpessoacliente
LEFT JOIN pessoa puo ON puo.idpessoa = pr.idpessoaunidade
LEFT JOIN pessoa puf ON puf.idpessoa = pr.idpessoaunidadefat
LEFT JOIN contacorrenteembarque cce ON cce.idprocesso = pr.idprocesso
LEFT JOIN itemdespesa it ON it.iditemdespesa = cce.iditem
LEFT JOIN pessoa pec ON pec.idpessoa = cce.idempresalancamento
LEFT JOIN servico se ON se.idservico = pr.idservico
inner JOIN contaspagaritem cpi on cpi.idprocesso = pr.idprocesso AND cpi.iditem = it.iditemdespesa
full JOIN contaspagar cp on cp.idcontaspagar = cpi.idcontaspagar  
LEFT JOIN (
    SELECT idprocesso, 
           MAX(nrdoctoitem) AS nrdoctoitem,
           MAX(observacao) AS observacao  
    FROM pagamentoitemembarque
    GROUP BY idprocesso
) pgi ON pgi.idprocesso = pr.idprocesso
LEFT JOIN usuario usco ON usco.idusuario = cce.idusuario
LEFT JOIN followupprocesso fpef ON fpef.idprocesso = pr.idprocesso AND fpef.idevento = 51
WHERE 
    pr.dtcancelamento IS NULL
    AND cce.dtcancelamento IS NULL
    AND cce.vritem IS NOT NULL
    --AND idcontafinanceiro IN ('4')
		--AND cce.idcontafinanceiro IN ('04')
    --AND cce.tpprocedencia IN ('S')
    --AND fpef.dtrealizacao IS NULL
    AND pr.dtabertura >= '2024-12-10'
		--AND PR.NRPROCESSO = 'NFIA2448076'
ORDER BY 
    cce.dtlancamento;
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
