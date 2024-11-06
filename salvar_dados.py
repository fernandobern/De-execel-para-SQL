import pandas as pd
import pyodbc
import random
import datetime

# Conectar ao SQL Server
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=FERNANDO\SQLFB;'
    'DATABASE=db_trovato;'
    'Trusted_Connection=yes'
)

# Ler a planilha usando pandas (substitua o caminho pelo local do arquivo)
df = pd.read_excel('dados.xlsm')

# Criar um cursor para executar comandos SQL
cursor = conn.cursor()

# Obter a data atual usando datetime
data_atual = datetime.datetime.now()

# Função para gerar um login aleatório (com 6 números aleatórios)
def gerar_login_aleatorio():
    return f"login_{random.randint(100000, 999999)}"

# Iterar sobre as linhas do dataframe começando da segunda linha (index 1)
for index, row in df.iloc[3:].iterrows():  # Começa a partir da segunda linha
    # Gerar um login aleatório para cada linha
    login_aleatorio = gerar_login_aleatorio()
    
    cursor.execute("""
        INSERT INTO dbo.Alunos (id_aluno, NOME, DATA_NASCIMENTO, SEXO, data_cadastro, login_cadastro)
        VALUES (?, ?, ?, ?, ?, ?)
    """, row['ID'], row['NOME'], row['DATA_NASCIMENTO'], row['SEXO'], data_atual, login_aleatorio)

# Confirmar a transação
conn.commit()

# Fechar a conexão
cursor.close()
conn.close()

print("Dados inseridos com sucesso.")
