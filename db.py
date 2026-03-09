import mysql.connector

def conectar():
    return mysql.connector.connect(
        host="localhost",
        user="root",      # substitua pelo usuário do MySQL
        password="root"     # substitua pela senha do MySQL
    )

def inicializar_banco():
    conn = conectar()
    cursor = conn.cursor()

    # Cria o banco se não existir
    cursor.execute("CREATE DATABASE IF NOT EXISTS sistema_escalas")
    cursor.execute("USE sistema_escalas")

    # Cria tabela de funcionários
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS funcionarios (
        id INT AUTO_INCREMENT PRIMARY KEY,
        nome VARCHAR(100) NOT NULL,
        cpf VARCHAR(11) UNIQUE NOT NULL,
        setor VARCHAR(50),
        turno VARCHAR(20)
    )
    """)

    # Cria tabela de escalas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS escalas (
            id INT AUTO_INCREMENT PRIMARY KEY,
            mes INT NOT NULL,              -- número do mês (1 a 12)
            ano INT NOT NULL,              -- ano da escala
            dias_no_mes INT NOT NULL,      -- quantidade de dias no mês
            feriados TEXT,                 -- lista de feriados, ex: "01-01,25-12"
            folgas TEXT                    -- string com os funcionários e seus dias de folga
        )
        """)

    conn.commit()
    conn.close()
    print("Banco e tabelas inicializados com sucesso!")