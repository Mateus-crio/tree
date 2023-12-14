import requests
from bs4 import BeautifulSoup
import pandas as pd
import mysql.connector

class TableScraper:
    def __init__(self, url, table_id):
        self.url = url
        self.headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36"}
        self.table_id = table_id
        self.data = None

    def scrape_data(self):
        try:
            response = requests.get(self.url, headers=self.headers)
            response.raise_for_status()  # Verifica se houve erro na requisição HTTP

            soup = BeautifulSoup(response.text, 'html.parser')
            # Localizar a tabela pelo ID
            table = soup.find('div')
            print(table)
            if table:
                # Extrair dados da tabela usando pandas
                self.data = pd.read_html(str(table))[0]
        except requests.exceptions.RequestException as e:
            print(f"Erro ao fazer a requisição: {e}")

    def store_in_database(self, cursor, table_name):
        if self.data is not None:
            try:
                # Criar a tabela se ela não existir
                self.create_table(cursor, table_name)

                # Inserir dados na tabela
                for _, row in self.data.iterrows():
                    cursor.execute(f"INSERT INTO {table_name} VALUES {tuple(row)}")

                # Commit the changes
                connection.commit()
                print("Dados inseridos com sucesso.")
            except mysql.connector.Error as err:
                print(f"Erro ao interagir com o banco de dados: {err}")
        else:
            print("Nenhum dado para inserir.")

    def create_table(self, cursor, table_name):
        try:
            # Obter os nomes das colunas
            column_names = list(self.data.columns)

            # Criar a tabela se ela não existir
            create_table_query = f"CREATE TABLE IF NOT EXISTS {table_name} ({', '.join([f'{col} VARCHAR(255)' for col in column_names])})"
            cursor.execute(create_table_query)
        except mysql.connector.Error as err:
            print(f"Erro ao criar a tabela: {err}")

if __name__ == "__main__":
    url = 'https://mateus-crio.github.io/tree/'
    host = 'localhost'
    user = 'senai'
    password = 'Senai125@'
    database = 'basedados'
    table_name = 'sala'
    table_id = 'content'

    # Conectar ao banco de dados MySQL
    try:
        connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database
        )

        cursor = connection.cursor()

        scraper = TableScraper(url, table_id)
        scraper.scrape_data()
        scraper.store_in_database(cursor, table_name)

    finally:
        # Fechar a conexão mesmo em caso de exceção
        if 'connection' in locals() and connection.is_connected():
            cursor.close()
            connection.close()


def export_to_excel(host, user, password, database, table_name, excel_filename):
    try:
        # Conectar ao banco de dados MySQL
        connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database
        )

        # Criar um DataFrame a partir dos dados da tabela
        query = f"SELECT * FROM {table_name}"
        query2 = f"SELECT * FROM {table_name2}"
        df = pd.read_sql(query, connection)
        df2 = pd.read_sql(query, connection)
        
        # Salvar o DataFrame em um arquivo Excel
        df.to_excel(excel_filename, index=False)
        df2.to_excel(excel_filename, index=False)

        print(f'Dados exportados para {excel_filename} com sucesso.')

    except mysql.connector.Error as err:
        print(f"Erro ao interagir com o banco de dados: {err}")

    finally:
        # Fechar a conexão mesmo em caso de exceção
        if 'connection' in locals() and connection.is_connected():
            connection.close()

if __name__ == "__main__":
    url = 'https://mateus-crio.github.io/tree/'
    host = 'localhost'
    user = 'senai'
    password = 'Senai125@'
    database = 'basedados'
    table_name = 'sala'
    table_name2 = 'estrutura'
    table_id = 'content'
    excel_filename = 'dados_sala.xlsx'
    # Extrair dados da tabela e armazenar no banco de dados MySQL
    scraper = TableScraper(url, table_id)
    scraper.scrape_data()

    connection = mysql.connector.connect(
        host=host,
        user=user,
        password=password,
        database=database
    )

    cursor = connection.cursor()
    scraper.store_in_database(cursor, table_name)

    # Exportar dados para o Excel
    export_to_excel(host, user, password, database, table_name, excel_filename)