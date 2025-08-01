import pandas as pd
import os
import time
import pyautogui
import csv

# Para o script se der algo de errado
pyautogui.FAILSAFE = True
# Pequena pausa entre cada comando do pyautogui para dar tempo ao sistema responder.
pyautogui.PAUSE = 0.5

# GUIA DE COMO PREENCHER O FIELD_MAPPINGS:

# 'nome_do_campo_no_seu_codigo': {
#     'csv_column': 'Nome Exato da Coluna no Seu CSV',
#     'x': Coordenada X (em pixels) na tela,
#     'y': Coordenada Y (em pixels) na tela,
#     'transform_func': Função para limpar/formatar o dado (opcional)
# }
"""Substituir as coordenadas x e y nos FIELD_MAPPINGS pelas coletadas do sistema interno.
conferir csv_column para garantir que batem exatamente com os cabeçalhos do seu tabela_extraida.csv."""
FIELD_MAPPINGS = {
    "razao_social": {"csv_column": "Razão Social", "x": 271, "y": 178, "transform_func": str.strip},
    "nome_fantasia": {"csv_column": "Nome Fantasia", "x": 261, "y": 197, "transform_func": str.strip},
    "cnpj": {"csv_column": "CNPJ", "x": 686, "y": 176, "transform_func": lambda x: str(x).replace('.', '').replace('/', '').replace('-', '') if pd.notna(x) else ''},
    "inscricao_estadual": {"csv_column": "Inscrição Estadual", "x": 207, "y": 215, "transform_func": lambda x: str(int(x)) if pd.notna(x) else ''}, # Converte para int antes de str
    "endereco": {"csv_column": "Endereço", "x": 300, "y": 270, "transform_func": str.strip},
    "responsavel_dados_cadastrais": {"csv_column": "Responsável", "x": 300, "y": 300, "transform_func": str.strip},
    "telefone_responsavel": {"csv_column": "Telefone", "x": 659, "y": 434, "transform_func": str.strip},
    "bairro": {"csv_column": "Bairro", "x": 150, "y": 420, "transform_func": str.strip},
    "cidade": {"csv_column": "Cidade", "x": 150, "y": 400, "transform_func": str.strip},
    "uf": {"csv_column": "UF", "x": 130, "y": 375, "transform_func": str.strip},
    "cep": {"csv_column": "CEP", "x": 244, "y": 375, "transform_func": lambda x: str(x).replace('-', '') if pd.notna(x) else ''},
}

# --- Classes de Suporte ---

class DataExtractor:
    """
    Gerencia a extração e o carregamento de dados de arquivos CSV.
    Inclui um passo de pré-processamento para formatar os dados.
    """
    def __init__(self, csv_file_path: str):
        self.csv_file_path = csv_file_path
        self.data = None

    def load_data(self) -> pd.DataFrame:
        """
        Carrega os dados do arquivo CSV bruto e os pós-processa.
        """
        if not os.path.exists(self.csv_file_path):
            print(f"Erro: O arquivo CSV '{self.csv_file_path}' não foi encontrado.")
            return None
        try:
            # Carrega o CSV bruto como veio do camelot
            raw_df = pd.read_csv(self.csv_file_path, header=None, dtype=str).fillna('') # Ler sem cabeçalho, tudo como string e NaN como vazio
            print(f"Dados brutos carregados de '{self.csv_file_path}'.")
            
            # --- CHAME O MÉTODO DE PÓS-PROCESSAMENTO AQUI ---
            processed_df = self._process_raw_data(raw_df)
            
            self.data = processed_df
            print(f"Dados pós-processados com sucesso. Total de registros: {len(processed_df)}")
            print("DataFrame pós-processado (primeiras 5 linhas):")
            print(processed_df.head())
            return self.data
        except Exception as e:
            print(f"Erro ao carregar ou processar dados do CSV '{self.csv_file_path}': {e}")
            return None

    def _process_raw_data(self, raw_df: pd.DataFrame) -> pd.DataFrame:
        """
        Transforma o DataFrame bruto extraído pelo camelot em um formato tabular limpo.
        Esta é a parte crucial que você precisa adaptar exatamente ao layout do seu PDF/CSV.
        """
        # Exemplo de como você pode extrair os dados para um único cliente.
        # Se o seu PDF tiver múltiplos clientes, você precisará de uma lógica para
        # identificar o início e fim de cada bloco de cliente.

        processed_records = []
        current_record = {}

        # Ajustes de como ler e limpar cada campo
        for index, row in raw_df.iterrows():
            # Acessa as colunas usando os índices numéricos que o pandas atribuiu (0, 1, 2, 3)
            col0 = row[0].strip() if 0 in row.index and pd.notna(row[0]) else ''
            col1 = row[1].strip() if 1 in row.index and pd.notna(row[1]) else ''
            col2 = row[2].strip() if 2 in row.index and pd.notna(row[2]) else ''
            col3 = row[3].strip() if 3 in row.index and pd.notna(row[3]) else ''

            # Lógica para UNIDADE DE NEGÓCIO, TIPO DE ANÁLISE, CADASTRO PAT, Nº DO CONTRATO
            # Estes estão na primeira linha após o cabeçalho 0,1,2,3
            if index == 0: # Assumindo que a primeira linha de dados é o índice 0 após o header=None
                current_record["UNIDADE DE NEGÓCIO"] = col0
                current_record["TIPO DE ANÁLISE"] = col1
                # O "CADASTRO PAT" tem que ser tratado para remover a linha extra "Nº.:"
                current_record["CADASTRO PAT"] = col2.replace('\nNº.:', '').strip()
                current_record["Nº DO CONTRATO"] = col3

            # --- DADOS CADASTRAIS ---
            elif "Razão Social:" in col0:
                current_record["Razão Social"] = col0.replace("Razão Social:", "").strip()
            elif "Nome Fantasia:" in col0:
                current_record["Nome Fantasia"] = col0.replace("Nome Fantasia:", "").strip()
            elif "CNPJ:" in col0:
                current_record["CNPJ"] = col0.replace("CNPJ:", "").strip()
                current_record["Inscrição Estadual"] = col1.replace("Inscrição Estadual:", "").strip()
                current_record["Responsável (Dados Cadastrais)"] = col2.replace("Responsável:", "").strip()
                current_record["Telefone (Dados Cadastrais)"] = col3.replace("Telefone:", "").strip()
            elif "Endereço:" in col0:
                 # Endereço pode ter quebras de linha que precisam ser tratadas
                current_record["Endereço"] = col0.replace("Endereço:", "").replace('\n', ' ').strip()
            elif "Bairro:" in col0:
                current_record["Bairro"] = col0.replace("Bairro:", "").strip()
                current_record["Cidade"] = col1.replace("Cidade:", "").strip()
                current_record["CEP"] = col2.replace("CEP:", "").strip()
                current_record["UF"] = col3.replace("UF:", "").strip()

            # --- DADOS OBRIGATÓRIOS (NFE) ---
            elif "Nota Fiscal Eletrônica - Responsável:" in col0:
                current_record["NFE Responsavel"] = col0.replace("Nota Fiscal Eletrônica - Responsável:", "").strip()
                current_record["NFE Email"] = col1.replace("E-mail:", "").strip()
                current_record["NFE Telefone"] = col2.replace("Telefone:", "").strip()
            # Pode haver mais seções como "Cobrança Eletrônica" ou "Crédito para Importação" aqui,
            # você precisará identificar e extrair da mesma forma. Ex:
            # elif "Cobrança Eletrônica - Responsável:" in col0:
            #     current_record["Cobranca Responsavel"] = col0.replace("Cobrança Eletrônica - Responsável:", "").strip()
            #     current_record["Cobranca Email"] = col1.replace("E-mail:", "").strip()
            #     current_record["Cobranca Telefone"] = col2.replace("Telefone:", "").strip()


            # --- DADOS COMERCIAIS ---
            elif "Produto:" in col0:
                current_record["Produto (Comercial)"] = col0.replace("Produto:", "").strip()
                # O campo Nome de Embossing está na terceira coluna (índice 2)
                current_record["Nome de Embossing"] = col2.replace("Nome de Embossing:", "").strip()
            elif "Confissão de Dívida:" in col0:
                current_record["Confissão de Dívida"] = col0.replace("Confissão de Dívida:", "").strip()
                current_record["Método de Pagamento"] = col1.replace("Método de Pagamento:", "").strip()
                current_record["Tipo de Plano"] = col2.replace("Tipo de Plano:", "").strip()
                current_record["Vencimento do Plano"] = col3.replace("Vencimento do Plano:", "").strip()
            elif "Taxa de Administração" in col0:
                current_record["Taxa de Administração"] = col0.replace("Taxa de Administração", "").strip()
                current_record["Valor p/vida Clube de Vantagens"] = col1.replace("Valor p/vida Clube de Vantagens", "").strip()
                current_record["Qtde de Vidas"] = col2.replace("Qtde de Vidas", "").strip()
                current_record["Limite de Crédito Cliente (R$)"] = col3.replace("Limite de Crédito Cliente (R$)", "").strip()
            # ... E assim por diante para Limite por Cartão, Emissão de Cartão, etc.

            # --- PESQUISA FINANCEIRA ---
            elif "Tipo da Pesquisa:" in col0:
                current_record["Tipo da Pesquisa"] = col0.replace("Tipo da Pesquisa:", "").strip()

            # --- DADOS DOS SÓCIOS (Estes são múltiplos por cliente, um desafio à parte se precisar de todos) ---
            # Se você precisar dos dados de múltiplos sócios, a lógica aqui será mais complexa,
            # possivelmente criando um array de sócios dentro do current_record ou gerando múltiplos registros de clientes.
            # Por simplicidade, vou ignorar múltiplos sócios neste exemplo, mas você poderia adaptar.

            # Adicione outras seções como PARECER - EXECUTIVO COMERCIAL, AUTORIZAÇÃO, etc.
            # a lógica será similar: identificar a linha pelo texto em col0 e extrair os valores.


        # Ao final do loop, se você processa um cliente por vez, adiciona o record
        # Se você tem múltiplos clientes no mesmo CSV, precisa de uma lógica para
        # identificar o "fim" de um cliente e o "início" do próximo.
        processed_records.append(current_record)

        return pd.DataFrame(processed_records)


# --- CLASSE DataMapper (sem alterações significativas na lógica de mapeamento) ---
class DataMapper:
    """
    Mapeia os dados extraídos para o formato esperado pela integração,
    incluindo coordenadas e transformações.
    """
    def __init__(self, field_mappings: dict):
        self.field_mappings = field_mappings

    def map_record(self, record: pd.Series) -> dict:
        """
        Mapeia um único registro (linha do DataFrame) para o formato de preenchimento.
        Aplica transformações e associa com as coordenadas X, Y.
        """
        mapped_data = {}
        for field_name, mapping_info in self.field_mappings.items():
            csv_column = mapping_info.get("csv_column")
            x_coord = mapping_info.get("x")
            y_coord = mapping_info.get("y")
            transform_func = mapping_info.get("transform_func")

            value = '' # Padrão: string vazia para evitar NONE

            if csv_column and csv_column in record:
                raw_value = record[csv_column]
                
                if pd.isna(raw_value):
                    value = ''
                else:
                    value = str(raw_value)

                if transform_func:
                    try:
                        value = transform_func(value)
                    except Exception as e:
                        print(f"Aviso: Erro ao transformar '{field_name}' (coluna '{csv_column}'). Valor original: '{raw_value}'. Erro: {e}")
                        value = ''
            else:
                print(f"Aviso: Coluna '{csv_column}' não encontrada no DataFrame pós-processado para o campo '{field_name}'. Usando valor vazio.")

            mapped_data[field_name] = {
                "value": value,
                "coords": {"x": x_coord, "y": y_coord}
            }
        return mapped_data

# --- CLASSE UISystemIntegrator (inalterada) ---
class UISystemIntegrator:
    """
    Integra os dados no sistema interno através de automação de UI (PyAutoGUI),
    usando coordenadas X, Y.
    """
    def __init__(self, initial_delay: int = 5):
        self.initial_delay = initial_delay
        print(f"Modo de integração: Automação de Interface (PyAutoGUI).")
        print(f"O script começará a interagir com a tela em {self.initial_delay} segundos.")
        print("Mova o mouse para a janela do seu sistema interno e esteja pronto!")
        print("Para parar o script a qualquer momento, mova o mouse para um dos cantos da tela.")

    def send_data(self, data: dict) -> bool:
        """
        Preenche os campos no sistema interno usando PyAutoGUI e as coordenadas fornecidas.
        Retorna True em caso de sucesso (simulado), False caso ocorra um erro de automação.
        """
        print(f"\nIniciando preenchimento de um novo registro. Você tem {self.initial_delay} segundos para posicionar a janela.")
        time.sleep(self.initial_delay)
        print("Iniciando automação...")

        try:
            for field_name, field_info in data.items():
                value_to_fill = field_info['value']
                x_coord = field_info['coords']['x']
                y_coord = field_info['coords']['y']

                if x_coord is not None and y_coord is not None:
                    print(f"Preenchendo '{field_name}' (X:{x_coord}, Y:{y_coord}) com '{value_to_fill}'")
                    pyautogui.click(x_coord, y_coord)
                    time.sleep(0.1)
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('delete')
                    pyautogui.write(str(value_to_fill))
                    time.sleep(0.2)
                else:
                    print(f"Aviso: Coordenadas X, Y ausentes para o campo '{field_name}'. Pulando.")

            print("Preenchimento de registro concluído (simulado).")
            return True
        except pyautogui.FailSafeException:
            print("\nScript abortado pelo usuário (mouse no canto da tela).")
            return False
        except Exception as e:
            print(f"Erro durante a automação com PyAutoGUI: {e}")
            return False

# --- Fluxo Principal de Execução ---

class DataIntegrationApp:
    """
    Orquestra todo o processo de extração, mapeamento e integração de dados via UI.
    """
    def __init__(self, pdf_file: str, output_csv: str, field_mappings: dict):
        self.pdf_file = pdf_file
        self.output_csv = output_csv
        self.field_mappings = field_mappings
        self.data_extractor = DataExtractor(self.output_csv) # DataExtractor agora com pré-processamento
        self.data_mapper = DataMapper(self.field_mappings)
        self.integrator = UISystemIntegrator()

    def run(self):
        print(f"Iniciando processo de integração para '{self.pdf_file}' via automação de UI...")
        
        if not os.path.exists(self.output_csv):
            print(f"Erro: O arquivo '{self.output_csv}' não foi encontrado. Por favor, execute o script de extração de tabelas (camelot) primeiro.")
            return

        df = self.data_extractor.load_data() # Carrega o DF pós-processado
        if df is None or df.empty:
            print("Nenhum dado válido para processar. Encerrando.")
            return

        print(f"\nProcessando {len(df)} registros para integração UI...")
        for index, record in df.iterrows():
            print(f"\n--- Processando Registro {index + 1} ---")
            mapped_record = self.data_mapper.map_record(record)
            
            # --- Adicione aqui sua lógica de navegação para um novo registro ---
            # Ex: pyautogui.click(x_novo_botao, y_novo_botao)
            # time.sleep(1)
            
            success = self.integrator.send_data(mapped_record)
            if success:
                print(f"Registro {index + 1} preenchido via UI com sucesso (simulado).")
                # --- Adicione aqui sua lógica para salvar o registro ---
                # Ex: pyautogui.click(x_salvar_botao, y_salvar_botao)
                # time.sleep(1)
            else:
                print(f"Falha ao preencher o registro {index + 1} via UI. Verifique a janela do sistema e as coordenadas.")
                break

        print("\nProcesso de integração de UI concluído.")

# --- Execução Principal ---
if __name__ == "__main__":
    PDF_FILE = 'Ploomes _ Clientes _ DAMASCO.pdf'
    OUTPUT_CSV_FILE = 'tabela_extraida.csv'

    app = DataIntegrationApp(PDF_FILE, OUTPUT_CSV_FILE, FIELD_MAPPINGS)
    app.run()
