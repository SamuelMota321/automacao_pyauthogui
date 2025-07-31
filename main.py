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

class DataExtractor:
    """
    Gerencia a extração e o carregamento de dados de arquivos CSV.
    """
    def __init__(self, csv_file_path: str):
        self.csv_file_path = csv_file_path
        self.data = None

    def load_data(self) -> pd.DataFrame:
        """
        Carrega os dados do arquivo CSV especificado em um DataFrame do pandas.
        """
        if not os.path.exists(self.csv_file_path):
            print(f"Erro: O arquivo CSV '{self.csv_file_path}' não foi encontrado.")
            return None
        try:
            self.data = pd.read_csv(self.csv_file_path)
            print(f"Dados carregados com sucesso de '{self.csv_file_path}'.")
            return self.data
        except Exception as e:
            print(f"Erro ao carregar dados do CSV '{self.csv_file_path}': {e}")
            return None

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

            value = None
            if csv_column and csv_column in record:
                value = record[csv_column]
                # Aplica a transformação se definida
                if transform_func:
                    try:
                        value = transform_func(value)
                    except Exception as e:
                        print(f"Aviso: Erro ao transformar '{field_name}' (coluna '{csv_column}'). Valor original: '{value}'. Erro: {e}")
                        value = '' # Em caso de erro, usa string vazia para não travar o pyautogui
                else:
                    value = str(value) if pd.notna(value) else '' # Garante que é string ou vazio

            # Armazena os dados com coordenadas
            mapped_data[field_name] = {
                "value": value,
                "coords": {"x": x_coord, "y": y_coord}
            }
        return mapped_data

# --- NOVA CLASSE: UISystemIntegrator ---
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
        # Dá um tempo para o usuário posicionar a janela do sistema
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
                    pyautogui.click(x_coord, y_coord) # Clica no campo
                    time.sleep(0.1) # Pequena pausa para o clique ser reconhecido
                    pyautogui.hotkey('ctrl', 'a') # Seleciona todo o texto existente (para limpar)
                    pyautogui.press('delete') # Deleta o texto existente
                    # Para textos longos ou com caracteres especiais, Ctrl+C / Ctrl+V é mais robusto:
                    # pyautogui.copy(value_to_fill)
                    # pyautogui.hotkey('ctrl', 'v')
                    pyautogui.write(str(value_to_fill)) # Digita o valor no campo
                    time.sleep(0.2) # Pausa após digitar

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
        self.data_extractor = DataExtractor(self.output_csv)
        self.data_mapper = DataMapper(self.field_mappings)
        self.integrator = UISystemIntegrator() # Agora usando o integrador de UI

    def run(self):
        print(f"Iniciando processo de integração para '{self.pdf_file}' via automação de UI...")
        
        # Passo 1: Verifica se o CSV gerado pelo seu script camelot existe.
        if not os.path.exists(self.output_csv):
            print(f"Erro: O arquivo '{self.output_csv}' não foi encontrado. Por favor, execute o script de extração de tabelas (camelot) primeiro.")
            return

        # Passo 2: Carrega os dados do CSV gerado.
        df = self.data_extractor.load_data()
        if df is None or df.empty:
            print("Nenhum dado válido para processar. Encerrando.")
            return

        # Passo 3: Itera pelos registros, mapeia e envia para o sistema interno via UI.
        print(f"\nProcessando {len(df)} registros para integração UI...")
        for index, record in df.iterrows():
            print(f"\n--- Processando Registro {index + 1} ---")
            mapped_record = self.data_mapper.map_record(record)
            
            # Aqui você adicionaria uma lógica para navegar para o "novo registro" no seu sistema interno.
            # Por exemplo, clicar em um botão "+ Novo" ou usar um atalho de teclado.
            # EX: pyautogui.click(x_novo_botao, y_novo_botao)
            # EX: pyautogui.hotkey('alt', 'n') # Exemplo de atalho para "Novo"
            
            success = self.integrator.send_data(mapped_record)
            if success:
                print(f"Registro {index + 1} preenchido via UI com sucesso (simulado).")
                # Se houver um botão de "Salvar" ou "Confirmar", você clicaria nele aqui.
                # EX: pyautogui.click(x_salvar_botao, y_salvar_botao)
                # EX: time.sleep(1) # Aguarda o salvamento
            else:
                print(f"Falha ao preencher o registro {index + 1} via UI. Verifique a janela do sistema e as coordenadas.")
                break # Interrompe se a automação falhar ou for abortada pelo usuário

        print("\nProcesso de integração de UI concluído.")

# --- Execução Principal (o 'main' do seu script) ---
if __name__ == "__main__":
    PDF_FILE = 'Ploomes _ Clientes _ DAMASCO.pdf'
    OUTPUT_CSV_FILE = 'tabela_extraida.csv' # Nome do CSV que o seu script camelot gera

    # Instancia e executa a aplicação
    app = DataIntegrationApp(PDF_FILE, OUTPUT_CSV_FILE, FIELD_MAPPINGS)
    app.run()
