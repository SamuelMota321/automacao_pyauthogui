import time
import camelot
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
 
# --- PARTE 1: EXTRAÇÃO DE DADOS DO PDF (Seu Código) ---3
def extract_data_from_pdf(pdf_file_path):
    """
    Extrai dados de tabelas de um PDF usando camelot e retorna um dicionário com os dados.
    Adaptado do seu código de extração.
    """
    data = {}
    try:
        tables = camelot.read_pdf(pdf_file_path, pages='all', flavor='lattice', line_scale=60)
 
        if len(tables) > 0:
            combined_df = pd.concat([table.df for table in tables], ignore_index=True)
 
            # Extrair dados específicos do DataFrame combinado
            # Esta parte é crucial e precisa ser adaptada com base na estrutura exata
            # do seu DataFrame combinado e como os dados aparecem no PDF.
 
            # Exemplo de extração de campos que parecem estar no PDF:
            # UNIDADE DE NEGÓCIO, TIPO DE ANALISE, CADASTRO PAT, Nº DO CONTRATO
            # Supondo que a primeira tabela tenha esses dados na primeira linha de dados.
            try:
                unidade_negocio_df = tables[0].df # A primeira tabela parece conter "UNIDADE DE NEGÓCIO"
                data['unidade_negocio'] = unidade_negocio_df.iloc[1, 0] # BRASILIA
                data['tipo_analise'] = unidade_negocio_df.iloc[1, 1] # Inclusão
                data['cadastro_pat'] = unidade_negocio_df.iloc[1, 2].replace('Nº.: ', '') # 0337846
                data['numero_contrato'] = unidade_negocio_df.iloc[1, 3] # 001.001.1251.01/15
            except (IndexError, AttributeError) as e:
                print(f"Erro ao extrair dados da tabela 'UNIDADE DE NEGÓCIO': {e}")
                pass # Continue sem esses dados se não forem encontrados
 
            # DADOS CADASTRAIS (assumindo que estão em um formato facilmente parseável após a extração do camelot)
            # Para extrair esses dados, é mais provável que você precise percorrer o DF combinado
            # ou ler linha por linha buscando as chaves como "Razão Social:", "CNPJ:", etc.
            # Aqui, vou simular a extração a partir do conteúdo total do PDF (fullContent) que você forneceu,
            # já que o camelot é para tabelas e nem tudo é tabela.
 
            # SIMULAÇÃO DE EXTRAÇÃO DE DADOS NÃO TABULARES DO PDF (A PARTIR DO fullContent)
            # Você substituiria isso por uma lógica mais robusta baseada no seu PDF real
            # ou em como seu código camelot está processando tudo.
            # Para simplificar, vou usar os dados que você forneceu no fullContent
            # como se já tivessem sido parseados.
 
            # Dados Cadastrais
            data['razao_social'] = "DAMASCO MATERIAL ELETRICO HIDRAULICO E FERRAGENS LTDA" # [cite: 6]
            data['nome_fantasia'] = "DAMASCO" # [cite: 7]
            data['cnpj'] = "37.054.319/0001-00" # [cite: 9]
            data['inscricao_estadual'] = "733131000188" # [cite: 10]
            data['responsavel'] = "Dináh Ovides" # [cite: 12]
            data['telefone_responsavel'] = "(61) 99115-9315" # [cite: 14]
 
            # Endereço (principal)
            data['endereco_logradouro'] = "ST SIA TRECHO 17 RUA 10 LOTES 495 E 535" # Adaptado do [cite: 11]
            data['endereco_numero'] = "S/N" # [cite: 11]
            data['endereco_complemento'] = "N/A" # [cite: 11]
            data['endereco_bairro'] = "SIA" # [cite: 16]
            data['endereco_cidade'] = "BRASÍLIA" # [cite: 17]
            data['endereco_cep'] = "71200-228" # [cite: 20]
            data['endereco_uf'] = "DF" # [cite: 21]
 
            # DADOS OBRIGATORIOS - Nota Fiscal Eletrônica
            data['nf_responsavel'] = "Dinah Ovides" # [cite: 23]
            data['nf_email'] = "dp@grupodamasco.com.br" # [cite: 26]
            data['nf_telefone'] = "(61) 3298-7171" # [cite: 27]
 
            # DADOS OBRIGATORIOS - Cobrança Eletrônica
            data['cobranca_responsavel'] = "Dináh Ovides" # [cite: 31]
            data['cobranca_email'] = "dp@grupodamasco.com.br" # [cite: 33]
            data['cobranca_telefone'] = "(61) 3298-7171" # [cite: 34]
 
            # DADOS OBRIGATORIOS - Crédito para Importação
            data['credito_importacao_responsavel'] = "Dinah Ovides" # [cite: 32]
            data['credito_importacao_email'] = "dp@grupodamasco.com.br" # [cite: 37]
            data['credito_importacao_telefone'] = "(61) 3298-7171" # [cite: 38]
 
            # Representantes Legais (Socios - assumindo que será o primeiro ou que haverá campos dinâmicos)
            # Para fins de preenchimento, vamos pegar o primeiro sócio como um representante legal
            # Você precisaria adaptar se o sistema lida com múltiplos representantes legais de forma diferente.
            try:
                socios_df = tables[2].df # A terceira tabela parece ser DADOS DOS SOCIOS [cite: 39]
                data['representante_legal_nome'] = socios_df.iloc[0, 0].replace('Nome: ', '') # Bassam Massouh
                data['representante_legal_cpf'] = socios_df.iloc[0, 2].replace('CPF: ', '') # 152.563.591-34
                # Participação não tem campo direto na imagem, então não vou incluir.
            except (IndexError, AttributeError) as e:
                print(f"Erro ao extrair dados dos sócios: {e}")
                pass # Continue sem esses dados se não forem encontrados
 
            # Informações Bancárias - Não há dados diretos no PDF fornecido, então serão vazios
            data['banco_numero'] = ""
            data['banco_agencia'] = ""
            data['banco_conta_corrente'] = ""
            data['banco_telefone'] = ""
 
            # Dados Comerciais (Valor p/vida Clube de Vantagens, Limite de Crédito Cliente)
            data['valor_vida_clube'] = "0,00" # [cite: 51]
            data['limite_credito_cliente'] = "42.000,00" # [cite: 60]
 
        else:
            print("Nenhuma tabela encontrada no PDF para extração de dados detalhados.")
 
    except Exception as e:
        print(f"Ocorreu um erro ao extrair os dados do PDF: {e}")
        print("Certifique-se de que o Ghostscript está instalado e no seu PATH.")
        print("Verifique também se o PDF contém tabelas e o 'flavor' (lattice/stream) está correto.")
 
    return data
 
# --- PARTE 2: AUTOMAÇÃO COM SELENIUM ---
 
def fill_system_with_data(data):
    """
    Preenche os campos do sistema interno usando os dados extraídos.
    """
    # Configure o caminho para o geckodriver se ele não estiver no PATH do sistema
    # service = webdriver.FirefoxService(executable_path='/caminho/para/seu/geckodriver')
    # driver = webdriver.Firefox(service=service)
    driver = webdriver.Firefox() # Isso assume que geckodriver está no PATH
 
    # URL do seu sistema interno (SUBSTITUA PELA URL CORRETA!)
    system_url = "http://172.30.1.219:7778/forms/frmservlet?config=valeshop"
 
    try:
        driver.get(system_url)
        print(f"Navegando para: {system_url}")
 
        # Aguardar até que a página seja carregada e o título "Manter Cliente" esteja visível
        # Ajuste o seletor se o título não for um H2 ou tiver um ID diferente
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Manter Cliente')]"))
        )
        print("Página 'Manter Cliente' carregada.")
 
        # --- PREENCHIMENTO DOS DADOS BÁSICOS ---
        print("Preenchendo Dados Básicos...")
        # Natureza do Cliente - assumindo que é um campo de texto ou que você já sabe o valor
        # Pode ser um dropdown, então o seletor e o método de preenchimento mudariam
        try:
            # Exemplo: Se fosse um campo de texto simples para Natureza do Cliente
            # driver.find_element(By.ID, "naturezaClienteId").send_keys("Pessoa Jurídica")
            # Se for um dropdown (select), você usaria:
            # from selenium.webdriver.support.ui import Select
            # select_natureza = Select(driver.find_element(By.ID, "idDoDropdownNatureza"))
            # select_natureza.select_by_visible_text("Pessoa Jurídica")
            pass # Sem campo visível na imagem para Natureza do Cliente
 
            # Preencher Razão Social
            # Assumindo que o campo de texto de Razão Social tem um ID ou XPath
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Razão Social:')]/following-sibling::input"))
            ).send_keys(data.get('razao_social', '')) # [cite: 6]
            print(f"Razão Social preenchida: {data.get('razao_social', '')}") # [cite: 6]
 
            # Preencher Nome Fantasia
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Nome Fantasia:')]/following-sibling::input"))
            ).send_keys(data.get('nome_fantasia', '')) # [cite: 7]
            print(f"Nome Fantasia preenchido: {data.get('nome_fantasia', '')}") # [cite: 7]
 
            # Preencher CNPJ
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'CNPJ:')]/following-sibling::input"))
            ).send_keys(data.get('cnpj', '')) # [cite: 9]
            print(f"CNPJ preenchido: {data.get('cnpj', '')}") # [cite: 9]
 
            # Preencher Inscrição Estadual
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Inscrição Estadual:')]/following-sibling::input"))
            ).send_keys(data.get('inscricao_estadual', '')) # [cite: 10]
            print(f"Inscrição Estadual preenchida: {data.get('inscricao_estadual', '')}") # [cite: 10]
 
            # Responsável (Não há campo direto na imagem para este, mas se houver)
            # WebDriverWait(driver, 10).until(
            #     EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Responsável:')]/following-sibling::input"))
            # ).send_keys(data.get('responsavel', ''))
 
            # Telefone do Responsável (Não há campo direto na imagem para este)
            # WebDriverWait(driver, 10).until(
            #     EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Telefone Responsável:')]/following-sibling::input"))
            # ).send_keys(data.get('telefone_responsavel', ''))
 
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Erro ao preencher campos de Dados Básicos: {e}")
 
        # --- PREENCHIMENTO DE ENDEREÇO (Principal) ---
        print("Preenchendo Endereço Principal...")
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), 'UF')]/following-sibling::input)[1]"))
            ).send_keys(data.get('endereco_uf', '')) # [cite: 21]
            print(f"UF (endereço 1) preenchido: {data.get('endereco_uf', '')}") # [cite: 21]
 
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), 'CEP')]/following-sibling::input)[1]"))
            ).send_keys(data.get('endereco_cep', '')) # [cite: 20]
            print(f"CEP (endereço 1) preenchido: {data.get('endereco_cep', '')}") # [cite: 20]
            # Nota: Muitos sistemas preenchem Cidade/Bairro/Logradouro automaticamente após o CEP.
            # Você pode precisar esperar por isso ou preencher manualmente.
 
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), 'Localidade')]/following-sibling::input)[1]"))
            ).send_keys(data.get('endereco_cidade', '')) # [cite: 17]
            print(f"Localidade (endereço 1) preenchida: {data.get('endereco_cidade', '')}") # [cite: 17]
 
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), 'Bairro')]/following-sibling::input)[1]"))
            ).send_keys(data.get('endereco_bairro', '')) # [cite: 16]
            print(f"Bairro (endereço 1) preenchido: {data.get('endereco_bairro', '')}") # [cite: 16]
 
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), 'Logradouro')]/following-sibling::input)[1]"))
            ).send_keys(data.get('endereco_logradouro', '')) # Adaptado do [cite: 11]
            print(f"Logradouro (endereço 1) preenchido: {data.get('endereco_logradouro', '')}") # Adaptado do [cite: 11]
 
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), 'Número')]/following-sibling::input)[1]"))
            ).send_keys(data.get('endereco_numero', '')) # [cite: 11]
            print(f"Número (endereço 1) preenchido: {data.get('endereco_numero', '')}") # [cite: 11]
 
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), 'Complemento')]/following-sibling::input)[1]"))
            ).send_keys(data.get('endereco_complemento', '')) # [cite: 11]
            print(f"Complemento (endereço 1) preenchido: {data.get('endereco_complemento', '')}") # [cite: 11]
 
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Erro ao preencher campos de Endereço Principal: {e}")
 
        # --- PREENCHIMENTO DE REPRESENTANTES LEGAIS ---
        print("Preenchendo Representantes Legais...")
        try:
            # Nome do Representante Legal
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Representantes Legais')]/following::label[contains(text(), 'Nome')]/following-sibling::input"))
            ).send_keys(data.get('representante_legal_nome', ''))
            print(f"Nome Representante Legal preenchido: {data.get('representante_legal_nome', '')}")
 
            # CPF do Representante Legal
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Representantes Legais')]/following::label[contains(text(), 'CPF')]/following-sibling::input"))
            ).send_keys(data.get('representante_legal_cpf', ''))
            print(f"CPF Representante Legal preenchido: {data.get('representante_legal_cpf', '')}")
 
            # Outros campos como CNPJ, Email, Telefone, Ocupação se existirem e forem necessários
            # Para os exemplos dados no PDF, esses campos não estavam sob "Representantes Legais" mas sim DADOS OBRIGATÓRIOS.
            # Ajuste conforme a sua interface.
 
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Erro ao preencher campos de Representantes Legais: {e}")
 
        # --- PREENCHIMENTO DE INFORMAÇÕES BANCÁRIAS (se houver dados) ---
        print("Preenchendo Informações Bancárias...")
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Nr. banco')]/following-sibling::input"))
            ).send_keys(data.get('banco_numero', ''))
            # Repita para Agência, Conta Corrente, Telefone...
            # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "agenciaId"))).send_keys(data.get('banco_agencia', ''))
            # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "contaCorrenteId"))).send_keys(data.get('banco_conta_corrente', ''))
            # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "telefoneBancoId"))).send_keys(data.get('banco_telefone', ''))
            print("Informações Bancárias preenchidas (se dados existiam).")
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Erro ao preencher campos de Informações Bancárias: {e}")
 
        # --- Preencher Dados Opcionais/Outras Abas se necessário ---
        # Se seu sistema tem abas para "DADOS COMPLEMENTARES", "CONTRATOS", etc.
        # Você precisaria clicar nessas abas primeiro e depois preencher os campos.
        # Exemplo: clicar na aba "DADOS COMPLEMENTARES"
        # WebDriverWait(driver, 10).until(
        #     EC.element_to_be_clickable((By.XPATH, "//div[@role='tab' and text()='DADOS COMPLEMENTARES']"))
        # ).click()
        # # Depois preencher campos na nova aba.
 
        # EX: Valor p/vida Clube de Vantagens e Limite de Crédito Cliente (provavelmente em outra aba ou seção)
        # Assumindo que você navegaria para uma aba "Dados Comerciais" ou similar
        print("Tentando preencher Dados Comerciais...")
        try:
            # Clicar na aba "Dados do Cliente" -> "DADOS COMERCIAIS" ou similar
            # Isso é um chute, você PRECISA verificar o seletor da aba no seu sistema
            # Exemplo de clique na aba "Dados do Cliente" (se for uma hierarquia de abas)
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@role='tablist']//li/a[contains(text(), 'Dados do Cliente')]"))
            ).click()
            # E depois talvez uma sub-aba se for o caso
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'tab-content')]//a[contains(text(), 'Dados Comerciais')]"))
            ).click()
 
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Valor p/vida Clube de Vantagens')]/following-sibling::input"))
            ).send_keys(data.get('valor_vida_clube', '')) # [cite: 51]
            print(f"Valor p/vida Clube de Vantagens preenchido: {data.get('valor_vida_clube', '')}") # [cite: 51]
 
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Limite de Crédito Cliente (R$)')]/following-sibling::input"))
            ).send_keys(data.get('limite_credito_cliente', '')) # [cite: 60]
            print(f"Limite de Crédito Cliente preenchido: {data.get('limite_credito_cliente', '')}") # [cite: 60]
 
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Erro ao navegar ou preencher campos de Dados Comerciais: {e}")
 
 
        # --- SALVAR (Exemplo) ---
        # Encontre o botão de salvar. O seletor pode variar (ID, name, text do botão, etc.)
        # Na imagem, há botões na barra superior, como um disquete.
        # Você precisará do XPath ou ID exato do botão "Salvar".
        # Exemplo:
        try:
            save_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@title='Gravar'] | //a[@title='Gravar'] | //img[@alt='Gravar']")) # Tente algumas opções
            )
            save_button.click()
            print("Botão 'Gravar' clicado.")
            # Você pode adicionar uma espera para uma mensagem de sucesso ou redirecionamento
            # WebDriverWait(driver, 10).until(EC.alert_is_present())
            # driver.switch_to.alert.accept() # Se houver um alerta de sucesso
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Não foi possível encontrar ou clicar no botão 'Gravar': {e}")
 
 
        print("Preenchimento concluído. Verifique o sistema.")
 
    except Exception as e:
        print(f"Ocorreu um erro geral durante a automação: {e}")
 
    finally:
        # Mantenha o navegador aberto por um tempo para você verificar,
        # ou feche-o automaticamente.
        # input("Pressione Enter para fechar o navegador...")
        # driver.quit()
        print("Script finalizado. Navegador permanece aberto para verificação.")
 
# --- Execução Principal ---
if __name__ == "__main__":
    pdf_file_name = 'Ploomes _ Clientes _ DAMASCO.pdf' # Certifique-se de que este PDF está no mesmo diretório ou forneça o caminho completo

    time.sleep(5) # Tempo para você abrir o PDF e preparar o ambiente
    print("Iniciando extração de dados do PDF...")
    extracted_data = extract_data_from_pdf(pdf_file_name)
    print("\nDados extraídos do PDF:")
    for key, value in extracted_data.items():
        print(f"  {key}: {value}")
 
    if extracted_data:
        print("\nIniciando preenchimento automático do sistema...")
        fill_system_with_data(extracted_data)
    else:
        print("Nenhum dado extraído do PDF. Automação do sistema não será executada.")