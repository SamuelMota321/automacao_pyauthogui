import camelot
import pandas as pd

# O caminho para o seu arquivo PDF
# ou forneça o caminho completo para o PDF.
pdf_file = 'Ploomes _ Clientes _ DAMASCO.pdf'

print(f"Tentando extrair tabelas do PDF: {pdf_file}")

try:
    # Lê as tabelas do PDF.
    # pages='all' tenta extrair tabelas de todas as páginas.
    # flavor='lattice' é para tabelas com bordas (linhas e colunas visíveis).
    # flavor='stream' é para tabelas sem bordas (separadas por espaços).
    # Se você não souber qual usar, pode tentar ambos ou começar com 'lattice'.
    tables = camelot.read_pdf(pdf_file, pages='all', flavor='lattice', line_scale=60) 

    print(f"Número de tabelas encontradas: {len(tables)}")
    if len(tables) > 0:
        # Imprimir cada tabela individualmente
        for i, table in enumerate(tables):
            print(f"\nDados da tabela {i+1}:")
            print(table.df) # Imprime o DataFrame de cada tabela
            table.to_csv(f'tabela_{i+1}.csv', index=False)

        # Combinar todas as tabelas em um único DataFrame
        # pd.concat() é a função correta para isso.
        # ignore_index=True redefine o índice do DataFrame combinado.
        combined_df = pd.concat([table.df for table in tables], ignore_index=True)

        print("\nDados de TODAS as tabelas combinadas:")
        print(combined_df)

        # Salvar o DataFrame combinado em um arquivo CSV
        combined_df.to_csv('tabela_extraida.csv', index=False)
        print("\nTabelas combinadas salvas em 'tabela_extraida.csv'")

    else:
        print("Nenhuma tabela encontrada no PDF.")

except Exception as e:
    print(f"Ocorreu um erro ao extrair as tabelas: {e}")
    print("Certifique-se de que o Ghostscript está instalado e no seu PATH.")
    print("Verifique também se o PDF contém tabelas e o 'flavor' (lattice/stream) está correto.")

print("Fim do script.")
