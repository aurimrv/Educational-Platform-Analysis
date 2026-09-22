import os
import re
import pandas as pd

def extrair_numeros_auto_subpastas():
    # Detecta a pasta onde o script está sendo executado
    pasta_raiz = os.getcwd()
    
    print(f"Iniciando varredura na pasta atual e subdiretórios:\n{pasta_raiz}\n")
    
    padrao = re.compile(r"S_(\d+)_", re.IGNORECASE)
    dados = []

    # Percorre a pasta atual e todas as subpastas
    for raiz, _dirs, arquivos in os.walk(pasta_raiz):
        subpasta_relativa = os.path.relpath(raiz, pasta_raiz)
        if subpasta_relativa == ".":
            subpasta_relativa = "(Pasta Raiz)"

        for arquivo in arquivos:
            match = padrao.search(arquivo)
            if match:
                numero = int(match.group(1))
                dados.append({
                    "Número": numero,
                    "Subpasta": subpasta_relativa,
                    "Nome do Arquivo": arquivo,
                    "Caminho Completo": os.path.join(raiz, arquivo)
                })

    if not dados:
        print("Nenhum arquivo correspondente ao padrão foi encontrado.")
        return

    # Estrutura os dados e ordena numericamente
    df = pd.DataFrame(dados)
    df = df.sort_values(by="Número")

    excel_saida = "numeros_beecrowd.xlsx"
    csv_saida = "numeros_beecrowd.csv"

    df.to_excel(excel_saida, index=False)
    df.to_csv(csv_saida, index=False, encoding="utf-8-sig")

    print("Varredura concluída com sucesso!")
    print(f"Total de arquivos encontrados: {len(df)}")
    print(f"Planilha gerada: {os.path.abspath(excel_saida)}")

if __name__ == "__main__":
    extrair_numeros_auto_subpastas()