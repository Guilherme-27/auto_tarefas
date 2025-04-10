import pandas as pd

def ler_csv_sed(caminho):
    """
    Lê e limpa um CSV exportado da SED, ignorando as duas primeiras linhas de metadados.
    Verifica se as colunas obrigatórias estão presentes.
    """
    try:
        df = pd.read_csv(caminho, sep=";", skiprows=2, encoding="utf-8")
    except Exception as e:
        raise ValueError(f"Erro ao ler o CSV: {e}")

    colunas_esperadas = ["Status", "Nome do Aluno"]
    for col in colunas_esperadas:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória '{col}' não encontrada no CSV.")

    return df


def converter_nota_para_escala_10(df):
    """
    Converte a coluna 'Status' de porcentagem para nota de 0 a 10.
    Cria a nova coluna 'Nota (%)' com nota em formato decimal.
    """
    df["Nota (%)"] = (
        df["Status"]
        .astype(str)
        .str.replace("%", "")
        .str.replace(",", ".")
        .astype(float)
        / 10
    )
    return df


def remover_duplicatas_maior_nota(df):
    """
    Remove nomes duplicados, mantendo apenas o registro com a maior nota.
    """
    df = df.sort_values("Nota (%)", ascending=False)
    df = df.drop_duplicates(subset="Nome do Aluno", keep="first")
    return df


def salvar_csv_limpo(df, caminho_saida="dados_limpos.csv"):
    """
    Exporta apenas as colunas essenciais ('Nota (%)' e 'Nome do Aluno') para um novo CSV.
    """
    df[["Nota (%)", "Nome do Aluno"]].to_csv(caminho_saida, sep=";", index=False)
    print(f"✅ CSV limpo salvo como: {caminho_saida}")