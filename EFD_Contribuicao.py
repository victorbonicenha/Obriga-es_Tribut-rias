import calendar
import os
import re
import unicodedata
from datetime import datetime
import pandas as pd


# Caminho do Excel (na mesma pasta deste arquivo)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_EXCEL = os.path.join(BASE_DIR, "anexo-ade-corat-no-2-de-27-01-26.xlsx")


MESES = {
    "janeiro": 1,
    "fevereiro": 2,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}

# Ajuste aqui se quiser match mais específico pela frase completa
DESCRICAO_EFD_CONTRIB_ALVO = "EFD-Contribuições"


def normalizar_texto(valor):
    texto = str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFD", texto)
    return "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")


def adicionar_meses(ano, mes, quantidade):
    total_meses = ano * 12 + (mes - 1) + quantidade
    novo_ano, novo_mes_zero = divmod(total_meses, 12)
    return novo_ano, novo_mes_zero + 1


def obter_nome_planilha_declaracoes(xlsx):
    for nome in xlsx.sheet_names:
        if normalizar_texto(nome) == "declaracoes":
            return nome
    raise ValueError("Planilha 'Declarações' não encontrada.")


def obter_colunas(df):
    colunas_norm = {normalizar_texto(col): col for col in df.columns}

    chave_nome = "declaracoes, demonstrativos e documentos"
    chave_periodo = "periodo de referencia"
    chave_prazo = "prazo de apresentacao"

    for chave in (chave_nome, chave_periodo, chave_prazo):
        if chave not in colunas_norm:
            raise ValueError(f"Coluna obrigatória não encontrada: {chave}")

    return colunas_norm[chave_nome], colunas_norm[chave_periodo], colunas_norm[chave_prazo]


def extrair_mes_ano_referencia(periodo_texto):
    periodo_norm = normalizar_texto(periodo_texto)

    padrao = (
        r"(janeiro|fevereiro|marco|abril|maio|junho|julho|agosto|setembro|"
        r"outubro|novembro|dezembro)\s*/\s*(\d{4})"
    )
    encontrados = re.findall(padrao, periodo_norm)
    if encontrados:
        mes_nome, ano_txt = encontrados[-1]
        return int(ano_txt), MESES[mes_nome]

    if "ano-calendario" in periodo_norm:
        m = re.search(r"\d{4}", periodo_norm)
        if m:
            return int(m.group()) + 1, 1

    return None, None


def converter_prazo_dia(prazo_valor):
    nums = re.findall(r"\d+", str(prazo_valor))
    return int(nums[0]) if nums else None


def calcular_vencimento(periodo, prazo, meses_dinamica=2):
    ano_base, mes_base = extrair_mes_ano_referencia(periodo)
    if not ano_base or not mes_base:
        return None

    dia = converter_prazo_dia(prazo)
    if not dia:
        return None

    ano_venc, mes_venc = adicionar_meses(ano_base, mes_base, meses_dinamica)
    ultimo_dia = calendar.monthrange(ano_venc, mes_venc)[1]
    dia_final = min(dia, ultimo_dia)

    return datetime(ano_venc, mes_venc, dia_final).date()


def eh_efd_contribuicoes(descricao):
    """
    Match flexível para EFD-Contribuições.
    Aceita variações com hífen, ponto, espaço e acentuação.
    """
    desc_norm = normalizar_texto(descricao)
    desc_compact = re.sub(r"[\s\-\.\u2013\u2014]+", "", desc_norm)  # remove espaços/pontuação comum

    alvos = [
        "efdcontribuicoes",
        "efdcontribuicao",  # se vier singular em algum caso
    ]
    return any(alvo in desc_compact for alvo in alvos)


def obter_dados_efd_contribuicoes_por_excel(caminho_excel=ARQUIVO_EXCEL):
    """
    Retorna dict com:
    {
      "descricao": <descricao completa>,
      "periodo": <periodo>,
      "prazo": <prazo>,
      "vencimento": <datetime.date>
    }
    ou None se não encontrar.
    """
    xlsx = pd.ExcelFile(caminho_excel)
    nome_aba = obter_nome_planilha_declaracoes(xlsx)
    df = pd.read_excel(xlsx, sheet_name=nome_aba)

    col_desc, col_periodo, col_prazo = obter_colunas(df)
    candidatos = []

    for _, row in df.iterrows():
        descricao = str(row.get(col_desc) or "").strip()
        periodo = row.get(col_periodo)
        prazo = row.get(col_prazo)

        if pd.isna(periodo) or pd.isna(prazo):
            continue

        if not eh_efd_contribuicoes(descricao):
            continue

        vencimento = calcular_vencimento(periodo, prazo, meses_dinamica=2)  # EFD-Contribuições = +2 meses
        if not vencimento:
            continue

        candidatos.append(
            {
                "descricao": descricao,   # descrição completa da planilha
                "periodo": str(periodo),
                "prazo": str(prazo),
                "vencimento": vencimento,
            }
        )

    if not candidatos:
        return None

    # mantém o vencimento mais próximo
    candidatos.sort(key=lambda x: x["vencimento"])
    return candidatos[0]


def obter_vencimento_efd_contribuicoes_por_excel(caminho_excel=ARQUIVO_EXCEL):
    """
    Retorna somente o vencimento (datetime.date) ou None.
    Útil para importar no script da Agenda.
    """
    dado = obter_dados_efd_contribuicoes_por_excel(caminho_excel)
    return dado["vencimento"] if dado else None


if __name__ == "__main__":
    dado = obter_dados_efd_contribuicoes_por_excel(ARQUIVO_EXCEL)

    if not dado:
        print("EFD-Contribuições não encontrada no Excel.")
    else:
        print("Descrição completa:", dado["descricao"])
        print("Período:", dado["periodo"])
        print("Prazo:", dado["prazo"])
        print("Vencimento EFD-Contribuições:", dado["vencimento"])
      
