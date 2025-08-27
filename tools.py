import os, re, zipfile, tempfile
import pandas as pd
import numpy as np
from datetime import datetime
from settings import DATA_DIR, RESULT_FILE, VALIDACAO_FILE

# cache global
_cached_bases = {}

# ===== Helpers =====
def salvar_excel_formatado(df, path):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Resultado", index=False)

        workbook  = writer.book
        worksheet = writer.sheets["Resultado"]

        # ===== FORMATAÇÕES =====
        header_fmt = workbook.add_format({
            "bold": True,
            "bg_color": "#000000",   # fundo preto
            "font_color": "#FFFFFF", # texto branco
            "align": "center",
            "valign": "vcenter"
        })
        date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy"})
        month_fmt = workbook.add_format({"num_format": "mm/yyyy"})
        money_fmt = workbook.add_format({"num_format": "R$ #,##0.00"})

        # Cabeçalho
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # Largura padrão
        worksheet.set_column(0, len(df.columns)-1, 15)

        # ===== FORMATO ESPECÍFICO POR COLUNA =====
        col_idx = {name: i for i, name in enumerate(df.columns)}

        if "Admissão" in col_idx:
            worksheet.set_column(col_idx["Admissão"], col_idx["Admissão"], 15, date_fmt)

        if "Competência" in col_idx:
            worksheet.set_column(col_idx["Competência"], col_idx["Competência"], 12, month_fmt)

        for col in ["VALOR DIÁRIO VR", "TOTAL", "Custo empresa", "Desconto profissional"]:
            if col in col_idx:
                worksheet.set_column(col_idx[col], col_idx[col], 18, money_fmt)

def parse_date(x):
    if pd.isna(x):
        return None
    if isinstance(x, (pd.Timestamp, datetime)):
        return pd.Timestamp(x).normalize()
    try:
        return pd.to_datetime(str(x), dayfirst=True, errors="coerce").normalize()
    except Exception:
        return None

def normalize_text(s):
    if pd.isna(s): return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def extract_uf_from_sindicato(s):
    UFs = ["SP","RJ","RS","PR","SC","MG","ES","BA","PE","CE","DF","GO","MT","MS",
           "PA","AM","RN","PB","AL","PI","RO","RR","AP","AC","TO"]
    s = normalize_text(s)
    for uf in UFs:
        if re.search(r"\b"+uf+r"\b", s):
            return uf
    return None

def uf_to_estado(uf):
    mapa = {"SP":"São Paulo","RJ":"Rio de Janeiro","RS":"Rio Grande do Sul","PR":"Paraná"}
    return mapa.get(uf)

# ===== Tools =====
def consolidar_bases(zip_path: str = None):
    """
    Carrega todas as planilhas no cache.
    - Se zip_path for informado, extrai antes de carregar.
    - Retorna apenas um resumo (qtd de registros por base).
    """
    global _cached_bases

    if zip_path and zip_path.endswith(".zip"):
        extract_dir = tempfile.mkdtemp(prefix="vr_bases_")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        base_dir = extract_dir
    else:
        base_dir = DATA_DIR  # pasta já extraída

    _cached_bases = {
        "ativos": pd.read_excel(os.path.join(base_dir, "ATIVOS.xlsx")),
        "admis": pd.read_excel(os.path.join(base_dir, "ADMISSÃO ABRIL.xlsx")),
        "deslig": pd.read_excel(os.path.join(base_dir, "DESLIGADOS.xlsx")),
        "ferias": pd.read_excel(os.path.join(base_dir, "FÉRIAS.xlsx")),
        "afast": pd.read_excel(os.path.join(base_dir, "AFASTAMENTOS.xlsx")),
        "aprend": pd.read_excel(os.path.join(base_dir, "APRENDIZ.xlsx")),
        "estagio": pd.read_excel(os.path.join(base_dir, "ESTÁGIO.xlsx")),
        "exterior": pd.read_excel(os.path.join(base_dir, "EXTERIOR.xlsx")),
        "dias_uteis": pd.read_excel(os.path.join(base_dir, "Base dias uteis.xlsx")),
        "sind_valor": pd.read_excel(os.path.join(base_dir, "Base sindicato x valor.xlsx"))
    }

    resumo = {nome: len(df) for nome, df in _cached_bases.items()}
    return f"Bases carregadas com sucesso. Resumo: {resumo}"

def aplicar_exclusoes(df):
    """Remove aprendizes, estagiários, afastados, expatriados e diretores."""
    bases = _cached_bases
    set_aprend = set(bases["aprend"]["MATRICULA"].dropna().astype(int))
    set_estagio = set(bases["estagio"]["MATRICULA"].dropna().astype(int))
    set_afast = set(bases["afast"]["MATRICULA"].dropna().astype(int))
    set_exterior = set(bases["exterior"]["Cadastro"].dropna().astype(int)) if "Cadastro" in bases["exterior"].columns else set()

    df["is_diretor"] = df["TITULO DO CARGO"].fillna("").str.contains("DIRETOR", case=False)
    excluidos = set_aprend | set_estagio | set_afast | set_exterior
    return df[~df["MATRICULA"].isin(excluidos) & ~df["is_diretor"]].copy()

def calcular_vr():
    """Aplica regra de cálculo do VR para maio/2025 usando _cached_bases."""
    bases = _cached_bases
    ativos = bases["ativos"].rename(columns=normalize_text)

    df = ativos[["MATRICULA","TITULO DO CARGO","Sindicato"]].copy()
    df["MATRICULA"] = df["MATRICULA"].astype(int)
    df["Sindicato"] = df["Sindicato"].apply(normalize_text)

    # Admissões
    admis = bases["admis"].rename(columns=normalize_text)
    if "Admissão" in admis.columns:
        admis_slim = admis[["MATRICULA","Admissão"]].copy()
        admis_slim["MATRICULA"] = admis_slim["MATRICULA"].astype(int)
        admis_slim["Admissão"] = admis_slim["Admissão"].apply(parse_date)
        df = df.merge(admis_slim, on="MATRICULA", how="left")

    # Desligados
    deslig = bases["deslig"].rename(columns=normalize_text)
    deslig["MATRICULA"] = deslig["MATRICULA"].astype(int)
    deslig["DATA DEMISSÃO"] = deslig["DATA DEMISSÃO"].apply(parse_date)
    deslig["COMUNICADO DE DESLIGAMENTO"] = deslig["COMUNICADO DE DESLIGAMENTO"].apply(parse_date)
    df = df.merge(deslig, on="MATRICULA", how="left")

    # Férias
    ferias = bases["ferias"].rename(columns=normalize_text)
    ferias["MATRICULA"] = ferias["MATRICULA"].astype(int)
    ferias["DIAS DE FÉRIAS"] = pd.to_numeric(ferias["DIAS DE FÉRIAS"], errors="coerce").fillna(0).astype(int)
    df = df.merge(ferias[["MATRICULA","DIAS DE FÉRIAS"]], on="MATRICULA", how="left")

    # Exclusões
    df = aplicar_exclusoes(df)

    # Dias úteis
    dias_uteis = bases["dias_uteis"].rename(columns=normalize_text)
    dias_uteis.columns = ["Sindicato","Dias_Uteis"]
    dias_uteis["Sindicato"] = dias_uteis["Sindicato"].apply(normalize_text)
    df = df.merge(dias_uteis, on="Sindicato", how="left")

    # Valor sindicato
    sind_valor = bases["sind_valor"].rename(columns=normalize_text)
    col_estado = [c for c in sind_valor.columns if "ESTADO" in c.upper()][0]
    sind_valor = sind_valor[[col_estado,"VALOR"]].copy()
    sind_valor.columns = ["Estado","Valor_Diario"]
    sind_valor["Estado"] = sind_valor["Estado"].apply(normalize_text)
    df["UF"] = df["Sindicato"].apply(extract_uf_from_sindicato)
    df["Estado"] = df["UF"].apply(uf_to_estado)
    df = df.merge(sind_valor, on="Estado", how="left")

    # Competência
    period_start, period_end = pd.Timestamp("2025-04-15"), pd.Timestamp("2025-05-16")
    bd_total = np.busday_count(period_start.date(), period_end.date())

    def compute_days(row):
        adm, dem, com = row.get("Admissão"), row.get("DATA DEMISSÃO"), row.get("COMUNICADO DE DESLIGAMENTO")
        dias_base = row.get("Dias_Uteis",0)
        if pd.notna(com) and com <= pd.Timestamp("2025-05-15"): return 0
        start = max(period_start, adm) if pd.notna(adm) else period_start
        end = min(period_end, dem+pd.Timedelta(days=1)) if pd.notna(dem) else period_end
        naive = np.busday_count(start.date(), end.date()) if end>start else 0
        return int(round((naive/bd_total)*dias_base)) if bd_total>0 else 0

    df["Dias_Trabalhados_Base"] = df.apply(compute_days, axis=1)
    df["DIAS DE FÉRIAS"] = df["DIAS DE FÉRIAS"].fillna(0)
    df["Dias"] = (df["Dias_Trabalhados_Base"] - df["DIAS DE FÉRIAS"]).clip(lower=0)

    df["Valor_Diario"] = df["Valor_Diario"].fillna(35.0)
    df["TOTAL"] = (df["Dias"]*df["Valor_Diario"]).round(2)
    df["Custo empresa"] = (df["TOTAL"]*0.8).round(2)
    df["Desconto profissional"] = (df["TOTAL"]*0.2).round(2)

    out = pd.DataFrame({
        "Matricula": df["MATRICULA"],
        "Admissão": df["Admissão"].dt.date if pd.api.types.is_datetime64_any_dtype(df["Admissão"]) else df["Admissão"],
        "Sindicato do Colaborador": df["Sindicato"],
        "Competência": pd.Timestamp("2025-05-01"),
        "Dias": df["Dias"],
        "VALOR DIÁRIO VR": df["Valor_Diario"],
        "TOTAL": df["TOTAL"],
        "Custo empresa": df["Custo empresa"],
        "Desconto profissional": df["Desconto profissional"],
        "OBS GERAL": ""
    })
    salvar_excel_formatado(out, RESULT_FILE)
    return f"Planilha gerada em {RESULT_FILE}"

def validar_resultado():
    """Gera relatório explicativo por matrícula usando _cached_bases."""
    bases = _cached_bases
    df = bases["ativos"].rename(columns=normalize_text)
    df = df[["MATRICULA","TITULO DO CARGO"]].copy()
    df["MATRICULA"] = df["MATRICULA"].astype(int)

    deslig = bases["deslig"].rename(columns=normalize_text)
    deslig["MATRICULA"] = deslig["MATRICULA"].astype(int)
    deslig["DATA DEMISSÃO"] = deslig["DATA DEMISSÃO"].apply(parse_date)
    deslig["COMUNICADO DE DESLIGAMENTO"] = deslig["COMUNICADO DE DESLIGAMENTO"].apply(parse_date)
    df = df.merge(deslig, on="MATRICULA", how="left")

    ferias = bases["ferias"].rename(columns=normalize_text)
    ferias["MATRICULA"] = ferias["MATRICULA"].astype(int)
    df = df.merge(ferias[["MATRICULA","DIAS DE FÉRIAS"]], on="MATRICULA", how="left")

    excl_aprend = set(bases["aprend"]["MATRICULA"].dropna().astype(int))
    excl_estagio = set(bases["estagio"]["MATRICULA"].dropna().astype(int))
    excl_afast = set(bases["afast"]["MATRICULA"].dropna().astype(int))
    excl_exterior = set(bases["exterior"]["Cadastro"].dropna().astype(int)) if "Cadastro" in bases["exterior"].columns else set()

    report = []
    for _, row in df.iterrows():
        m = row["MATRICULA"]
        motivos = []
        if m in excl_aprend: motivos.append("Aprendiz")
        if m in excl_estagio: motivos.append("Estagiário")
        if m in excl_afast: motivos.append("Afastamento")
        if m in excl_exterior: motivos.append("Exterior")
        if "DIRETOR" in str(row["TITULO DO CARGO"]).upper(): motivos.append("Diretor")
        if pd.notna(row.get("COMUNICADO DE DESLIGAMENTO")) and row["COMUNICADO DE DESLIGAMENTO"] <= pd.Timestamp("2025-05-15"):
            motivos.append("Comunicado até 15/05 → zerado")
        if row.get("DIAS DE FÉRIAS",0) > 0:
            motivos.append(f"Férias {row['DIAS DE FÉRIAS']} dias")
        if not motivos: motivos.append("Sem ajustes")
        report.append({"Matricula": m, "Motivos": "; ".join(motivos)})

    report_df = pd.DataFrame(report)
    report_df.to_excel(VALIDACAO_FILE, index=False)
    return f"Relatório de validação gerado em {VALIDACAO_FILE}"