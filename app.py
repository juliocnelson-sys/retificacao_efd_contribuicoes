"""
EFD Contribuições – Retificação
Aplicativo Streamlit para uso corporativo
"""

import io
import re
import zipfile
from collections import defaultdict
from pathlib import Path

import openpyxl
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# Configuração da página
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="EFD Contribuições – Retificação",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    .main .block-container { padding-top: 2rem; max-width: 960px; }
    .stAlert { border-radius: 8px; }
    div[data-testid="stMetric"] { background: #f8f9fa; padding: 1rem; border-radius: 8px; border: 1px solid #e9ecef; }
    .status-ok   { color: #0F6E56; font-weight: 500; }
    .status-info { color: #185FA5; }
    .status-warn { color: #854F0B; font-weight: 500; }
    .status-err  { color: #A32D2D; font-weight: 500; }
    .log-box { background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px;
               padding: 1rem; font-family: monospace; font-size: 13px;
               max-height: 340px; overflow-y: auto; line-height: 1.8; }
    .rule-box { background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 0.75rem 1rem; margin-bottom: 8px; }
    .rule-title { font-weight: 600; font-size: 14px; margin-bottom: 4px; }
    .rule-desc  { font-size: 13px; color: #555; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Funções auxiliares de formato SPED
# ─────────────────────────────────────────────────────────────────────────────
ENCODINGS_TO_TRY = ["latin-1", "windows-1252", "utf-8", "utf-8-sig"]

def detect_encoding(data: bytes) -> str:
    """Tenta decodificar com múltiplos encodings e retorna o primeiro que funcionar."""
    for enc in ENCODINGS_TO_TRY:
        try:
            data.decode(enc)
            return enc
        except (UnicodeDecodeError, LookupError):
            continue
    return "latin-1"  # fallback seguro

def parse_pipe(line: str) -> list:
    return line.strip().strip("|").split("|")

def to_pipe(fields: list) -> str:
    return "|" + "|".join(str(f) for f in fields) + "|\n"

def fmt_valor(v) -> str:
    if v is None or str(v).strip() == "":
        return "0,00"
    s = str(v).strip().replace(".", ",")
    if "," not in s:
        s += ",00"
    return s

def soma_valores(*vals) -> float:
    total = 0.0
    for v in vals:
        s = str(v).strip().replace(".", "").replace(",", ".")
        try:
            total += float(s)
        except ValueError:
            pass
    return total

def fmt_br(f: float) -> str:
    return f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def gc(row: dict, *keys) -> str:
    for k in keys:
        if k in row and str(row[k]).strip() not in ("", "None"):
            return str(row[k]).strip()
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# Leitura da planilha
# ─────────────────────────────────────────────────────────────────────────────
def sheet_to_dicts(ws) -> list:
    headers = []
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if not headers:
            headers = [str(c).strip() if c else f"COL{j}" for j, c in enumerate(row)]
            continue
        if all(c is None for c in row):
            continue
        rows.append({headers[j]: (row[j] if row[j] is not None else "")
                     for j in range(len(headers))})
    return rows

@st.cache_data(show_spinner=False)
def load_planilha(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    data = {}
    for name in wb.sheetnames:
        if name in ("INSTRUÇÕES", "LISTAS"):
            continue
        data[name] = sheet_to_dicts(wb[name])
    return data


# ─────────────────────────────────────────────────────────────────────────────
# Parse do TXT
# ─────────────────────────────────────────────────────────────────────────────
def parse_txt(lines: list) -> dict:
    result = {"cnpj_0140": {}, "cod_0200": set(), "bloco_0_end": None}
    for idx, line in enumerate(lines):
        ls = line.strip()
        if not ls.startswith("|"):
            continue
        f = parse_pipe(ls)
        reg = f[0]
        if reg == "0140":
            cnpj = f[3] if len(f) > 3 else ""
            result["cnpj_0140"][cnpj] = f
        elif reg == "0200":
            result["cod_0200"].add(f[1] if len(f) > 1 else "")
        elif reg == "0990":
            result["bloco_0_end"] = idx
    return result

def get_periodo(lines: list) -> str | None:
    for line in lines:
        if line.strip().startswith("|0000|"):
            f = parse_pipe(line.strip())
            if len(f) > 5 and f[5]:
                return f[5][2:]  # MMAAAA
    return None

def filter_periodo(rows: list, periodo: str | None) -> list:
    if not periodo:
        return rows
    result = []
    for r in rows:
        p = str(r.get("PERÍODO\n(MMAAAA)", r.get("PERÍODO", r.get("PERIODO", "")))).strip()
        if not p or p == periodo:
            result.append(r)
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Construtores de registros
# ─────────────────────────────────────────────────────────────────────────────
def build_0140(row):
    cnpj = gc(row, "CNPJ_ESTAB", "CNPJ")
    return to_pipe(["0140", cnpj, gc(row, "NOME / RAZÃO SOCIAL"), cnpj, gc(row, "UF"), "", gc(row, "COD_MUN"), "", ""])

def build_0150(r):
    return to_pipe(["0150", gc(r,"COD_PART"), gc(r,"NOME"), gc(r,"COD_PAIS") or "1058",
                    gc(r,"CNPJ"), gc(r,"CPF"), gc(r,"IE"), gc(r,"COD_MUN"),
                    gc(r,"SUFRAMA"), gc(r,"END"), gc(r,"NUM"), gc(r,"COMPL"), gc(r,"BAIRRO")])

def build_0190(r):
    return to_pipe(["0190", gc(r,"UNID"), gc(r,"DESCR")])

def build_0200(r):
    return to_pipe(["0200", gc(r,"COD_ITEM"), gc(r,"DESCR_ITEM"), gc(r,"COD_BARRA"), "",
                    gc(r,"UNID_INV"), gc(r,"TIPO_ITEM"), gc(r,"COD_NCM"),
                    gc(r,"EX_IPI"), gc(r,"COD_GEN"), gc(r,"COD_LST"), fmt_valor(r.get("ALIQ_ICMS",""))])

def build_0400(r):
    return to_pipe(["0400", gc(r,"COD_NAT_REC"), gc(r,"DESCR_NAT_REC")])

def build_0500(r):
    return to_pipe(["0500", gc(r,"DT_ALT"), gc(r,"COD_CTA"), gc(r,"NÍVEL"),
                    gc(r,"IND_CTA"), gc(r,"NOME_CTA"), gc(r,"COD_CTA_SUP"), gc(r,"COD_CTA_REF"), ""])

def build_a_block(grupo: list) -> list:
    r0 = grupo[0]
    vl_doc    = soma_valores(*[r.get("VL_ITEM", 0) for r in grupo])
    vl_desc   = soma_valores(*[r.get("VL_DESC_ITEM", 0) for r in grupo])
    vl_bc_pis = soma_valores(*[r.get("VL_BC_PIS_ITEM", 0) for r in grupo])
    vl_pis    = soma_valores(*[r.get("VL_PIS_ITEM", 0) for r in grupo])
    vl_bc_cof = soma_valores(*[r.get("VL_BC_COFINS_ITEM", 0) for r in grupo])
    vl_cof    = soma_valores(*[r.get("VL_COFINS_ITEM", 0) for r in grupo])
    lines = [to_pipe(["A100",
        gc(r0,"IND_OPER"), gc(r0,"IND_EMIT"), gc(r0,"COD_PART\n(Fornecedor/Cliente)"),
        gc(r0,"COD_MOD"), gc(r0,"COD_SIT"), gc(r0,"SER"), gc(r0,"SUB"),
        gc(r0,"NUM_DOC"), gc(r0,"CHV_DOC"), gc(r0,"DT_DOC"), gc(r0,"DT_EXE_SERV"),
        fmt_br(vl_doc), gc(r0,"IND_PGTO"), fmt_br(vl_desc),
        fmt_br(vl_bc_pis), fmt_valor(r0.get("ALIQ_PIS_ITEM","")), fmt_br(vl_pis),
        fmt_br(vl_bc_cof), fmt_valor(r0.get("ALIQ_COFINS_ITEM","")), fmt_br(vl_cof),
        gc(r0,"COD_CTA"), gc(r0,"COD_MUN")])]
    for i, r in enumerate(grupo, 1):
        lines.append(to_pipe(["A170", str(i), gc(r,"COD_ITEM"), gc(r,"DESCR_COMPL"),
            fmt_valor(r.get("VL_ITEM","")), fmt_valor(r.get("VL_DESC_ITEM","")),
            gc(r,"NAT_BC_CRED"), gc(r,"IND_ORIG_CRED"), gc(r,"CST_PIS"),
            fmt_valor(r.get("VL_BC_PIS_ITEM","")), fmt_valor(r.get("ALIQ_PIS_ITEM","")), fmt_valor(r.get("VL_PIS_ITEM","")),
            gc(r,"CST_COFINS"), fmt_valor(r.get("VL_BC_COFINS_ITEM","")),
            fmt_valor(r.get("ALIQ_COFINS_ITEM","")), fmt_valor(r.get("VL_COFINS_ITEM","")),
            gc(r,"COD_CTA_ITEM"), gc(r,"COD_CCUS")]))
    return lines

def build_c_block(grupo: list) -> list:
    r0 = grupo[0]
    vl_doc  = soma_valores(*[r.get("VL_ITEM", 0) for r in grupo])
    vl_desc = soma_valores(*[r.get("VL_DESC_ITEM", 0) for r in grupo])
    vl_pis  = soma_valores(*[r.get("VL_PIS_ITEM", 0) for r in grupo])
    vl_cof  = soma_valores(*[r.get("VL_COFINS_ITEM", 0) for r in grupo])
    vl_icms = soma_valores(*[r.get("VL_ICMS_ITEM", 0) for r in grupo])
    vl_ipi  = soma_valores(*[r.get("VL_IPI_ITEM", 0) for r in grupo])
    lines = [to_pipe(["C100",
        gc(r0,"IND_OPER"), gc(r0,"IND_EMIT"), gc(r0,"COD_PART\n(Fornecedor/Cliente)"),
        gc(r0,"COD_MOD"), gc(r0,"COD_SIT"), gc(r0,"SER"), gc(r0,"NUM_DOC"),
        gc(r0,"CHV_NFE"), gc(r0,"DT_DOC"), gc(r0,"DT_E_S"),
        fmt_br(vl_doc), gc(r0,"IND_PGTO"), fmt_br(vl_desc),
        fmt_valor(r0.get("VL_ABAT_NT","")), fmt_br(vl_doc - vl_desc),
        gc(r0,"IND_FRT"), fmt_valor(r0.get("VL_FRT","")), fmt_valor(r0.get("VL_SEG","")),
        fmt_valor(r0.get("VL_OUT_DA","")), "0,00", fmt_br(vl_icms),
        "0,00", "0,00", fmt_br(vl_ipi), fmt_br(vl_pis), fmt_br(vl_cof), "0,00", "0,00",
        gc(r0,"COD_CTA")])]
    for i, r in enumerate(grupo, 1):
        lines.append(to_pipe(["C170", str(i), gc(r,"COD_ITEM"), gc(r,"DESCR_COMPL"),
            fmt_valor(r.get("QTD","")), gc(r,"UNID"), fmt_valor(r.get("VL_ITEM","")),
            fmt_valor(r.get("VL_DESC_ITEM","")), gc(r,"IND_MOV"), gc(r,"CST_ICMS"),
            gc(r,"CFOP"), gc(r,"COD_NAT"), fmt_valor(r.get("VL_BC_ICMS_ITEM","")),
            fmt_valor(r.get("ALIQ_ICMS","")), fmt_valor(r.get("VL_ICMS_ITEM","")),
            fmt_valor(r.get("VL_BC_ICMS_ST_ITEM","")), fmt_valor(r.get("ALIQ_ST","")),
            fmt_valor(r.get("VL_ICMS_ST_ITEM","")), "", gc(r,"CST_IPI"), "",
            fmt_valor(r.get("VL_BC_IPI","")), fmt_valor(r.get("ALIQ_IPI","")),
            fmt_valor(r.get("VL_IPI_ITEM","")), gc(r,"CST_PIS"),
            fmt_valor(r.get("VL_BC_PIS_ITEM","")), fmt_valor(r.get("ALIQ_PIS_%","")),
            fmt_valor(r.get("QUANT_BC_PIS","")), fmt_valor(r.get("ALIQ_PIS_R$","")),
            fmt_valor(r.get("VL_PIS_ITEM","")), gc(r,"CST_COFINS"),
            fmt_valor(r.get("VL_BC_COFINS_ITEM","")), fmt_valor(r.get("ALIQ_COFINS_%","")),
            fmt_valor(r.get("QUANT_BC_COFINS","")), fmt_valor(r.get("ALIQ_COFINS_R$","")),
            fmt_valor(r.get("VL_COFINS_ITEM","")), gc(r,"COD_CTA_ITEM")]))
    return lines

def build_d_block(r) -> list:
    return [
        to_pipe(["D100", gc(r,"IND_OPER"), gc(r,"IND_EMIT"), gc(r,"COD_PART\n(Transportadora)"),
            gc(r,"COD_MOD"), gc(r,"COD_SIT"), gc(r,"SER"), gc(r,"SUB"), gc(r,"NUM_DOC"),
            gc(r,"CHV_CTE"), gc(r,"DT_DOC"), gc(r,"DT_A_P"), gc(r,"TP_CT_e"), gc(r,"CHV_CTE_REF"),
            fmt_valor(r.get("VL_DOC","")), fmt_valor(r.get("VL_DESC","")), gc(r,"IND_FRT"),
            fmt_valor(r.get("VL_SERV","")), fmt_valor(r.get("VL_BC_ICMS","")),
            fmt_valor(r.get("VL_ICMS","")), fmt_valor(r.get("VL_NT","")),
            gc(r,"COD_CTA"), gc(r,"COD_INF")]),
        to_pipe(["D101", gc(r,"IND_NAT_FRT"), fmt_valor(r.get("VL_ITEM_PIS","")),
            gc(r,"CST_PIS"), gc(r,"NAT_BC_CRED"), fmt_valor(r.get("VL_BC_PIS","")),
            fmt_valor(r.get("ALIQ_PIS_%","")), fmt_valor(r.get("QUANT_BC_PIS","")),
            fmt_valor(r.get("ALIQ_PIS_R$","")), fmt_valor(r.get("VL_PIS","")),
            gc(r,"COD_CTA_PIS"), gc(r,"COD_CCUS_PIS")]),
        to_pipe(["D105", gc(r,"IND_NAT_FRT"), fmt_valor(r.get("VL_ITEM_COF","")),
            gc(r,"CST_COFINS"), gc(r,"NAT_BC_CRED"), fmt_valor(r.get("VL_BC_COFINS","")),
            fmt_valor(r.get("ALIQ_COFINS_%","")), fmt_valor(r.get("QUANT_BC_COF","")),
            fmt_valor(r.get("ALIQ_COFINS_R$","")), fmt_valor(r.get("VL_COFINS","")),
            gc(r,"COD_CTA_COF"), gc(r,"COD_CCUS_COF")]),
    ]

def build_f100(r):
    return to_pipe(["F100", gc(r,"IND_OPER"), gc(r,"COD_PART"), gc(r,"DT_OPER"),
        fmt_valor(r.get("VL_OPER","")), gc(r,"COD_ITEM"), gc(r,"DES_COMPL"), gc(r,"COD_CTA"),
        gc(r,"CST_PIS"), fmt_valor(r.get("VL_BC_PIS","")), fmt_valor(r.get("ALIQ_PIS_%","")),
        fmt_valor(r.get("QUANT_BC_PIS","")), fmt_valor(r.get("ALIQ_PIS_R$","")),
        fmt_valor(r.get("VL_PIS","")), gc(r,"CST_COFINS"), fmt_valor(r.get("VL_BC_COFINS","")),
        fmt_valor(r.get("ALIQ_COFINS_%","")), fmt_valor(r.get("QUANT_BC_COF","")),
        fmt_valor(r.get("ALIQ_COFINS_R$","")), fmt_valor(r.get("VL_COFINS","")),
        gc(r,"NAT_BC_CRED"), gc(r,"IND_ORIG_CRED"), gc(r,"COD_CCUS"), gc(r,"COD_MOD"), gc(r,"GRUPO_TENSAO")])

def build_f120(r):
    return to_pipe(["F120", gc(r,"NAT_BC_CRED"), gc(r,"IDENT_BEM_IMOB"), gc(r,"IND_ORIG_CRED"),
        gc(r,"IND_UTIL_BEM_IMOB"), fmt_valor(r.get("VL_OPER_DEP","")), fmt_valor(r.get("PARC_NAO_BC","")),
        gc(r,"CST_PIS"), fmt_valor(r.get("VL_BC_PIS","")), fmt_valor(r.get("ALIQ_PIS_%","")),
        fmt_valor(r.get("QUANT_BC_PIS","")), fmt_valor(r.get("ALIQ_PIS_R$","")),
        fmt_valor(r.get("VL_PIS","")), gc(r,"CST_COFINS"), fmt_valor(r.get("VL_BC_COFINS","")),
        fmt_valor(r.get("ALIQ_COFINS_%","")), fmt_valor(r.get("QUANT_BC_COF","")),
        fmt_valor(r.get("ALIQ_COFINS_R$","")), fmt_valor(r.get("VL_COFINS","")),
        gc(r,"COD_CTA"), gc(r,"COD_CCUS"), gc(r,"DES_COMPL"), gc(r,"COD_MOD")])

def build_f130(r):
    return to_pipe(["F130", gc(r,"NAT_BC_CRED"), gc(r,"IDENT_BEM_IMOB"), gc(r,"IND_ORIG_CRED"),
        gc(r,"IND_UTIL_BEM_IMOB"), gc(r,"MES_OPER_AQUIS"), fmt_valor(r.get("VL_OPER_AQUIS","")),
        fmt_valor(r.get("PARC_NAO_BC","")), fmt_valor(r.get("VL_BC_CRED","")), gc(r,"IND_NR_PARC"),
        gc(r,"CST_PIS"), fmt_valor(r.get("VL_BC_PIS","")), fmt_valor(r.get("ALIQ_PIS_%","")),
        fmt_valor(r.get("QUANT_BC_PIS","")), fmt_valor(r.get("ALIQ_PIS_R$","")),
        fmt_valor(r.get("VL_PIS","")), gc(r,"CST_COFINS"), fmt_valor(r.get("VL_BC_COFINS","")),
        fmt_valor(r.get("ALIQ_COFINS_%","")), fmt_valor(r.get("QUANT_BC_COF","")),
        fmt_valor(r.get("ALIQ_COFINS_R$","")), fmt_valor(r.get("VL_COFINS","")),
        gc(r,"COD_CTA"), gc(r,"COD_CCUS"), gc(r,"DES_COMPL"), gc(r,"COD_MOD")])


# ─────────────────────────────────────────────────────────────────────────────
# Bloco 9 e injeção
# ─────────────────────────────────────────────────────────────────────────────
def recalc_9900(lines: list) -> list:
    counter = defaultdict(int)
    new_lines = []
    skip = False
    for line in lines:
        ls = line.strip()
        if re.match(r"^\|(9001|9900|9990|9999)\|", ls):
            skip = True
            continue
        if skip:
            continue
        new_lines.append(line)
        if ls.startswith("|"):
            reg = parse_pipe(ls)[0]
            counter[reg] += 1
    counter["9001"] = 1
    bloco9 = ["|9001|0|\n"]
    for reg in sorted(counter.keys()):
        bloco9.append(f"|9900|{reg}|{counter[reg]}|\n")
    bloco9.append(f"|9900|9900|{len(bloco9)}|\n")
    bloco9.append("|9900|9990|1|\n")
    bloco9.append("|9900|9999|1|\n")
    bloco9.append("|9990|1|\n")
    bloco9.append(f"|9999|{len(new_lines) + len(bloco9) + 1}|\n")
    return new_lines + bloco9

def find_x010_ranges(lines: list, prefix: str) -> dict:
    ranges = {}
    positions = []
    x990 = None
    for i, line in enumerate(lines):
        ls = line.strip()
        if ls.startswith(f"|{prefix}010|"):
            f = parse_pipe(ls)
            positions.append((i, f[1] if len(f) > 1 else ""))
        if ls.startswith(f"|{prefix}990|"):
            x990 = i
    for k, (pos, cnpj) in enumerate(positions):
        fim = positions[k+1][0] if k+1 < len(positions) else (x990 or len(lines))
        ranges[cnpj] = fim
    return ranges

def inject_by_cnpj(lines: list, prefix: str, cnpj_map: dict, log_lines: list) -> list:
    offset = 0
    ranges = find_x010_ranges(lines, prefix)
    for cnpj, novas in cnpj_map.items():
        if not novas:
            continue
        if cnpj not in ranges:
            log_lines.append(f"⚠ CNPJ {cnpj} sem {prefix}010 no TXT — pulando")
            continue
        pos = ranges[cnpj] + offset
        lines[pos:pos] = novas
        offset += len(novas)
        log_lines.append(f"✔ Bloco {prefix}: {len(novas)} linha(s) → CNPJ {cnpj}")
    return lines

def agrupar_por_doc(rows: list) -> dict:
    grupos = {}
    for row in rows:
        cnpj  = gc(row, "CNPJ_ESTAB")
        ndoc  = gc(row, "NUM_DOC")
        if not cnpj or not ndoc:
            continue
        key = f"{cnpj}||{ndoc}"
        if key not in grupos:
            grupos[key] = {"cnpj": cnpj, "rows": []}
        grupos[key]["rows"].append(row)
    return grupos


# ─────────────────────────────────────────────────────────────────────────────
# Motor principal
# ─────────────────────────────────────────────────────────────────────────────
def retificar(txt_bytes: bytes, planilha: dict) -> tuple[bytes, list]:
    log = []
    enc = detect_encoding(txt_bytes)
    lines = txt_bytes.decode(enc, errors="replace").splitlines(keepends=True)
    periodo = get_periodo(lines)
    log.append(f"→ Período detectado: {periodo or 'não encontrado'}")
    parsed = parse_txt(lines)

    # 0000 → IND_RET = 1
    for i, line in enumerate(lines):
        if line.strip().startswith("|0000|"):
            f = parse_pipe(line.strip())
            f[2] = "1"
            lines[i] = to_pipe(f)
            log.append("✔ 0000 marcado como retificador (IND_RET=1)")
            break

    # ── Hierarquia do bloco 0 conforme Manual EFD Contribuições ─────────────
    # Ordem: 0140 → 0150 → 0190 → 0200 → 0400 → 0450 → 0500 → 0990
    # Cada novo registro é inserido APÓS os registros existentes do mesmo tipo,
    # mantendo a hierarquia e nunca ultrapassando o 0990.

    def find_last_reg(lines_list, reg_prefix, before_idx):
        """Encontra o índice da última linha do registro, limitado a before_idx."""
        last = None
        for i, line in enumerate(lines_list):
            if i >= before_idx:
                break
            if line.strip().startswith(f"|{reg_prefix}|"):
                last = i
        return last

    b0end = parsed["bloco_0_end"]
    offset = 0  # acumula deslocamento a cada inserção

    def insert_after_last(reg_prefix, new_lines):
        nonlocal offset
        if not new_lines:
            return
        b0 = b0end + offset
        last_pos = find_last_reg(lines, reg_prefix, b0)
        if last_pos is not None:
            insert_at = last_pos + 1
        else:
            # Não existe nenhum registro desse tipo: insere antes do 0990
            insert_at = b0
        lines[insert_at:insert_at] = new_lines
        offset += len(new_lines)

    # 0140 — novos estabelecimentos
    novos_0140 = []
    for emp in filter_periodo(planilha.get("0_EMPRESA", []), periodo):
        cnpj = gc(emp, "CNPJ_ESTAB", "CNPJ")
        if not cnpj:
            continue
        if cnpj in parsed["cnpj_0140"]:
            log.append(f"→ 0140 CNPJ {cnpj} já existe — mantém")
        else:
            novos_0140.append(build_0140(emp))
            log.append(f"✔ 0140 CNPJ {cnpj} criado")
    insert_after_last("0140", novos_0140)

    # 0150 — participantes
    novos_0150 = [build_0150(r) for r in filter_periodo(planilha.get("0150_PARTICIPANTES", []), periodo)]
    insert_after_last("0150", novos_0150)

    # 0190 — unidades
    ex_0190 = set()
    for line in lines:
        if line.strip().startswith("|0190|"):
            f = parse_pipe(line.strip()); ex_0190.add(f[1] if len(f) > 1 else "")
    novos_0190 = []
    for r in filter_periodo(planilha.get("0190_UNIDADES", []), periodo):
        u = gc(r, "UNID")
        if u and u not in ex_0190:
            novos_0190.append(build_0190(r))
    insert_after_last("0190", novos_0190)

    # 0200 — itens (centralizado: sem duplicar COD_ITEM)
    novos_0200 = []
    for r in filter_periodo(planilha.get("0200_ITENS", []), periodo):
        cod = gc(r, "COD_ITEM")
        if not cod:
            continue
        if cod not in parsed["cod_0200"]:
            novos_0200.append(build_0200(r))
            parsed["cod_0200"].add(cod)
            log.append(f"✔ 0200 item {cod} inserido")
        else:
            log.append(f"→ 0200 item {cod} já existe — mantém")
    insert_after_last("0200", novos_0200)

    # 0400 — naturezas de receita
    ex_0400 = set()
    for line in lines:
        if line.strip().startswith("|0400|"):
            f = parse_pipe(line.strip()); ex_0400.add(f[1] if len(f) > 1 else "")
    novos_0400 = []
    for r in filter_periodo(planilha.get("0400_NAT_REC", []), periodo):
        c = gc(r, "COD_NAT_REC")
        if c and c not in ex_0400:
            novos_0400.append(build_0400(r))
    insert_after_last("0400", novos_0400)

    # 0500 — plano de contas (inserido após o último 0500, antes do 0990)
    ex_0500 = set()
    for line in lines:
        if line.strip().startswith("|0500|"):
            f = parse_pipe(line.strip()); ex_0500.add(f[1] if len(f) > 1 else "")
    novos_0500 = []
    for r in filter_periodo(planilha.get("0500_PLANO_CONTAS", []), periodo):
        c = gc(r, "COD_CTA")
        if c and c not in ex_0500:
            novos_0500.append(build_0500(r))
    insert_after_last("0500", novos_0500)

    total_ins = len(novos_0140) + len(novos_0150) + len(novos_0190) + len(novos_0200) + len(novos_0400) + len(novos_0500)
    if total_ins:
        log.append(f"✔ {total_ins} linha(s) inseridas no bloco 0 (ordem hierárquica)") 

    # Bloco A
    a_map = defaultdict(list)
    for g in agrupar_por_doc(filter_periodo(planilha.get("A100_A170", []), periodo)).values():
        a_map[g["cnpj"]].extend(build_a_block(g["rows"]))
    lines = inject_by_cnpj(lines, "A", dict(a_map), log)

    # Bloco C
    c_map = defaultdict(list)
    for g in agrupar_por_doc(filter_periodo(planilha.get("C100_C170", []), periodo)).values():
        c_map[g["cnpj"]].extend(build_c_block(g["rows"]))
    lines = inject_by_cnpj(lines, "C", dict(c_map), log)

    # Bloco D
    d_map = defaultdict(list)
    for r in filter_periodo(planilha.get("D100_D101_D105", []), periodo):
        cnpj = gc(r, "CNPJ_ESTAB")
        if cnpj:
            d_map[cnpj].extend(build_d_block(r))
    lines = inject_by_cnpj(lines, "D", dict(d_map), log)

    # Bloco F — F100, F120 e F130 dentro do F010 do estabelecimento correto
    # Ordem dentro do sub-bloco F010: F100 → F120 → F130
    f_map = defaultdict(list)
    for aba, fn in [("F100", build_f100), ("F120", build_f120), ("F130", build_f130)]:
        for r in filter_periodo(planilha.get(aba, []), periodo):
            cnpj = gc(r, "CNPJ_ESTAB")
            if cnpj:
                f_map[cnpj].append(fn(r))
    lines = inject_by_cnpj(lines, "F", dict(f_map), log)

    lines = recalc_9900(lines)
    log.append(f"✔ Bloco 9 recalculado — {len(lines)} linhas no total")

    result_bytes = "".join(lines).encode(enc, errors="replace")
    return result_bytes, log


# ─────────────────────────────────────────────────────────────────────────────
# Interface Streamlit
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("## 📄 EFD Contribuições — Retificação")
st.markdown("Faça upload da planilha modelo e dos TXTs originais. O sistema processa cada período automaticamente e gera os arquivos retificados para download.")
st.divider()

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 1. Planilha modelo")
    xlsx_file = st.file_uploader(
        "Arquivo .xlsx com os novos registros",
        type=["xlsx"],
        help="Use o modelo EFD_Contribuicoes_Modelo_v3.xlsx"
    )

with col2:
    st.markdown("### 2. TXTs originais")
    txt_files = st.file_uploader(
        "Um ou mais arquivos .txt (um por período)",
        type=["txt"],
        accept_multiple_files=True,
        help="Ex.: EFD_012024.txt, EFD_022024.txt"
    )

st.divider()

# Regras
with st.expander("📋 Regras aplicadas automaticamente", expanded=False):
    r1, r2, r3 = st.columns(3)
    with r1:
        st.markdown("**0140 — Estabelecimentos**")
        st.caption("CNPJ já existente no TXT é reutilizado. Novo 0140 criado apenas se não encontrado.")
        st.markdown("**0200 — Itens centralizados**")
        st.caption("Novo item inserido uma vez no bloco 0. COD_ITEM duplicado é ignorado.")
    with r2:
        st.markdown("**A100 / C100 — Totais automáticos**")
        st.caption("VL_DOC, VL_PIS, VL_COFINS calculados somando os itens A170 / C170.")
        st.markdown("**Bloco 9 — Contadores**")
        st.caption("Todos os registros 9900 recalculados. 0000 marcado IND_RET = 1.")
    with r3:
        st.markdown("**0150 / 0190 / 0400 / 0500**")
        st.caption("Inseridos sem duplicar registros já presentes no TXT original.")
        st.markdown("**Blocos A / C / D / F**")
        st.caption("Cada documento inserido no sub-bloco A010 / C010 do CNPJ correto.")

# Botão processar
if xlsx_file and txt_files:
    m1, m2, m3 = st.columns(3)
    m1.metric("Planilha", xlsx_file.name)
    m2.metric("TXTs enviados", len(txt_files))
    m3.metric("Arquivos a gerar", len(txt_files))
    st.divider()

    if st.button("🚀 Processar retificação", type="primary", use_container_width=True):
        with st.spinner("Lendo planilha…"):
            try:
                planilha = load_planilha(xlsx_file.read())
                st.success(f"Planilha lida — {len(planilha)} abas carregadas")
            except Exception as e:
                st.error(f"Erro ao ler planilha: {e}")
                st.stop()

        resultados = []
        todos_logs = []

        progress = st.progress(0, text="Iniciando…")

        for i, txt_file in enumerate(txt_files):
            nome_saida = txt_file.name.replace(".txt", "_RET.txt").replace(".TXT", "_RET.txt")
            st.markdown(f"**Processando: {txt_file.name}**")

            try:
                result_bytes, log_lines = retificar(txt_file.read(), planilha)
                resultados.append((nome_saida, result_bytes))
                todos_logs.extend([f"[{txt_file.name}] {l}" for l in log_lines])

                log_html = "<br>".join(
                    f'<span class="{"status-ok" if l.startswith("✔") else "status-warn" if l.startswith("⚠") else "status-info"}">{l}</span>'
                    for l in log_lines
                )
                st.markdown(f'<div class="log-box">{log_html}</div>', unsafe_allow_html=True)
                st.success(f"✔ {nome_saida} gerado — {len(result_bytes):,} bytes")

            except Exception as e:
                st.error(f"Erro ao processar {txt_file.name}: {e}")
                todos_logs.append(f"[{txt_file.name}] ✖ Erro: {e}")

            progress.progress((i + 1) / len(txt_files), text=f"{i+1}/{len(txt_files)} arquivo(s) processado(s)")

        progress.progress(1.0, text="Concluído!")
        st.divider()
        st.markdown("### 📥 Downloads")

        if len(resultados) == 1:
            nome, dados = resultados[0]
            st.download_button(
                label=f"⬇ Baixar {nome}",
                data=dados,
                file_name=nome,
                mime="text/plain",
                use_container_width=True,
            )
        elif len(resultados) > 1:
            # ZIP com todos
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for nome, dados in resultados:
                    zf.writestr(nome, dados)
            zip_buffer.seek(0)
            st.download_button(
                label=f"⬇ Baixar todos ({len(resultados)} arquivos em .zip)",
                data=zip_buffer,
                file_name="EFD_Retificados.zip",
                mime="application/zip",
                use_container_width=True,
            )
            st.markdown("Ou baixe individualmente:")
            for nome, dados in resultados:
                st.download_button(
                    label=f"⬇ {nome}",
                    data=dados,
                    file_name=nome,
                    mime="text/plain",
                    key=nome,
                )
else:
    st.info("Faça upload da planilha .xlsx e de um ou mais TXTs originais para habilitar o processamento.")
