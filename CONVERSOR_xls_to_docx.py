# -*- coding: utf-8 -*-

"""
Lê uma planilha (Excel ou CSV), processa despesas e gera um relatório Word.

Regras desta versão:
- Excel: usa sempre a 2ª aba (índice 1); falha se o arquivo tiver menos de 2 abas.
- Remove linhas não-transação (vazias, totais e valor inválido).
- Mantém regras de cálculo de Parcela/% do fluxo existente.
- Aplica cores na coluna Parcela: vermelho para negativo, azul para positivo (tags carol/m&c).
- Inclui gráfico donut de 2 níveis (categoria + subdivisão por tag).
"""

import os
import re
import tempfile

import docx
import numpy as np
import pandas as pd
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

COLUNAS_ESPERADAS = ["Data", "Descrição", "Conta", "Categoria", "Tags", "Valor", "%", "Parcela", "Situação"]
TAGS_GRAFICO = ["mauricio", "carol", "m&c"]
TAGS_TOTAL_PARCELA = ["carol", "m&c"]


# --- Helpers para forçar larguras fixas no Word (tblGrid + tcW) ---
def _twips_from_cm(cm_value: float) -> int:
    # 1 cm ≈ 566.929 twips
    return int(round(float(cm_value) * 567))


def _ensure_tblPr(table):
    """Garante a existência de w:tblPr e o retorna."""
    tbl = table._element
    tbl_pr_list = tbl.xpath("./w:tblPr")
    if tbl_pr_list:
        return tbl_pr_list[0]
    tbl_pr = OxmlElement("w:tblPr")
    tbl.insert(0, tbl_pr)
    return tbl_pr


def _set_table_layout_fixed(table):
    """Força w:tblLayout type='fixed' dentro de w:tblPr."""
    tbl_pr = _ensure_tblPr(table)
    for child in list(tbl_pr):
        if child.tag == qn("w:tblLayout"):
            tbl_pr.remove(child)
    tbl_layout = OxmlElement("w:tblLayout")
    tbl_layout.set(qn("w:type"), "fixed")
    tbl_pr.append(tbl_layout)


def _set_cell_width(cell, cm_value: float):
    """Define a largura da célula via w:tcW (dxa) e mantém cell.width (Cm)."""
    try:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for child in list(tcPr):
            if child.tag == qn("w:tcW"):
                tcPr.remove(child)
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:type"), "dxa")
        tcW.set(qn("w:w"), str(_twips_from_cm(cm_value)))
        tcPr.append(tcW)
        cell.width = Cm(cm_value)
    except Exception:
        cell.width = Cm(cm_value)


def _apply_table_grid(table, column_names, widths_map_cm: dict):
    """Define w:tblGrid com larguras em twips e ajusta w:tblW."""
    tbl = table._element
    tbl_pr = _ensure_tblPr(table)

    for grid in tbl.xpath("./w:tblGrid"):
        tbl.remove(grid)

    grid = OxmlElement("w:tblGrid")
    total_twips = 0
    for name in column_names:
        width_cm = float(widths_map_cm.get(name, 2.0))
        twips = _twips_from_cm(width_cm)
        total_twips += twips
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(twips))
        grid.append(grid_col)

    try:
        tbl.insert(1, grid)
    except Exception:
        tbl.append(grid)

    for child in list(tbl_pr):
        if child.tag == qn("w:tblW"):
            tbl_pr.remove(child)
    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(total_twips))
    tbl_pr.append(tbl_w)


def _formatar_numero_br(valor):
    return f"{float(valor):,.2f}".replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")


def _is_number(valor):
    return isinstance(valor, (int, float, np.integer, np.floating)) and not pd.isna(valor)


def _limpar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def _normalizar_tag(valor):
    return _limpar_texto(valor).lower()


def set_cell_font(cell, text, bold=False, size=14, color=None, align="left"):
    p = cell.paragraphs[0]
    p.text = text
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = p.runs[0]
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color


def _verificar_dependencia_excel(caminho_arquivo):
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    if extensao in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        try:
            import openpyxl  # noqa: F401
        except Exception:
            return "openpyxl"
    elif extensao == ".xls":
        try:
            import xlrd  # noqa: F401
        except Exception:
            return "xlrd"
    return None


def _verificar_dependencia_grafico():
    try:
        import matplotlib  # noqa: F401
    except Exception:
        return "matplotlib"
    return None


def _obter_titulo_mes(nome_base):
    meses_pt = {
        "JAN": "Janeiro",
        "FEV": "Fevereiro",
        "MAR": "Março",
        "ABR": "Abril",
        "MAI": "Maio",
        "JUN": "Junho",
        "JUL": "Julho",
        "AGO": "Agosto",
        "SET": "Setembro",
        "OUT": "Outubro",
        "NOV": "Novembro",
        "DEZ": "Dezembro",
    }
    match = re.match(r"^([A-Za-z]{3})_(\d{4})$", nome_base.strip())
    if not match:
        return nome_base

    mes_abrev = match.group(1).upper()
    ano = match.group(2)
    if mes_abrev in meses_pt:
        return f"{meses_pt[mes_abrev]} {ano}"
    return nome_base


def _ler_dataframe_excel_segunda_aba(caminho_arquivo):
    dependencia_faltando = _verificar_dependencia_excel(caminho_arquivo)
    if dependencia_faltando:
        raise RuntimeError(
            "ERRO FATAL ao tentar ler a planilha: "
            f"Dependência '{dependencia_faltando}' não instalada. "
            f"Instale com: pip install {dependencia_faltando}"
        )

    excel_file = pd.ExcelFile(caminho_arquivo)
    if len(excel_file.sheet_names) < 2:
        raise RuntimeError(
            "ERRO FATAL: arquivo Excel precisa ter pelo menos 2 abas. "
            "Este fluxo usa sempre a 2ª aba (índice 1)."
        )

    aba_escolhida = excel_file.sheet_names[1]
    dataframe = pd.read_excel(caminho_arquivo, sheet_name=aba_escolhida)
    return dataframe, aba_escolhida, excel_file.sheet_names


def _filtrar_linhas_nao_transacao(df):
    colunas_base = ["Data", "Descrição", "Valor", "Conta", "Situação", "Categoria", "Tags"]
    for col in colunas_base:
        if col not in df.columns:
            df[col] = np.nan

    total_inicial = len(df)
    df_filtrado = df.dropna(how="all", subset=colunas_base).copy()
    removidas_vazias = total_inicial - len(df_filtrado)

    data_str = df_filtrado["Data"].astype(str).str.strip().str.lower()
    mask_total = data_str.str.startswith("total")
    removidas_total = int(mask_total.sum())
    df_filtrado = df_filtrado.loc[~mask_total].copy()

    df_filtrado["Valor"] = pd.to_numeric(df_filtrado["Valor"], errors="coerce")
    mask_valor_invalido = df_filtrado["Valor"].isna()
    removidas_valor_invalido = int(mask_valor_invalido.sum())
    df_filtrado = df_filtrado.loc[~mask_valor_invalido].copy()

    return df_filtrado, {
        "inicial": total_inicial,
        "removidas_vazias": removidas_vazias,
        "removidas_total": removidas_total,
        "removidas_valor_invalido": removidas_valor_invalido,
        "final": len(df_filtrado),
    }


def _ajustar_luminosidade(rgb, fator):
    if fator < 1:
        return tuple(max(0.0, min(1.0, c * fator)) for c in rgb)

    ganho = fator - 1
    return tuple(max(0.0, min(1.0, c + (1 - c) * ganho)) for c in rgb)


def _gerar_grafico_donut_categoria_tag(df_processado, caminho_png):
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Patch

    df_chart = df_processado.copy()
    df_chart["Tags"] = df_chart["Tags"].apply(_normalizar_tag)
    df_chart["Categoria"] = df_chart["Categoria"].apply(_limpar_texto)

    df_chart = df_chart[df_chart["Tags"].isin(TAGS_GRAFICO)].copy()
    df_chart = df_chart[df_chart["Categoria"] != ""].copy()
    df_chart["metric"] = df_chart["Parcela"].abs()
    df_chart = df_chart[df_chart["metric"] > 0].copy()

    if df_chart.empty:
        raise RuntimeError("ERRO FATAL: não há dados válidos para gerar o gráfico donut.")

    inner = df_chart.groupby("Categoria", as_index=True)["metric"].sum().sort_values(ascending=False)
    outer = df_chart.groupby(["Categoria", "Tags"], as_index=True)["metric"].sum()

    categorias = inner.index.tolist()
    mapa_cor_base = {}
    cmap = plt.get_cmap("tab20")
    for indice, categoria in enumerate(categorias):
        mapa_cor_base[categoria] = cmap(indice % cmap.N)[:3]

    inner_valores = inner.values.tolist()
    inner_cores = [mapa_cor_base[categoria] for categoria in categorias]

    fatores_tag = {"mauricio": 0.72, "carol": 1.0, "m&c": 1.28}
    outer_valores = []
    outer_cores = []
    for categoria in categorias:
        cor_base = mapa_cor_base[categoria]
        for tag in TAGS_GRAFICO:
            valor = float(outer.get((categoria, tag), 0.0))
            if valor <= 0:
                continue
            outer_valores.append(valor)
            outer_cores.append(_ajustar_luminosidade(cor_base, fatores_tag[tag]))

    if not outer_valores:
        raise RuntimeError("ERRO FATAL: não foi possível montar as subdivisões do gráfico donut.")

    fig, ax = plt.subplots(figsize=(12, 7.2), dpi=180)

    ax.pie(
        inner_valores,
        radius=1.0,
        colors=inner_cores,
        startangle=90,
        counterclock=False,
        wedgeprops={"width": 0.34, "edgecolor": "white", "linewidth": 1.2},
    )

    ax.pie(
        outer_valores,
        radius=1.34,
        colors=outer_cores,
        startangle=90,
        counterclock=False,
        wedgeprops={"width": 0.30, "edgecolor": "white", "linewidth": 0.8},
    )

    total = float(sum(inner_valores))
    ax.text(
        0,
        0,
        f"Total\nR$ {_formatar_numero_br(total)}",
        ha="center",
        va="center",
        fontsize=12,
        fontweight="bold",
    )

    legenda_categoria = [
        Patch(facecolor=inner_cores[i], label=f"{cat} (R$ {_formatar_numero_br(inner_valores[i])})")
        for i, cat in enumerate(categorias)
    ]
    legenda_tag = [
        Patch(facecolor=(0.35, 0.35, 0.35), label="mauricio (tom escuro)"),
        Patch(facecolor=(0.6, 0.6, 0.6), label="carol (tom base)"),
        Patch(facecolor=(0.85, 0.85, 0.85), label="m&c (tom claro)"),
    ]

    legenda_1 = ax.legend(
        handles=legenda_categoria,
        title="Categorias",
        loc="center left",
        bbox_to_anchor=(1.02, 0.50),
        fontsize=9,
        title_fontsize=10,
        frameon=False,
    )
    ax.add_artist(legenda_1)

    ax.legend(
        handles=legenda_tag,
        title="Subdivisão por pessoa",
        loc="upper center",
        bbox_to_anchor=(0.5, -0.05),
        ncol=3,
        fontsize=9,
        title_fontsize=10,
        frameon=False,
    )

    ax.set_title("Distribuição de despesas por categoria e pessoa (base: abs(Parcela))", fontsize=12)
    fig.subplots_adjust(left=0.04, right=0.78, top=0.90, bottom=0.14)
    fig.savefig(caminho_png, dpi=220, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def selecionar_arquivo():
    """Abre o seletor de arquivos para escolher a planilha de entrada."""
    try:
        from tkinter import Tk, filedialog

        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha de entrada",
            filetypes=[
                ("Planilhas Excel", "*.xlsx *.xls"),
                ("Arquivos CSV", "*.csv"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        root.destroy()
        return caminho if caminho else None
    except Exception as exc:
        print(f"ERRO ao abrir o seletor de arquivos: {exc}")
        return None


def processar_e_gerar_docx(caminho_arquivo, verbose=False):
    """
    Executa leitura, cálculo e geração do relatório .docx.

    Args:
        caminho_arquivo (str): caminho do arquivo de entrada (.xlsx, .xls ou .csv).
        verbose (bool): ativa logs detalhados.
    """
    if not os.path.exists(caminho_arquivo):
        print(f"ERRO: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return

    dependencia_grafico = _verificar_dependencia_grafico()
    if dependencia_grafico:
        print(
            "ERRO FATAL ao gerar relatório: "
            f"Dependência '{dependencia_grafico}' não instalada. "
            f"Instale com: pip install {dependencia_grafico}"
        )
        return

    nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    titulo_relatorio = _obter_titulo_mes(nome_base)

    try:
        if caminho_arquivo.lower().endswith(".csv"):
            try:
                df = pd.read_csv(caminho_arquivo, sep=None, engine="python")
            except Exception:
                df = pd.read_csv(caminho_arquivo, sep=";")
            aba_origem = "CSV"
            if verbose:
                print("Entrada CSV: leitura direta sem seleção de abas.")
        else:
            df, aba_origem, abas_disponiveis = _ler_dataframe_excel_segunda_aba(caminho_arquivo)
            print(f"Aba utilizada (índice 1): '{aba_origem}'. Abas disponíveis: {abas_disponiveis}")

        for col in COLUNAS_ESPERADAS:
            if col not in df.columns:
                df[col] = np.nan

        df, info_filtro = _filtrar_linhas_nao_transacao(df)
        print(
            "Filtro de linhas concluído: "
            f"inicial={info_filtro['inicial']}, "
            f"vazias={info_filtro['removidas_vazias']}, "
            f"totais={info_filtro['removidas_total']}, "
            f"valor_inválido={info_filtro['removidas_valor_invalido']}, "
            f"final={info_filtro['final']}."
        )

    except Exception as exc:
        print(f"ERRO FATAL ao tentar ler/filtrar a planilha: {exc}")
        return

    if df.empty:
        print("ERRO FATAL: não há linhas válidas para processar após a filtragem.")
        return

    for col in ["Data", "Descrição", "Conta", "Categoria", "Tags", "Situação"]:
        df[col] = df[col].apply(_limpar_texto)

    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce")
    df = df.loc[df["Valor"].notna()].copy()
    if df.empty:
        print("ERRO FATAL: não há valores numéricos válidos na coluna 'Valor'.")
        return

    df["Parcela"] = df["Valor"].copy()
    df["%"] = "1"
    logs = []

    contas_excluidas_inversao = ["Itaú - C.Corrente", "Banco do Brasil - C.Corrente"]

    for index, row in df.iterrows():
        valor = row["Valor"]
        tags = _normalizar_tag(row["Tags"])
        conta = _limpar_texto(row["Conta"])
        situacao = _limpar_texto(row["Situação"])

        deve_inverter_sinal_negativo = False
        if situacao == "Paga":
            if conta not in contas_excluidas_inversao:
                deve_inverter_sinal_negativo = True

        if tags == "carol" and situacao == "Paga":
            novo_valor = abs(valor)
            logs.append(
                f"[linha {index:04d}] Regra: Paga, carol. "
                f"Valor {valor:8.2f} -> {novo_valor:8.2f} (invertido para POSITIVO)."
            )
            df.loc[index, "Parcela"] = novo_valor
        elif tags == "m&c":
            novo_valor = valor / 2
            df.loc[index, "%"] = "0,5"
            if deve_inverter_sinal_negativo:
                novo_valor = -novo_valor
                logs.append(
                    f"[linha {index:04d}] Regra: Paga, m&c. "
                    f"Valor {valor:8.2f} -> {novo_valor:8.2f} (invertido e dividido)."
                )
            else:
                logs.append(
                    f"[linha {index:04d}] Regra: m&c (sem inversão). "
                    f"Valor {valor:8.2f} -> {novo_valor:8.2f} (dividido)."
                )
            df.loc[index, "Parcela"] = novo_valor
        elif tags == "mauricio":
            novo_valor = valor
            if deve_inverter_sinal_negativo and valor < 0:
                novo_valor = -valor
                logs.append(
                    f"[linha {index:04d}] Regra: Paga, mauricio. "
                    f"Valor {valor:8.2f} -> {novo_valor:8.2f} (invertido para POSITIVO)."
                )
            else:
                logs.append(f"[linha {index:04d}] Regra: mauricio (sem inversão). Valor {valor:8.2f} mantido.")
            df.loc[index, "Parcela"] = novo_valor
        else:
            novo_valor = valor
            if deve_inverter_sinal_negativo and valor < 0:
                novo_valor = -valor
                logs.append(
                    f"[linha {index:04d}] Regra: Paga, outrem. "
                    f"Valor {valor:8.2f} -> {novo_valor:8.2f} (invertido para POSITIVO)."
                )
            else:
                logs.append(f"[linha {index:04d}] Regra: {tags} (sem inversão). Valor {valor:8.2f} mantido.")
            df.loc[index, "Parcela"] = novo_valor

    if verbose:
        for linha in logs[:20]:
            print(linha)
        if len(logs) > 20:
            print(f"... {len(logs) - 20} logs adicionais omitidos.")

    df["Conta"] = df["Conta"].str.replace("Itaú - C.Corrente", "Itaú", regex=False)
    df.sort_values(by=["Tags", "Valor"], ascending=[True, True], inplace=True)

    colunas_ordenadas = ["Data", "Descrição", "Conta", "Categoria", "Tags", "Valor", "%", "Parcela", "Situação"]
    df_final = df.reindex(columns=colunas_ordenadas).copy()

    soma_valor_original = float(df["Valor"].sum())
    filtro_parcela_tags = df["Tags"].apply(_normalizar_tag).isin(TAGS_TOTAL_PARCELA)
    soma_parcela_especifica = float(df.loc[filtro_parcela_tags, "Parcela"].sum())
    print(
        "Totais calculados: "
        f"soma Valor={_formatar_numero_br(soma_valor_original)}, "
        f"soma Parcela(carol+m&c)={_formatar_numero_br(soma_parcela_especifica)}."
    )

    doc = docx.Document()

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    doc.add_heading(titulo_relatorio, level=1)
    doc.add_paragraph()

    larguras_fixas_cm = {
        "Data": 3.1,
        "Descrição": 6.2,
        "Conta": 2.0,
        "Categoria": 3.5,
        "Tags": 2.3,
        "Valor": 3.0,
        "%": 1.5,
        "Parcela": 2.6,
        "Situação": 2.4,
    }

    table = doc.add_table(rows=1, cols=len(df_final.columns))
    table.style = "Table Grid"

    try:
        table.autofit = False
        table.allow_autofit = False
    except Exception:
        table.autofit = False

    _set_table_layout_fixed(table)
    _apply_table_grid(table, list(df_final.columns), larguras_fixas_cm)

    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_final.columns):
        _set_cell_width(hdr_cells[i], larguras_fixas_cm.get(col_name, 2.0))
        set_cell_font(hdr_cells[i], col_name, bold=True, align="center")

    for _, row in df_final.iterrows():
        row_cells = table.add_row().cells

        for i, col_name in enumerate(df_final.columns):
            _set_cell_width(row_cells[i], larguras_fixas_cm.get(col_name, 2.0))
            font_color = None
            alinhamento = "justify"

            if col_name == "Parcela":
                tags = _normalizar_tag(row["Tags"])
                parcela = row["Parcela"]
                if tags in TAGS_TOTAL_PARCELA and _is_number(parcela):
                    if parcela < 0:
                        font_color = RGBColor(255, 0, 0)
                    elif parcela > 0:
                        font_color = RGBColor(0, 0, 255)

            valor_celula = row[col_name]
            if _is_number(valor_celula):
                texto_formatado = _formatar_numero_br(valor_celula)
                alinhamento = "right"
            else:
                texto_formatado = str(valor_celula if pd.notna(valor_celula) else "")
                if col_name == "Tags":
                    alinhamento = "center"

            set_cell_font(row_cells[i], texto_formatado, color=font_color, align=alinhamento)

    total_cells = table.add_row().cells
    for i, col_name in enumerate(df_final.columns):
        _set_cell_width(total_cells[i], larguras_fixas_cm.get(col_name, 2.0))

    valor_col_index = list(df_final.columns).index("Valor")
    parcela_col_index = list(df_final.columns).index("Parcela")

    total_parcela_color = None
    if soma_parcela_especifica < 0:
        total_parcela_color = RGBColor(238, 0, 0)
    elif soma_parcela_especifica > 0:
        total_parcela_color = RGBColor(0, 0, 255)

    set_cell_font(total_cells[0], "TOTAIS", bold=True, align="center")
    set_cell_font(total_cells[valor_col_index], _formatar_numero_br(soma_valor_original), bold=True, align="right")
    set_cell_font(
        total_cells[parcela_col_index],
        _formatar_numero_br(soma_parcela_especifica),
        bold=True,
        color=total_parcela_color,
        align="right",
    )

    doc.add_paragraph()
    subtitulo = doc.add_paragraph("Distribuição de despesas por categoria e pessoa")
    subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    caminho_grafico_tmp = None
    try:
        with tempfile.NamedTemporaryFile(prefix="grafico_donut_", suffix=".png", delete=False) as arquivo_tmp:
            caminho_grafico_tmp = arquivo_tmp.name

        _gerar_grafico_donut_categoria_tag(df, caminho_grafico_tmp)
        doc.add_picture(caminho_grafico_tmp, width=Cm(20))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        print("Gráfico donut gerado e inserido no documento.")

    except Exception as exc:
        print(f"ERRO FATAL ao gerar/inserir gráfico: {exc}")
        if caminho_grafico_tmp and os.path.exists(caminho_grafico_tmp):
            os.remove(caminho_grafico_tmp)
        return

    nome_saida = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    pasta_saida = os.path.dirname(caminho_arquivo)
    caminho_saida = os.path.join(pasta_saida, f"{nome_saida}.docx")
    caminho_tmp_docx = f"{caminho_saida}.tmp"

    try:
        doc.save(caminho_tmp_docx)
        try:
            if os.path.exists(caminho_saida):
                os.remove(caminho_saida)
        except Exception:
            pass
        os.replace(caminho_tmp_docx, caminho_saida)
        print(f"Arquivo Word gerado com sucesso: {caminho_saida}")

    finally:
        if caminho_grafico_tmp and os.path.exists(caminho_grafico_tmp):
            os.remove(caminho_grafico_tmp)


if __name__ == "__main__":
    nome_do_arquivo_de_entrada = selecionar_arquivo()
    if nome_do_arquivo_de_entrada:
        print(f"\nArquivo selecionado: '{nome_do_arquivo_de_entrada}'. Iniciando processamento...")
        modo_verbose = True
        processar_e_gerar_docx(nome_do_arquivo_de_entrada, verbose=modo_verbose)
    else:
        print("Nenhum arquivo foi selecionado.")

    try:
        input("\nPressione Enter para sair...")
    except Exception:
        pass
