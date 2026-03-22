# -*- coding: utf-8 -*-

from __future__ import annotations
import builtins
import re
import csv
import os
import asyncio
import sys
import traceback
import unicodedata
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from collections import deque

class C:
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    END = '\033[0m'
    BOLD = '\033[1m'
# --------------------------------------------------------------------


def print(*args, sep=" ", end="\n", file=None, flush=False):
    destino = file or sys.stdout
    encoding = getattr(destino, "encoding", None)
    if not encoding:
        builtins.print(*args, sep=sep, end=end, file=destino, flush=flush)
        return

    safe_args = []
    for arg in args:
        texto = str(arg)
        try:
            texto.encode(encoding)
        except UnicodeEncodeError:
            texto = texto.encode(encoding, errors="replace").decode(encoding)
        safe_args.append(texto)

    builtins.print(*safe_args, sep=sep, end=end, file=destino, flush=flush)

fitz = None
openai = None
tqdm_asyncio = None

try:
    import fitz  # PyMuPDF
except ImportError:
    pass

try:
    import openai
except ImportError:
    pass

try:
    from tqdm.asyncio import tqdm_asyncio
except ImportError:
    pass

try:
    from dotenv import load_dotenv  # opcional
    load_dotenv()
except ImportError:
    print("dotenv não instalado, pulando o carregamento de variáveis de .env")
except Exception as e:
    print(f"Erro ao carregar .env: {e}")

missing_obrigatorias = []
missing_opcionais = []
if fitz is None:
    missing_obrigatorias.append("pymupdf")
if openai is None:
    missing_opcionais.append("openai")
if tqdm_asyncio is None:
    missing_opcionais.append("tqdm")

if missing_obrigatorias:
    print(f"{C.RED}{C.BOLD}ERRO: Dependência obrigatória não encontrada: {', '.join(missing_obrigatorias)}.{C.END}")
    print(f"{C.YELLOW}Instale com: python -m pip install {' '.join(missing_obrigatorias)}{C.END}")
if missing_opcionais:
    print(f"{C.YELLOW}AVISO: Dependências opcionais ausentes: {', '.join(missing_opcionais)}.{C.END}")
    print(f"{C.YELLOW}Para habilitar a classificação por IA: python -m pip install {' '.join(missing_opcionais)}{C.END}")


def _ler_primeira_linha(path: Path) -> Optional[str]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            for linha in f:
                valor = linha.strip()
                if valor and not valor.startswith("#"):
                    if "=" in valor:
                        chave, possivel_valor = valor.split("=", 1)
                        if chave.strip().upper() in {"OPENAI_API_KEY", "API_KEY"}:
                            valor = possivel_valor.strip()
                    valor = valor.strip("'\"")
                    return valor
    except Exception:
        return None
    return None


def _carregar_openai_api_key() -> Optional[str]:
    key_env = (os.getenv("OPENAI_API_KEY") or "").strip()
    if key_env:
        return key_env

    arquivos_candidatos = [
        Path.cwd() / "OPENAI_API_KEY.txt",
        Path.cwd() / "CHAVE_SECRETA_API_Mauricio_local.txt",
        Path.cwd() / ".openai_api_key",
        Path.home() / ".openai_api_key",
    ]
    for arquivo in arquivos_candidatos:
        if not arquivo.exists():
            continue
        key_arquivo = _ler_primeira_linha(arquivo)
        if key_arquivo:
            print(f"{C.GREEN}✓ OPENAI_API_KEY carregada de {arquivo.name}.{C.END}")
            return key_arquivo

    return None


OPENAI_API_KEY = _carregar_openai_api_key()
PDF_PASSWORD = "02533"

MESES_SIGLA_PT = {
    1: "JAN",
    2: "FEV",
    3: "MAR",
    4: "ABR",
    5: "MAI",
    6: "JUN",
    7: "JUL",
    8: "AGO",
    9: "SET",
    10: "OUT",
    11: "NOV",
    12: "DEZ",
}

MODELO_IA = "gpt-5-mini"

CATEGORIAS_PERMITIDAS = [
    "Casa", "Educação", "Eletrônicos", "Lazer", "Outros", "Restaurante",
    "Saúde", "Serviços", "Supermercado", "Transporte", "Vestuário", "Viagem"
]

MAX_RETRIES = 3
INITIAL_RETRY_DELAY = 5
MAX_CONCURRENT_REQUESTS = 10


def _abrir_pdf_com_senha(pdf_path: Path):
    if fitz is None:
        raise RuntimeError("PyMuPDF (pymupdf) não está instalado.")

    doc = fitz.open(pdf_path)
    if doc.needs_pass or doc.is_encrypted:
        ok = doc.authenticate(PDF_PASSWORD)
        if not ok:
            doc.close()
            raise RuntimeError("Não foi possível abrir o PDF com a senha configurada.")
    return doc


def _normalizar_hifens(texto: str) -> str:
    return texto.replace("−", "-").replace("–", "-").replace("—", "-")


def _nome_pasta_competencia(mes: int, ano: int) -> str:
    sigla = MESES_SIGLA_PT.get(mes, f"M{mes:02d}")
    return f"{sigla}_{ano}"

async def classificar_transacao(transacao: Dict[str, Any], cliente: Any) -> str:
    """Usa a API da OpenAI para classificar a categoria de uma única transação."""
    descricao = transacao['descricao']
    prompt = f"""
    Classifique a seguinte descrição de despesa de cartão de crédito em uma das categorias abaixo.
    Responda APENAS com o nome exato da categoria.

    Categorias Válidas: {', '.join(CATEGORIAS_PERMITIDAS)}
    Descrição da Despesa: "{descricao}"

    Categoria:
    """
    
    for attempt in range(MAX_RETRIES):
        try:
            response = await cliente.chat.completions.create(
                model=MODELO_IA,
                messages=[
                    {"role": "system", "content": "Você é um assistente de finanças pessoais preciso e eficiente."},
                    {"role": "user", "content": prompt}
                ],
            )
            categoria = response.choices[0].message.content.strip().replace('"', '').replace('.', '')
            if categoria in CATEGORIAS_PERMITIDAS:
                return categoria
            else:
                return "Outros"
        except (openai.APITimeoutError, openai.APIConnectionError) as e:
            delay = INITIAL_RETRY_DELAY * (2 ** attempt)
            print(f"{C.YELLOW}AVISO: Erro de API para '{descricao}' (tentativa {attempt + 1}/{MAX_RETRIES}). Tentando novamente em {delay}s...{C.END}")
            await asyncio.sleep(delay)
        except Exception as e:
            print(f"{C.RED}ERRO INESPERADO na API para '{descricao}': {e}{C.END}")
            return "Outros"

    return "Outros"

async def processar_categorias_em_lote(transacoes_por_cartao: Dict[str, List[Dict[str, Any]]]):
    """Orquestra a classificação de todas as transações de forma concorrente."""
    if openai is None:
        print(f"{C.YELLOW}Classificação por IA indisponível: biblioteca 'openai' não instalada.{C.END}")
        return
    if not OPENAI_API_KEY:
        print(f"\n{C.RED}{C.BOLD}ERRO CRÍTICO: A variável OPENAI_API_KEY não foi configurada.{C.END}")
        print(f"{C.YELLOW}Defina OPENAI_API_KEY no ambiente ou preencha o arquivo CHAVE_SECRETA_API_Mauricio_local.txt.{C.END}")
        return

    cliente_openai = openai.AsyncOpenAI(api_key=OPENAI_API_KEY)
    
    tarefas_para_ia = []
    for cartao, transacoes in transacoes_por_cartao.items():
        for transacao in transacoes:
            if float(transacao['valor']) < 0:
                tarefas_para_ia.append(transacao)

    if not tarefas_para_ia:
        print(f"{C.YELLOW}Nenhuma despesa encontrada para ser classificada.{C.END}")
        return
        
    print(f"\n{C.CYAN}{C.BOLD}🤖 Iniciando classificação de {len(tarefas_para_ia)} despesas com IA (usando {MODELO_IA})...{C.END}")
    print(f"{C.CYAN}   Isso pode levar alguns minutos. Serão feitas até {MAX_CONCURRENT_REQUESTS} chamadas simultâneas à API.{C.END}")

    tasks = [classificar_transacao(transacao, cliente_openai) for transacao in tarefas_para_ia]
    if tqdm_asyncio:
        resultados_categorias = await tqdm_asyncio.gather(*tasks, desc="Classificando despesas")
    else:
        resultados_categorias = await asyncio.gather(*tasks)
    
    for transacao, categoria in zip(tarefas_para_ia, resultados_categorias):
        transacao['categoria'] = categoria
        
    print(f"{C.GREEN}✓ Classificação por IA concluída.{C.END}")
# ==============================================================================
# UTILITÁRIOS
# ==============================================================================

def selecionar_arquivo_entrada() -> Optional[Path]:
    """Abre uma janela para selecionar PDF ou CSV; faz fallback para busca local se não houver GUI."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as e:
        print(f"{C.YELLOW}AVISO: Não foi possível abrir a janela de seleção ({e}).{C.END}")
        return encontrar_arquivo_entrada()

    try:
        root = tk.Tk()
        root.withdraw()
        root.update()
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo PDF ou CSV",
            filetypes=[
                ("Arquivos PDF e CSV", "*.pdf *.csv"),
                ("Arquivos PDF", "*.pdf"),
                ("Arquivos CSV", "*.csv"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        root.destroy()
    except Exception as e:
        print(f"{C.YELLOW}AVISO: Erro ao abrir a janela de seleção ({e}).{C.END}")
        return encontrar_arquivo_entrada()

    if not file_path:
        print(f"{C.YELLOW}Nenhum arquivo selecionado.{C.END}")
        return None

    entrada_path = Path(file_path)
    if entrada_path.suffix.lower() not in (".pdf", ".csv"):
        print(f"{C.RED}ERRO: O arquivo selecionado não é PDF nem CSV.{C.END}")
        return None

    print(f"{C.GREEN}✓ Arquivo selecionado: {entrada_path}{C.END}")
    return entrada_path

def _priorizar_csv(csvs: List[Path]) -> Path:
    csvs = sorted(csvs, key=lambda p: p.name.lower())
    for csv_path in csvs:
        stem = csv_path.stem.lower()
        if stem.startswith(("fatura-", "fatura_")):
            return csv_path
    return csvs[0]

def encontrar_arquivo_entrada() -> Optional[Path]:
    print(f"{C.CYAN}🔎 Procurando por arquivos .pdf ou .csv na pasta atual...{C.END}")
    pdfs = sorted(Path.cwd().glob("*.pdf"))
    csvs = sorted(Path.cwd().glob("*.csv"))
    if not pdfs and not csvs:
        print(f"{C.RED}ERRO: Nenhum arquivo PDF ou CSV encontrado no diretório: {Path.cwd()}{C.END}")
        return None

    if csvs and not pdfs:
        csv_path = _priorizar_csv(csvs)
        print(f"{C.GREEN}✓ Arquivo CSV encontrado: {csv_path.name}{C.END}")
        return csv_path

    if pdfs and not csvs:
        pdf_path = pdfs[0]
        if len(pdfs) > 1:
            print(f"{C.YELLOW}AVISO: Múltiplos arquivos PDF encontrados. Usando o primeiro: {pdf_path.name}{C.END}")
        else:
            print(f"{C.GREEN}✓ Arquivo PDF encontrado: {pdf_path.name}{C.END}")
        return pdf_path

    csv_path = _priorizar_csv(csvs)
    print(f"{C.YELLOW}AVISO: PDFs e CSVs encontrados. Usando o CSV: {csv_path.name}{C.END}")
    return csv_path

def extrair_texto_do_pdf(pdf_path: Path) -> str:
    print(f"{C.CYAN}📄 Extraindo texto do arquivo PDF...{C.END}")
    if fitz is None:
        print(f"{C.RED}ERRO: PyMuPDF (pymupdf) não está instalado. Não é possível ler o PDF.{C.END}")
        return ""
    texto_completo = ""
    try:
        with _abrir_pdf_com_senha(pdf_path) as doc:
            for i, page in enumerate(doc):
                texto_completo += page.get_text("text")
                print(f"   Lendo página {i+1}/{len(doc)}...")
        print(f"{C.GREEN}✓ Extração de texto concluída com sucesso.{C.END}")
        return texto_completo
    except Exception as e:
        print(f"{C.RED}ERRO ao tentar ler o arquivo PDF: {e}{C.END}")
        return ""

def extrair_data_vencimento(texto_pdf: str) -> tuple[int, int]:
    print(f"{C.CYAN}🗓️ Identificando data de vencimento da fatura...{C.END}")
    texto_pdf = _normalizar_hifens(texto_pdf)

    match = re.search(r"vencimento\s+em:\s*(\d{2}/\d{2}/\d{4})", texto_pdf, re.IGNORECASE)
    if not match:
        linhas = texto_pdf.splitlines()
        for i, linha in enumerate(linhas):
            if re.search(r"vencimento\s+em", linha, re.IGNORECASE):
                for prox in linhas[i:i + 15]:
                    m_data = re.search(r"\b\d{2}/\d{2}/\d{4}\b", prox)
                    if m_data:
                        match = m_data
                        break
                if match:
                    break

    if not match:
        match = re.search(r"\b\d{2}/\d{2}/\d{4}\b", texto_pdf[:3000])

    if match:
        data_txt = match.group(1) if match.lastindex else match.group(0)
        _, mes_txt, ano_txt = data_txt.split("/")
        mes = int(mes_txt)
        ano = int(ano_txt)
        print(f"{C.GREEN}✓ Vencimento identificado: Mês {mes}, Ano {ano}.{C.END}")
        return mes, ano

    print(f"{C.YELLOW}AVISO: Não foi possível encontrar a data de vencimento. Usando data atual como fallback.{C.END}")
    hoje = datetime.now()
    return hoje.month, hoje.year

def _limpar_descricao(desc: str) -> str:
    desc = _normalizar_hifens(desc)
    d = re.sub(r'\s+', ' ', desc).strip()
    # remove frases de metadado caso venham coladas na linha
    d = re.sub(r'Parcela\s+Lojista\s+Visa\s+Parc\.?\s*\d{1,2}\s*/\s*\d{1,2}', '', d, flags=re.IGNORECASE)
    d = re.sub(r'Compra\s+a\s+Vista', '', d, flags=re.IGNORECASE)
    d = re.sub(r'\bPARC\.?\s*\d{1,2}\s*/\s*\d{1,2}\b', '', d, flags=re.IGNORECASE)
    d = re.sub(r'IOF\s+Transacoes\s+Exterior\s*R\$', '', d, flags=re.IGNORECASE)
    d = re.sub(r'Compra\s+Exterior\s*R\$\s*[-]\s*Visa', '', d, flags=re.IGNORECASE)
    d = re.sub(r'\s{2,}', ' ', d).strip(' -')
    return d

def _append_frac(base: str, frac: Optional[str]) -> str:
    if not frac:
        return base
    if re.search(r'\(\s*\d{1,2}\s*/\s*\d{1,2}\s*\)\s*$', base):
        return base
    return f"{base} ({frac})"

def _normalizar_chave_csv(chave: str) -> str:
    chave = chave.strip().lower().replace("\ufeff", "")
    chave = unicodedata.normalize("NFKD", chave)
    return "".join(c for c in chave if not unicodedata.combining(c))

def _normalizar_data_csv(data_raw: str) -> str:
    if not data_raw:
        return ""
    data_raw = data_raw.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(data_raw, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return data_raw

def _parse_valor_csv(valor_raw: Optional[str]) -> Optional[float]:
    if valor_raw is None:
        return None
    valor_txt = str(valor_raw).strip()
    if not valor_txt:
        return None
    valor_txt = valor_txt.replace("R$", "").strip()
    negativo = False
    if valor_txt.startswith("(") and valor_txt.endswith(")"):
        negativo = True
        valor_txt = valor_txt[1:-1].strip()
    if "," in valor_txt and "." in valor_txt:
        valor_txt = valor_txt.replace(".", "").replace(",", ".")
    elif "," in valor_txt:
        valor_txt = valor_txt.replace(",", ".")
    try:
        valor_num = float(valor_txt)
    except Exception:
        return None
    return -valor_num if negativo else valor_num

def _inferir_cartao_csv(csv_path: Path) -> str:
    match = re.search(r"(\d{4})(?!.*\d)", csv_path.stem)
    if match:
        return f"CSV.{match.group(1)}"
    return "CSV"

def _decidir_inversao_sinal_csv(registros: List[Dict[str, Any]]) -> bool:
    positivos = 0
    negativos = 0
    for item in registros:
        valor = item.get("valor_num")
        if valor is None:
            continue
        if valor > 0:
            positivos += 1
        elif valor < 0:
            negativos += 1

    for item in registros:
        descricao = (item.get("descricao_raw") or "").upper()
        valor = item.get("valor_num")
        if valor is None:
            continue
        if "PAGAMENTO" in descricao:
            return valor < 0

    if positivos and not negativos:
        return True
    if negativos and not positivos:
        return False
    return positivos >= negativos

def parsear_transacoes_csv(csv_path: Path) -> Dict[str, List[Dict[str, Any]]]:
    print(f"{C.CYAN}🧾 Lendo transações do CSV...{C.END}")
    registros: List[Dict[str, Any]] = []

    try:
        with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            if not reader.fieldnames:
                print(f"{C.RED}ERRO: CSV sem cabeçalho reconhecível.{C.END}")
                return {}

            for row in reader:
                norm_row = {}
                for chave, valor in row.items():
                    if chave is None:
                        continue
                    norm_row[_normalizar_chave_csv(chave)] = valor

                data_raw = norm_row.get("data") or norm_row.get("date")
                desc_raw = norm_row.get("lancamento") or norm_row.get("descricao") or norm_row.get("descricao_da_transacao")
                valor_raw = norm_row.get("valor") or norm_row.get("amount")

                if not data_raw and not desc_raw and not valor_raw:
                    continue

                registros.append({
                    "data_raw": data_raw or "",
                    "descricao_raw": desc_raw or "",
                    "valor_num": _parse_valor_csv(valor_raw),
                })
    except Exception as e:
        print(f"{C.RED}ERRO ao ler o CSV: {e}{C.END}")
        return {}

    if not registros:
        print(f"{C.YELLOW}AVISO: Nenhuma linha válida encontrada no CSV.{C.END}")
        return {}

    inverter_sinal = _decidir_inversao_sinal_csv(registros)
    cartao_key = _inferir_cartao_csv(csv_path)
    transacoes_por_cartao: Dict[str, List[Dict[str, Any]]] = {cartao_key: []}

    for item in registros:
        valor_num = item.get("valor_num")
        if valor_num is None:
            continue
        if inverter_sinal:
            valor_num = -valor_num

        data_norm = _normalizar_data_csv(item.get("data_raw", ""))
        desc_raw = item.get("descricao_raw", "")
        m_local_frac = re.search(r'\bPARC\.?\s*(\d{1,2}\s*/\s*\d{1,2})\b', desc_raw, flags=re.IGNORECASE)
        local_frac = re.sub(r'\s+', '', m_local_frac.group(1)) if m_local_frac else None
        descricao_base = _limpar_descricao(desc_raw)
        descricao_final = _append_frac(descricao_base, local_frac)

        valor_final = str(int(valor_num)) if valor_num == int(valor_num) else f"{valor_num:.2f}"
        transacoes_por_cartao[cartao_key].append({
            "data": data_norm,
            "descricao": descricao_final,
            "valor": valor_final,
            "conta": "BRB",
            "categoria": "A classificar",
        })

    total = len(transacoes_por_cartao[cartao_key])
    print(f"{C.GREEN}✓ Análise concluída. {total} transações extraídas.{C.END}")
    return transacoes_por_cartao

# ==============================================================================
# PARSER COM ALINHAMENTO DE METADADOS ("PARCELA LOJISTA..." / "COMPRA A VISTA")
# ==============================================================================

def parsear_transacoes(pdf_path: Path, mes_venc: int, ano_venc: int) -> Dict[str, List[Dict[str, Any]]]:
    """
    Motor com associação estável entre metadados e linhas de transação:
    • Varre por página, preservando a ordem de aparição.
    • Alinha cada metadado ("Parcela Lojista Visa Parc.x/y", "Compra a Vista", "IOF Transacoes Exterior R$",
      "Compra Exterior R$ - Visa") à PRÓXIMA linha de transação pendente.
    • Se o metadado vier ANTES das linhas (como ocorre em alguns blocos do PDF), ele é enfileirado
      e aplicado à primeira linha que surgir (uma por uma).
    • Se vier DEPOIS das linhas, consumirá a mais antiga pendente (o que mantém o pareamento na ordem).
    """
    print(f"{C.CYAN}💳 Extraindo transações com alinhamento de metadados por página...{C.END}")
    if fitz is None:
        print(f"{C.RED}ERRO: PyMuPDF (pymupdf) não está instalado. Não é possível processar o PDF.{C.END}")
        return {}

    transacoes_por_cartao: Dict[str, List[Dict[str, Any]]] = {}

    def inferir_ano(mes_tx: int) -> int:
        return ano_venc if mes_tx <= mes_venc else (ano_venc - 1)

    try:
        with _abrir_pdf_com_senha(pdf_path) as doc:
            pat_card = re.compile(r"(\d{4}(?:\.\*{4}\.\*{4}|\*{2}\.\*{6})\.\d{4})")
            pat_tx   = re.compile(r"(?m)^(\d{2}/\d{2})\s+([\s\S]+?)\s*(R\$\s*[\d\.,]+)\s*([+\-\u2212])")
            pat_mk_parc = re.compile(r"Parcela\s+Lojista\s+Visa\s+Parc\.?\s*(\d{1,2}\s*/\s*\d{1,2})", re.IGNORECASE)
            pat_mk_vista = re.compile(r"Compra\s+a\s+Vista", re.IGNORECASE)
            pat_mk_ignore = re.compile(r"IOF\s+Transacoes\s+Exterior\s*R\$|Compra\s+Exterior\s*R\$\s*[-−]\s*Visa", re.IGNORECASE)

            current_card_global: Optional[str] = None
            transacoes_sem_cartao = 0

            for page in doc:
                texto_pagina = _normalizar_hifens(page.get_text("text"))

                tokens = []
                tokens += [('card', m.start(), m.group(1), m) for m in pat_card.finditer(texto_pagina)]
                tokens += [('tx',   m.start(), m, m) for m in pat_tx.finditer(texto_pagina)]
                tokens += [('mkp',  m.start(), m.group(1), m) for m in pat_mk_parc.finditer(texto_pagina)]
                tokens += [('mkv',  m.start(), None, m) for m in pat_mk_vista.finditer(texto_pagina)]
                tokens += [('mki',  m.start(), None, m) for m in pat_mk_ignore.finditer(texto_pagina)]
                tokens.sort(key=lambda x: x[1])

                current_card = current_card_global
                pendentes_idx: List[int] = []
                pagina_transacoes: List[Dict[str, Any]] = []
                atribuicoes: Dict[int, Optional[str]] = {}  # None => consumiu metadado sem fração
                queue_markers: deque[Tuple[str, Optional[str]]] = deque()

                for typ, pos, content, mobj in tokens:
                    if typ == 'card':
                        current_card = content
                        current_card_global = content
                        continue

                    if typ == 'tx':
                        data_str, desc_raw, valor_raw, sinal = mobj.groups()
                        dia_str, mes_str = data_str.split('/')
                        ano_tx = inferir_ano(int(mes_str))

                        valor_txt = valor_raw.replace("R$", "").strip().replace(".", "").replace(",", ".")
                        try:
                            valor_num = float(valor_txt)
                        except Exception:
                            continue
                        if sinal == '+':
                            valor_num = -valor_num
                        valor_final = str(int(valor_num)) if valor_num == int(valor_num) else f"{valor_num:.2f}"

                        # Captura fração inline "PARC.x/y" antes da limpeza
                        m_local_frac = re.search(r'\bPARC\.?\s*(\d{1,2}\s*/\s*\d{1,2})\b', desc_raw, flags=re.IGNORECASE)
                        local_frac = re.sub(r'\s+', '', m_local_frac.group(1)) if m_local_frac else None
                        descricao_base = _limpar_descricao(desc_raw)

                        item = {
                            "data": f"{dia_str}/{mes_str}/{ano_tx}",
                            "descricao": descricao_base,
                            "valor": valor_final,
                            "conta": "BRB",
                            "categoria": "A classificar",
                            "_card": current_card,
                        }
                        idx = len(pagina_transacoes)
                        pagina_transacoes.append(item)

                        # Se há metadados enfileirados (vieram antes), consome 1 para esta linha
                        if queue_markers:
                            mk_typ, mk_frac = queue_markers.popleft()
                            atribuicoes[idx] = (re.sub(r'\s+', '', mk_frac) if mk_frac else None) if mk_typ == 'mkp' else None
                        else:
                            pendentes_idx.append(idx)
                        continue

                    # metadados: aplica ao próximo pendente; se não houver, enfileira
                    if typ in ('mkp', 'mkv', 'mki'):
                        frac = re.sub(r'\s+', '', content) if typ == 'mkp' else None
                        if pendentes_idx:
                            alvo = pendentes_idx.pop(0)
                            atribuicoes[alvo] = frac  # None para 'mkv'/'mki' => apenas consumir alinhamento
                        else:
                            queue_markers.append((typ, frac))

                # aplica as frações atribuídas
                for i, tx in enumerate(pagina_transacoes):
                    frac = atribuicoes.get(i, None)
                    # aplica fração vinda do metadado e/ou a fração inline (PARC.x/y)
                    if frac:
                        tx['descricao'] = _append_frac(tx['descricao'], frac)
                    if 'local_frac' in locals() and local_frac:
                        tx['descricao'] = _append_frac(tx['descricao'], local_frac)

                    card_key = tx.pop("_card", None)
                    if not card_key:
                        transacoes_sem_cartao += 1
                        continue
                    transacoes_por_cartao.setdefault(card_key, []).append(tx)

        total = sum(len(v) for v in transacoes_por_cartao.values())
        if transacoes_sem_cartao:
            print(f"{C.YELLOW}AVISO: {transacoes_sem_cartao} transações fora do bloco de cartão foram ignoradas.{C.END}")
        print(f"{C.GREEN}✓ Análise concluída. {total} transações extraídas.{C.END}")
        return transacoes_por_cartao

    except Exception as e:
        print(f"{C.RED}ERRO durante a extração: {e}{C.END}")
        return {}

# ==============================================================================
# IA (Opcional) — Classificação de categoria
# ==============================================================================

async def classificar_transacao(transacao: Dict[str, Any], cliente) -> str:
    descricao = transacao['descricao']
    prompt = f"""
    Classifique a seguinte descrição de despesa de cartão de crédito em uma das categorias abaixo.
    Responda APENAS com o nome exato da categoria.

    Categorias Válidas: {', '.join(CATEGORIAS_PERMITIDAS)}

    Descrição da Despesa: "{descricao}"

    Categoria:
    """
    if not cliente:
        return "Outros"

    for attempt in range(MAX_RETRIES):
        try:
            response = await cliente.chat.completions.create(
                model=MODELO_IA,
                messages=[
                    {"role": "system", "content": "Você é um assistente de finanças pessoais preciso e eficiente."},
                    {"role": "user", "content": prompt}
                ],
            )
            categoria = response.choices[0].message.content.strip().replace('"', '').replace('.', '')
            return categoria if categoria in CATEGORIAS_PERMITIDAS else "Outros"
        except Exception:
            delay = INITIAL_RETRY_DELAY * (2 ** attempt)
            print(f"{C.YELLOW}AVISO: Erro de API para '{descricao}' (tentativa {attempt + 1}/{MAX_RETRIES}). Tentando novamente em {delay}s...{C.END}")
            await asyncio.sleep(delay)
    return "Outros"



def salvar_csvs(
    transacoes_por_cartao: Dict[str, List[Dict[str, Any]]],
    base_name: str = "fatura_cartao_final",
    output_dir: Optional[Path] = None,
):
    print(f"\n{C.CYAN}{C.BOLD}💾 Salvando arquivos CSV...{C.END}")
    if not any(transacoes_por_cartao.values()):
        print(f"{C.YELLOW}Nenhuma transação para salvar.{C.END}")
        return
    destino = output_dir or Path.cwd()
    destino.mkdir(parents=True, exist_ok=True)
    for cartao, transacoes in transacoes_por_cartao.items():
        if not transacoes:
            print(f"{C.YELLOW}AVISO: Sem dados para o cartão {cartao}.{C.END}")
            continue
        if cartao:
            m_last4 = re.search(r"(\d{4})(?!.*\d)", cartao)
            last4 = m_last4.group(1) if m_last4 else "None"
        else:
            last4 = "None"
        nome_arquivo = f"{base_name}_{last4}.csv"
        try:
            caminho_saida = destino / nome_arquivo
            with open(caminho_saida, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_ALL)
                writer.writerow(["Data", "Descrição", "Valor", "Conta", "Categoria"])
                for t in transacoes:
                    writer.writerow([t.get('data',''), t.get('descricao',''), t.get('valor',''), t.get('conta',''), t.get('categoria','')])
            print(f"{C.GREEN}✓ Arquivo salvo: {caminho_saida}{C.END}")
        except IOError as e:
            print(f"{C.RED}ERRO ao salvar {nome_arquivo}: {e}{C.END}")

# ==============================================================================
# MAIN
# ==============================================================================

async def main():
    inicio = datetime.now()
    print(f"{C.BOLD}{C.BLUE}{'='*80}{C.END}")
    print(f"{C.BOLD}{C.BLUE}  EXTRATOR DE FATURA — Parcelas por Metadado v2.0{C.END}")
    print(f"{C.BOLD}{C.BLUE}{'='*80}{C.END}\n")

    try:
        entrada_path = selecionar_arquivo_entrada()
        if not entrada_path:
            return
        output_dir = entrada_path.parent
        if entrada_path.suffix.lower() == ".pdf":
            texto_pdf = extrair_texto_do_pdf(entrada_path)
            if not texto_pdf:
                return

            mes_venc, ano_venc = extrair_data_vencimento(texto_pdf)
            transacoes_por_cartao = parsear_transacoes(entrada_path, mes_venc, ano_venc)
            output_dir = entrada_path.parent / _nome_pasta_competencia(mes_venc, ano_venc)
            print(f"{C.CYAN}📁 Pasta de saída definida para: {output_dir}{C.END}")
        elif entrada_path.suffix.lower() == ".csv":
            transacoes_por_cartao = parsear_transacoes_csv(entrada_path)
        else:
            print(f"{C.RED}ERRO: Formato de arquivo não suportado.{C.END}")
            return

        await processar_categorias_em_lote(transacoes_por_cartao)
        salvar_csvs(transacoes_por_cartao, output_dir=output_dir)

    except Exception:
        print(f"\n{C.RED}{C.BOLD}{'!'*80}{C.END}")
        print(f"{C.RED}{C.BOLD}Ocorreu um erro inesperado durante a execução:{C.END}")
        print(f"{C.RED}{traceback.format_exc()}{C.END}")
        print(f"{C.RED}{C.BOLD}{'!'*80}{C.END}")
    finally:
        fim = datetime.now()
        print(f"\n{C.BOLD}{C.GREEN}Execução finalizada em: {fim - inicio}{C.END}")

if __name__ == "__main__":
    asyncio.run(main())
