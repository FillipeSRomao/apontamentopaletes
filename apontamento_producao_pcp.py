"""Apontamento de produção PCP (versão 5 - lote da planilha, sem GS1).

Gera a etiqueta de palete em PDF, no layout do ROMANEIO_DE_PALETE.xlsx (logo, produto,
validade, lote, pesos...), e abre o PDF numa nova aba do navegador: a impressão é feita
a partir dele (Ctrl+P). O lote segue a mesma regra da planilha. Registra a produção em
SQLite e exporta relatório para Excel.

Os textos, a validade e os pesos de cada produto vêm da aba "Base" do
ROMANEIO_DE_PALETE.xlsx, que deve ficar na mesma pasta deste arquivo.
O logo é lido de logo_oasis.png (mesma pasta).

Para alterar TURNOS, PRODUTOS/QUANTIDADES ou CORES, procure pelo bloco
">>> ÁREA DE CONFIGURAÇÃO" mais abaixo neste arquivo.

Variáveis de ambiente opcionais:
    PCP_DB            caminho do banco SQLite (padrão: ./pcp_producao.db)
    PCP_ROMANEIO      caminho do ROMANEIO_DE_PALETE.xlsx (padrão: mesma pasta)
    PCP_SECRET_KEY    chave da sessão do Flask (troque em produção)
    PCP_HOST/PCP_PORT endereço e porta do servidor (padrão: 0.0.0.0:5000)
"""

from __future__ import annotations

import io
import logging
import os
import re
import sqlite3
import time as tempo
import unicodedata
from contextlib import closing
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path

from flask import (
    Flask,
    abort,
    flash,
    g,
    redirect,
    render_template_string,
    request,
    send_file,
    url_for,
)
from openpyxl import Workbook, load_workbook
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

# --------------------------------------------------------------------------
# Configuração técnica (raramente muda)
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("PCP_DB", BASE_DIR / "pcp_producao.db"))
LABELS_DIR = BASE_DIR / "etiquetas"
PAGE_SIZE = landscape(A4)  # o romaneio é A4 deitado
MARGEM = 28
ROMANEIO_XLSX = Path(os.environ.get("PCP_ROMANEIO", BASE_DIR / "ROMANEIO_DE_PALETE.xlsx"))
LOGO_PATH = BASE_DIR / "logo_oasis.png"
DIAS_GUARDAR_PDF = 30
LIMITE_PAINEL = 200  # linhas exibidas no painel do PCP (o Excel exporta tudo)


FORMATO_DATA = "%d/%m/%Y"
FORMATO_DATA_HORA = "%d/%m/%Y %H:%M:%S"  # mantido por compatibilidade com o banco antigo

STATUS_PENDENTE = "Pendente"
STATUS_PRODUCAO = "Em Produção"
STATUS_REIMPRESSAO = "Reimpressão"
STATUS_ERRO = "Erro de Impressão"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("pcp")


# --------------------------------------------------------------------------
# Modelos (Turno e Produto) — não precisa mexer aqui
# --------------------------------------------------------------------------
@dataclass(frozen=True)
class Turno:
    nome: str
    inicio: time
    fim: time

    @property
    def periodo(self) -> str:
        return f"{self.inicio:%H:%M} às {self.fim:%H:%M}"

    @property
    def digito(self) -> str:
        """Número do turno ("Turno 2" -> "2"): é o último dígito do lote."""
        achado = re.search(r"\d+", self.nome)
        if not achado:
            raise ValueError(f"O nome do turno {self.nome!r} precisa ter o número do turno (ex.: 'Turno 2').")
        return achado.group()

    def contem(self, hora: time) -> bool:
        if self.inicio <= self.fim:
            return self.inicio <= hora <= self.fim
        return hora >= self.inicio or hora <= self.fim  # turno que vira a meia-noite


def _slug(texto: str) -> str:
    sem_acento = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", sem_acento.lower()).strip("-")


# --------------------------------------------------------------------------
# Base do romaneio: aba "Base" do ROMANEIO_DE_PALETE.xlsx (somente leitura)
# --------------------------------------------------------------------------
@dataclass(frozen=True)
class ItemBase:
    """Uma linha da aba Base: textos da etiqueta, validade e pesos de um produto."""

    cod: str
    desc1: str  # linha de cima da etiqueta (ex.: AGUA MINERAL LA PRIORI)
    desc2: str  # linha grande da etiqueta (ex.: 500 ML SEM GAS)
    validade_dias: int
    qtd_base: float  # quantidade a que os dois pesos abaixo se referem (palete cheio)
    peso_liquido_base: float
    peso_bruto_base: float


def _cod(valor: object) -> str:
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()


def _numero(valor: object, onde: str) -> float:
    try:
        return float(valor)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        raise ValueError(f"Aba Base, {onde}: esperava um número e veio {valor!r}.") from None


def _dias_validade(valor: object, cod: str) -> int:
    """A coluna VALIDADE traz a fórmula '=$E$1+365' (vale o 365) ou direto o nº de dias."""
    if isinstance(valor, (int, float)) and not isinstance(valor, bool) and 0 < valor < 3650:
        return int(valor)
    if isinstance(valor, str) and valor.startswith("="):
        achado = re.search(r"\+\s*(\d+)\s*$", valor)
        if achado:
            return int(achado.group(1))
    raise ValueError(
        f"Aba Base, código {cod}: a coluna VALIDADE deve ser '=$E$1+365' "
        f"(ou outro nº de dias) e veio {valor!r}."
    )


def carregar_base_romaneio(caminho: Path = ROMANEIO_XLSX) -> dict[str, ItemBase]:
    """Lê a aba 'Base' do romaneio, pelo NOME das colunas. Nunca altera a planilha."""
    if not caminho.is_file():
        raise FileNotFoundError(
            f"Planilha do romaneio não encontrada: {caminho}\n"
            "Coloque o ROMANEIO_DE_PALETE.xlsx na mesma pasta do programa "
            "(ou informe o caminho na variável PCP_ROMANEIO)."
        )
    try:
        wb = load_workbook(caminho, read_only=True, data_only=False)
    except OSError as exc:
        raise OSError(f"Não consegui abrir {caminho} ({exc}). Ela está aberta em outro programa?") from exc
    try:
        if "Base" not in wb.sheetnames:
            raise ValueError(f"A planilha {caminho.name} não tem a aba 'Base'.")
        linhas = list(wb["Base"].iter_rows(values_only=True))
    finally:
        wb.close()

    indices: dict[str, int] = {}
    inicio = None
    for i, linha in enumerate(linhas[:10]):  # acha a linha de títulos (a que tem "CÓD")
        chaves = [_slug(str(v)) if v is not None else "" for v in linha]
        if "cod" in chaves:
            indices = {k: n for n, k in enumerate(chaves) if k}
            inicio = i + 1
            break
    if inicio is None:
        raise ValueError("Aba Base: não achei a linha de títulos (coluna 'CÓD').")
    obrigatorias = {
        "cod": "CÓD", "descricao-01": "DESCRIÇÃO - 01", "descricao-02": "DESCRIÇÃO - 02",
        "validade": "VALIDADE", "qnt": "QNT",
        "peso-liquido-dos-produtos": "PESO LIQUIDO DOS PRODUTOS", "peso-bruto-pallet": "PESO BRUTO - PALLET",
    }
    faltam = [nome for chave, nome in obrigatorias.items() if chave not in indices]
    if faltam:
        raise ValueError(f"Aba Base: faltam as colunas {', '.join(faltam)}.")

    def pega(linha: tuple, chave: str) -> object:
        i = indices[chave]
        return linha[i] if i < len(linha) else None

    base: dict[str, ItemBase] = {}
    for linha in linhas[inicio:]:
        bruto = pega(linha, "cod")
        cod = "" if bruto in (None, "") else _cod(bruto)
        if not cod.isdigit():
            continue  # linha vazia ou de totais
        if cod in base:
            raise ValueError(f"Aba Base: o código {cod} aparece duas vezes.")
        onde = f"código {cod}"
        qtd = _numero(pega(linha, "qnt"), f"{onde}, coluna QNT")
        liq = _numero(pega(linha, "peso-liquido-dos-produtos"), f"{onde}, PESO LIQUIDO")
        bru = _numero(pega(linha, "peso-bruto-pallet"), f"{onde}, PESO BRUTO")
        if qtd <= 0 or liq < 0 or bru < liq:
            raise ValueError(f"Aba Base, {onde}: QNT deve ser maior que 0 e o peso bruto não pode ser menor que o líquido.")
        base[cod] = ItemBase(
            cod=cod,
            desc1=str(pega(linha, "descricao-01") or "").strip(),
            desc2=str(pega(linha, "descricao-02") or "").strip(),
            validade_dias=_dias_validade(pega(linha, "validade"), cod),
            qtd_base=qtd,
            peso_liquido_base=liq,
            peso_bruto_base=bru,
        )
    if not base:
        raise ValueError("Aba Base: nenhum produto encontrado.")
    return base


BASE_ROMANEIO = carregar_base_romaneio()


@dataclass(frozen=True)
class Produto:
    codigo: str  # identificador interno (usado nas URLs e formulários)
    numero: str  # número da linha de produção, ex.: "05"
    nome: str
    unidade: str
    qtd_padrao: int
    qtd_max: int
    cod: str  # código do produto no romaneio/ERP (coluna CÓD da aba Base)
    item: ItemBase  # dados desse código na aba Base
    usa_selo: bool = True  # False = produto que nunca leva selo (quadro do selo fica em branco)

    @property
    def linha(self) -> str:
        return f"Linha {self.numero}"

    @property
    def cor(self) -> str:
        return CORES_LINHA.get(self.numero, "#6c757d")


FATOR_QTD_MAX = 3  # a tela aceita até 3x a quantidade padrão (se não informar qtd_max)


def produto(
    numero: str,
    nome: str,
    unidade: str,
    qtd_padrao: int,
    cod: str,
    qtd_max: int | None = None,
    usa_selo: bool = True,
) -> Produto:
    item = BASE_ROMANEIO.get(str(cod).strip())
    if item is None:
        raise ValueError(
            f"Produto {nome!r}: o código {cod!r} não existe na aba Base de {ROMANEIO_XLSX.name}. "
            f"Cadastre-o lá primeiro. Códigos existentes: {', '.join(BASE_ROMANEIO)}."
        )
    codigo = f"{numero}-{_slug(nome)}"
    return Produto(
        codigo=codigo,
        numero=numero,
        nome=nome,
        unidade=unidade,
        qtd_padrao=qtd_padrao,
        qtd_max=qtd_max or qtd_padrao * FATOR_QTD_MAX,
        cod=item.cod,
        item=item,
        usa_selo=usa_selo,
    )


# ==========================================================================
# >>> ÁREA DE CONFIGURAÇÃO — é aqui que você edita no dia a dia
# ==========================================================================

# --- 1) TURNOS -------------------------------------------------------------
# Formato: Turno("Nome", time(hora, minuto) de início, time(hora, minuto) de fim)
# Um turno que passa da meia-noite (início maior que o fim) é tratado sozinho.
# O nome PRECISA conter o número do turno (ex.: "Turno 2"): ele é o último dígito do lote.
TURNOS = (
    Turno("Turno 2", time(7, 0), time(16, 46)),
    Turno("Turno 3", time(16, 48), time(2, 16)),
)

# --- 2) PRODUTOS POR LINHA -------------------------------------------------
# Uma linha por produto:
#   produto("nº da linha", "Nome do produto", "unidade", quantidade padrão, cod="CÓD do romaneio")
# O cod TEM de existir na coluna CÓD da aba Base do ROMANEIO_DE_PALETE.xlsx: de lá vêm
# a descrição da etiqueta, a validade e os pesos. Produto novo = linha nova na Base + linha aqui.
# Opcional: qtd_max=... (limite da tela).
# usa_selo=False -> produto que NUNCA leva selo: não aparece o botão de selo no cartão
# e o quadro do selo da etiqueta mostra TEXTO_SELO_NAO_APLICA. Sem isso, o cartão mostra SEM SELO / COM SELO.
CATALOGO = (
    produto("02", "Copo 200mL", "Caixas", 80, cod="5", usa_selo=False),
    produto("03", "Garrafa 5L", "Garrafas", 126, cod="3"),
    produto("05", "Pacote 1,5L Com Gás", "Pacotes", 84, cod="4626"),
    produto("05", "Pacote 1,5L Sem Gás", "Pacotes", 84, cod="4625"),
    produto("05", "Pacote Saborizada Limão", "Pacotes", 126, cod="4671", usa_selo=False),
    produto("05", "Pacote Saborizada Tangerina", "Pacotes", 126, cod="4673", usa_selo=False),
    produto("05", "Pacote Saborizada Frutas Vermelhas", "Pacotes", 126, cod="4670", usa_selo=False),
    produto("05", "Pacote Saborizada Maçã Verde", "Pacotes", 126, cod="4672", usa_selo=False),
    produto("05", "Pacote Saborizada Abacaxi com Hortelã", "Pacotes", 126, cod="4422", usa_selo=False),
    produto("05", "Pacote Saborizada Uva", "Pacotes", 126, cod="4674", usa_selo=False),
    produto("06", "Pacote 500mL Com Gás", "Pacotes", 140, cod="4668"),
    produto("06", "Pacote 500mL Sem Gás", "Pacotes", 140, cod="4666"),
    produto("06", "Pacote 355mL Com Gás", "Pacotes", 160, cod="4669"),
    produto("06", "Pacote 355mL Sem Gás", "Pacotes", 160, cod="4667"),
)

# --- 3) COR DE CADA LINHA NA TELA ------------------------------------------
CORES_LINHA = {
    "02": "#007bff",
    "03": "#17a2b8",
    "05": "#28a745",
    "06": "#d39e00",
}

# --- 4) LAYOUT DA TELA DO OPERADOR -----------------------------------------
# Quantos cartões de produto ficam lado a lado, no máximo, dentro de cada linha.
# As próprias linhas (02, 03, 05, 06) também ficam lado a lado quando couberem.
COLUNAS_POR_LINHA = 4
LARGURA_CARTAO_PX = 205

# --- 5) ETIQUETA (layout do ROMANEIO_DE_PALETE.xlsx) -----------------------
TITULO_ETIQUETA = "LIQUIDA"  # faixa grande do topo da etiqueta
SELOS = ("SEM SELO", "COM SELO")  # opções do botão de selo de cada cartão (a primeira é a padrão)
SELO_COM = SELOS[-1]  # qual das opções é a "com selo" (ganha o destaque visual na etiqueta)
TEXTO_SELO_NAO_APLICA = "NÃO SE APLICA"  # quadro do selo nos produtos com usa_selo=False
CAMPO_NAO_DEFINIDO = "N/D"  # texto de RESERVADO / Nº PEDIDO / CLIENTE

# --- 6) LOTE (mesma regra da planilha ROMANEIO_DE_PALETE.xlsx) --------------
# Lote impresso = DDD + AA + linha + código + final + " - H:MM". Exemplo: 31926646662 - 10:58
#   DDD   dia do ano (001 a 366) do DIA DE PRODUÇÃO     AA   ano com 2 dígitos
#   linha nº da linha sem zero à esquerda (06 -> 6)     código  CÓD da aba Base (mínimo 3 dígitos)
#   final número do turno (Turno 2 -> 2, Turno 3 -> 3), o "2" da célula Q7 da planilha
#   H:MM  hora e minuto em que a etiqueta foi impressa
# Nada a configurar aqui: a regra é fixa; só o nome dos TURNOS precisa ter o número.

# ==========================================================================
# >>> FIM DA ÁREA DE CONFIGURAÇÃO
# ==========================================================================

PRODUTOS: dict[str, Produto] = {p.codigo: p for p in CATALOGO}
NOMES_TURNOS = {t.nome for t in TURNOS}
TURNO_POR_NOME = {t.nome: t for t in TURNOS}
for _turno in TURNOS:
    _turno.digito  # falha cedo se o nome do turno não tiver número

if len(PRODUTOS) != len(CATALOGO):
    raise ValueError("CATALOGO tem o mesmo produto repetido na mesma linha.")
if len(NOMES_TURNOS) != len(TURNOS) or not TURNOS:
    raise ValueError("TURNOS vazio ou com nomes repetidos.")


def turno_em_andamento(agora: datetime | None = None) -> Turno:
    """Turno em andamento. Comparação em minutos (16:46:59 ainda é 16:46).

    Fora de qualquer janela (16:47 e 02:17–06:59) vale o turno que começou
    mais recentemente; o operador pode trocar no seletor da tela.
    """
    hora = (agora or datetime.now()).replace(second=0, microsecond=0).time()
    for turno in TURNOS:
        if turno.contem(hora):
            return turno
    minutos = hora.hour * 60 + hora.minute

    def desde_o_inicio(t: Turno) -> int:
        return (minutos - (t.inicio.hour * 60 + t.inicio.minute)) % 1440

    return min(TURNOS, key=desde_o_inicio)


def turno_atual(agora: datetime | None = None) -> str:
    return turno_em_andamento(agora).nome


def dia_producao(agora: datetime) -> date:
    """Dia a que a produção pertence.

    Se o turno em andamento atravessa a meia-noite e já passamos dela (ou
    estamos na folga antes do próximo início), o palete conta no dia em que
    o turno COMEÇOU. Segue automaticamente os horários de TURNOS.
    """
    turno = turno_em_andamento(agora)
    hora = agora.replace(second=0, microsecond=0).time()
    if turno.inicio > turno.fim and hora < turno.inicio:
        return agora.date() - timedelta(days=1)
    return agora.date()


def lote_prefixo(prod: Produto, agora: datetime) -> str:
    """Lote sem o último dígito: DDD + AA + linha + código, igual à planilha.

    DDD e AA vêm do DIA DE PRODUÇÃO (turno que passa da meia-noite conta no dia em que
    começou). O código tem no mínimo 3 dígitos (3 -> 003), como na planilha.
    """
    dia = dia_producao(agora)
    ddd = f"{dia.timetuple().tm_yday:03d}"
    return f"{ddd}{dia:%y}{int(prod.numero)}{prod.cod.zfill(3)}"


def lote_do_dia(prod: Produto, agora: datetime, turno: str) -> str:
    """Lote SEM a hora: prefixo + número do turno (o "2" da célula Q7 da planilha)."""
    return lote_prefixo(prod, agora) + TURNO_POR_NOME[turno].digito


def lote_completo(prod: Produto, agora: datetime, turno: str) -> str:
    """Lote impresso e gravado: lote do dia + " - H:MM" (hora sem zero à esquerda)."""
    return f"{lote_do_dia(prod, agora, turno)} - {agora.hour}:{agora.minute:02d}"


def dados_romaneio(prod: Produto, quantidade: int, dia: date) -> tuple[str, float, float]:
    """(validade, peso líquido, peso bruto) do palete.

    Validade = dia de produção + dias de validade da aba Base. A Base informa o peso do
    palete CHEIO: o líquido é proporcional à quantidade e o bruto soma a tara do palete
    (peso bruto - peso líquido da Base).
    """
    item = prod.item
    validade = (dia + timedelta(days=item.validade_dias)).strftime(FORMATO_DATA)
    peso_liquido = round(item.peso_liquido_base / item.qtd_base * quantidade, 2)
    tara = item.peso_bruto_base - item.peso_liquido_base
    return validade, peso_liquido, round(peso_liquido + tara, 2)


# --------------------------------------------------------------------------
# Banco de dados
# --------------------------------------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("PCP_SECRET_KEY", "troque-esta-chave-em-producao")


def get_db() -> sqlite3.Connection:
    """Uma conexão por requisição, fechada automaticamente no final."""
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH, timeout=10)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def fechar_db(_exc: BaseException | None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db() -> None:
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS producao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_hora TEXT,
                turno TEXT,
                linha TEXT,
                sku TEXT,
                lote TEXT,
                quantidade INTEGER,
                palete INTEGER,
                status TEXT,
                dia_producao TEXT,
                cod TEXT,
                validade TEXT,
                peso_liquido REAL,
                peso_bruto REAL,
                selo TEXT
            )
            """
        )
        # Bancos criados por versões anteriores ganham a coluna nova.
        colunas = {linha[1] for linha in conn.execute("PRAGMA table_info(producao)")}
        novas = {"dia_producao": "TEXT", "cod": "TEXT", "validade": "TEXT",
                 "peso_liquido": "REAL", "peso_bruto": "REAL", "selo": "TEXT", "palete": "INTEGER"}
        for nome_coluna, tipo in novas.items():
            if nome_coluna not in colunas:
                conn.execute(f"ALTER TABLE producao ADD COLUMN {nome_coluna} {tipo}")
        # Versões antigas numeravam o palete pelo SSCC (as colunas gtin/sscc continuam no banco,
        # só como histórico). Cada SSCC vira o nº do 1º registro dele; reimpressões herdam o nº.
        if "sscc" in colunas:
            conn.execute(
                "UPDATE producao SET palete = (SELECT MIN(p2.id) FROM producao p2 "
                "WHERE p2.sscc = producao.sscc) WHERE palete IS NULL AND sscc IS NOT NULL"
            )
        conn.execute("UPDATE producao SET palete = id WHERE palete IS NULL")
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_producao_linha_sku_status "
            "ON producao (linha, sku, status, id)"
        )
        try:  # um nº de palete nunca pode aparecer em dois "novos paletes"
            conn.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS ux_producao_palete_novo "
                "ON producao (palete) WHERE status = 'Em Produção'"
            )
        except sqlite3.IntegrityError:
            logger.warning("Banco antigo tem nº de palete duplicado; índice único não foi criado.")
        conn.commit()


SQL_INSERT = """
    INSERT INTO producao
        (data_hora, turno, linha, sku, lote, quantidade, palete, status, dia_producao,
         cod, validade, peso_liquido, peso_bruto, selo)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""


# --------------------------------------------------------------------------
# Etiqueta: PDF no layout do romaneio
# --------------------------------------------------------------------------
@dataclass(frozen=True)
class Etiqueta:
    linha: str
    produto: str
    turno: str
    lote: str
    quantidade: int
    unidade: str
    palete: int  # nº sequencial do palete (id do registro original)
    data_hora: str
    cod: str = ""
    desc1: str = ""  # linha de cima (ex.: AGUA MINERAL LA PRIORI)
    desc2: str = ""  # linha grande (ex.: 500 ML SEM GAS)
    fabricacao: str = ""
    validade: str = ""
    peso_liquido: float = 0.0
    peso_bruto: float = 0.0
    selo: str = SELOS[0]
    reimpressao: bool = False
    teste: bool = False


# Medidas copiadas da aba "Imprimir romaneio BLOK" do ROMANEIO_DE_PALETE.xlsx:
# colunas B..O (em pixels) e linhas 2..27 (em pontos). Assim o layout usa os MESMOS
# endereços de célula do Excel (ex.: "B7:F7" = campo CÓD).
_COLUNAS_PX = [80] * 13 + [120]
_LINHAS_PT = [15] * 5 + [51] + [15] * 12 + [45] * 3 + [24.75, 26.25, 21.75, 24.75, 27]
_LARGURA_ROMANEIO = sum(_COLUNAS_PX) * 0.75  # pontos
_ALTURA_ROMANEIO = sum(_LINHAS_PT)  # pontos
_LETRAS = "BCDEFGHIJKLMNO"


class _Grade:
    """Converte endereços de célula do Excel em retângulos no PDF."""

    def __init__(self, x0: float, y_topo: float, escala: float) -> None:
        self.escala = escala
        self.xs = [x0]
        for px in _COLUNAS_PX:
            self.xs.append(self.xs[-1] + px * 0.75 * escala)
        self.ys = [y_topo]
        for pt in _LINHAS_PT:
            self.ys.append(self.ys[-1] - pt * escala)

    def caixa(self, endereco: str) -> tuple[float, float, float, float]:
        """(x, y, largura, altura) do intervalo, com origem no canto inferior esquerdo."""
        achado = re.fullmatch(r"([B-O])(\d+)(?::([B-O])(\d+))?", endereco)
        if not achado:
            raise ValueError(f"Endereço de célula inválido: {endereco!r}")
        c1, r1, c2, r2 = achado.groups()
        i1, i2 = _LETRAS.index(c1), _LETRAS.index(c2 or c1)
        topo, base = int(r1) - 2, int(r2 or r1) - 1
        return self.xs[i1], self.ys[base], self.xs[i2 + 1] - self.xs[i1], self.ys[topo] - self.ys[base]


def _texto(
    c: canvas.Canvas,
    caixa: tuple[float, float, float, float],
    texto: str,
    tamanho: float,
    alinhar: str = "centro",
    cor: tuple[float, float, float] = (0, 0, 0),
    margem: float = 4,
    fonte: str = "Helvetica-Bold",
) -> None:
    """Escreve o texto dentro da caixa, diminuindo a fonte se ele não couber."""
    x, y, w, h = caixa
    util = max(w - 2 * margem, 1)
    largura = stringWidth(texto, fonte, tamanho)
    if largura > util:
        tamanho *= util / largura
    base = y + h / 2 - tamanho * 0.35  # centraliza na vertical
    c.setFillColorRGB(*cor)
    c.setFont(fonte, tamanho)
    if alinhar == "direita":
        c.drawRightString(x + w - margem, base, texto)
    elif alinhar == "esquerda":
        c.drawString(x + margem, base, texto)
    else:
        c.drawCentredString(x + w / 2, base, texto)
    c.setFillColorRGB(0, 0, 0)


def _hachurar(c: canvas.Canvas, caixa: tuple[float, float, float, float],
              passo: float = 9, espessura: float = 2.2) -> None:
    """Listras diagonais dentro da caixa: chamam a atenção de longe usando pouco toner."""
    x, y, w, h = caixa
    c.saveState()
    recorte = c.beginPath()
    recorte.rect(x, y, w, h)
    c.clipPath(recorte, stroke=0, fill=0)
    c.setLineWidth(espessura)
    deslocamento = -h
    while deslocamento < w:
        c.line(x + deslocamento, y, x + deslocamento + h, y + h)
        deslocamento += passo
    c.restoreState()


def _kg(valor: float) -> str:
    return f"{valor:.2f}".replace(".", ",") + " kg"


def gerar_pdf(et: Etiqueta, destino: Path) -> None:
    """Etiqueta no layout do romaneio."""
    pag_l, pag_a = PAGE_SIZE
    c = canvas.Canvas(str(destino), pagesize=PAGE_SIZE)
    c.setTitle(f"Etiqueta de palete {et.palete}")

    faixa_aviso, faixa_rodape = 18, 22
    area_l = pag_l - 2 * MARGEM
    area_a = pag_a - 2 * MARGEM - faixa_aviso - faixa_rodape
    esc = min(area_l / _LARGURA_ROMANEIO, area_a / _ALTURA_ROMANEIO)  # tamanho do romaneio na página
    x0 = (pag_l - _LARGURA_ROMANEIO * esc) / 2
    y_topo = pag_a - MARGEM - faixa_aviso
    g = _Grade(x0, y_topo, esc)

    aviso = "REIMPRESSÃO" if et.reimpressao else ("TESTE - NÃO USAR" if et.teste else "")
    if aviso:
        c.setFont("Helvetica-Bold", 13)
        c.drawRightString(x0 + _LARGURA_ROMANEIO * esc, y_topo + 9, aviso)

    # --- molduras (finas nas células, grossas nos blocos principais) ---------
    def moldura(endereco: str, espessura: float) -> None:
        x, y, w, h = g.caixa(endereco)
        c.setLineWidth(espessura)
        c.rect(x, y, w, h, stroke=1, fill=0)

    for endereco in ("B20:D20", "E20:G20", "B21:D21", "E21:G21", "B22:D22", "E22:G22",
                     "I23:L23", "M23:O23", "I24:L24", "M24:O24", "I25:L25", "M25:O25",
                     "I26:L26", "M26:O26", "I27:L27", "M27:O27", "B26:D27", "E26:E27"):
        moldura(endereco, 0.8)
    com_selo = et.selo == SELO_COM
    for endereco in ("B2:C6", "D2:O6", "B7:F7", "G7:O7", "B8:O19", "B20:G22", "B2:O27"):
        moldura(endereco, 1.6)
    moldura("H20:O22", 4 if com_selo else 1.6)  # quadro do selo: ocupa também a linha 22

    # --- cabeçalho -----------------------------------------------------------
    x, y, w, h = g.caixa("B2:C6")
    if LOGO_PATH.is_file():
        c.drawImage(ImageReader(str(LOGO_PATH)), x + 3, y + 3, w - 6, h - 6,
                    mask="auto", preserveAspectRatio=True, anchor="c")
    else:  # sem o arquivo do logo, imprime o nome no lugar
        _texto(c, g.caixa("B2:C6"), "OASIS", 24 * esc)
    _texto(c, g.caixa("D2:O6"), TITULO_ETIQUETA, 70 * esc)
    _texto(c, g.caixa("B7:F7"), f"CÓD: {et.cod}", 36 * esc)
    _texto(c, g.caixa("G7:O7"), et.desc1, 36 * esc)
    _texto(c, g.caixa("B8:O19"), et.desc2, 110 * esc)

    # --- reservado / pedido / cliente / selo --------------------------------
    for linha, rotulo in ((20, "RESERVADO:"), (21, "Nº PEDIDO:"), (22, "CLIENTE:")):
        _texto(c, g.caixa(f"B{linha}:D{linha}"), rotulo, 24 * esc, "direita")
        _texto(c, g.caixa(f"E{linha}:G{linha}"), CAMPO_NAO_DEFINIDO, 30 * esc)

    # --- quadro do selo (sem fundo preto: o destaque usa só listras e moldura) ---
    caixa_selo = g.caixa("H20:O22")
    sx, sy, sw, sh = caixa_selo
    if com_selo:
        # Listras ao redor + placa branca com o texto + moldura grossa em volta da etiqueta toda.
        _hachurar(c, caixa_selo)
        folga_x, folga_y = 24 * esc, 17 * esc
        placa = (sx + folga_x, sy + folga_y, sw - 2 * folga_x, sh - 2 * folga_y)
        c.setFillColorRGB(1, 1, 1)
        c.setLineWidth(3)
        c.rect(*placa, stroke=1, fill=1)
        c.setFillColorRGB(0, 0, 0)
        _texto(c, placa, et.selo, 90 * esc, margem=10)
        c.setLineWidth(6)
        c.rect(x0 - 3.5, y_topo - _ALTURA_ROMANEIO * esc - 3.5,
               _LARGURA_ROMANEIO * esc + 7, _ALTURA_ROMANEIO * esc + 7, stroke=1, fill=0)
    elif et.selo:  # SEM SELO: quadro simples, sem listras
        _texto(c, caixa_selo, et.selo, 80 * esc)
    else:  # produto que não usa selo
        _texto(c, caixa_selo, TEXTO_SELO_NAO_APLICA, 46 * esc, cor=(0.4, 0.4, 0.4), fonte="Helvetica")

    # --- tabela de datas, lote e pesos --------------------------------------
    tabela = (
        (23, "FABRICAÇÃO:", et.fabricacao),
        (24, "VALIDADE:", et.validade),
        (25, "LOTE:", et.lote),
        (26, "PESO BRUTO:", _kg(et.peso_bruto)),
        (27, "PESO LIQUIDO:", _kg(et.peso_liquido)),
    )
    for linha, rotulo, valor in tabela:
        _texto(c, g.caixa(f"I{linha}:L{linha}"), rotulo, 19 * esc, "direita")
        _texto(c, g.caixa(f"M{linha}:O{linha}"), valor, 19 * esc, "esquerda")

    # --- quantidade e nº da linha (caixa preta) ------------------------------
    _texto(c, g.caixa("B26:D27"), "QUANTIDADE:", 26 * esc)
    _texto(c, g.caixa("E26:E27"), str(et.quantidade), 26 * esc)
    caixa_linha = g.caixa("G26:G27")
    c.setFillColorRGB(0, 0, 0)
    c.rect(*caixa_linha, stroke=0, fill=1)
    numero = et.linha.split()[-1]
    _texto(c, caixa_linha, str(int(numero)) if numero.isdigit() else numero, 36 * esc, cor=(1, 1, 1))

    status = "REIMPRESSÃO" if et.reimpressao else ("TESTE" if et.teste else "EM PRODUÇÃO")
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(MARGEM, MARGEM, f"{et.turno} - {et.data_hora} - Palete nº {et.palete} - Status: {status}")
    c.save()


def caminho_pdf(et: Etiqueta, registro_id: int) -> Path:
    return LABELS_DIR / f"palete{et.palete:06d}_{registro_id}.pdf"


def gerar_etiqueta(et: Etiqueta, registro_id: int) -> bool:
    """Gera o PDF da etiqueta em etiquetas/. A impressão é feita a partir dele, no navegador."""
    try:
        LABELS_DIR.mkdir(exist_ok=True)
        gerar_pdf(et, caminho_pdf(et, registro_id))
    except Exception:  # fronteira do sistema: registra o erro e avisa o operador
        logger.exception("Erro ao gerar a etiqueta do palete %s", et.palete)
        return False
    return True


def limpar_etiquetas_antigas() -> None:
    if not LABELS_DIR.exists():
        return
    limite = tempo.time() - DIAS_GUARDAR_PDF * 86400
    for pdf in LABELS_DIR.glob("*.pdf"):
        if pdf.stat().st_mtime < limite:
            pdf.unlink(missing_ok=True)


def concluir_impressao(
    db: sqlite3.Connection, registro_id: int, et: Etiqueta, status_ok: str
) -> bool:
    """Gera o PDF e só então marca o registro como válido (ou como erro)."""
    ok = gerar_etiqueta(et, registro_id)
    with db:
        db.execute(
            "UPDATE producao SET status = ? WHERE id = ?",
            (status_ok if ok else STATUS_ERRO, registro_id),
        )
    return ok


# --------------------------------------------------------------------------
# Template
# --------------------------------------------------------------------------
TEMPLATE_HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Apontamento PCP - Linha de Produção</title>
<style>
  body { font-family: Arial, sans-serif; background: #f4f7f6; margin: 0; padding: 15px; }
  .container { width: 98%; max-width: 2000px; margin: auto; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,.1); text-align: center; }
  h1 { color: #333; margin-top: 0; }
  .nav-links { margin-bottom: 20px; }
  .nav-links a { margin: 0 10px; text-decoration: none; font-weight: bold; color: #007bff; }
  .flash-area { position: sticky; top: 8px; z-index: 10; }
  .flash { padding: 12px 15px; border-radius: 5px; margin: 0 auto 10px; font-weight: bold; max-width: 900px; box-shadow: 0 2px 6px rgba(0,0,0,.15); }
  .flash.ok { background: #d4edda; color: #155724; }
  .flash.erro { background: #f8d7da; color: #721c24; }
  .turno-box { background: #e9ecef; padding: 15px; border-radius: 5px; margin: 0 auto 25px; max-width: 420px; }
  .turno-box select { margin-top: 8px; }

  .secoes { display: flex; flex-wrap: wrap; justify-content: center; align-items: flex-start; gap: 30px 45px; }
  .secao-linha { flex: 0 1 auto; max-width: 100%;
                 width: calc(var(--n) * var(--larg) + (var(--n) - 1) * 15px); }
  .secao-linha h2 { display: inline-block; color: var(--cor); border-bottom: 4px solid var(--cor); padding: 0 25px 4px; margin: 0 0 15px; }
  .grade { display: flex; flex-wrap: wrap; justify-content: center; gap: 15px; }
  .form-linha { flex: 0 0 var(--larg); max-width: 100%; display: flex; scroll-margin-top: 90px; }
  .card-linha { width: 100%; display: flex; flex-direction: column; background: #fff; border: 2px solid #ddd; border-top: 6px solid var(--cor); padding: 14px 12px; border-radius: 6px; }
  .card-linha h3 { margin: 0 0 6px; color: var(--cor); font-size: 16px; min-height: 2.6em; display: flex; align-items: center; justify-content: center; }
  .meta { font-size: 12px; color: #666; margin: 2px 0; word-break: break-all; }
  .aviso { font-size: 12px; color: #856404; background: #fff3cd; border-radius: 4px; padding: 3px 5px; margin: 6px 0 0; }
  .input-group { margin: 10px 0; }
  .input-group label { display: block; font-weight: bold; font-size: 13px; margin-bottom: 5px; }
  input, select { width: 100%; padding: 10px; font-size: 16px; box-sizing: border-box; text-align: center; }
  .input-group input { font-size: 22px; font-weight: bold; }
  .dica { font-size: 11px; color: #888; margin-top: 3px; }
  .btn-flex { display: flex; flex-direction: column; gap: 8px; margin-top: auto; }
  button { border: none; padding: 12px 10px; font-size: 15px; border-radius: 4px; cursor: pointer; font-weight: bold; }
  button:disabled { opacity: .5; cursor: wait; }
  .btn-imprimir { background: var(--cor); color: #fff; }
  .btn-reimprimir { background: #ffc107; color: #333; }
  .selo-group { display: flex; gap: 6px; margin: 4px 0 10px; }
  .selo-op { position: relative; flex: 1; }
  .selo-op input { position: absolute; opacity: 0; width: 100%; height: 100%; margin: 0; cursor: pointer; }
  .selo-op span { display: block; padding: 9px 2px; border: 2px solid #ccc; border-radius: 4px; background: #f8f9fa; color: #666; font-size: 13px; font-weight: bold; }
  .selo-op.sem input:checked + span { background: #495057; border-color: #495057; color: #fff; }
  .selo-op.com input:checked + span { background: #c0392b; border-color: #c0392b; color: #fff; }
  .selo-op input:focus-visible + span { outline: 3px solid #007bff; }
  .sem-selo { font-size: 12px; color: #888; margin: 4px 0 10px; padding: 9px 2px; border: 2px dashed #ddd; border-radius: 4px; }

  .rodape { margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; font-size: 13px; }
  .modo { display: inline-block; padding: 5px 10px; border-radius: 4px; margin-right: 10px; }
  .modo.ok { background: #d4edda; color: #155724; }
  .modo.aviso { background: #fff3cd; color: #856404; }
  .btn-teste { background: #6c757d; color: #fff; padding: 7px 12px; font-size: 13px; }
  .nota { font-size: 12px; color: #777; margin: -4px 0 8px; }
  .btn-excel { background: #17a2b8; color: #fff; padding: 10px 15px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block; margin-bottom: 15px; }
  table { width: 100%; border-collapse: collapse; margin: 10px auto; }
  table.resumo { width: auto; min-width: 60%; }
  th, td { border: 1.5px solid #ddd; padding: 10px; text-align: center; font-size: 14px; }
  th { background: #f2f2f2; }
  .badge { padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
  .badge.producao { background: #d4edda; color: #155724; }
  .badge.reimpressao { background: #ffc107; color: #333; }
  .badge.erro { background: #f8d7da; color: #721c24; }
  .badge.pendente { background: #e2e3e5; color: #383d41; }
  .badge.selo-com { background: #c0392b; color: #fff; }
  .selo-totais { display: flex; gap: 14px; justify-content: center; flex-wrap: wrap; margin: 12px 0 16px; }
  .selo-total { min-width: 230px; padding: 12px 18px; border-radius: 8px; border: 3px solid; text-align: center; }
  .selo-total .titulo { font-size: 15px; font-weight: bold; letter-spacing: 1px; }
  .selo-total .numero { font-size: 34px; font-weight: bold; line-height: 1.1; }
  .selo-total .detalhe { font-size: 13px; }
  .selo-total.com { border-color: #c0392b; color: #c0392b; background: #fdf1ef; }
  .selo-total.sem { border-color: #495057; color: #495057; background: #f1f3f5; }
  table.resumo th.com { background: #f5c6c0; color: #7b241c; }
  table.resumo th.sem { background: #ced4da; color: #343a40; }
  table.resumo td.com { background: #fdf1ef; }
  table.resumo td.sem { background: #f1f3f5; }
  table.resumo tr.total td { border-top: 3px solid #555; font-weight: bold; font-size: 15px; }
  .zero { color: #bbb; }
</style>
</head>
<body>
<div class="container">
  <h1>Apontamento Automático - PCP</h1>

  <div class="flash-area">
    {% for categoria, texto in get_flashed_messages(with_categories=true) %}
      <div class="flash {{ categoria }}">{{ texto }}</div>
    {% endfor %}
  </div>

  {% if view == 'operador' %}
    <div class="nav-links">
      <a href="{{ url_for('pcp_painel') }}">📊 Ver Relatório e Exportar PCP</a>
    </div>

    <div class="turno-box">
      <label for="turno_global"><strong>Turno atual de operação:</strong></label>
      <select id="turno_global">
        {% for t in turnos %}
          <option value="{{ t.nome }}" data-digito="{{ t.digito }}" {{ 'selected' if t.nome == turno_padrao }}>{{ t.nome }} ({{ t.periodo }})</option>
        {% endfor %}
      </select>
    </div>

    <div class="secoes" style="--larg: {{ largura }}px">
    {% for numero, itens in grupos.items() %}
      <section class="secao-linha" style="--cor: {{ itens[0].produto.cor }}; --n: {{ [itens|length, colunas]|min }}">
        <h2>Linha {{ numero }}</h2>
        <div class="grade">
          {% for c in itens %}
            {% set p = c.produto %}
            <form action="{{ url_for('imprimir') }}" method="POST" class="form-linha" id="card-{{ p.codigo }}" target="_blank">
              <input type="hidden" name="produto" value="{{ p.codigo }}">
              <input type="hidden" name="turno" class="hidden-turno" value="{{ turno_padrao }}">
              <div class="card-linha" style="--cor: {{ p.cor }}">
                <h3>{{ p.nome }}</h3>
                <p class="meta">Cód: {{ p.cod }}</p>
                <p class="meta">Lote: {{ c.prefixo }}<strong class="lote-turno">{{ turno_digito }}</strong></p>
                <p class="meta">+ hora da impressão</p>
                <div class="input-group">
                  <label for="qtd_{{ p.codigo }}">Quantidade de {{ p.unidade|lower }} no palete</label>
                  <input id="qtd_{{ p.codigo }}" type="number" name="quantidade" value="{{ p.qtd_padrao }}" min="1" max="{{ p.qtd_max }}" inputmode="numeric" required>
                  <div class="dica">Padrão {{ p.qtd_padrao }} · máx. {{ p.qtd_max }} (ajuste na ponta de lote)</div>
                </div>
                {% if p.usa_selo %}
                  <div class="selo-group" role="radiogroup" aria-label="Selo">
                    {% for s in selos %}
                      <label class="selo-op {{ 'sem' if loop.first else 'com' }}">
                        <input type="radio" name="selo" value="{{ s }}" {{ 'checked' if s == c.selo }} required>
                        <span>{{ s }}</span>
                      </label>
                    {% endfor %}
                  </div>
                {% else %}
                  <div class="sem-selo">Produto sem selo</div>
                {% endif %}
                <div class="btn-flex">
                  <button type="submit" class="btn-imprimir">Imprimir Novo Palete</button>
                  <button type="submit" class="btn-reimprimir" formnovalidate
                          formaction="{{ url_for('reimprimir', codigo=p.codigo) }}">🔄 Reimprimir Última</button>
                </div>
              </div>
            </form>
          {% endfor %}
        </div>
      </section>
    {% endfor %}
    </div>

    <div class="rodape" id="rodape-impressao">
      <span class="modo ok">🖨 A etiqueta abre em PDF numa nova aba: imprima com Ctrl+P</span>
      <button type="submit" form="form-teste" class="btn-teste">Gerar etiqueta de teste</button>
    </div>
    <form id="form-teste" action="{{ url_for('teste_impressao') }}" method="POST" target="_blank"></form>

    <script>
      const sel = document.getElementById('turno_global');
      const sincronizar = () => {
        document.querySelectorAll('.hidden-turno').forEach(t => t.value = sel.value);
        // O último dígito do lote é o número do turno: atualiza em todos os cartões.
        const digito = sel.options[sel.selectedIndex].dataset.digito;
        document.querySelectorAll('.lote-turno').forEach(s => s.textContent = digito);
      };
      sel.addEventListener('change', sincronizar);
      sincronizar();
      // Evita clique duplo gerar dois paletes/etiquetas. O PDF abre em outra aba e esta página
      // continua aberta, então os botões voltam ao normal depois de alguns segundos.
      document.querySelectorAll('.form-linha').forEach(f =>
        f.addEventListener('submit', () => {
          setTimeout(() => f.querySelectorAll('button').forEach(b => b.disabled = true), 0);
          setTimeout(() => f.querySelectorAll('button').forEach(b => b.disabled = false), 3000);
        }));
      window.addEventListener('pageshow', e => { if (e.persisted) location.reload(); });
    </script>

  {% elif view == 'pcp' %}
    <div class="nav-links"><a href="{{ url_for('index') }}">⬅ Voltar para Painel do Operador</a></div>

    <h3>Resumo do dia de produção ({{ hoje }}) — somente novos paletes</h3>
    <p class="nota">Paletes feitos depois da meia-noite contam no dia em que o turno começou.
      Produtos que não usam selo entram em "Sem selo".</p>

    <div class="selo-totais">
      <div class="selo-total com">
        <div class="titulo">COM SELO</div>
        <div class="numero">{{ totais['unid_com'] }}</div>
        <div class="detalhe">unidades em <strong>{{ totais['paletes_com'] }}</strong> palete(s)</div>
      </div>
      <div class="selo-total sem">
        <div class="titulo">SEM SELO</div>
        <div class="numero">{{ totais['unid_sem'] }}</div>
        <div class="detalhe">unidades em <strong>{{ totais['paletes_sem'] }}</strong> palete(s)</div>
      </div>
    </div>

    <table class="resumo">
      <tr>
        <th rowspan="2">Linha</th><th rowspan="2">Produto</th>
        <th colspan="2" class="com">COM SELO</th><th colspan="2" class="sem">SEM SELO</th>
        <th colspan="2">TOTAL</th>
      </tr>
      <tr>
        <th class="com">Paletes</th><th class="com">Unidades</th>
        <th class="sem">Paletes</th><th class="sem">Unidades</th>
        <th>Paletes</th><th>Unidades</th>
      </tr>
      {% for r in resumo %}
        <tr>
          <td>{{ r['linha'] }}</td><td>{{ r['sku'] }}</td>
          <td class="com">{% if r['paletes_com'] %}{{ r['paletes_com'] }}{% else %}<span class="zero">–</span>{% endif %}</td>
          <td class="com">{% if r['unid_com'] %}<strong>{{ r['unid_com'] }}</strong>{% else %}<span class="zero">–</span>{% endif %}</td>
          <td class="sem">{% if r['paletes_sem'] %}{{ r['paletes_sem'] }}{% else %}<span class="zero">–</span>{% endif %}</td>
          <td class="sem">{% if r['unid_sem'] %}<strong>{{ r['unid_sem'] }}</strong>{% else %}<span class="zero">–</span>{% endif %}</td>
          <td>{{ r['paletes'] }}</td><td><strong>{{ r['total'] }}</strong></td>
        </tr>
      {% else %}
        <tr><td colspan="8">Nenhum palete apontado hoje.</td></tr>
      {% endfor %}
      {% if resumo %}
        <tr class="total">
          <td colspan="2">TOTAL DO DIA</td>
          <td class="com">{{ totais['paletes_com'] }}</td><td class="com">{{ totais['unid_com'] }}</td>
          <td class="sem">{{ totais['paletes_sem'] }}</td><td class="sem">{{ totais['unid_sem'] }}</td>
          <td>{{ totais['paletes'] }}</td><td>{{ totais['total'] }}</td>
        </tr>
      {% endif %}
    </table>

    <h3 style="margin-top: 30px;">Últimos {{ limite }} registros</h3>
    <a href="{{ url_for('exportar_excel') }}" class="btn-excel">📥 Baixar Relatório Completo em Excel (.xlsx)</a>
    <table>
      <tr>
        <th>Data/Hora</th><th>Dia de produção</th><th>Turno</th><th>Linha</th><th>Produto</th>
        <th>Qtd Palete</th><th>Selo</th><th>Nº Palete</th><th>Status / Tipo</th><th>Etiqueta</th>
      </tr>
      {% for r in registros %}
        {% set rotulo, classe = badges.get(r['status'], (r['status'], 'pendente')) %}
        <tr>
          <td>{{ r['data_hora'] }}</td>
          <td>{{ r['dia_producao'] or r['data_hora'][:10] }}</td>
          <td>{{ r['turno'] }}</td>
          <td>{{ r['linha'] }}</td>
          <td>{{ r['sku'] }}</td>
          <td><strong>{{ r['quantidade'] }}</strong></td>
          <td>{% if r['selo'] == selo_com %}<span class="badge selo-com">{{ r['selo'] }}</span>{% elif r['selo'] %}{{ r['selo'] }}{% else %}<span class="zero">–</span>{% endif %}</td>
          <td>{{ r['palete'] }}</td>
          <td><span class="badge {{ classe }}">{{ rotulo }}</span></td>
          <td><a href="{{ url_for('etiqueta_pdf', registro_id=r['id']) }}" target="_blank">Abrir PDF</a></td>
        </tr>
      {% endfor %}
    </table>
  {% endif %}
</div>
</body>
</html>
"""

BADGES = {
    STATUS_PRODUCAO: ("Novo Palete", "producao"),
    STATUS_REIMPRESSAO: ("Reimpressão (Não Soma)", "reimpressao"),
    STATUS_ERRO: ("Falha na etiqueta (Não Soma)", "erro"),
    STATUS_PENDENTE: ("Pendente (Não Soma)", "pendente"),
}


# --------------------------------------------------------------------------
# Rotas
# --------------------------------------------------------------------------
def ler_quantidade(valor: str | None, prod: Produto) -> int | None:
    try:
        quantidade = int(valor or "")
    except ValueError:
        return None
    return quantidade if 1 <= quantidade <= prod.qtd_max else None


def selo_valido(valor: str | None) -> str:
    return valor if valor in SELOS else SELOS[0]


def voltar_ao_cartao(prod: Produto, turno: str, selo: str | None = None):
    """Volta para a tela do operador, já rolando até o cartão usado."""
    # Guarda a escolha do selo só no cartão usado (os outros voltam para o padrão).
    return redirect(url_for("index", turno=turno, selo=selo or None,
                            selo_produto=prod.codigo if selo else None,
                            _anchor=f"card-{prod.codigo}"))


@app.get("/")
def index():
    agora = datetime.now()
    turno = request.args.get("turno", "")
    selo = selo_valido(request.args.get("selo"))
    selo_de = request.args.get("selo_produto", "")
    turno_sel = turno if turno in NOMES_TURNOS else turno_atual(agora)
    grupos: dict[str, list[dict]] = {}
    for p in PRODUTOS.values():
        grupos.setdefault(p.numero, []).append(
            {"produto": p, "prefixo": lote_prefixo(p, agora), "selo": selo if p.codigo == selo_de else SELOS[0]}
        )
    return render_template_string(
        TEMPLATE_HTML,
        view="operador",
        grupos=grupos,
        colunas=COLUNAS_POR_LINHA,
        largura=LARGURA_CARTAO_PX,
        turnos=TURNOS,
        turno_padrao=turno_sel,
        turno_digito=TURNO_POR_NOME[turno_sel].digito,
        selos=SELOS,
    )


@app.post("/imprimir")
def imprimir():
    prod = PRODUTOS.get(request.form.get("produto", ""))
    if prod is None:
        abort(400, "Produto inválido.")

    turno = request.form.get("turno", "")
    if turno not in NOMES_TURNOS:
        turno = turno_atual()
    selo = selo_valido(request.form.get("selo")) if prod.usa_selo else ""

    quantidade = ler_quantidade(request.form.get("quantidade"), prod)
    if quantidade is None:
        flash(
            f"Quantidade inválida para {prod.nome}: informe um número de 1 a {prod.qtd_max}.",
            "erro",
        )
        return voltar_ao_cartao(prod, turno, selo)

    agora = datetime.now()
    data_hora = agora.strftime(FORMATO_DATA_HORA)
    lote = lote_completo(prod, agora, turno)
    dia_data = dia_producao(agora)
    dia = dia_data.strftime(FORMATO_DATA)
    validade, peso_liq, peso_bru = dados_romaneio(prod, quantidade, dia_data)
    db = get_db()

    # O id gerado pelo banco vira o nº do palete: único e sequencial.
    with db:
        cur = db.execute(
            SQL_INSERT,
            (data_hora, turno, prod.linha, prod.nome, lote,
             quantidade, None, STATUS_PENDENTE, dia,
             prod.cod, validade, peso_liq, peso_bru, selo),
        )
        registro_id = cur.lastrowid
        palete = registro_id
        db.execute("UPDATE producao SET palete = ? WHERE id = ?", (palete, registro_id))

    etiqueta = Etiqueta(
        prod.linha, prod.nome, turno, lote, quantidade,
        prod.unidade, palete, data_hora,
        cod=prod.cod, desc1=prod.item.desc1, desc2=prod.item.desc2,
        fabricacao=dia, validade=validade, peso_liquido=peso_liq, peso_bruto=peso_bru,
        selo=selo,
    )
    if concluir_impressao(db, registro_id, etiqueta, STATUS_PRODUCAO):
        # Redireciona para o PDF (nova aba): atualizar essa aba não cria outro palete.
        return redirect(url_for("etiqueta_pdf", registro_id=registro_id))
    flash(
        f"FALHA ao gerar a etiqueta de {prod.nome}. O palete NÃO foi contabilizado; tente novamente.",
        "erro",
    )
    return voltar_ao_cartao(prod, turno, selo)


@app.post("/reimprimir/<codigo>")
def reimprimir(codigo: str):
    prod = PRODUTOS.get(codigo)
    if prod is None:
        abort(404)

    db = get_db()
    ultimo = db.execute(
        """
        SELECT * FROM producao
        WHERE linha = ? AND sku = ? AND status = ?
        ORDER BY id DESC LIMIT 1
        """,
        (prod.linha, prod.nome, STATUS_PRODUCAO),
    ).fetchone()
    if ultimo is None:
        flash(f"Não há palete anterior para reimprimir de {prod.nome}.", "erro")
        return voltar_ao_cartao(prod, turno_atual())

    data_hora = datetime.now().strftime(FORMATO_DATA_HORA)
    # Validade, pesos e selo foram gravados na impressão original. Paletes anteriores à
    # versão 4 não têm esses dados: nesse caso são recalculados a partir da aba Base.
    dia_txt = ultimo["dia_producao"] or ultimo["data_hora"][:10]
    calc_val, calc_liq, calc_bru = dados_romaneio(
        prod, ultimo["quantidade"], datetime.strptime(dia_txt, FORMATO_DATA).date()
    )
    validade = ultimo["validade"] or calc_val
    peso_liq = calc_liq if ultimo["peso_liquido"] is None else ultimo["peso_liquido"]
    peso_bru = calc_bru if ultimo["peso_bruto"] is None else ultimo["peso_bruto"]
    selo = selo_valido(ultimo["selo"]) if prod.usa_selo else ""
    with db:
        cur = db.execute(
            SQL_INSERT,
            (data_hora, ultimo["turno"], ultimo["linha"], ultimo["sku"],
             ultimo["lote"], ultimo["quantidade"], ultimo["palete"], STATUS_PENDENTE,
             ultimo["dia_producao"], ultimo["cod"] or prod.cod, validade, peso_liq, peso_bru, selo),
        )
    # A reimpressão usa os dados GRAVADOS no palete original (inclusive o lote da época).
    etiqueta = Etiqueta(
        ultimo["linha"], ultimo["sku"], ultimo["turno"],
        ultimo["lote"], ultimo["quantidade"], prod.unidade, ultimo["palete"], data_hora,
        cod=ultimo["cod"] or prod.cod, desc1=prod.item.desc1, desc2=prod.item.desc2,
        fabricacao=dia_txt, validade=validade, peso_liquido=peso_liq, peso_bruto=peso_bru,
        selo=selo, reimpressao=True,
    )
    if concluir_impressao(db, cur.lastrowid, etiqueta, STATUS_REIMPRESSAO):
        return redirect(url_for("etiqueta_pdf", registro_id=cur.lastrowid))
    flash(f"FALHA ao gerar a reimpressão de {prod.nome}.", "erro")
    return voltar_ao_cartao(prod, ultimo["turno"], selo)


@app.post("/teste-impressao")
def teste_impressao():
    """Gera uma etiqueta de teste (não grava nada no banco) e abre o PDF."""
    amostra = next(iter(PRODUTOS.values()))
    agora = datetime.now()
    dia_data = dia_producao(agora)
    validade, peso_liq, peso_bru = dados_romaneio(amostra, 1, dia_data)
    etiqueta = Etiqueta(
        amostra.linha, "TESTE DE IMPRESSÃO", turno_atual(agora),
        lote_completo(amostra, agora, turno_atual(agora)), 1, amostra.unidade, 0,
        agora.strftime(FORMATO_DATA_HORA),
        cod=amostra.cod, desc1="TESTE DE IMPRESSÃO", desc2="NÃO USAR",
        fabricacao=dia_data.strftime(FORMATO_DATA), validade=validade,
        peso_liquido=peso_liq, peso_bruto=peso_bru, teste=True,
        selo=SELOS[0] if amostra.usa_selo else "",
    )
    if gerar_etiqueta(etiqueta, 0):
        return redirect(url_for("etiqueta_pdf", registro_id=0))
    flash("FALHA ao gerar a etiqueta de teste. Veja a janela preta do programa.", "erro")
    return redirect(url_for("index", _anchor="rodape-impressao"))


@app.get("/etiqueta/<int:registro_id>")
def etiqueta_pdf(registro_id: int):
    """Abre o PDF de um registro no navegador (de lá se imprime com Ctrl+P)."""
    arquivos = sorted(LABELS_DIR.glob(f"palete*_{registro_id}.pdf")) if LABELS_DIR.exists() else []
    if not arquivos:
        abort(404, f"PDF não encontrado. As etiquetas ficam guardadas por {DIAS_GUARDAR_PDF} dias; "
                   "para uma etiqueta antiga, use \"Reimprimir Última\" no painel do operador.")
    return send_file(arquivos[-1], mimetype="application/pdf", max_age=0)


@app.get("/pcp")
def pcp_painel():
    db = get_db()
    hoje = dia_producao(datetime.now()).strftime(FORMATO_DATA)
    registros = db.execute(
        "SELECT * FROM producao ORDER BY id DESC LIMIT ?", (LIMITE_PAINEL,)
    ).fetchall()
    resumo = db.execute(
        """
        SELECT linha, sku,
               SUM(CASE WHEN selo = ? THEN 1 ELSE 0 END) AS paletes_com,
               SUM(CASE WHEN selo = ? THEN quantidade ELSE 0 END) AS unid_com,
               SUM(CASE WHEN COALESCE(selo, '') <> ? THEN 1 ELSE 0 END) AS paletes_sem,
               SUM(CASE WHEN COALESCE(selo, '') <> ? THEN quantidade ELSE 0 END) AS unid_sem,
               COUNT(*) AS paletes, SUM(quantidade) AS total
        FROM producao
        WHERE status = ? AND COALESCE(dia_producao, substr(data_hora, 1, 10)) = ?
        GROUP BY linha, sku ORDER BY linha, sku
        """,
        (SELO_COM, SELO_COM, SELO_COM, SELO_COM, STATUS_PRODUCAO, hoje),
    ).fetchall()
    chaves = ("paletes_com", "unid_com", "paletes_sem", "unid_sem", "paletes", "total")
    totais = {k: sum(r[k] for r in resumo) for k in chaves}
    return render_template_string(
        TEMPLATE_HTML, view="pcp", registros=registros, resumo=resumo, totais=totais,
        hoje=hoje, limite=LIMITE_PAINEL, badges=BADGES, selo_com=SELO_COM,
    )


@app.get("/exportar-excel")
def exportar_excel():
    linhas = get_db().execute(
        """
        SELECT id, data_hora, COALESCE(dia_producao, substr(data_hora, 1, 10)) AS dia,
               turno, linha, sku, lote, quantidade, palete, status,
               cod, validade, peso_liquido, peso_bruto, selo
        FROM producao ORDER BY id DESC
        """
    ).fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Apontamento"
    ws.append(["ID", "Data/Hora", "Dia de Produção", "Turno", "Linha", "Produto", "Lote",
               "Quantidade", "Nº Palete", "Status", "Cód", "Validade", "Peso Líquido (kg)",
               "Peso Bruto (kg)", "Selo"])
    for linha in linhas:
        ws.append(list(linha))
    ws.freeze_panes = "A2"
    for coluna in ws.columns:
        ws.column_dimensions[coluna[0].column_letter].width = (
            max(len(str(c.value or "")) for c in coluna) + 2
        )

    # Segunda aba: contagem pronta por dia, produto e selo (só novos paletes).
    resumo = get_db().execute(
        """
        SELECT COALESCE(dia_producao, substr(data_hora, 1, 10)) AS dia, linha, sku,
               CASE WHEN selo = ? THEN ? ELSE ? END AS grupo_selo,
               COUNT(*) AS paletes, SUM(quantidade) AS unidades
        FROM producao WHERE status = ?
        GROUP BY dia, linha, sku, grupo_selo
        ORDER BY substr(dia, 7, 4) DESC, substr(dia, 4, 2) DESC, substr(dia, 1, 2) DESC, linha, sku, grupo_selo
        """,
        (SELO_COM, SELO_COM, SELOS[0], STATUS_PRODUCAO),
    ).fetchall()
    ws2 = wb.create_sheet("Resumo por selo")
    ws2.append(["Dia de Produção", "Linha", "Produto", "Selo", "Paletes", "Unidades"])
    for linha in resumo:
        ws2.append(list(linha))
    ws2.freeze_panes = "A2"
    for coluna in ws2.columns:
        ws2.column_dimensions[coluna[0].column_letter].width = (
            max(len(str(c.value or "")) for c in coluna) + 2
        )

    buffer = io.BytesIO()  # em memória: sem arquivo fixo, sem corrida entre downloads
    wb.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"Apontamento_PCP_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# Inicialização no import: funciona com `python`, `flask run` e `waitress`.
init_db()
limpar_etiquetas_antigas()
logger.info("Etiquetas em PDF na pasta: %s", LABELS_DIR)
logger.info("Base do romaneio: %d produtos lidos de %s", len(BASE_ROMANEIO), ROMANEIO_XLSX.name)

if __name__ == "__main__":
    # Nunca use debug=True com host 0.0.0.0: o depurador permite executar código remotamente.
    app.run(
        host=os.environ.get("PCP_HOST", "0.0.0.0"),
        port=int(os.environ.get("PCP_PORT", "5000")),
        debug=False,
    )
