"""
Entre Linhas — App Web Completo (Streamlit)
Gerador de grade interativa com times e placar ao vivo.
"""

import io
import random
import hashlib
from datetime import datetime

import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from PIL import Image, ImageDraw, ImageFont


# =====================================================
# CONFIGURAÇÃO DA PÁGINA
# =====================================================
st.set_page_config(
    page_title="Entre Linhas — Gerador de Grade",
    page_icon="🎲",
    layout="wide",
)


# =====================================================
# BIBLIOTECA DE 343 PALAVRAS
# =====================================================
PALAVRAS = [
    # Natureza (30)
    "sol", "lua", "estrela", "céu", "nuvem", "chuva", "vento", "neve", "gelo", "fogo",
    "água", "terra", "ar", "mar", "rio", "lago", "floresta", "montanha", "deserto", "praia",
    "ilha", "vulcão", "caverna", "cachoeira", "campo", "vale", "colina", "duna", "pântano", "savana",
    # Animais (40)
    "cachorro", "gato", "pássaro", "peixe", "cavalo", "vaca", "porco", "galinha", "pato", "coelho",
    "leão", "tigre", "elefante", "macaco", "cobra", "aranha", "abelha", "formiga", "borboleta", "urso",
    "lobo", "raposa", "coruja", "águia", "tartaruga", "jacaré", "sapo", "golfinho", "baleia", "tubarão",
    "polvo", "caranguejo", "camelo", "girafa", "zebra", "panda", "koala", "pinguim", "esquilo", "morcego",
    # Alimentos (60)
    "arroz", "feijão", "macarrão", "pão", "carne", "frango", "ovo", "queijo", "leite", "café",
    "chá", "suco", "refrigerante", "cerveja", "vinho", "sal", "açúcar", "óleo", "farinha", "manteiga",
    "mel", "chocolate", "bolo", "sorvete", "doce", "pizza", "sanduíche", "hambúrguer", "batata", "tomate",
    "cebola", "alho", "cenoura", "alface", "brócolis", "milho", "ervilha", "abóbora", "beterraba", "pepino",
    "pimentão", "maçã", "banana", "laranja", "uva", "morango", "abacaxi", "manga", "melancia", "limão",
    "pêssego", "pera", "kiwi", "cereja", "ameixa", "coco", "abacate", "goiaba", "maracujá", "pitanga",
    # Casa (45)
    "casa", "apartamento", "quarto", "sala", "cozinha", "banheiro", "quintal", "jardim", "garagem", "sótão",
    "porão", "varanda", "escritório", "biblioteca", "porta", "janela", "parede", "chão", "teto", "telhado",
    "escada", "corredor", "cama", "mesa", "cadeira", "sofá", "armário", "geladeira", "fogão", "pia",
    "espelho", "lâmpada", "chave", "fechadura", "relógio", "tapete", "cortina", "almofada", "cobertor",
    "travesseiro", "lençol", "toalha", "sabonete", "escova", "porta-retrato",
    # Eletrônicos (20)
    "computador", "celular", "televisão", "rádio", "telefone", "câmera", "impressora", "teclado", "mouse", "fone",
    "ventilador", "aspirador", "liquidificador", "batedeira", "torradeira", "cafeteira", "micro-ondas",
    "ventilador de teto", "caixa de som", "videogame",
    # Profissões (30)
    "médico", "professor", "motorista", "policial", "bombeiro", "cozinheiro", "padeiro", "advogado",
    "engenheiro", "dentista", "enfermeiro", "aluno", "ator", "cantor", "dançarino", "pintor", "escritor",
    "jornalista", "fotógrafo", "piloto", "soldado", "agricultor", "pescador", "carteiro", "garçom",
    "eletricista", "encanador", "mecânico", "veterinário", "cientista",
    # Lugares (35)
    "escola", "hospital", "mercado", "banco", "igreja", "praça", "rua", "cidade", "fazenda", "parque",
    "zoológico", "museu", "cinema", "teatro", "restaurante", "hotel", "aeroporto", "porto", "estádio",
    "shopping", "clube", "padaria", "farmácia", "posto", "rodoviária", "praia", "campo", "vila", "aldeia",
    "bairro", "avenida", "beco", "jardim", "estação", "viela",
    # Transportes (30)
    "carro", "moto", "bicicleta", "ônibus", "caminhão", "trem", "metrô", "avião", "helicóptero", "navio",
    "barco", "canoa", "submarino", "foguete", "skate", "patins", "trator", "ambulância", "táxi", "van",
    "caminhonete", "balão", "jangada", "iate", "carruagem", "triciclo", "patinete", "teleférico", "bonde",
    "furgão",
    # Corpo (20)
    "cabeça", "cabelo", "olho", "nariz", "boca", "orelha", "braço", "mão", "dedo", "perna",
    "pé", "coração", "pulmão", "estômago", "sangue", "osso", "pele", "dente", "língua", "unha",
    # Cultura e ciência (33)
    "música", "dança", "arte", "pintura", "escultura", "poesia", "história", "ciência", "matemática",
    "física", "química", "biologia", "medicina", "tecnologia", "internet", "filosofia", "geografia",
    "astronomia", "literatura", "gramática",
    "amor", "amizade", "saudade", "alegria", "tristeza", "medo", "coragem", "esperança", "fé", "paz",
    "tempo", "espaço", "vida",
]


# =====================================================
# CORES E ESTILOS
# =====================================================
COR_COLUNAS = "C5E0B4"
COR_LINHAS  = "FFE699"
COR_CARTA   = "D9E1F2"
COR_COORD   = "F0F0F0"
COR_CANTO   = "E6E6E6"


# =====================================================
# FUNÇÕES UTILITÁRIAS
# =====================================================
def sortear(palavras: list, tamanho: int):
    """Sorteia 2*tamanho palavras sem repetir."""
    if len(palavras) < tamanho * 2:
        return None, None
    sorteadas = random.sample(palavras, tamanho * 2)
    return sorteadas[:tamanho], sorteadas[tamanho:]


def _grid_hash(linhas, colunas):
    """Hash único das palavras da grade — muda quando o sorteio muda."""
    s = "|".join(linhas) + "##" + "|".join(colunas)
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:12]


# =====================================================
# GRID INTERATIVO COM TIMES, PLACAR E PERSISTÊNCIA
# =====================================================
def render_grid_interativo(linhas, colunas, nome_time_1="Time 1", nome_time_2="Time 2",
                            game_key="partida"):
    """
    Grid clicável com:
    - 3 estados: ⚪ neutro → 🔵 Time 1 → 🔴 Time 2 → ⚪
    - Placar ao vivo com nomes editáveis
    - Persistência (cores + nomes) via localStorage
    - Botão "Limpar cores" interno
    """
    tam = len(linhas)
    grid_hash = _grid_hash(linhas, colunas)
    storage_key = f"entre_linhas_{game_key}_{grid_hash}"
    storage_nomes = f"entre_linhas_nomes_{game_key}"

    # Sanitiza os nomes para não quebrar o JS
    nome_1_safe = nome_time_1.replace("\\", "\\\\").replace("'", "\\'").replace('"', '\\"')
    nome_2_safe = nome_time_2.replace("\\", "\\\\").replace("'", "\\'").replace('"', '\\"')

    # Cabeçalho 1: letras A, B, C...
    header1 = '<div></div><div></div>'
    for j in range(1, tam + 1):
        header1 += f'<div class="cel cel-coord">{chr(64 + j)}</div>'

    # Cabeçalho 2: canto × + palavras das colunas
    header2 = '<div></div><div class="cel cel-canto">×</div>'
    for col in colunas:
        header2 += f'<div class="cel cel-col">{col}</div>'

    # Linhas com cartas clicáveis
    rows_html = ""
    for i, lin in enumerate(linhas, 1):
        rows_html += f'<div class="cel cel-coord">{i}</div>'
        rows_html += f'<div class="cel cel-lin">{lin}</div>'
        for j in range(1, tam + 1):
            coord = f"{chr(64 + j)}{i}"
            rows_html += f'<div class="cel cel-carta" data-coord="{coord}">{coord}</div>'

    html = f"""
<!DOCTYPE html>
<html>
<head>
<style>
    body {{ margin: 0; font-family: system-ui, sans-serif; }}

    /* ---- PLACAR ---- */
    .placar {{
        display: flex; gap: 24px; align-items: center;
        padding: 12px 16px;
        background: #f8fafc;
        border: 2px solid #cbd5e1;
        border-radius: 10px;
        margin-bottom: 14px;
        font-size: 16px;
        font-weight: 600;
        flex-wrap: wrap;
    }}
    .placar-time {{ display: flex; align-items: center; gap: 8px; }}
    .placar-dot {{
        width: 22px; height: 22px; border-radius: 4px;
        border: 2px solid #4a5568;
        flex-shrink: 0;
    }}
    .dot-b {{ background: #1e3c78; }}
    .dot-r {{ background: #dc2626; }}
    .placar-nome {{
        font-size: 15px;
        max-width: 180px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }}
    .placar-num {{ font-size: 26px; min-width: 34px; text-align: center; font-weight: bold; }}
    .placar-num.blue {{ color: #1e3c78; }}
    .placar-num.red  {{ color: #dc2626; }}
    .placar-hint {{ margin-left: auto; font-size: 13px; color: #64748b; font-weight: normal; }}

    /* ---- GRID ---- */
    .grid-el {{
        display: grid;
        grid-template-columns: 45px 130px repeat({tam}, 130px);
        gap: 4px;
    }}
    .cel {{
        border: 2px solid #4a5568;
        padding: 12px 8px;
        text-align: center;
        display: flex;
        align-items: center;
        justify-content: center;
        min-height: 45px;
        font-size: 13px;
        border-radius: 4px;
        user-select: none;
    }}
    .cel-coord {{ background: #f0f0f0; font-weight: bold; font-size: 16px; color: #323232; }}
    .cel-canto {{ background: #e6e6e6; font-weight: bold; font-size: 20px; }}
    .cel-col   {{ background: #c5e0b4; font-weight: bold; }}
    .cel-lin   {{ background: #ffe699; font-weight: bold; }}

    /* ---- CARTAS ---- */
    .cel-carta {{
        font-weight: bold;
        font-size: 22px;
        min-height: 55px;
        cursor: pointer;
        transition: transform 0.12s ease, background 0.15s ease, color 0.15s ease;
        position: relative;
    }}
    .cel-carta:hover {{
        transform: scale(1.06);
        box-shadow: 0 3px 10px rgba(0,0,0,0.25);
        z-index: 2;
    }}
    .cel-carta:active {{ transform: scale(0.96); }}

    .cel-carta.state-0 {{ background: #d9e1f2; color: #1e3c78; }}
    .cel-carta.state-1 {{ background: #1e3c78; color: #ffffff; }}
    .cel-carta.state-2 {{ background: #dc2626; color: #ffffff; }}

    /* ---- BOTÃO RESET ---- */
    .btn-reset {{
        margin-top: 14px;
        padding: 8px 18px;
        background: #64748b;
        color: white;
        border: none;
        border-radius: 6px;
        font-size: 14px;
        font-weight: 600;
        cursor: pointer;
        transition: background 0.15s;
    }}
    .btn-reset:hover {{ background: #475569; }}
</style>
</head>
<body>

<div class="placar">
    <div class="placar-time">
        <span class="placar-dot dot-b"></span>
        <span class="placar-nome" id="nome-blue">{nome_time_1}</span>
        <span class="placar-num blue" id="score-blue">0</span>
    </div>
    <div class="placar-time">
        <span class="placar-dot dot-r"></span>
        <span class="placar-nome" id="nome-red">{nome_time_2}</span>
        <span class="placar-num red" id="score-red">0</span>
    </div>
    <div class="placar-hint">Clique: 🔵 → 🔴 → ⚪</div>
</div>

<div class="grid-el">
    {header1}
    {header2}
    {rows_html}
</div>

<button class="btn-reset" onclick="resetCores()">🔄 Limpar cores</button>

<script>
    const STORAGE_KEY = "{storage_key}";
    const STORAGE_NOMES = "{storage_nomes}";
    const NOME_1_DEFAULT = '{nome_1_safe}';
    const NOME_2_DEFAULT = '{nome_2_safe}';

    function getState() {{
        try {{ return JSON.parse(localStorage.getItem(STORAGE_KEY) || "{{}}"); }}
        catch (e) {{ return {{}}; }}
    }}
    function saveState(state) {{
        try {{ localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }} catch (e) {{}}
    }}

    function updatePlacar() {{
        const azuis = document.querySelectorAll(".cel-carta.state-1").length;
        const vermelhos = document.querySelectorAll(".cel-carta.state-2").length;
        document.getElementById("score-blue").textContent = azuis;
        document.getElementById("score-red").textContent = vermelhos;
    }}

    function applyState() {{
        const state = getState();
        document.querySelectorAll(".cel-carta").forEach(el => {{
            const coord = el.dataset.coord;
            const s = state[coord] || 0;
            el.classList.remove("state-0", "state-1", "state-2");
            el.classList.add("state-" + s);
        }});
        updatePlacar();
    }}

    function cycleColor(el) {{
        const coord = el.dataset.coord;
        const state = getState();
        const atual = state[coord] || 0;
        const novo = (atual + 1) % 3;
        state[coord] = novo;
        saveState(state);

        el.classList.remove("state-0", "state-1", "state-2");
        el.classList.add("state-" + novo);
        updatePlacar();
    }}

    function resetCores() {{
        try {{ localStorage.removeItem(STORAGE_KEY); }} catch (e) {{}}
        document.querySelectorAll(".cel-carta").forEach(el => {{
            el.classList.remove("state-0", "state-1", "state-2");
            el.classList.add("state-0");
        }});
        updatePlacar();
    }}

    function updateNomes() {{
        const salvos = JSON.parse(localStorage.getItem(STORAGE_NOMES) || "{{}}");
        const n1 = NOME_1_DEFAULT || salvos.nome1 || "Time 1";
        const n2 = NOME_2_DEFAULT || salvos.nome2 || "Time 2";
        document.getElementById("nome-blue").textContent = n1;
        document.getElementById("nome-red").textContent = n2;
        try {{
            localStorage.setItem(STORAGE_NOMES, JSON.stringify({{nome1: n1, nome2: n2}}));
        }} catch (e) {{}}
    }}

    document.querySelectorAll(".cel-carta").forEach(el => {{
        el.addEventListener("click", () => cycleColor(el));
    }});

    updateNomes();
    applyState();
</script>
</body>
</html>
"""

    altura = int(260 + tam * 105)
    components.html(html, height=altura, scrolling=False)


# =====================================================
# GERAÇÃO DE EXCEL
# =====================================================
def gerar_excel(linhas, colunas, nome_time_1="Time 1", nome_time_2="Time 2"):
    """Gera .xlsx com grade + aba de cartas + placar dos times."""
    tamanho = len(linhas)
    wb = Workbook()

    # ---- Estilos ----
    fill_col = PatternFill("solid", fgColor=COR_COLUNAS)
    fill_lin = PatternFill("solid", fgColor=COR_LINHAS)
    fill_car = PatternFill("solid", fgColor=COR_CARTA)
    fill_coo = PatternFill("solid", fgColor=COR_COORD)
    fill_can = PatternFill("solid", fgColor=COR_CANTO)
    fill_time1 = PatternFill("solid", fgColor="1E3C78")
    fill_time2 = PatternFill("solid", fgColor="DC2626")

    f_cinza = Font(bold=True, size=14, color="323232")
    f_canto = Font(bold=True, size=14)
    f_palavra = Font(bold=True, size=11)
    f_coord_med = Font(bold=True, size=16, color="1E3C78")
    f_coord_gde = Font(bold=True, size=40, color="1E3C78")
    f_time_branco = Font(bold=True, size=13, color="FFFFFF")

    alin = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fina = Border(*[Side(style="thin")] * 4)
    media = Border(*[Side(style="medium")] * 4)

    # ---- Aba Grade ----
    ws = wb.active
    ws.title = "Grade"
    linha_ini, col_ini = 4, 4

    c = ws.cell(row=linha_ini - 1, column=col_ini, value="×")
    c.font, c.fill, c.alignment, c.border = f_canto, fill_can, alin, media

    for j in range(1, tamanho + 1):
        c = ws.cell(row=linha_ini - 2, column=col_ini + j, value=chr(64 + j))
        c.font, c.fill, c.alignment, c.border = f_cinza, fill_coo, alin, fina

    for i in range(1, tamanho + 1):
        c = ws.cell(row=linha_ini + i - 1, column=col_ini - 1, value=i)
        c.font, c.fill, c.alignment, c.border = f_cinza, fill_coo, alin, fina

    for j, p in enumerate(colunas, 1):
        c = ws.cell(row=linha_ini - 1, column=col_ini + j, value=p)
        c.font, c.fill, c.alignment, c.border = f_palavra, fill_col, alin, media

    for i, p in enumerate(linhas, 1):
        c = ws.cell(row=linha_ini + i - 1, column=col_ini, value=p)
        c.font, c.fill, c.alignment, c.border = f_palavra, fill_lin, alin, media

    for i in range(1, tamanho + 1):
        for j in range(1, tamanho + 1):
            coord = f"{chr(64 + j)}{i}"
            c = ws.cell(row=linha_ini + i - 1, column=col_ini + j, value=coord)
            c.font, c.fill, c.alignment, c.border = f_coord_med, fill_car, alin, fina

    ws.column_dimensions[get_column_letter(col_ini - 1)].width = 5
    ws.column_dimensions[get_column_letter(col_ini)].width = 14
    for j in range(1, tamanho + 1):
        ws.column_dimensions[get_column_letter(col_ini + j)].width = 16
    ws.row_dimensions[linha_ini - 2].height = 22
    ws.row_dimensions[linha_ini - 1].height = 30
    for i in range(1, tamanho + 1):
        ws.row_dimensions[linha_ini + i - 1].height = 45

    # Placar abaixo da grade
    linha_placar = linha_ini + tamanho + 2
    c = ws.cell(row=linha_placar, column=col_ini, value=nome_time_1)
    c.font, c.fill, c.alignment, c.border = f_time_branco, fill_time1, alin, media
    ws.merge_cells(start_row=linha_placar, start_column=col_ini,
                   end_row=linha_placar, end_column=col_ini + 1)
    c2 = ws.cell(row=linha_placar, column=col_ini + 1)
    c2.fill, c2.border = fill_time1, media

    c = ws.cell(row=linha_placar, column=col_ini + 2, value=nome_time_2)
    c.font, c.fill, c.alignment, c.border = f_time_branco, fill_time2, alin, media
    ws.merge_cells(start_row=linha_placar, start_column=col_ini + 2,
                   end_row=linha_placar, end_column=col_ini + 3)
    c2 = ws.cell(row=linha_placar, column=col_ini + 3)
    c2.fill, c2.border = fill_time2, media

    p1 = f"{get_column_letter(col_ini - 1)}{linha_ini - 2}"
    p2 = f"{get_column_letter(col_ini + tamanho)}{linha_placar}"
    ws.print_area = f"{p1}:{p2}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ---- Aba Cartas ----
    ws2 = wb.create_sheet("Cartas")
    linha_cart = 1
    for i in range(1, tamanho + 1):
        for j in range(1, tamanho + 1):
            coord = f"{chr(64 + j)}{i}"
            ws2.merge_cells(start_row=linha_cart, start_column=1,
                            end_row=linha_cart + 3, end_column=2)
            c = ws2.cell(row=linha_cart, column=1, value=coord)
            c.font, c.fill, c.alignment = f_coord_gde, fill_car, alin
            for r in range(linha_cart, linha_cart + 4):
                for col in range(1, 3):
                    ws2.cell(row=r, column=col).border = media
            linha_cart += 5

    ws2.column_dimensions["A"].width = 20
    ws2.column_dimensions["B"].width = 20
    for i in range(1, linha_cart, 5):
        for off in range(4):
            ws2.row_dimensions[i + off].height = 22

    if linha_cart > 1:
        ws2.print_area = f"A1:B{linha_cart - 2}"
        ws2.page_setup.orientation = "portrait"
        ws2.page_setup.fitToWidth = 1
        ws2.page_setup.fitToHeight = 0
        ws2.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# =====================================================
# GERAÇÃO DE PNG
# =====================================================
def _carregar_fonte(tamanho: int):
    for nome in ("DejaVuSans-Bold.ttf", "arialbd.ttf", "Arial Bold.ttf"):
        try:
            return ImageFont.truetype(nome, tamanho)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def gerar_png_grade(linhas, colunas, nome_time_1="Time 1", nome_time_2="Time 2"):
    """Gera PNG da grade visual + placar."""
    tamanho = len(linhas)
    larg_cel, alt_cel = 130, 90
    larg_cab, alt_cab = 60, 40
    alt_placar = 60

    largura = larg_cab * 2 + larg_cel * tamanho
    altura = alt_cab * 2 + alt_cel * tamanho + alt_placar

    img = Image.new("RGB", (largura, altura), "white")
    d = ImageDraw.Draw(img)

    fonte_pal = _carregar_fonte(16)
    fonte_coord = _carregar_fonte(22)
    fonte_cab = _carregar_fonte(18)
    fonte_time = _carregar_fonte(18)

    def celula(x, y, w, h, texto, cor_fundo, cor_texto="black", fonte=None):
        d.rectangle([x, y, x + w, y + h], fill=cor_fundo, outline="black", width=2)
        if texto:
            f = fonte or fonte_pal
            bbox = d.textbbox((0, 0), texto, font=f)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            d.text((x + (w - tw) / 2 - bbox[0], y + (h - th) / 2 - bbox[1]),
                   texto, fill=cor_texto, font=f)

    cores = {"col": "#C5E0B4", "lin": "#FFE699", "car": "#D9E1F2",
             "coo": "#F0F0F0", "can": "#E6E6E6"}

    # Placar no topo
    metade = largura // 2
    d.rectangle([0, 0, metade, alt_placar], fill="#1E3C78", outline="black", width=2)
    d.rectangle([metade, 0, largura, alt_placar], fill="#DC2626", outline="black", width=2)
    for texto, x_ini, x_fim in [(nome_time_1, 0, metade), (nome_time_2, metade, largura)]:
        bbox = d.textbbox((0, 0), texto, font=fonte_time)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        d.text(((x_ini + x_fim - tw) / 2 - bbox[0], (alt_placar - th) / 2 - bbox[1]),
               texto, fill="white", font=fonte_time)

    y_base = alt_placar

    for j in range(1, tamanho + 1):
        x = larg_cab * 2 + (j - 1) * larg_cel
        celula(x, y_base, larg_cel, alt_cab, chr(64 + j), cores["coo"], fonte=fonte_cab)

    for i in range(1, tamanho + 1):
        y = y_base + alt_cab * 2 + (i - 1) * alt_cel
        celula(0, y, larg_cab, alt_cel, str(i), cores["coo"], fonte=fonte_cab)

    celula(larg_cab, y_base + alt_cab, larg_cab, alt_cab, "×", cores["can"], fonte=fonte_cab)

    for j, p in enumerate(colunas, 1):
        x = larg_cab * 2 + (j - 1) * larg_cel
        celula(x, y_base + alt_cab, larg_cel, alt_cab, p, cores["col"])

    for i, p in enumerate(linhas, 1):
        y = y_base + alt_cab * 2 + (i - 1) * alt_cel
        celula(larg_cab, y, larg_cab, alt_cel, p, cores["lin"])

    for i in range(1, tamanho + 1):
        for j in range(1, tamanho + 1):
            x = larg_cab * 2 + (j - 1) * larg_cel
            y = y_base + alt_cab * 2 + (i - 1) * alt_cel
            coord = f"{chr(64 + j)}{i}"
            celula(x, y, larg_cel, alt_cel, coord, cores["car"],
                   cor_texto="#1E3C78", fonte=fonte_coord)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


def gerar_png_cartas(linhas, colunas, cartas_por_linha=4):
    """PNG com todas as cartas (apenas coordenadas)."""
    tamanho = len(linhas)
    cartas = [f"{chr(64 + j)}{i}" for i in range(1, tamanho + 1) for j in range(1, tamanho + 1)]

    largura_carta, altura_carta = 200, 140
    gap, margem = 15, 20

    n = len(cartas)
    n_lin = (n + cartas_por_linha - 1) // cartas_por_linha

    largura = margem * 2 + cartas_por_linha * largura_carta + (cartas_por_linha - 1) * gap
    altura = margem * 2 + n_lin * altura_carta + (n_lin - 1) * gap

    img = Image.new("RGB", (largura, altura), "white")
    d = ImageDraw.Draw(img)
    fonte = _carregar_fonte(60)

    for idx, coord in enumerate(cartas):
        linha = idx // cartas_por_linha
        col = idx % cartas_por_linha
        x = margem + col * (largura_carta + gap)
        y = margem + linha * (altura_carta + gap)

        d.rectangle([x, y, x + largura_carta, y + altura_carta],
                    fill="#D9E1F2", outline="#1E3C78", width=3)

        bbox = d.textbbox((0, 0), coord, font=fonte)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        d.text((x + (largura_carta - tw) / 2 - bbox[0],
                y + (altura_carta - th) / 2 - bbox[1]),
               coord, fill="#1E3C78", font=fonte)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


# =====================================================
# INICIALIZAÇÃO DE ESTADO
# =====================================================
if "linhas" not in st.session_state:
    st.session_state.linhas = None
    st.session_state.colunas = None
    st.session_state.tamanho_atual = None
if "nome_time_1" not in st.session_state:
    st.session_state.nome_time_1 = "Time 1"
if "nome_time_2" not in st.session_state:
    st.session_state.nome_time_2 = "Time 2"


# =====================================================
# TÍTULO
# =====================================================
st.title("🎲 Entre Linhas — Gerador de Grade")
st.caption(f"Biblioteca com **{len(PALAVRAS)} palavras**. "
           "Sorteie, monte a grade interativa e jogue com dois times.")


# =====================================================
# SIDEBAR
# =====================================================
with st.sidebar:
    st.header("⚙️ Configurações")

    tamanho = st.select_slider(
        "Tamanho da grade",
        options=[3, 4, 5, 6],
        value=5,
    )
    st.caption(f"Precisa de **{tamanho * 2} palavras** sorteadas.")

    st.divider()
    st.subheader("👥 Times")
    nome_time_1 = st.text_input(
        "🔵 Time 1 (azul)",
        value=st.session_state.nome_time_1,
        max_chars=20,
        key="input_time_1",
    )
    nome_time_2 = st.text_input(
        "🔴 Time 2 (vermelho)",
        value=st.session_state.nome_time_2,
        max_chars=20,
        key="input_time_2",
    )
    st.session_state.nome_time_1 = nome_time_1.strip() or "Time 1"
    st.session_state.nome_time_2 = nome_time_2.strip() or "Time 2"

    st.divider()
    st.subheader("📚 Biblioteca de palavras")
    st.caption(f"Total: **{len(PALAVRAS)} palavras**")

    editar = st.toggle("✏️ Editar lista de palavras", value=False)
    if editar:
        texto_atual = "\n".join(PALAVRAS)
        texto_novo = st.text_area(
            "Uma palavra por linha:",
            value=texto_atual,
            height=300,
        )
        palavras_usuario = [p.strip() for p in texto_novo.split("\n") if p.strip()]
    else:
        palavras_usuario = PALAVRAS[:]

    st.divider()
    sortear_btn = st.button("🎲 Sortear", type="primary", use_container_width=True)

    if st.button("🔄 Limpar sorteio", use_container_width=True):
        st.session_state.linhas = None
        st.session_state.colunas = None
        st.session_state.tamanho_atual = None
        st.rerun()

    st.caption("💡 Clique em Sortear novamente para gerar uma nova grade.")


# =====================================================
# LÓGICA DE SORTEIO
# =====================================================
if sortear_btn:
    if len(palavras_usuario) < tamanho * 2:
        st.error(
            f"❌ Você tem apenas **{len(palavras_usuario)} palavras**. "
            f"Precisa de pelo menos **{tamanho * 2}** para uma grade {tamanho}×{tamanho}."
        )
    else:
        linhas, colunas = sortear(palavras_usuario, tamanho)
        st.session_state.linhas = linhas
        st.session_state.colunas = colunas
        st.session_state.tamanho_atual = tamanho
        st.success(f"✅ Grade **{tamanho}×{tamanho}** sorteada com sucesso!")


# =====================================================
# CONTEÚDO PRINCIPAL
# =====================================================
if st.session_state.linhas is None:
    st.info("👈 Configure na barra lateral e clique em **🎲 Sortear** para começar.")
    st.stop()

linhas = st.session_state.linhas
colunas = st.session_state.colunas
tam = st.session_state.tamanho_atual
nome_t1 = st.session_state.nome_time_1
nome_t2 = st.session_state.nome_time_2


# ---------- GRADE INTERATIVA ----------
st.subheader(f"🎯 Grade {tam}×{tam}")
st.caption(
    f"💡 Clique em cada carta para alternar entre os times. "
    f"O placar atualiza em tempo real e o estado é salvo automaticamente."
)

render_grid_interativo(
    linhas,
    colunas,
    nome_time_1=nome_t1,
    nome_time_2=nome_t2,
    game_key="partida1",
)


# ---------- PALAVRAS SORTEADAS ----------
with st.expander("📋 Ver palavras sorteadas"):
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**Linhas (1-{tam})**")
        for i, p in enumerate(linhas, 1):
            st.markdown(f"- `{i}` → **{p}**")
    with col2:
        st.markdown(f"**Colunas (A-{chr(64 + tam)})**")
        for j, p in enumerate(colunas, 1):
            st.markdown(f"- `{chr(64 + j)}` → **{p}**")


# ---------- CARTAS PARA IMPRESSÃO ----------
st.subheader(f"🃏 Cartas ({tam * tam} no total)")
st.caption("Use Ctrl+P no navegador para imprimir esta seção.")

cartas_html = """
<style>
.cartas-wrap {
    display: flex; flex-wrap: wrap; gap: 10px; margin: 15px 0;
}
.carta-item {
    border: 3px solid #1e3c78;
    background: #d9e1f2;
    color: #1e3c78;
    font-weight: bold;
    font-size: 36px;
    width: 130px; height: 90px;
    display: flex; align-items: center; justify-content: center;
    border-radius: 8px;
    font-family: system-ui, sans-serif;
}
@media print {
    .carta-item { page-break-inside: avoid; }
}
</style>
<div class="cartas-wrap">
"""
for i in range(1, tam + 1):
    for j in range(1, tam + 1):
        cartas_html += f'<div class="carta-item">{chr(64 + j)}{i}</div>'
cartas_html += "</div>"
st.markdown(cartas_html, unsafe_allow_html=True)


# ---------- DOWNLOADS ----------
st.divider()
st.subheader("📥 Exportar")

col1, col2, col3, col4 = st.columns(4)

with col1:
    excel_bytes = gerar_excel(linhas, colunas, nome_t1, nome_t2)
    st.download_button(
        "📊 Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"entre_linhas_{tam}x{tam}_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col2:
    png_grade = gerar_png_grade(linhas, colunas, nome_t1, nome_t2)
    st.download_button(
        "🖼️ Grade (PNG)",
        data=png_grade,
        file_name=f"grade_{tam}x{tam}.png",
        mime="image/png",
        use_container_width=True,
    )

with col3:
    png_cartas = gerar_png_cartas(linhas, colunas)
    st.download_button(
        "🃏 Cartas (PNG)",
        data=png_cartas,
        file_name=f"cartas_{tam}x{tam}.png",
        mime="image/png",
        use_container_width=True,
    )

with col4:
    txt = f"TIMES:\n🔵 {nome_t1}\n🔴 {nome_t2}\n\n"
    txt += "LINHAS:\n" + "\n".join(f"{i}. {p}" for i, p in enumerate(linhas, 1))
    txt += "\n\nCOLUNAS:\n" + "\n".join(f"{chr(64+j)}. {p}" for j, p in enumerate(colunas, 1))
    st.download_button(
        "📝 Palavras (TXT)",
        data=txt.encode("utf-8"),
        file_name=f"palavras_{tam}x{tam}.txt",
        mime="text/plain",
        use_container_width=True,
    )


# ---------- RODAPÉ ----------
st.divider()
st.caption(
    "🎲 **Entre Linhas caseiro** — cada jogador recebe uma coordenada (ex.: **B3**) "
    "e dá uma dica de uma única palavra. Os outros tentam descobrir qual é o cruzamento."
)
