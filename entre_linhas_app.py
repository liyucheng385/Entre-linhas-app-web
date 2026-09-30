"""
Entre Linhas — App Web (Streamlit)
Gerador de grade e cartas para o jogo caseiro.
"""

import random
import io
from datetime import datetime

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from PIL import Image, ImageDraw, ImageFont


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
    "espelho", "lâmpada", "chave", "fechadura", "relógio", "tapete", "cortina", "almofada", "cobertor", "travesseiro",
    "lençol", "toalha", "sabonete", "escova", "porta-retrato",
    # Eletrônicos (20)
    "computador", "celular", "televisão", "rádio", "telefone", "câmera", "impressora", "teclado", "mouse", "fone",
    "ventilador", "aspirador", "liquidificador", "batedeira", "torradeira", "cafeteira", "micro-ondas",
    "ventilador de teto", "caixa de som", "videogame",
    # Profissões (30)
    "médico", "professor", "motorista", "policial", "bombeiro", "cozinheiro", "padeiro", "advogado", "engenheiro", "dentista",
    "enfermeiro", "aluno", "ator", "cantor", "dançarino", "pintor", "escritor", "jornalista", "fotógrafo", "piloto",
    "soldado", "agricultor", "pescador", "carteiro", "garçom", "eletricista", "encanador", "mecânico", "veterinário", "cientista",
    # Lugares (35)
    "escola", "hospital", "mercado", "banco", "igreja", "praça", "rua", "cidade", "fazenda", "parque",
    "zoológico", "museu", "cinema", "teatro", "restaurante", "hotel", "aeroporto", "porto", "estádio", "shopping",
    "clube", "padaria", "farmácia", "posto", "rodoviária", "praia", "campo", "vila", "aldeia", "bairro",
    "avenida", "beco", "jardim", "estação", "viela",
    # Transportes (30)
    "carro", "moto", "bicicleta", "ônibus", "caminhão", "trem", "metrô", "avião", "helicóptero", "navio",
    "barco", "canoa", "submarino", "foguete", "skate", "patins", "trator", "ambulância", "táxi", "van",
    "caminhonete", "balão", "jangada", "iate", "carruagem", "triciclo", "patinete", "teleférico", "bonde", "furgão",
    # Corpo (20)
    "cabeça", "cabelo", "olho", "nariz", "boca", "orelha", "braço", "mão", "dedo", "perna",
    "pé", "coração", "pulmão", "estômago", "sangue", "osso", "pele", "dente", "língua", "unha",
    # Cultura e ciência (33)
    "música", "dança", "arte", "pintura", "escultura", "poesia", "história", "ciência", "matemática", "física",
    "química", "biologia", "medicina", "tecnologia", "internet", "filosofia", "geografia", "astronomia",
    "literatura", "gramática",
    "amor", "amizade", "saudade", "alegria", "tristeza", "medo", "coragem", "esperança", "fé", "paz",
    "tempo", "espaço", "vida",
]


# =====================================================
# CORES E ESTILOS
# =====================================================
COR_COLUNAS = "C5E0B4"   # verde claro
COR_LINHAS  = "FFE699"   # amarelo claro
COR_CARTA   = "D9E1F2"   # azul claro
COR_COORD   = "F0F0F0"   # cinza claro
COR_CANTO   = "E6E6E6"   # cinza médio


# =====================================================
# FUNÇÕES DE SORTEIO
# =====================================================
def sortear(palavras: list[str], tamanho: int):
    """Sorteia 2*tamanho palavras sem repetir."""
    if len(palavras) < tamanho * 2:
        return None, None
    sorteadas = random.sample(palavras, tamanho * 2)
    return sorteadas[:tamanho], sorteadas[tamanho:]


# =====================================================
# GERAÇÃO DE EXCEL
# =====================================================
def gerar_excel(linhas, colunas):
    """Gera arquivo .xlsx com grade visual + aba de cartas."""
    tamanho = len(linhas)
    wb = Workbook()
    ws = wb.active
    ws.title = "Grade"
    
    # Estilos
    fill_col = PatternFill("solid", fgColor=COR_COLUNAS)
    fill_lin = PatternFill("solid", fgColor=COR_LINHAS)
    fill_car = PatternFill("solid", fgColor=COR_CARTA)
    fill_coo = PatternFill("solid", fgColor=COR_COORD)
    fill_can = PatternFill("solid", fgColor=COR_CANTO)
    
    f_cinza = Font(bold=True, size=14, color="323232")
    f_canto = Font(bold=True, size=14)
    f_palavra = Font(bold=True, size=11)
    f_coord_med = Font(bold=True, size=16, color="1E3C78")
    f_coord_gde = Font(bold=True, size=40, color="1E3C78")
    
    alin = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fina = Border(*[Side(style="thin")] * 4)
    media = Border(*[Side(style="medium")] * 4)
    
    linha_ini, col_ini = 4, 4
    
    # Canto
    c = ws.cell(row=linha_ini - 1, column=col_ini, value="×")
    c.font, c.fill, c.alignment, c.border = f_canto, fill_can, alin, media
    
    # Letras A, B, C...
    for j in range(1, tamanho + 1):
        c = ws.cell(row=linha_ini - 2, column=col_ini + j, value=chr(64 + j))
        c.font, c.fill, c.alignment, c.border = f_cinza, fill_coo, alin, fina
    
    # Números 1, 2, 3...
    for i in range(1, tamanho + 1):
        c = ws.cell(row=linha_ini + i - 1, column=col_ini - 1, value=i)
        c.font, c.fill, c.alignment, c.border = f_cinza, fill_coo, alin, fina
    
    # Palavras das colunas
    for j, p in enumerate(colunas, 1):
        c = ws.cell(row=linha_ini - 1, column=col_ini + j, value=p)
        c.font, c.fill, c.alignment, c.border = f_palavra, fill_col, alin, media
    
    # Palavras das linhas
    for i, p in enumerate(linhas, 1):
        c = ws.cell(row=linha_ini + i - 1, column=col_ini, value=p)
        c.font, c.fill, c.alignment, c.border = f_palavra, fill_lin, alin, media
    
    # Cartas (só coordenada)
    for i in range(1, tamanho + 1):
        for j in range(1, tamanho + 1):
            coord = f"{chr(64 + j)}{i}"
            c = ws.cell(row=linha_ini + i - 1, column=col_ini + j, value=coord)
            c.font, c.fill, c.alignment, c.border = f_coord_med, fill_car, alin, fina
    
    # Larguras e alturas
    ws.column_dimensions[get_column_letter(col_ini - 1)].width = 5
    ws.column_dimensions[get_column_letter(col_ini)].width = 14
    for j in range(1, tamanho + 1):
        ws.column_dimensions[get_column_letter(col_ini + j)].width = 16
    ws.row_dimensions[linha_ini - 2].height = 22
    ws.row_dimensions[linha_ini - 1].height = 30
    for i in range(1, tamanho + 1):
        ws.row_dimensions[linha_ini + i - 1].height = 45
    
    # Área de impressão
    p1 = f"{get_column_letter(col_ini - 1)}{linha_ini - 2}"
    p2 = f"{get_column_letter(col_ini + tamanho)}{linha_ini + tamanho - 1}"
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
    
    # Salva em memória
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# =====================================================
# GERAÇÃO DE PNG (grade + cartas)
# =====================================================
def _carregar_fonte(tamanho: int):
    """Tenta carregar uma fonte decente; cai para a padrão se falhar."""
    for nome in ("DejaVuSans-Bold.ttf", "arialbd.ttf", "Arial Bold.ttf"):
        try:
            return ImageFont.truetype(nome, tamanho)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def gerar_png_grade(linhas, colunas):
    """Gera imagem PNG da grade visual."""
    tamanho = len(linhas)
    largura_cel = 130
    altura_cel = 90
    largura_cab = 60
    altura_cab = 40
    
    largura = largura_cab * 2 + largura_cel * tamanho
    altura = altura_cab * 2 + altura_cel * tamanho
    
    img = Image.new("RGB", (largura, altura), "white")
    d = ImageDraw.Draw(img)
    
    fonte_pal = _carregar_fonte(16)
    fonte_coord = _carregar_fonte(22)
    fonte_cab = _carregar_fonte(18)
    
    def desenhar_celula(x, y, w, h, texto, cor_fundo, cor_texto="black", fonte=None):
        d.rectangle([x, y, x + w, y + h], fill=cor_fundo, outline="black", width=2)
        if texto:
            fonte = fonte or fonte_pal
            bbox = d.textbbox((0, 0), texto, font=fonte)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            d.text((x + (w - tw) / 2 - bbox[0], y + (h - th) / 2 - bbox[1]),
                   texto, fill=cor_texto, font=fonte)
    
    cores = {
        "col": "#C5E0B4",
        "lin": "#FFE699",
        "car": "#D9E1F2",
        "coo": "#F0F0F0",
        "can": "#E6E6E6",
    }
    
    # Cabeçalho superior (letras)
    for j in range(1, tamanho + 1):
        x = largura_cab * 2 + (j - 1) * largura_cel
        desenhar_celula(x, 0, largura_cel, altura_cab, chr(64 + j), cores["coo"], fonte=fonte_cab)
    
    # Cabeçalho esquerdo (números)
    for i in range(1, tamanho + 1):
        y = altura_cab * 2 + (i - 1) * altura_cel
        desenhar_celula(0, y, largura_cab, altura_cel, str(i), cores["coo"], fonte=fonte_cab)
    
    # Canto ×
    desenhar_celula(largura_cab, altura_cab, largura_cab, altura_cab, "×", cores["can"], fonte=fonte_cab)
    
    # Palavras das colunas
    for j, p in enumerate(colunas, 1):
        x = largura_cab * 2 + (j - 1) * largura_cel
        desenhar_celula(x, altura_cab, largura_cel, altura_cab, p, cores["col"])
    
    # Palavras das linhas
    for i, p in enumerate(linhas, 1):
        y = altura_cab * 2 + (i - 1) * altura_cel
        desenhar_celula(largura_cab, y, largura_cab, altura_cel, p, cores["lin"])
    
    # Cartas
    for i in range(1, tamanho + 1):
        for j in range(1, tamanho + 1):
            x = largura_cab * 2 + (j - 1) * largura_cel
            y = altura_cab * 2 + (i - 1) * altura_cel
            coord = f"{chr(64 + j)}{i}"
            desenhar_celula(x, y, largura_cel, altura_cel, coord, cores["car"],
                            cor_texto="#1E3C78", fonte=fonte_coord)
    
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


def gerar_png_cartas(linhas, colunas, cartas_por_linha=4):
    """Gera PNG com todas as cartas em grid."""
    tamanho = len(linhas)
    cartas = [f"{chr(64 + j)}{i}" for i in range(1, tamanho + 1) for j in range(1, tamanho + 1)]
    
    largura_carta = 200
    altura_carta = 140
    gap = 15
    margem = 20
    
    n_cartas = len(cartas)
    n_linhas = (n_cartas + cartas_por_linha - 1) // cartas_por_linha
    
    largura = margem * 2 + cartas_por_linha * largura_carta + (cartas_por_linha - 1) * gap
    altura = margem * 2 + n_linhas * altura_carta + (n_linhas - 1) * gap
    
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
# INTERFACE STREAMLIT
# =====================================================
st.set_page_config(
    page_title="Entre Linhas — Gerador de Grade",
    page_icon="🎲",
    layout="wide",
)

st.title("🎲 Entre Linhas — Gerador de Grade")
st.caption(f"Biblioteca com **{len(PALAVRAS)} palavras**. "
           "Sorteia, monta a grade colorida e gera cartas para imprimir.")

# ---------- SIDEBAR ----------
with st.sidebar:
    st.header("⚙️ Configurações")
    
    tamanho = st.select_slider(
        "Tamanho da grade",
        options=[3, 4, 5, 6],
        value=5,
    )
    st.caption(f"Precisa de **{tamanho * 2} palavras** sorteadas.")
    
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
        st.session_state.pop("linhas", None)
        st.session_state.pop("colunas", None)
        st.session_state.pop("tamanho_atual", None)
        st.rerun()
    
    st.divider()
    st.caption("💡 **Dica:** clique em Sortear novamente para gerar uma nova grade.")


# ---------- ESTADO ----------
if "linhas" not in st.session_state:
    st.session_state.linhas = None
    st.session_state.colunas = None
    st.session_state.tamanho_atual = None

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


# ---------- CONTEÚDO PRINCIPAL ----------
if st.session_state.linhas is None:
    st.info("👈 Configure na barra lateral e clique em **🎲 Sortear** para começar.")
    st.stop()

linhas = st.session_state.linhas
colunas = st.session_state.colunas
tam = st.session_state.tamanho_atual


# ---------- GRADE VISUAL (HTML) ----------
st.subheader(f"🎯 Grade {tam}×{tam}")

# CSS
st.markdown("""
<style>
.grade-el {
    display: grid;
    gap: 4px;
    margin: 10px 0 20px 0;
    width: fit-content;
}
.cel {
    border: 2px solid #4a5568;
    padding: 12px 8px;
    text-align: center;
    font-family: system-ui, sans-serif;
    display: flex;
    align-items: center;
    justify-content: center;
    min-height: 45px;
    font-size: 13px;
    border-radius: 4px;
}
.cel-coord {
    background: #f0f0f0;
    font-weight: bold;
    font-size: 16px;
    color: #323232;
}
.cel-canto { background: #e6e6e6; font-weight: bold; font-size: 20px; }
.cel-col   { background: #c5e0b4; font-weight: bold; font-size: 13px; }
.cel-lin   { background: #ffe699; font-weight: bold; font-size: 13px; }
.cel-carta {
    background: #d9e1f2;
    color: #1e3c78;
    font-weight: bold;
    font-size: 20px;
    min-height: 55px;
}
.cartas-wrap {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin: 15px 0;
}
.carta-item {
    border: 3px solid #1e3c78;
    background: #d9e1f2;
    color: #1e3c78;
    font-weight: bold;
    font-size: 36px;
    width: 130px;
    height: 90px;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 8px;
    font-family: system-ui, sans-serif;
}
@media print {
    .no-print { display: none; }
    .carta-item { page-break-inside: avoid; }
}
</style>
""", unsafe_allow_html=True)

# Monta o grid
n_cols = tam + 2  # 1 coluna p/ números + 1 coluna p/ palavras-linha + N colunas
html = f'<div class="grade-el" style="grid-template-columns: 45px 130px repeat({tam}, 130px);">'

# Linha 1 (letras A-F): célula vazia + célula vazia + letras
html += '<div></div><div></div>'
for j in range(1, tam + 1):
    html += f'<div class="cel cel-coord">{chr(64 + j)}</div>'

# Linha 2: vazio + canto × + palavras-coluna
html += '<div></div>'
html += '<div class="cel cel-canto">×</div>'
for col in colunas:
    html += f'<div class="cel cel-col">{col}</div>'

# Linhas da grade
for i, lin in enumerate(linhas, 1):
    html += f'<div class="cel cel-coord">{i}</div>'
    html += f'<div class="cel cel-lin">{lin}</div>'
    for j in range(1, tam + 1):
        html += f'<div class="cel cel-carta">{chr(64 + j)}{i}</div>'

html += '</div>'
st.markdown(html, unsafe_allow_html=True)


# ---------- PALAVRAS SORTEADAS ----------
with st.expander("📋 Ver palavras sorteadas"):
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Linhas (1-{})**".format(tam))
        for i, p in enumerate(linhas, 1):
            st.markdown(f"- `{i}` → **{p}**")
    with col2:
        st.markdown("**Colunas (A-{})**".format(chr(64 + tam)))
        for j, p in enumerate(colunas, 1):
            st.markdown(f"- `{chr(64 + j)}` → **{p}**")


# ---------- CARTAS ----------
st.subheader(f"🃏 Cartas ({tam * tam} no total)")
st.caption("Clique em Ctrl+P no navegador para imprimir esta seção.")

cartas_html = '<div class="cartas-wrap">'
for i in range(1, tam + 1):
    for j in range(1, tam + 1):
        cartas_html += f'<div class="carta-item">{chr(64 + j)}{i}</div>'
cartas_html += '</div>'
st.markdown(cartas_html, unsafe_allow_html=True)


# ---------- DOWNLOADS ----------
st.divider()
st.subheader("📥 Exportar")

col1, col2, col3 = st.columns(3)

with col1:
    excel_bytes = gerar_excel(linhas, colunas)
    st.download_button(
        "📊 Baixar Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"entre_linhas_{tam}x{tam}_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col2:
    png_grade = gerar_png_grade(linhas, colunas)
    st.download_button(
        "🖼️ Baixar grade (PNG)",
        data=png_grade,
        file_name=f"grade_{tam}x{tam}.png",
        mime="image/png",
        use_container_width=True,
    )

with col3:
    png_cartas = gerar_png_cartas(linhas, colunas)
    st.download_button(
        "🃏 Baixar cartas (PNG)",
        data=png_cartas,
        file_name=f"cartas_{tam}x{tam}.png",
        mime="image/png",
        use_container_width=True,
    )

# Extra: TXT com palavras
txt = "LINHAS:\n" + "\n".join(f"{i}. {p}" for i, p in enumerate(linhas, 1))
txt += "\n\nCOLUNAS:\n" + "\n".join(f"{chr(64+j)}. {p}" for j, p in enumerate(colunas, 1))
st.download_button(
    "📝 Baixar palavras (TXT)",
    data=txt.encode("utf-8"),
    file_name=f"palavras_{tam}x{tam}.txt",
    mime="text/plain",
)


# ---------- RODAPÉ ----------
st.divider()
st.caption(
    "🎲 **Entre Linhas caseiro** — cada jogador recebe uma coordenada (ex.: **B3**) "
    "e dá uma dica de uma única palavra. Os outros tentam descobrir qual é o cruzamento."
)