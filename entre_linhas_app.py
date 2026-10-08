"""
Entre Linhas — Desktop + QR único + Recuperação de Sala + Manual com imagens
e regra de comunicação detalhada (dica única sem radical).
"""

import io
import time
import random
import string
from datetime import datetime

import requests
import qrcode
import streamlit as st
import streamlit.components.v1 as components
from streamlit_autorefresh import st_autorefresh
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont


# =====================================================
# CONFIGURAÇÃO DA PÁGINA
# =====================================================
st.set_page_config(
    page_title="Entre Linhas — Jogo",
    page_icon="🎲",
    layout="wide",
)


# =====================================================
# CSS GLOBAL
# =====================================================
st.markdown("""
<style>
    [data-testid="column"] { padding: 0 12px; }
    [data-testid="stMetricValue"] { font-size: 26px !important; }
    [data-testid="stMetricLabel"] { font-size: 13px !important; }
    h1 { font-size: 26px !important; margin-bottom: 4px !important; }
    h2 { font-size: 22px !important; }
    .stButton > button { font-size: 15px !important; }
</style>
""", unsafe_allow_html=True)


# =====================================================
# SECRETS
# =====================================================
def _get_secret(key, default=""):
    try:
        return st.secrets[key]
    except (KeyError, FileNotFoundError):
        return default
    except Exception:
        return default


FIREBASE_URL = _get_secret("FIREBASE_URL", "")
APP_URL = _get_secret("APP_URL", "http://localhost:8501")


# =====================================================
# BIBLIOTECA DE 343 PALAVRAS
# =====================================================
PALAVRAS = [
    "sol", "lua", "estrela", "céu", "nuvem", "chuva", "vento", "neve", "gelo", "fogo",
    "água", "terra", "ar", "mar", "rio", "lago", "floresta", "montanha", "deserto", "praia",
    "ilha", "vulcão", "caverna", "cachoeira", "campo", "vale", "colina", "duna", "pântano", "savana",
    "cachorro", "gato", "pássaro", "peixe", "cavalo", "vaca", "porco", "galinha", "pato", "coelho",
    "leão", "tigre", "elefante", "macaco", "cobra", "aranha", "abelha", "formiga", "borboleta", "urso",
    "lobo", "raposa", "coruja", "águia", "tartaruga", "jacaré", "sapo", "golfinho", "baleia", "tubarão",
    "polvo", "caranguejo", "camelo", "girafa", "zebra", "panda", "koala", "pinguim", "esquilo", "morcego",
    "arroz", "feijão", "macarrão", "pão", "carne", "frango", "ovo", "queijo", "leite", "café",
    "chá", "suco", "refrigerante", "cerveja", "vinho", "sal", "açúcar", "óleo", "farinha", "manteiga",
    "mel", "chocolate", "bolo", "sorvete", "doce", "pizza", "sanduíche", "hambúrguer", "batata", "tomate",
    "cebola", "alho", "cenoura", "alface", "brócolis", "milho", "ervilha", "abóbora", "beterraba", "pepino",
    "pimentão", "maçã", "banana", "laranja", "uva", "morango", "abacaxi", "manga", "melancia", "limão",
    "pêssego", "pera", "kiwi", "cereja", "ameixa", "coco", "abacate", "goiaba", "maracujá", "pitanga",
    "casa", "apartamento", "quarto", "sala", "cozinha", "banheiro", "quintal", "jardim", "garagem", "sótão",
    "porão", "varanda", "escritório", "biblioteca", "porta", "janela", "parede", "chão", "teto", "telhado",
    "escada", "corredor", "cama", "mesa", "cadeira", "sofá", "armário", "geladeira", "fogão", "pia",
    "espelho", "lâmpada", "chave", "fechadura", "relógio", "tapete", "cortina", "almofada", "cobertor",
    "travesseiro", "lençol", "toalha", "sabonete", "escova", "porta-retrato",
    "computador", "celular", "televisão", "rádio", "telefone", "câmera", "impressora", "teclado", "mouse", "fone",
    "ventilador", "aspirador", "liquidificador", "batedeira", "torradeira", "cafeteira", "micro-ondas",
    "ventilador de teto", "caixa de som", "videogame",
    "médico", "professor", "motorista", "policial", "bombeiro", "cozinheiro", "padeiro", "advogado",
    "engenheiro", "dentista", "enfermeiro", "aluno", "ator", "cantor", "dançarino", "pintor", "escritor",
    "jornalista", "fotógrafo", "piloto", "soldado", "agricultor", "pescador", "carteiro", "garçom",
    "eletricista", "encanador", "mecânico", "veterinário", "cientista",
    "escola", "hospital", "mercado", "banco", "igreja", "praça", "rua", "cidade", "fazenda", "parque",
    "zoológico", "museu", "cinema", "teatro", "restaurante", "hotel", "aeroporto", "porto", "estádio",
    "shopping", "clube", "padaria", "farmácia", "posto", "rodoviária", "praia", "campo", "vila", "aldeia",
    "bairro", "avenida", "beco", "jardim", "estação", "viela",
    "carro", "moto", "bicicleta", "ônibus", "caminhão", "trem", "metrô", "avião", "helicóptero", "navio",
    "barco", "canoa", "submarino", "foguete", "skate", "patins", "trator", "ambulância", "táxi", "van",
    "caminhonete", "balão", "jangada", "iate", "carruagem", "triciclo", "patinete", "teleférico", "bonde",
    "furgão",
    "cabeça", "cabelo", "olho", "nariz", "boca", "orelha", "braço", "mão", "dedo", "perna",
    "pé", "coração", "pulmão", "estômago", "sangue", "osso", "pele", "dente", "língua", "unha",
    "música", "dança", "arte", "pintura", "escultura", "poesia", "história", "ciência", "matemática",
    "física", "química", "biologia", "medicina", "tecnologia", "internet", "filosofia", "geografia",
    "astronomia", "literatura", "gramática",
    "amor", "amizade", "saudade", "alegria", "tristeza", "medo", "coragem", "esperança", "fé", "paz",
    "tempo", "espaço", "vida",
]


# =====================================================
# FONTE ÚNICA
# =====================================================
def _carregar_fonte(tamanho):
    for nome in ("DejaVuSans-Bold.ttf", "arialbd.ttf", "Arial Bold.ttf",
                 "DejaVuSans.ttf", "arial.ttf", "Arial.ttf"):
        try:
            return ImageFont.truetype(nome, tamanho)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


# =====================================================
# FIREBASE
# =====================================================
def fb_url(sala_id, path=""):
    base = FIREBASE_URL.rstrip("/")
    if path:
        return f"{base}/salas/{sala_id}/{path}.json"
    return f"{base}/salas/{sala_id}.json"


def fb_get(sala_id):
    try:
        r = requests.get(fb_url(sala_id), timeout=5)
        return r.json() or {}
    except Exception:
        return {}


def fb_patch(sala_id, dados):
    try:
        requests.patch(fb_url(sala_id), json=dados, timeout=5)
    except Exception:
        pass


def fb_put(sala_id, dados):
    try:
        requests.put(fb_url(sala_id), json=dados, timeout=5)
    except Exception:
        pass


def fb_delete(sala_id):
    try:
        requests.delete(fb_url(sala_id), timeout=5)
    except Exception:
        pass


# =====================================================
# UTILITÁRIOS
# =====================================================
def gerar_sala_id():
    chars = string.ascii_uppercase + string.digits
    return "".join(random.choices(chars, k=8))


def sortear(palavras, tamanho):
    if len(palavras) < tamanho * 2:
        return None, None
    sorteadas = random.sample(palavras, tamanho * 2)
    return sorteadas[:tamanho], sorteadas[tamanho:]


def gerar_coordenadas(tamanho):
    coords = [f"{chr(64 + j)}{i}" for i in range(1, tamanho + 1) for j in range(1, tamanho + 1)]
    random.shuffle(coords)
    return coords


def _agora_ms():
    return int(time.time() * 1000)


def gerar_qr_code(url):
    qr = qrcode.QRCode(version=1, box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


# =====================================================
# RECUPERAÇÃO DE SALA
# =====================================================
def recuperar_sala(sala_id):
    dados = fb_get(sala_id)
    if not dados or "linhas_fb" not in dados:
        return False
    try:
        st.session_state.sala_id = sala_id
        st.session_state.tamanho_atual = dados.get("tamanho", 5)
        st.session_state.linhas = dados.get("linhas_fb", [])
        st.session_state.colunas = dados.get("colunas_fb", [])
        st.session_state.deck = dados.get("deck_fb", [])
        st.session_state.sorteadas = dados.get("sorteadas_fb", [])
        st.session_state.estados = dados.get("estados_fb", {})
        st.session_state.nome_time_1 = dados.get("nome_time_1_fb", "Time 1")
        st.session_state.nome_time_2 = dados.get("nome_time_2_fb", "Time 2")
        st.session_state.time_atual = dados.get("time_atual", 1)
        st.session_state.carta_atual = dados.get("coord")
        st.session_state.fase = dados.get("estado", "aguardando")
        st.session_state.timer_ativo = dados.get("timer_ativo", False)
        st.session_state.timer_minutos = dados.get("timer_minutos_fb", 2)
        st.session_state.timer_segundos = dados.get("timer_segundos_fb", 0)
        st.session_state.timer_mudo = dados.get("timer_mudo_fb", False)
        st.session_state.timer_rodando = False
        st.session_state.tempo_pausado_segundos = None
        st.session_state.turno_iniciado_em = _agora_ms()
        st.session_state.turno_contador += 1
        st.session_state.msg_sucesso = f"✅ Sala {sala_id} recuperada!"
        return True
    except Exception:
        return False


# =====================================================
# DETECÇÃO DE VITÓRIA
# =====================================================
def detectar_vencedor(estados, nome_t1, nome_t2):
    p1 = sum(1 for v in estados.values() if v == 1)
    p2 = sum(1 for v in estados.values() if v == 2)
    desc = sum(1 for v in estados.values() if v == 3)
    if p1 == 0 and p2 == 0:
        return "sem_dados", p1, p2, desc, "#64748b"
    if p1 > p2:
        return "t1", p1, p2, desc, "#1e3c78"
    if p2 > p1:
        return "t2", p1, p2, desc, "#dc2626"
    return "empate", p1, p2, desc, "#eab308"


def render_banner_vitoria(vencedor, p1, p2, desc, cor, nome_t1, nome_t2):
    total_computado = p1 + p2
    pct1 = (p1 / total_computado * 100) if total_computado > 0 else 0
    pct2 = (p2 / total_computado * 100) if total_computado > 0 else 0
    if vencedor == "t1":
        titulo, subtitulo, emoji = f"🏆 {nome_t1} VENCEU!", f"{p1} × {p2} cartas conquistadas", "🔵"
    elif vencedor == "t2":
        titulo, subtitulo, emoji = f"🏆 {nome_t2} VENCEU!", f"{p2} × {p1} cartas conquistadas", "🔴"
    elif vencedor == "empate":
        titulo, subtitulo, emoji = "🤝 EMPATE!", f"{p1} × {p2} cartas conquistadas", "⚪"
    else:
        titulo, subtitulo, emoji = "🏁 FIM DE JOGO", "Nenhuma carta foi conquistada", "🎲"

    st.markdown(
        f"""
        <div style="padding: 32px 24px;
                    background: linear-gradient(135deg, {cor}18, {cor}08);
                    border: 4px solid {cor}; border-radius: 20px;
                    text-align: center; font-family: system-ui;
                    margin: 16px 0 24px 0;
                    box-shadow: 0 8px 32px {cor}30;">
            <div style="font-size: 64px; line-height: 1; margin-bottom: 8px;">{emoji}</div>
            <div style="font-size: 38px; font-weight: bold; color: {cor};
                        letter-spacing: 1px; line-height: 1.1;">{titulo}</div>
            <div style="font-size: 20px; color: #64748b; margin-top: 12px;
                        font-weight: 500;">{subtitulo}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col_e, col_d = st.columns(2)
    with col_e:
        st.markdown(
            f"""
            <div style="padding: 16px;
                        background: {('#1e3c78' if vencedor == 't1' else '#f8fafc')};
                        border: 2px solid #1e3c78; border-radius: 12px;
                        text-align: center; font-family: system-ui;">
                <div style="font-size: 13px;
                            color: {('#ffffff' if vencedor == 't1' else '#64748b')};
                            text-transform: uppercase; letter-spacing: 1.5px;">
                    🔵 {nome_t1}</div>
                <div style="font-size: 42px; font-weight: bold;
                            color: {('#ffffff' if vencedor == 't1' else '#1e3c78')};
                            margin-top: 4px;">{p1}</div>
                <div style="font-size: 12px;
                            color: {('#e0e7ff' if vencedor == 't1' else '#94a3b8')};
                            margin-top: 4px;">{pct1:.0f}% das cartas</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_d:
        st.markdown(
            f"""
            <div style="padding: 16px;
                        background: {('#dc2626' if vencedor == 't2' else '#f8fafc')};
                        border: 2px solid #dc2626; border-radius: 12px;
                        text-align: center; font-family: system-ui;">
                <div style="font-size: 13px;
                            color: {('#ffffff' if vencedor == 't2' else '#64748b')};
                            text-transform: uppercase; letter-spacing: 1.5px;">
                    🔴 {nome_t2}</div>
                <div style="font-size: 42px; font-weight: bold;
                            color: {('#ffffff' if vencedor == 't2' else '#dc2626')};
                            margin-top: 4px;">{p2}</div>
                <div style="font-size: 12px;
                            color: {('#fee2e2' if vencedor == 't2' else '#94a3b8')};
                            margin-top: 4px;">{pct2:.0f}% das cartas</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.write("")
    st.caption(f"Total: **{total_computado + desc}** cartas — Descartadas: **{desc}**")


# =====================================================
# TIMER VISUAL
# =====================================================
def render_timer(tempo_total, iniciado_em_ms, mudo=False, key="timer"):
    if tempo_total <= 0:
        return
    mudo_js = "true" if mudo else "false"
    html = f"""
    <!DOCTYPE html><html><head><style>
        body {{ margin: 0; font-family: system-ui, sans-serif; }}
        .timer-wrap {{ padding: 20px; border-radius: 16px;
            background: #22c55e15; border: 4px solid #22c55e;
            text-align: center; transition: background 0.3s, border-color 0.3s; }}
        .timer-tempo {{ font-size: 110px; font-weight: 900;
            font-variant-numeric: tabular-nums; line-height: 1;
            margin-bottom: 8px; color: #22c55e; transition: color 0.3s;
            letter-spacing: 3px; text-shadow: 0 3px 14px rgba(0,0,0,0.06); }}
        .timer-label {{ font-size: 13px; color: #64748b;
            text-transform: uppercase; letter-spacing: 2px; margin-top: 6px; }}
        .timer-bar {{ width: 100%; height: 14px; background: #e2e8f0;
            border-radius: 7px; overflow: hidden; margin-top: 16px; }}
        .timer-fill {{ height: 100%; width: 100%; background: #22c55e;
            transition: width 1s linear, background 0.3s; }}
    </style></head><body>
        <div class="timer-wrap" id="wrap">
            <div class="timer-tempo" id="tempo">--:--</div>
            <div class="timer-label">Tempo restante</div>
            <div class="timer-bar"><div class="timer-fill" id="fill"></div></div>
        </div>
        <script>
        (function() {{
            const total = {tempo_total};
            const inicio = {iniciado_em_ms};
            const mudo = {mudo_js};
            const key = "beeps_{key}_" + inicio;
            let beeps = [];
            try {{ beeps = JSON.parse(sessionStorage.getItem(key) || "[]"); }} catch(e) {{}}
            let ctx = null;
            function getCtx() {{
                if (!ctx) {{ try {{ ctx = new (window.AudioContext || window.webkitAudioContext)(); }}
                            catch(e) {{ return null; }} }}
                if (ctx && ctx.state === 'suspended') ctx.resume();
                return ctx;
            }}
            function beep(f, d, v) {{
                if (mudo) return;
                const c = getCtx(); if (!c) return;
                try {{
                    const o = c.createOscillator(); const g = c.createGain();
                    o.type = 'sine'; o.frequency.value = f;
                    g.gain.setValueAtTime(v, c.currentTime);
                    g.gain.exponentialRampToValueAtTime(0.001, c.currentTime + d);
                    o.connect(g); g.connect(c.destination);
                    o.start(c.currentTime); o.stop(c.currentTime + d);
                }} catch(e) {{}}
            }}
            function tocar(r) {{
                if (r <= 0) beep(220, 0.8, 0.30);
                else if (r <= 3) beep(880, 0.15, 0.25);
                else if (r <= 10) beep(440, 0.10, 0.15);
            }}
            function fmt(s) {{
                if (s < 0) s = 0;
                return String(Math.floor(s/60)).padStart(2,"0") + ":" + String(s%60).padStart(2,"0");
            }}
            function tick() {{
                let r = Math.ceil(total - (Date.now() - inicio) / 1000);
                if (r < 0) r = 0;
                const et = document.getElementById("tempo");
                const ef = document.getElementById("fill");
                const ew = document.getElementById("wrap");
                et.textContent = fmt(r);
                const pct = total > 0 ? (r / total) * 100 : 0;
                ef.style.width = pct + "%";
                let cor;
                if (pct > 50) cor = "#22c55e";
                else if (pct > 20) cor = "#eab308";
                else cor = "#dc2626";
                et.style.color = cor; ef.style.background = cor;
                ew.style.background = cor + "15"; ew.style.borderColor = cor;
                if (r > 0 && r <= 10 && !beeps.includes(r)) {{
                    tocar(r); beeps.push(r);
                    try {{ sessionStorage.setItem(key, JSON.stringify(beeps)); }} catch(e) {{}}
                }}
                if (r === 0 && !beeps.includes(0)) {{
                    tocar(0); beeps.push(0);
                    try {{ sessionStorage.setItem(key, JSON.stringify(beeps)); }} catch(e) {{}}
                }}
                if (r > 0) setTimeout(tick, 250);
            }}
            function unlock() {{
                const c = getCtx();
                if (c && c.state === 'running') {{
                    document.removeEventListener('click', unlock);
                    document.removeEventListener('touchstart', unlock);
                }}
            }}
            document.addEventListener('click', unlock);
            document.addEventListener('touchstart', unlock);
            tick();
        }})();
        </script></body></html>
    """
    components.html(html, height=310, scrolling=False)


# =====================================================
# GRID COM AUTO-SCALE
# =====================================================
def render_grid_estados(linhas, colunas, estados, nome_time_1, nome_time_2):
    tam = len(linhas)
    tamanhos = {
        3: {"cell": 230, "coord": 65, "min_h": 100, "f_letter": 30, "f_word": 24, "f_coord": 48},
        4: {"cell": 200, "coord": 60, "min_h": 90, "f_letter": 28, "f_word": 22, "f_coord": 42},
        5: {"cell": 180, "coord": 55, "min_h": 82, "f_letter": 26, "f_word": 20, "f_coord": 38},
        6: {"cell": 160, "coord": 50, "min_h": 74, "f_letter": 24, "f_word": 18, "f_coord": 34},
    }
    sz = tamanhos.get(tam, tamanhos[5])
    cell, coord_w, min_h = sz["cell"], sz["coord"], sz["min_h"]
    f_letter, f_word, f_coord = sz["f_letter"], sz["f_word"], sz["f_coord"]

    header1 = '<div></div><div></div>'
    for j in range(1, tam + 1):
        header1 += f'<div class="cel cel-coord">{chr(64 + j)}</div>'

    header2 = '<div></div><div class="cel cel-canto">×</div>'
    for col in colunas:
        header2 += f'<div class="cel cel-col">{col}</div>'

    rows = ""
    for i, lin in enumerate(linhas, 1):
        rows += f'<div class="cel cel-coord">{i}</div>'
        rows += f'<div class="cel cel-lin">{lin}</div>'
        for j in range(1, tam + 1):
            coord = f"{chr(64 + j)}{i}"
            estado = estados.get(coord, 0)
            rows += f'<div class="cel cel-carta state-{estado}">{coord}</div>'

    p1 = sum(1 for v in estados.values() if v == 1)
    p2 = sum(1 for v in estados.values() if v == 2)
    desc = sum(1 for v in estados.values() if v == 3)

    html = f"""
<!DOCTYPE html><html><head><style>
    html, body {{ margin: 0; padding: 0; font-family: system-ui, sans-serif; }}
    .placar {{ display: flex; gap: 24px; align-items: center;
        padding: 12px 16px; background: #f8fafc;
        border: 2px solid #cbd5e1; border-radius: 10px;
        margin-bottom: 14px; font-size: 16px; font-weight: 600; flex-wrap: wrap; }}
    .placar-time {{ display: flex; align-items: center; gap: 8px; }}
    .placar-dot {{ width: 22px; height: 22px; border-radius: 4px;
        border: 2px solid #4a5568; flex-shrink: 0; }}
    .dot-b {{ background: #1e3c78; }} .dot-r {{ background: #dc2626; }}
    .dot-x {{ background: #94a3b8; }}
    .placar-nome {{ font-size: 15px; max-width: 180px;
        overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    .placar-num {{ font-size: 26px; min-width: 34px; text-align: center; font-weight: bold; }}
    .placar-num.blue {{ color: #1e3c78; }} .placar-num.red {{ color: #dc2626; }}
    .placar-num.grey {{ color: #64748b; }}
    .grid-wrap {{ width: 100%; overflow: visible; position: relative; }}
    .grid-el {{ display: grid;
        grid-template-columns: {coord_w}px {cell}px repeat({tam}, {cell}px);
        gap: 6px; width: fit-content;
        transform-origin: top left; transition: transform 0.15s ease-out; }}
    .cel {{ border: 3px solid #4a5568; padding: 10px 6px; text-align: center;
        display: flex; align-items: center; justify-content: center;
        min-height: {min_h}px; font-size: {f_word}px;
        border-radius: 6px; user-select: none;
        overflow: hidden; word-break: break-word; line-height: 1.15; }}
    .cel-coord {{ background: #f0f0f0; font-weight: bold;
        font-size: {f_letter}px; color: #323232; }}
    .cel-canto {{ background: #e6e6e6; font-weight: bold; font-size: {f_letter + 2}px; }}
    .cel-col {{ background: #c5e0b4; font-weight: bold; font-size: {f_word}px; }}
    .cel-lin {{ background: #ffe699; font-weight: bold; font-size: {f_word}px; }}
    .cel-carta {{ font-weight: 900; font-size: {f_coord}px;
        min-height: {min_h}px; transition: background 0.3s, color 0.3s;
        letter-spacing: 0.5px; }}
    .cel-carta.state-0 {{ background: #d9e1f2; color: #1e3c78; }}
    .cel-carta.state-1 {{ background: #1e3c78; color: #ffffff; }}
    .cel-carta.state-2 {{ background: #dc2626; color: #ffffff; }}
    .cel-carta.state-3 {{ background: #cbd5e1; color: #64748b;
        text-decoration: line-through; opacity: 0.7; }}
</style></head><body>
<div class="placar">
    <div class="placar-time"><span class="placar-dot dot-b"></span>
        <span class="placar-nome">{nome_time_1}</span>
        <span class="placar-num blue">{p1}</span></div>
    <div class="placar-time"><span class="placar-dot dot-r"></span>
        <span class="placar-nome">{nome_time_2}</span>
        <span class="placar-num red">{p2}</span></div>
    <div class="placar-time"><span class="placar-dot dot-x"></span>
        <span class="placar-nome">Descartadas</span>
        <span class="placar-num grey">{desc}</span></div>
</div>
<div class="grid-wrap" id="grid-wrap">
    <div class="grid-el" id="grid-el">
        {header1}
        {header2}
        {rows}
    </div>
</div>
<script>
(function() {{
    function autoScale() {{
        const wrap = document.getElementById('grid-wrap');
        const grid = document.getElementById('grid-el');
        if (!wrap || !grid) return;
        grid.style.transform = 'scale(1)';
        wrap.style.height = 'auto';
        const naturalWidth = grid.offsetWidth;
        const naturalHeight = grid.offsetHeight;
        const containerWidth = wrap.clientWidth;
        let scale = 1;
        if (naturalWidth > containerWidth && containerWidth > 0) {{
            scale = containerWidth / naturalWidth;
        }}
        grid.style.transform = 'scale(' + scale + ')';
        wrap.style.height = (naturalHeight * scale) + 'px';
        try {{
            window.parent.postMessage({{
                type: 'streamlit:setFrameHeight',
                height: document.body.scrollHeight + 20
            }}, '*');
        }} catch(e) {{}}
    }}
    autoScale();
    setTimeout(autoScale, 30);
    setTimeout(autoScale, 100);
    setTimeout(autoScale, 300);
    setTimeout(autoScale, 800);
    window.addEventListener('resize', autoScale);
}})();
</script>
</body></html>
"""
    alturas = {3: 950, 4: 1050, 5: 1150, 6: 1200}
    components.html(html, height=alturas.get(tam, 1150), scrolling=False)


# =====================================================
# INTERFACE MOBILE
# =====================================================
def render_mobile(sala_id):
    st.markdown("""
    <style>
        #MainMenu, header, footer { visibility: hidden; }
        .block-container { padding-top: 1rem; padding-bottom: 1rem; max-width: 500px; }
    </style>
    """, unsafe_allow_html=True)

    if not FIREBASE_URL:
        st.error("⚠️ Firebase não configurado.")
        return

    st_autorefresh(interval=1500, key="mobile_poll")
    dados = fb_get(sala_id)

    if not dados:
        st.warning("⏳ Aguardando o jogo iniciar no dispositivo principal...")
        st.caption(f"Sala: **{sala_id}**")
        return

    coord = dados.get("coord")
    estado = dados.get("estado", "aguardando")
    evento = dados.get("evento")
    time_atual = dados.get("time_atual", 1)
    nome_t1 = dados.get("nome_time_1_fb", "Time 1")
    nome_t2 = dados.get("nome_time_2_fb", "Time 2")
    baralho_vazio = dados.get("baralho_vazio", False)

    nome_turno = nome_t1 if time_atual == 1 else nome_t2
    cor_turno = "#1e3c78" if time_atual == 1 else "#dc2626"
    emoji_turno = "🔵" if time_atual == 1 else "🔴"

    st.markdown(
        f"""
        <div style="padding:20px 24px; background:{cor_turno};
                    border-radius:16px; text-align:center;
                    font-family:system-ui; color:white;
                    margin-bottom:20px; box-shadow: 0 6px 24px {cor_turno}55;">
            <div style="font-size:12px; text-transform:uppercase;
                        letter-spacing:3px; opacity:0.85; font-weight:bold;">
                Vez de
            </div>
            <div style="font-size:32px; font-weight:900;
                        margin-top:6px; letter-spacing:1px;">
                {emoji_turno} {nome_turno}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if baralho_vazio and not coord:
        st.success("🏁 Jogo encerrado! Todas as cartas foram sorteadas.")
        return

    if not coord or estado == "aguardando":
        st.markdown(
            f"""
            <div style="text-align:center; padding:40px 20px;
                        background:#f8fafc; border:3px dashed #cbd5e1;
                        border-radius:16px; font-family:system-ui;">
                <div style="font-size:52px;">🎴</div>
                <div style="font-size:16px; color:#64748b;
                            margin-top:16px; font-weight:bold;">
                    Pronto para sortear
                </div>
                <div style="font-size:13px; color:#94a3b8; margin-top:12px;">
                    Toque no botão abaixo para sortear sua carta secreta
                </div>
                <div style="font-size:12px; color:#cbd5e1; margin-top:20px;">
                    🔒 Só {nome_turno} deve estar vendo esta tela
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.write("")
        if st.button("🎴 Sortear carta secreta", type="primary",
                     use_container_width=True, key="btn_sortear_mobile"):
            fb_patch(sala_id, {"solicitar_sorteio": _agora_ms()})
            st.rerun()
        st.caption("💡 Ou aguarde o operador do desktop sortear.")
        return

    if evento:
        st.success("✅ Decisão registrada!")
        st.markdown(
            f"""
            <div style="text-align:center; padding:40px 20px;
                        background:#f1f5f9; border-radius:12px;
                        font-family:system-ui;">
                <div style="font-size:14px;color:#64748b;">{nome_turno} marcou</div>
                <div style="font-size:32px;font-weight:bold;
                            color:{'#22c55e' if evento == 'acertou' else '#dc2626'};
                            margin-top:12px;">
                    {'✅ ACERTOU' if evento == 'acertou' else '❌ ERROU'}
                </div>
                <div style="font-size:13px;color:#94a3b8;margin-top:16px;">
                    🔄 Passe o celular para o próximo time
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        f"""
        <div style="text-align:center; padding:30px 20px;
                    background:#d9e1f2; border:4px solid {cor_turno};
                    border-radius:16px; margin-bottom:20px;
                    font-family:system-ui;">
            <div style="font-size:13px;color:#64748b;
                        text-transform:uppercase;letter-spacing:1.5px;">
                Carta secreta de {nome_turno}
            </div>
            <div style="font-size:120px;font-weight:bold;color:#1e3c78;
                        letter-spacing:8px; line-height:1; margin:16px 0;">
                {coord}
            </div>
            <div style="font-size:12px;color:#94a3b8;">
                Dê uma dica de UMA palavra e aguarde seu time responder
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    timer_ativo = dados.get("timer_ativo", False)
    timer_iniciado_em = dados.get("timer_iniciado_em")
    tempo_total_seg = dados.get("tempo_total", 0)

    if timer_ativo and tempo_total_seg > 0:
        if not timer_iniciado_em:
            st.markdown("---")
            st.caption("⏱️ Inicie o cronômetro quando seu time começar a pensar.")
            if st.button("▶️ Iniciar cronômetro", type="primary", use_container_width=True):
                fb_patch(sala_id, {"timer_iniciado_em": _agora_ms()})
                st.rerun()
        else:
            decorrido_ms = _agora_ms() - timer_iniciado_em
            restante_seg = max(0, int(tempo_total_seg - decorrido_ms / 1000))
            mm, ss = restante_seg // 60, restante_seg % 60
            if restante_seg == 0:
                cor, icone, texto_estado = "#dc2626", "⏰", "Tempo esgotado"
            elif restante_seg <= tempo_total_seg * 0.2:
                cor, icone, texto_estado = "#dc2626", "🔴", "Correndo"
            elif restante_seg <= tempo_total_seg * 0.5:
                cor, icone, texto_estado = "#eab308", "🟡", "Correndo"
            else:
                cor, icone, texto_estado = "#22c55e", "🟢", "Correndo"
            st.markdown(
                f"""
                <div style="padding:14px 18px; background:{cor}15;
                            border:2px solid {cor}; border-radius:10px;
                            text-align:center; font-family:system-ui; margin:14px 0;">
                    <div style="font-size:12px;color:#64748b;
                                text-transform:uppercase;letter-spacing:1px;">
                        {icone} Cronômetro {texto_estado}
                    </div>
                    <div style="font-size:42px;font-weight:bold;color:{cor};
                                font-variant-numeric:tabular-nums;line-height:1; margin-top:6px;">
                        {mm:02d}:{ss:02d}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.caption("🔄 Atualizando a cada 1.5s...")

    st.markdown("---")
    col_ok, col_err = st.columns(2)
    with col_ok:
        if st.button("✅ Acertou", type="primary", use_container_width=True):
            fb_patch(sala_id, {"evento": "acertou"})
            st.rerun()
    with col_err:
        if st.button("❌ Errou", use_container_width=True):
            fb_patch(sala_id, {"evento": "errou"})
            st.rerun()
    st.caption(f"Clique quando {nome_turno} der a resposta.")


# =====================================================
# GERAÇÃO DE IMAGENS PARA O MANUAL PDF
# =====================================================
def _img_to_bytes(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def _desenhar_celula(d, x, y, w, h, texto, cor_fundo, cor_texto="black",
                     fonte=None, borda="black", largura_borda=2, riscado=False):
    d.rectangle([x, y, x + w, y + h], fill=cor_fundo, outline=borda, width=largura_borda)
    if texto:
        f = fonte or _carregar_fonte(16)
        bbox = d.textbbox((0, 0), texto, font=f)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        tx = x + (w - tw) / 2 - bbox[0]
        ty = y + (h - th) / 2 - bbox[1]
        d.text((tx, ty), texto, fill=cor_texto, font=f)
        if riscado:
            d.line([tx, ty + th / 2, tx + tw, ty + th / 2], fill=cor_texto, width=2)


def gerar_imagem_capa():
    w, h = 1200, 500
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, h], fill="#1e3c78")
    d.rectangle([150, 200, 290, 340], fill="#ffffff", outline="#0f2557", width=4)
    for cx, cy in [(185, 235), (255, 235), (185, 305), (255, 305), (220, 270)]:
        d.ellipse([cx - 8, cy - 8, cx + 8, cy + 8], fill="#1e3c78")
    d.rectangle([310, 200, 450, 340], fill="#ffffff", outline="#0f2557", width=4)
    for cx, cy in [(345, 235), (415, 235), (380, 270), (345, 305), (415, 305)]:
        d.ellipse([cx - 8, cy - 8, cx + 8, cy + 8], fill="#1e3c78")
    f_titulo = _carregar_fonte(72)
    f_sub = _carregar_fonte(24)
    d.text((500, 220), "ENTRE", fill="white", font=f_titulo)
    d.text((500, 290), "LINHAS", fill="white", font=f_titulo)
    d.text((500, 380), "Manual Completo do Jogo", fill="#c5e0b4", font=f_sub)
    return _img_to_bytes(img)


def gerar_imagem_grade_exemplo():
    tam = 3
    cell, coord_w, cab_h = 110, 90, 45
    w = coord_w * 2 + cell * tam + 40
    h = cab_h * 2 + cell * tam + 80
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)

    f_pal = _carregar_fonte(14)
    f_coord = _carregar_fonte(24)
    f_cab = _carregar_fonte(18)

    linhas = ["cachorro", "praia", "música"]
    colunas = ["verão", "escola", "café"]

    ox = 20
    oy = cab_h + 30

    for j in range(tam):
        x = ox + coord_w + j * cell
        _desenhar_celula(d, x, oy - cab_h - cab_h, cell, cab_h, colunas[j],
                         "#c5e0b4", "black", f_pal, largura_borda=2)

    for j in range(tam):
        x = ox + coord_w + j * cell
        _desenhar_celula(d, x, oy - cab_h, cell, cab_h, chr(65 + j),
                         "#f0f0f0", "black", f_cab, largura_borda=2)

    for i in range(tam):
        y = oy + i * cell
        _desenhar_celula(d, ox, y, coord_w, cell, linhas[i],
                         "#ffe699", "black", f_pal, largura_borda=2)

    _desenhar_celula(d, ox + coord_w, oy - cab_h, coord_w, cab_h, "×",
                     "#e6e6e6", "black", f_cab, largura_borda=2)
    _desenhar_celula(d, ox, oy - cab_h, coord_w, cab_h, "",
                     "#e6e6e6", "black", f_cab, largura_borda=2)

    for i in range(tam):
        for j in range(tam):
            x = ox + coord_w + j * cell
            y = oy + i * cell
            coord = f"{chr(65 + j)}{i + 1}"
            cor_bg = "#d9e1f2"
            cor_txt = "#1e3c78"
            if coord == "B2":
                cor_bg = "#1e3c78"
                cor_txt = "white"
            _desenhar_celula(d, x, y, cell, cell, coord, cor_bg, cor_txt,
                             f_coord, largura_borda=2)

    f_leg = _carregar_fonte(16)
    d.text((30, oy + tam * cell + 15),
           "Carta B2 = praia (linha 2) × escola (coluna B)",
           fill="#1e3c78", font=f_leg)
    return _img_to_bytes(img)


def gerar_imagem_estados():
    w, h = 1000, 260
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    f_coord = _carregar_fonte(56)
    f_leg = _carregar_fonte(18)
    cartas = [
        ("A1", "#d9e1f2", "#1e3c78", "Neutra"),
        ("B1", "#1e3c78", "#ffffff", "Time 1"),
        ("C1", "#dc2626", "#ffffff", "Time 2"),
        ("D1", "#cbd5e1", "#64748b", "Descartada"),
    ]
    cw, ch = 200, 140
    gap = 25
    total_w = 4 * cw + 3 * gap
    ox = (w - total_w) // 2
    oy = 40
    for idx, (coord, bg, txt, leg) in enumerate(cartas):
        x = ox + idx * (cw + gap)
        _desenhar_celula(d, x, oy, cw, ch, coord, bg, txt, f_coord,
                         largura_borda=3, riscado=(leg == "Descartada"))
        bbox = d.textbbox((0, 0), leg, font=f_leg)
        tw = bbox[2] - bbox[0]
        d.text((x + (cw - tw) / 2, oy + ch + 20), leg, fill="#334155", font=f_leg)
    return _img_to_bytes(img)


def gerar_imagem_cronometro():
    w, h = 1000, 300
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    f_tempo = _carregar_fonte(70)
    f_leg = _carregar_fonte(18)
    tempos = [
        ("01:30", "#22c55e", "#dcfce7", "> 50% do tempo"),
        ("00:45", "#eab308", "#fef9c3", "20% a 50%"),
        ("00:15", "#dc2626", "#fee2e2", "< 20%"),
    ]
    cw, ch = 280, 180
    gap = 30
    total_w = 3 * cw + 2 * gap
    ox = (w - total_w) // 2
    oy = 40
    for idx, (tempo, cor, bg, leg) in enumerate(tempos):
        x = ox + idx * (cw + gap)
        d.rectangle([x, oy, x + cw, oy + ch], fill=bg, outline=cor, width=4)
        bbox = d.textbbox((0, 0), tempo, font=f_tempo)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        d.text((x + (cw - tw) / 2 - bbox[0], oy + (ch - th) / 2 - bbox[1] - 10),
               tempo, fill=cor, font=f_tempo)
        bbox = d.textbbox((0, 0), leg, font=f_leg)
        tw = bbox[2] - bbox[0]
        d.text((x + (cw - tw) / 2, oy + ch + 15), leg, fill="#334155", font=f_leg)
    return _img_to_bytes(img)


def gerar_imagem_fluxo_rodada():
    w, h = 1100, 320
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    f_num = _carregar_fonte(30)
    f_txt = _carregar_fonte(14)
    f_seta = _carregar_fonte(30)
    passos = [
        ("1", "Tirador\nescaneia QR"),
        ("2", "Sorteia\ncarta secreta"),
        ("3", "Vê a carta\n(ex: B2)"),
        ("4", "Dá uma\ndica"),
        ("5", "Time\ntenta acertar"),
    ]
    cw, ch = 170, 150
    gap = 40
    total_w = 5 * cw + 4 * gap
    ox = (w - total_w) // 2
    oy = 80
    for idx, (num, texto) in enumerate(passos):
        x = ox + idx * (cw + gap)
        d.ellipse([x + cw / 2 - 25, oy - 45, x + cw / 2 + 25, oy + 5], fill="#1e3c78")
        bbox = d.textbbox((0, 0), num, font=f_num)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        d.text((x + cw / 2 - tw / 2 - bbox[0], oy - 45 + (50 - th) / 2 - bbox[1]),
               num, fill="white", font=f_num)
        d.rectangle([x, oy + 15, x + cw, oy + ch + 15],
                    fill="#f8fafc", outline="#1e3c78", width=3)
        for li, linha in enumerate(texto.split("\n")):
            bbox = d.textbbox((0, 0), linha, font=f_txt)
            tw = bbox[2] - bbox[0]
            d.text((x + (cw - tw) / 2, oy + 45 + li * 25),
                   linha, fill="#1e3c78", font=f_txt)
        if idx < len(passos) - 1:
            d.text((x + cw + 5, oy + ch / 2 - 5), "→", fill="#64748b", font=f_seta)
    return _img_to_bytes(img)


# =====================================================
# MANUAL — TEXTO (para o modal)
# =====================================================
MANUAL_TEXTO = """
# ENTRE LINHAS — MANUAL DO JOGO

## 1. SOBRE O JOGO

Entre Linhas é um jogo de dedução e criatividade para 2 times. Cada time tenta
conquistar cartas do tabuleiro dando dicas de UMA ÚNICA PALAVRA. Quem conquistar
mais cartas até o baralho esvaziar vence a partida.

Baseado no jogo de tabuleiro "Entre Linhas" da PaperGames, esta versão é uma
adaptação digital para jogar em família ou com amigos.

## 2. COMPONENTES

- 1 dispositivo principal (notebook, tablet ou TV espelhada) com o app aberto
- 1 celular opcional para o tirador da vez
- Tabuleiro digital com grade 3x3, 4x4, 5x5 ou 6x6
- 343 palavras na biblioteca (podem ser editadas)
- Cronômetro de turno opcional

## 3. COMO O TABULEIRO FUNCIONA

A grade é composta por linhas e colunas:
- As LINHAS são numeradas de 1 a N
- As COLUNAS têm letras de A a N

Cada célula é uma CARTA identificada por uma coordenada (ex.: A1, B2, C3).
Cada célula representa o cruzamento de uma palavra-linha com uma palavra-coluna.

Exemplo prático:
- Linha 2 = praia
- Coluna B = escola
- Carta B2 = praia x escola

## 4. PREPARAÇÃO

1. Abra o app em um dispositivo principal
2. Configure o tamanho da grade (3x3, 4x4, 5x5 ou 6x6)
3. Defina os nomes dos times
4. Opcional: ative o cronômetro
5. Clique em "Sortear grade"
6. Um QR code aparece na barra lateral

## 5. COMO JOGAR (FLUXO DE UMA RODADA)

### Fase 1 — Sortear a carta
- O time da vez escolhe um integrante para ser o TIRADOR
- Ele escaneia o QR code com o celular
- O time adversário DEVE olhar para o outro lado
- O tirador clica em "Sortear carta secreta"

### Fase 2 — Dar a dica
- O tirador vê a carta secreta (ex.: B2) no celular
- Consulta o tabuleiro para descobrir as duas palavras do cruzamento
- Dá uma dica de UMA ÚNICA PALAVRA para o seu próprio time

REGRA DA DICA ÚNICA:
- A palavra NÃO pode conter o RADICAL de nenhuma das duas palavras da carta
- Exemplo: se a carta é B2 (praia × escola), não vale dizer "praiano"
  (radical de praia) nem "escolar" (radical de escola)
- A palavra pode ser COMPOSTA com hífen, desde que seja válida em português
  (ex.: guarda-chuva, beija-flor, arco-íris)
- Mesmo com hífen, a regra do radical se aplica em cada parte

### Fase 3 — Adivinhar
- O time do tirador discute entre si para acertar a coordenada
- Podem falar à vontade entre eles
- Apenas UMA resposta final é dada
- O tirador NÃO pode participar da discussão após dar a dica
- O tirador NÃO pode fazer gestos, caras ou sons que indiquem certo/errado
- O tirador NÃO pode responder perguntas como "é essa?" ou "tá quente?"
- Qualquer sinal do tirador é considerado trapaça

### Fase 4 — Registrar o resultado
- ACERTOU: a carta passa a pertencer ao time
- ERROU: a carta é descartada do jogo

### Fase 5 — Passar o turno
- O turno passa automaticamente para o outro time
- Passe o celular para o próximo tirador

## 6. COMO OS JOGADORES DEVEM AGIR

### Escolha do tirador
- Cada rodada tem um NOVO tirador, escolhido pelo time
- Recomendado: revezar entre todos os integrantes

### Disciplina com o celular
- O celular fica com o TIRADOR da vez
- Ao passar o turno, entregue o celular para o próximo tirador
- Nunca olhe a tela do celular do outro time

### Durante a fase de sorteio
- O time adversário DEVE virar-se ou fechar os olhos
- Aguarde até que todos confirmem que não estão olhando

### Regra da dica única
- O tirador informa APENAS UMA PALAVRA em voz alta
- Essa palavra NÃO pode conter o RADICAL de nenhuma das palavras da carta
- A palavra pode ser composta (com hífen) desde que seja uma palavra válida
- O tirador deve EVITAR INDUZIR o time quando sugestões são discutidas

### O que é radical
Radical é a parte da palavra que carrega o significado principal.
Palavras da mesma família compartilham o radical. Exemplos:
- cachorro → radical cachorr- (proibido "cachorrada", "cachorrinho")
- praia → radical prai- (proibido "praiano", "praiana")
- escola → radical escol- (proibido "escolar", "escolinha")
- música → radical music- (proibido "musical", "musicista")

Exemplo prático: se a carta é B2 (praia × escola):
- Dica VÁLIDA: "férias" (não compartilha radical)
- Dica INVÁLIDA: "praiano" (contém radical de praia)
- Dica INVÁLIDA: "escolar" (contém radical de escola)

### Palavras compostas com hífen
Permitido usar palavras compostas com hífen, desde que sejam válidas:
- guarda-chuva, beija-flor, segunda-feira, arco-íris, couve-flor
Cuidado: a regra do radical se aplica em cada parte da palavra.

### Como se portar durante a discussão do time
- Após dar a dica, o tirador deve permanecer em SILÊNCIO
- O time discute livremente para tentar acertar
- O tirador NÃO pode fazer caras, gestos ou sons
- O tirador NÃO pode responder "é essa?" ou "tá quente?"
- O tirador NÃO pode repetir a dica com entonação diferente
- O tirador NÃO pode reforçar a dica com sinônimos
- Qualquer sinal que influencie o time é considerado TRAPAÇA

### Durante a decisão
- O tirador é o único que pode clicar em "Acertou" ou "Errou"
- A decisão é definitiva — não há segunda chance

### Após o registro
- Ao acertar: a carta fica com o time, e o turno passa
- Ao errar: a carta sai do jogo, e o turno passa
- Aguardar o app atualizar antes da próxima rodada

### Ao passar o turno
- O tirador entrega o celular para o próximo time
- Aguarda-se que o app mostre "Vez de [Nome do Time]"

### Comunicação entre times
- Permitido: comemorações discretas, gestos de incentivo
- Proibido: revelar a dica do time adversário
- Proibido: dar pistas sobre a carta secreta do outro time

### Em caso de dúvida ou impasse
- Se a dica for ambígua, o time do tirador decide se aceita
- Se houver suspeita de trapaça, o grupo vota (maioria decide)
- Se o app apresentar erro, use o botão de emergência do desktop

## 7. PONTUAÇÃO

- Cada carta conquistada vale 1 ponto para o time
- Cartas descartadas não valem pontos
- Vence quem tiver mais cartas conquistadas

## 8. CRONÔMETRO (OPCIONAL)

- Cada turno tem um tempo limite configurável
- Pode ser iniciado pelo celular do tirador ou pelo desktop
- Cores indicam o tempo restante:
  - Verde: mais de 50% do tempo
  - Amarelo: entre 20% e 50%
  - Vermelho: menos de 20%
- Bipes tocam nos últimos 10 segundos

## 9. CELULAR DO TIRADOR

O celular mostra sempre de quem é a vez, em um banner colorido:
- Azul: vez do Time 1
- Vermelho: vez do Time 2

Funcionalidades:
- Sortear carta secreta
- Ver a carta secreta (apenas quando é a sua vez)
- Iniciar o cronômetro
- Registrar acerto ou erro

## 10. FIM DE JOGO

Quando o baralho esvaziar:
- Se um time conquistou mais cartas: "Nome do Time VENCEU!"
- Se houve empate: "EMPATE!"

Opções ao final:
- Nova partida (mesma grade): reembaralha as cartas
- Nova grade: sorteia palavras novas

## 11. RECUPERAÇÃO DE SALA

Se o navegador fechar ou travar:
- A URL contém o código da sala (ex.: ?sala=ABC12345)
- Ao reabrir pela mesma URL, o jogo é recuperado automaticamente

Recuperação manual:
1. Abra o app normalmente
2. Na barra lateral, procure "Recuperar Sala"
3. Cole o código da sala
4. Clique em "Recuperar"

## 12. DICAS PARA BOAS PARTIDAS

- Escolha dicas que combinem as duas palavras sem entregar demais
- Evite dicas óbvias que o adversário também entenderia
- Use o cronômetro para dar ritmo ao jogo
- Comece com grades menores (3x3) para aprender
- Varie os tiradores para manter todos participando

## 13. PERGUNTAS FREQUENTES

P: Posso jogar sem o celular?
R: Sim. O operador do desktop sorteia e mostra para o time da vez.

P: O que acontece se a dica for muito ambígua?
R: O próprio time avalia. Se errar, a carta é descartada.

P: O cronômetro é obrigatório?
R: Não. É opcional.

P: Como funciona a alternância de turnos?
R: Automaticamente. Após cada resposta, o próximo time pega o celular.

P: Posso jogar em dois celulares diferentes?
R: Não. Use apenas UM celular.

## 14. ATALHOS E BOTÕES ÚTEIS

No desktop:
- Sortear grade: cria uma nova partida
- Zerar partida: apaga a sala atual
- Recuperar Sala: volta pelo código
- Sortear carta secreta: ativa o tirador sem usar o celular
- Concluir turno (manual): em caso de emergência

No celular (tirador):
- Sortear carta secreta
- Iniciar cronômetro
- Acertou / Errou

## 15. SOBRE O APP

- Baseado no jogo "Entre Linhas" da PaperGames
- Biblioteca interna com 343 palavras em português
- Compatível com desktop, tablet e celular
- Otimizado para TV espelhada em modo tela cheia (F11)

Bom jogo!
"""


# =====================================================
# GERAÇÃO DO MANUAL EM PDF
# =====================================================
def gerar_manual_pdf():
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image as RLImage
    )
    from reportlab.lib.colors import HexColor

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=2 * cm, leftMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
        title="Manual Entre Linhas", author="Entre Linhas Jogo",
    )

    styles = getSampleStyleSheet()
    titulo_principal = ParagraphStyle("TituloPrincipal", parent=styles["Heading1"],
        fontSize=26, textColor=HexColor("#1e3c78"), spaceAfter=20,
        alignment=TA_CENTER, leading=30)
    subtitulo_home = ParagraphStyle("SubtituloHome", parent=styles["Normal"],
        fontSize=11, textColor=HexColor("#64748b"), spaceAfter=30,
        alignment=TA_CENTER)
    titulo_secao = ParagraphStyle("TituloSecao", parent=styles["Heading2"],
        fontSize=16, textColor=HexColor("#1e3c78"), spaceBefore=18,
        spaceAfter=10, leading=20)
    titulo_sub = ParagraphStyle("TituloSub", parent=styles["Heading3"],
        fontSize=13, textColor=HexColor("#334155"), spaceBefore=12, spaceAfter=6)
    corpo = ParagraphStyle("Corpo", parent=styles["BodyText"],
        fontSize=10.5, leading=15, alignment=TA_JUSTIFY, spaceAfter=6)
    item_lista = ParagraphStyle("ItemLista", parent=corpo,
        leftIndent=15, bulletIndent=5, spaceAfter=3)
    destaque = ParagraphStyle("Destaque", parent=corpo,
        backColor=HexColor("#f1f5f9"), borderColor=HexColor("#cbd5e1"),
        borderWidth=1, borderPadding=8, spaceBefore=8, spaceAfter=12)
    destaque_alerta = ParagraphStyle("DestaqueAlerta", parent=corpo,
        backColor=HexColor("#fef2f2"), borderColor=HexColor("#dc2626"),
        borderWidth=1, borderPadding=8, spaceBefore=8, spaceAfter=12)
    legenda_img = ParagraphStyle("LegendaImg", parent=styles["Normal"],
        fontSize=9, textColor=HexColor("#64748b"), alignment=TA_CENTER,
        spaceBefore=4, spaceAfter=16)
    rodape = ParagraphStyle("Rodape", parent=styles["Normal"],
        fontSize=9, textColor=HexColor("#64748b"), alignment=TA_CENTER,
        spaceBefore=30)

    story = []

    # ---- CAPA ----
    story.append(Spacer(1, 1 * cm))
    story.append(RLImage(gerar_imagem_capa(), width=17 * cm, height=7.08 * cm))
    story.append(Spacer(1, 1.5 * cm))
    story.append(Paragraph("Um jogo de dedução e criatividade para 2 times",
        ParagraphStyle("cap", parent=corpo, alignment=TA_CENTER,
                       fontSize=13, textColor=HexColor("#64748b"))))
    story.append(Spacer(1, 3 * cm))
    story.append(Paragraph(
        "Versão 1.0 — Adaptação digital do jogo de tabuleiro da PaperGames",
        rodape))
    story.append(PageBreak())

    # ---- SUMÁRIO ----
    story.append(Paragraph("Sumário", titulo_secao))
    for item in [
        "1. Sobre o jogo", "2. Componentes", "3. Como o tabuleiro funciona",
        "4. Preparação", "5. Como jogar (fluxo de uma rodada)",
        "6. Como os jogadores devem agir", "7. Pontuação",
        "8. Estados das cartas", "9. Cronômetro (opcional)",
        "10. Celular do tirador", "11. Fim de jogo",
        "12. Recuperação de sala", "13. Dicas para boas partidas",
        "14. Perguntas frequentes", "15. Atalhos e botões úteis",
        "16. Sobre o app",
    ]:
        story.append(Paragraph(f"• {item}", item_lista))
    story.append(PageBreak())

    # ---- 1. SOBRE ----
    story.append(Paragraph("1. Sobre o jogo", titulo_secao))
    story.append(Paragraph(
        "Entre Linhas é um jogo de dedução e criatividade para 2 times. "
        "Cada time tenta conquistar cartas do tabuleiro dando dicas de "
        "<b>uma única palavra</b>. Quem conquistar mais cartas até o baralho "
        "esvaziar vence a partida.", corpo))
    story.append(Paragraph(
        "Baseado no jogo de tabuleiro <i>Entre Linhas</i> da PaperGames, esta "
        "versão é uma adaptação digital para jogar em família ou com amigos.",
        corpo))

    # ---- 2. COMPONENTES ----
    story.append(Paragraph("2. Componentes", titulo_secao))
    for c in [
        "1 dispositivo principal (notebook, tablet ou TV espelhada) com o app aberto",
        "1 celular opcional para o tirador da vez",
        "Tabuleiro digital com grade 3×3, 4×4, 5×5 ou 6×6",
        "343 palavras na biblioteca (podem ser editadas)",
        "Cronômetro de turno opcional",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- 3. TABULEIRO ----
    story.append(PageBreak())
    story.append(Paragraph("3. Como o tabuleiro funciona", titulo_secao))
    story.append(Paragraph(
        "A grade é composta por linhas e colunas. As linhas são numeradas de "
        "1 a N e as colunas têm letras de A a N.", corpo))
    story.append(Spacer(1, 0.3 * cm))
    story.append(RLImage(gerar_imagem_grade_exemplo(), width=12 * cm, height=10 * cm))
    story.append(Paragraph(
        "Exemplo de grade 3×3. A carta B2 é o cruzamento de <b>praia</b> "
        "(linha 2) com <b>escola</b> (coluna B).", legenda_img))
    story.append(Paragraph(
        "O tirador da vez recebe uma coordenada secreta e deve dar uma dica "
        "que remeta ao cruzamento daquelas duas palavras.", corpo))

    # ---- 4. PREPARAÇÃO ----
    story.append(Paragraph("4. Preparação", titulo_secao))
    for i, c in enumerate([
        "Abra o app em um dispositivo principal",
        "Configure o tamanho da grade: 3×3, 4×4, 5×5 ou 6×6",
        "Defina os nomes dos times",
        "Opcional: ative o cronômetro e escolha o tempo por turno",
        "Clique em \"Sortear grade\"",
        "Um QR code aparece na barra lateral",
    ], 1):
        story.append(Paragraph(f"{i}. {c}", item_lista))

    # ---- 5. COMO JOGAR ----
    story.append(PageBreak())
    story.append(Paragraph("5. Como jogar (fluxo de uma rodada)", titulo_secao))
    story.append(Paragraph("Cada rodada tem 5 fases bem definidas:", corpo))
    story.append(Spacer(1, 0.2 * cm))
    story.append(RLImage(gerar_imagem_fluxo_rodada(), width=17 * cm, height=4.95 * cm))
    story.append(Paragraph("Fluxo visual de uma rodada completa", legenda_img))

    story.append(Paragraph("Fase 1 — Sortear a carta", titulo_sub))
    for c in [
        "O time da vez escolhe um integrante para ser o <b>tirador</b>",
        "Ele escaneia o QR code com o celular",
        "O time adversário deve olhar para o outro lado ou fechar os olhos",
        "O tirador clica em \"Sortear carta secreta\"",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Fase 2 — Dar a dica", titulo_sub))
    for c in [
        "O tirador vê a carta secreta (ex.: B2) no celular",
        "Consulta o tabuleiro para descobrir as duas palavras do cruzamento",
        "Dá uma dica de <b>uma única palavra</b> para o seu próprio time",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Fase 3 — Adivinhar", titulo_sub))
    for c in [
        "O time do tirador discute entre si para tentar acertar a coordenada",
        "Apenas UMA resposta final é dada",
        "O tirador NÃO pode participar da discussão após dar a dica",
        "O tirador NÃO pode fazer gestos, caras ou sons",
        "O tirador NÃO pode responder perguntas como \"é essa?\"",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Fase 4 — Registrar o resultado", titulo_sub))
    for c in [
        "<b>Acertou</b>: a carta passa a pertencer ao time",
        "<b>Errou</b>: a carta é descartada do jogo",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Fase 5 — Passar o turno", titulo_sub))
    for c in [
        "O turno passa automaticamente para o outro time",
        "Passe o celular para o próximo tirador",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # =====================================================
    # 6. COMO OS JOGADORES DEVEM AGIR
    # =====================================================
    story.append(PageBreak())
    story.append(Paragraph("6. Como os jogadores devem agir", titulo_secao))

    story.append(Paragraph("Escolha do tirador", titulo_sub))
    for c in [
        "Cada rodada tem um <b>novo tirador</b>, escolhido pelo próprio time",
        "Recomendado: revezar entre todos os integrantes",
        "Em times com 2+ jogadores: o tirador muda a cada rodada",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Disciplina com o celular", titulo_sub))
    for c in [
        "O celular fica com o <b>tirador da vez</b>, nunca com o time adversário",
        "Ao passar o turno, entregue o celular para o próximo tirador",
        "Nunca olhe a tela do celular do outro time",
        "Se a tela mostrar a carta secreta por acidente, avise o grupo",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Durante a fase de sorteio", titulo_sub))
    for c in [
        "O time adversário <b>deve virar-se ou fechar os olhos</b>",
        "Aguarde até que todos confirmem que não estão olhando",
        "Só então o tirador sorteia a carta",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- REGRA DA DICA ÚNICA (NOVA) ----
    story.append(Paragraph("Regra da dica única", titulo_sub))
    for c in [
        "O tirador deve informar <b>apenas UMA PALAVRA</b> em voz alta",
        "Essa palavra <b>não pode conter o radical</b> de nenhuma das palavras "
        "das coordenadas sorteadas",
        "A palavra pode ser <b>composta com hífen</b> desde que seja uma "
        "palavra válida em português",
        "O tirador deve <b>evitar induzir</b> o time quando sugestões estão "
        "sendo discutidas entre os participantes",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("O que é radical", titulo_sub))
    story.append(Paragraph(
        "Radical é a parte da palavra que carrega o significado principal. "
        "Palavras da mesma família compartilham o radical. Exemplos:", corpo))
    for c in [
        "<b>cachorro</b> → radical <i>cachorr-</i> → proibido usar "
        "\"cachorrada\", \"cachorrinho\"",
        "<b>praia</b> → radical <i>prai-</i> → proibido usar \"praiano\"",
        "<b>escola</b> → radical <i>escol-</i> → proibido usar \"escolar\"",
        "<b>música</b> → radical <i>music-</i> → proibido usar \"musical\"",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph(
        "<b>Exemplo prático:</b> se a carta é B2 e as palavras são <i>praia</i> "
        "(linha 2) e <i>escola</i> (coluna B):<br/>"
        "• Dica válida: <i>\"férias\"</i> (não compartilha radical)<br/>"
        "• Dica inválida: <i>\"praiano\"</i> (contém radical de \"praia\")<br/>"
        "• Dica inválida: <i>\"escolar\"</i> (contém radical de \"escola\")",
        destaque))

    story.append(Paragraph("Palavras compostas com hífen", titulo_sub))
    story.append(Paragraph(
        "É permitido usar palavras compostas com hífen, desde que sejam "
        "válidas em português. Exemplos:", corpo))
    for c in [
        "<b>guarda-chuva</b> — palavra composta válida",
        "<b>beija-flor</b> — palavra composta válida",
        "<b>segunda-feira</b> — palavra composta válida",
        "<b>arco-íris</b> — palavra composta válida",
        "<b>couve-flor</b> — palavra composta válida",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))
    story.append(Paragraph(
        "Cuidado: mesmo com hífen, a regra do radical se aplica em cada "
        "parte da palavra.", corpo))

    # ---- COMO SE PORTAR NA DISCUSSÃO (NOVA) ----
    story.append(Paragraph("Como se portar durante a discussão do time", titulo_sub))
    for c in [
        "Após dar a dica, o tirador <b>deve permanecer em silêncio</b>",
        "O time discute livremente para tentar acertar a coordenada",
        "O tirador <b>não pode</b> fazer caras, gestos ou sons que indiquem "
        "certo/errado",
        "O tirador <b>não pode</b> responder perguntas como \"é essa?\" ou "
        "\"tá quente?\"",
        "O tirador <b>não pode</b> repetir a dica com entonação diferente",
        "O tirador <b>não pode</b> reforçar a dica com sinônimos ou frases",
        "O tirador deve apenas ouvir a discussão até o time bater o martelo",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph(
        "⚠️ <b>Importante:</b> qualquer sinal do tirador que influencie a "
        "resposta do time é considerado <b>trapaça</b>. Nesse caso, o time "
        "adversário pode contestar e o grupo decide por votação.",
        destaque_alerta))

    story.append(Paragraph("Durante a decisão", titulo_sub))
    for c in [
        "O tirador é o único que pode clicar em \"Acertou\" ou \"Errou\"",
        "A decisão é definitiva — não há segunda chance na mesma rodada",
        "Se houve discordância interna, o tirador decide",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Após o registro", titulo_sub))
    for c in [
        "Ao acertar: a carta fica com o time, e o turno passa",
        "Ao errar: a carta sai do jogo, e o turno passa",
        "Recomendado: aguardar o app atualizar antes da próxima rodada",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Ao passar o turno", titulo_sub))
    for c in [
        "O tirador entrega o celular para o próximo time",
        "O próximo tirador pega o dispositivo",
        "Aguarda-se que o app mostre \"Vez de [Nome do Time]\"",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Comunicação entre times", titulo_sub))
    for c in [
        "<b>Permitido</b>: comemorações discretas, gestos de incentivo",
        "<b>Proibido</b>: revelar a dica do time adversário",
        "<b>Proibido</b>: dar pistas sobre a carta secreta do outro time",
        "<b>Proibido</b>: olhar a tela do celular do outro time",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Em caso de dúvida ou impasse", titulo_sub))
    for c in [
        "Se a dica for ambígua, o <b>time do tirador</b> decide se aceita",
        "Se houver suspeita de trapaça, o grupo vota (maioria decide)",
        "Se o app apresentar erro, use o botão de emergência do desktop",
        "Se o cronômetro travar, clique em \"Reiniciar\" no desktop",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Paragraph("Ao final da partida", titulo_sub))
    for c in [
        "Confirme o resultado no banner de vitória",
        "Aguarde todos verem o placar antes de reiniciar",
        "Escolha em conjunto: \"Nova partida\" ou \"Nova grade\"",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- 7. PONTUAÇÃO ----
    story.append(Paragraph("7. Pontuação", titulo_secao))
    for c in [
        "Cada carta conquistada vale 1 ponto para o time",
        "Cartas descartadas não valem pontos",
        "Ao final do baralho, vence quem tiver mais cartas conquistadas",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- 8. ESTADOS ----
    story.append(PageBreak())
    story.append(Paragraph("8. Estados das cartas", titulo_secao))
    story.append(Paragraph(
        "Durante o jogo, cada carta pode assumir 4 estados visuais diferentes:",
        corpo))
    story.append(Spacer(1, 0.3 * cm))
    story.append(RLImage(gerar_imagem_estados(), width=16 * cm, height=4.16 * cm))
    story.append(Paragraph("Os 4 estados possíveis das cartas no tabuleiro",
        legenda_img))
    for c in [
        "<b>Neutra (azul claro)</b>: carta ainda não sorteada",
        "<b>Time 1 (azul escuro)</b>: o Time 1 conquistou a carta",
        "<b>Time 2 (vermelho)</b>: o Time 2 conquistou a carta",
        "<b>Descartada (cinza riscada)</b>: alguém errou, carta fora do jogo",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- 9. CRONÔMETRO ----
    story.append(Paragraph("9. Cronômetro (opcional)", titulo_secao))
    story.append(Paragraph(
        "Cada turno tem um tempo limite configurável (ex.: 1 minuto e 30 "
        "segundos). Pode ser iniciado pelo celular do tirador ou pelo desktop.",
        corpo))
    story.append(Spacer(1, 0.3 * cm))
    story.append(RLImage(gerar_imagem_cronometro(), width=16 * cm, height=4.8 * cm))
    story.append(Paragraph("As cores do cronômetro indicam o tempo restante",
        legenda_img))
    for c in [
        "<b>Verde</b>: mais de 50% do tempo",
        "<b>Amarelo</b>: entre 20% e 50%",
        "<b>Vermelho</b>: menos de 20%",
        "Bipes tocam nos últimos 10 segundos (médio, agudo, alarme final)",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- 10. CELULAR ----
    story.append(Paragraph("10. Celular do tirador", titulo_secao))
    story.append(Paragraph(
        "O celular mostra sempre de quem é a vez, em um banner colorido:",
        corpo))
    for c in [
        "<b>Azul</b>: vez do Time 1",
        "<b>Vermelho</b>: vez do Time 2",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))
    story.append(Paragraph("Funcionalidades do celular:", corpo))
    for c in [
        "Sortear carta secreta",
        "Ver a carta secreta (apenas quando é a sua vez)",
        "Iniciar o cronômetro",
        "Registrar acerto ou erro",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- 11. FIM DE JOGO ----
    story.append(Paragraph("11. Fim de jogo", titulo_secao))
    for c in [
        "Se um time conquistou mais cartas: <b>\"Nome do Time VENCEU!\"</b>",
        "Se houve empate: <b>\"EMPATE!\"</b>",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))
    story.append(Paragraph("Opções ao final:", corpo))
    for c in [
        "<b>Nova partida (mesma grade)</b>: reembaralha as cartas",
        "<b>Nova grade</b>: sorteia palavras novas",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- 12. RECUPERAÇÃO ----
    story.append(PageBreak())
    story.append(Paragraph("12. Recuperação de sala", titulo_secao))
    for c in [
        "A URL contém o código da sala (ex.: ?sala=ABC12345)",
        "Ao reabrir o app pela mesma URL, o jogo é recuperado automaticamente",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))
    story.append(Paragraph("Também é possível recuperar manualmente:", corpo))
    for i, c in enumerate([
        "Abra o app normalmente",
        "Na barra lateral, procure \"Recuperar Sala\"",
        "Cole o código da sala",
        "Clique em \"Recuperar\"",
    ], 1):
        story.append(Paragraph(f"{i}. {c}", item_lista))

    # ---- 13. DICAS ----
    story.append(Paragraph("13. Dicas para boas partidas", titulo_secao))
    for c in [
        "Escolha dicas que combinem as duas palavras sem entregar demais",
        "Evite dicas óbvias que o adversário também entenderia",
        "Use o cronômetro para dar ritmo ao jogo",
        "Comece com grades menores (3×3) para aprender",
        "Varie os tiradores para manter todos participando",
        "Nunca esconda nem mostre a carta secreta para o outro time",
        "Respeite as decisões do tirador — ele é o único que viu a carta",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- 14. FAQ ----
    story.append(Paragraph("14. Perguntas frequentes", titulo_secao))
    for p, r in [
        ("Posso jogar sem o celular?",
         "Sim. O operador do desktop sorteia e mostra para o time da vez."),
        ("O que acontece se a dica for muito ambígua?",
         "O próprio time avalia. Se errar, a carta é descartada."),
        ("O cronômetro é obrigatório?",
         "Não. É opcional e pode ser desativado."),
        ("Como funciona a alternância de turnos?",
         "Automaticamente. Após cada resposta, o próximo time pega o celular."),
        ("Posso jogar em dois celulares diferentes?",
         "Não. Use apenas UM celular."),
    ]:
        story.append(Paragraph(f"<b>P: {p}</b>", corpo))
        story.append(Paragraph(f"R: {r}", corpo))
        story.append(Spacer(1, 4))

    # ---- 15. ATALHOS ----
    story.append(Paragraph("15. Atalhos e botões úteis", titulo_secao))
    story.append(Paragraph("<b>No desktop:</b>", corpo))
    for c in [
        "Sortear grade: cria uma nova partida",
        "Zerar partida: apaga a sala atual",
        "Recuperar Sala: volta pelo código",
        "Sortear carta secreta: ativa o tirador sem usar o celular",
        "Concluir turno (manual): em caso de emergência",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))
    story.append(Paragraph("<b>No celular (tirador):</b>", corpo))
    for c in ["Sortear carta secreta", "Iniciar cronômetro", "Acertou / Errou"]:
        story.append(Paragraph(f"• {c}", item_lista))

    # ---- 16. SOBRE ----
    story.append(Paragraph("16. Sobre o app", titulo_secao))
    for c in [
        "Baseado no jogo \"Entre Linhas\" da PaperGames",
        "Biblioteca interna com 343 palavras em português",
        "Compatível com desktop, tablet e celular",
        "Otimizado para TV espelhada em modo tela cheia (F11)",
    ]:
        story.append(Paragraph(f"• {c}", item_lista))

    story.append(Spacer(1, 1 * cm))
    story.append(Paragraph("Bom jogo! 🎲", ParagraphStyle(
        "BomJogo", parent=titulo_principal, fontSize=20, alignment=TA_CENTER)))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# =====================================================
# MODAL DO MANUAL
# =====================================================
@st.dialog("📖 Manual do Jogo", width="large")
def modal_manual():
    st.markdown(
        "<div style='text-align:center; font-family:system-ui; margin-bottom:16px;'>"
        "<div style='font-size:48px;'>🎲</div>"
        "<div style='font-size:24px; font-weight:bold; color:#1e3c78;'>"
        "Entre Linhas — Manual Completo</div>"
        "<div style='font-size:14px; color:#64748b; margin-top:6px;'>"
        "Como jogar, regras, dicas e recuperação de sala</div>"
        "</div>", unsafe_allow_html=True)
    st.divider()
    with st.container(height=520, border=False):
        st.markdown(MANUAL_TEXTO)
    st.divider()
    col_pdf, col_fechar = st.columns(2)
    with col_pdf:
        pdf_bytes = gerar_manual_pdf()
        st.download_button(
            label="⬇️ Baixar Manual em PDF",
            data=pdf_bytes,
            file_name="manual_entre_linhas.pdf",
            mime="application/pdf", type="primary",
            use_container_width=True, key="btn_baixar_manual_pdf")
    with col_fechar:
        if st.button("✖️ Fechar", use_container_width=True, key="btn_fechar_manual"):
            st.rerun()


# =====================================================
# INICIALIZAÇÃO
# =====================================================
defaults = {
    "linhas": None, "colunas": None, "tamanho_atual": None,
    "nome_time_1": "Time 1", "nome_time_2": "Time 2",
    "deck": [], "sorteadas": [], "estados": {},
    "time_atual": 1, "carta_atual": None,
    "fase": "aguardando", "sala_id": None,
    "timer_ativo": False, "timer_minutos": 2, "timer_segundos": 0,
    "timer_mudo": False, "turno_iniciado_em": _agora_ms(), "turno_contador": 0,
    "timer_rodando": False, "tempo_pausado_segundos": None,
    "msg_sucesso": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# =====================================================
# ROTEADOR
# =====================================================
qp = st.query_params
sala_qp = qp.get("sala")
role_qp = qp.get("role")

if role_qp == "tirador" and sala_qp:
    st.title("🎴 Controle do Tirador")
    render_mobile(sala_qp.upper())
    st.stop()

if sala_qp and st.session_state.sala_id is None:
    recuperar_sala(sala_qp.upper())


# =====================================================
# DESKTOP
# =====================================================
st.title("🎲 Entre Linhas — Jogo Principal")
st.caption("Biblioteca com **343 palavras**. Sorteie a grade, exiba o QR code e jogue com 2 times.")

if st.session_state.msg_sucesso:
    st.success(st.session_state.msg_sucesso)
    st.session_state.msg_sucesso = None


# =====================================================
# SIDEBAR
# =====================================================
with st.sidebar:
    st.header("⚙️ Configurações")

    if st.button("📖 Manual do Jogo", use_container_width=True, key="btn_manual_sidebar"):
        modal_manual()

    st.divider()
    tamanho = st.select_slider("Tamanho da grade", options=[3, 4, 5, 6], value=5)
    st.caption(f"Precisa de **{tamanho * 2} palavras** sorteadas.")

    st.divider()
    st.subheader("👥 Times")
    nome_time_1 = st.text_input("🔵 Time 1 (azul)",
                                 value=st.session_state.nome_time_1,
                                 max_chars=20, key="input_time_1")
    nome_time_2 = st.text_input("🔴 Time 2 (vermelho)",
                                 value=st.session_state.nome_time_2,
                                 max_chars=20, key="input_time_2")
    st.session_state.nome_time_1 = nome_time_1.strip() or "Time 1"
    st.session_state.nome_time_2 = nome_time_2.strip() or "Time 2"

    st.divider()
    st.subheader("⏱️ Timer de turno")
    st.session_state.timer_ativo = st.checkbox("Ativar cronômetro",
                                                value=st.session_state.timer_ativo)
    if st.session_state.timer_ativo:
        c1, c2 = st.columns(2)
        with c1:
            st.session_state.timer_minutos = st.number_input(
                "Minutos", 0, 60, st.session_state.timer_minutos, 1)
        with c2:
            st.session_state.timer_segundos = st.number_input(
                "Segundos", 0, 59, st.session_state.timer_segundos, 5)
        cm, _ = st.columns(2)
        with cm:
            icone = "🔇 Som off" if st.session_state.timer_mudo else "🔊 Som on"
            if st.button(icone, use_container_width=True):
                st.session_state.timer_mudo = not st.session_state.timer_mudo
                st.rerun()

    st.divider()
    st.subheader("📚 Biblioteca")
    editar = st.toggle("✏️ Editar palavras", value=False)
    if editar:
        texto_novo = st.text_area("Uma palavra por linha:",
                                   value="\n".join(PALAVRAS), height=200)
        palavras_usuario = [p.strip() for p in texto_novo.split("\n") if p.strip()]
    else:
        palavras_usuario = PALAVRAS[:]

    st.divider()
    sortear_btn = st.button("🎲 Sortear grade", type="primary", use_container_width=True)

    if st.button("🔄 Zerar partida", use_container_width=True):
        if st.session_state.sala_id:
            fb_delete(st.session_state.sala_id)
        for k in defaults:
            st.session_state[k] = defaults[k]
        st.query_params.clear()
        st.rerun()

    st.divider()
    st.subheader("🔁 Recuperar Sala")
    st.caption("Perdeu a sala? Cole o código para voltar.")
    codigo_rec = st.text_input("Código da sala", max_chars=8,
                                placeholder="Ex.: ABC12345",
                                key="input_recuperar").upper().strip()
    if st.button("🔓 Recuperar", use_container_width=True, key="btn_recuperar"):
        if not codigo_rec:
            st.warning("Digite o código da sala.")
        else:
            dados_teste = fb_get(codigo_rec)
            if not dados_teste or "linhas_fb" not in dados_teste:
                st.error("Sala não encontrada ou dados incompletos.")
            else:
                if recuperar_sala(codigo_rec):
                    st.query_params["sala"] = codigo_rec
                    st.rerun()
                else:
                    st.error("Erro ao recuperar sala.")

    if st.session_state.sala_id:
        st.divider()
        st.subheader("📱 QR do Tirador")
        url_mobile = f"{APP_URL}?sala={st.session_state.sala_id}&role=tirador"
        st.image(gerar_qr_code(url_mobile), use_column_width=True)
        st.caption(f"Sala: **{st.session_state.sala_id}**")
        st.caption("Escaneie **uma vez** com o celular.")


# =====================================================
# SORTEIO DA GRADE
# =====================================================
if sortear_btn:
    if len(palavras_usuario) < tamanho * 2:
        st.error(f"❌ Você tem apenas **{len(palavras_usuario)} palavras**.")
    elif not FIREBASE_URL:
        st.error("⚠️ Configure o **FIREBASE_URL** em Secrets antes de iniciar.")
    else:
        linhas, colunas = sortear(palavras_usuario, tamanho)
        sala_id = gerar_sala_id()
        coords = gerar_coordenadas(tamanho)

        st.session_state.linhas = linhas
        st.session_state.colunas = colunas
        st.session_state.tamanho_atual = tamanho
        st.session_state.deck = coords
        st.session_state.estados = {c: 0 for c in coords}
        st.session_state.sorteadas = []
        st.session_state.time_atual = 1
        st.session_state.carta_atual = None
        st.session_state.fase = "aguardando"
        st.session_state.sala_id = sala_id
        st.session_state.turno_iniciado_em = _agora_ms()
        st.session_state.turno_contador += 1
        st.session_state.timer_rodando = False
        st.session_state.tempo_pausado_segundos = None

        tempo_total_seg = 0
        if st.session_state.timer_ativo:
            tempo_total_seg = (
                st.session_state.timer_minutos * 60
                + st.session_state.timer_segundos
            )

        fb_put(sala_id, {
            "coord": None, "estado": "aguardando", "evento": None,
            "time_atual": 1, "nome_turno": st.session_state.nome_time_1,
            "nome_time_1_fb": st.session_state.nome_time_1,
            "nome_time_2_fb": st.session_state.nome_time_2,
            "criada_em": _agora_ms(),
            "timer_ativo": st.session_state.timer_ativo and tempo_total_seg > 0,
            "tempo_total": tempo_total_seg, "timer_iniciado_em": None,
            "solicitar_sorteio": None, "baralho_vazio": False,
            "tamanho": tamanho, "linhas_fb": linhas, "colunas_fb": colunas,
            "deck_fb": coords, "sorteadas_fb": [],
            "estados_fb": {c: 0 for c in coords},
            "timer_minutos_fb": st.session_state.timer_minutos,
            "timer_segundos_fb": st.session_state.timer_segundos,
            "timer_mudo_fb": st.session_state.timer_mudo,
        })
        st.session_state.msg_sucesso = f"✅ Grade {tamanho}×{tamanho} pronta! Sala {sala_id}."
        st.query_params["sala"] = sala_id
        st.rerun()


if st.session_state.linhas is None:
    st.info("👈 Configure na barra lateral e clique em **🎲 Sortear grade** para começar.")
    st.stop()

linhas = st.session_state.linhas
colunas = st.session_state.colunas
tam = st.session_state.tamanho_atual
nome_t1 = st.session_state.nome_time_1
nome_t2 = st.session_state.nome_time_2
deck = st.session_state.deck
estados = st.session_state.estados
time_atual = st.session_state.time_atual
carta_atual = st.session_state.carta_atual
fase = st.session_state.fase
sala_id = st.session_state.sala_id
total_cartas = tam * tam
nome_turno = nome_t1 if time_atual == 1 else nome_t2
cor_turno = "#1e3c78" if time_atual == 1 else "#dc2626"
jogo_acabou = (len(deck) == 0 and carta_atual is None)


# =====================================================
# POLLING
# =====================================================
if sala_id and (fase == "decidindo" or fase == "aguardando"):
    st_autorefresh(interval=1500, key="desktop_poll")
    dados_fb = fb_get(sala_id)
    evento = dados_fb.get("evento")
    timer_fb = dados_fb.get("timer_iniciado_em")
    solicitacao = dados_fb.get("solicitar_sorteio")

    baralho_vazio_fb = dados_fb.get("baralho_vazio", False)
    baralho_vazio_local = (len(deck) == 0)
    if baralho_vazio_fb != baralho_vazio_local:
        fb_patch(sala_id, {"baralho_vazio": baralho_vazio_local})

    if solicitacao and fase == "aguardando":
        if len(deck) > 0:
            nova = deck[0]
            st.session_state.carta_atual = nova
            st.session_state.fase = "decidindo"
            st.session_state.timer_rodando = False
            st.session_state.tempo_pausado_segundos = None
            fb_patch(sala_id, {
                "coord": nova, "estado": "sorteada", "evento": None,
                "time_atual": time_atual, "nome_turno": nome_turno,
                "timer_iniciado_em": None, "solicitar_sorteio": None,
            })
        else:
            fb_patch(sala_id, {"solicitar_sorteio": None})
        st.rerun()

    if timer_fb and not st.session_state.timer_rodando and st.session_state.timer_ativo:
        st.session_state.turno_iniciado_em = timer_fb
        st.session_state.tempo_pausado_segundos = None
        st.session_state.timer_rodando = True
        st.session_state.turno_contador += 1
        st.rerun()

    if evento == "acertou":
        estados[carta_atual] = time_atual
        st.session_state.estados = estados
        st.session_state.deck = deck[1:]
        st.session_state.sorteadas = st.session_state.sorteadas + [carta_atual]

        novo_time = 2 if time_atual == 1 else 1
        novo_nome = nome_t2 if novo_time == 2 else nome_t1

        st.session_state.time_atual = novo_time
        st.session_state.carta_atual = None
        st.session_state.fase = "aguardando"
        st.session_state.turno_iniciado_em = _agora_ms()
        st.session_state.turno_contador += 1
        st.session_state.timer_rodando = False
        st.session_state.tempo_pausado_segundos = None

        fb_patch(sala_id, {
            "evento": None, "coord": None, "estado": "aguardando",
            "timer_iniciado_em": None,
            "time_atual": novo_time, "nome_turno": novo_nome,
            "deck_fb": st.session_state.deck,
            "sorteadas_fb": st.session_state.sorteadas,
            "estados_fb": estados,
        })
        st.rerun()

    elif evento == "errou":
        estados[carta_atual] = 3
        st.session_state.estados = estados
        st.session_state.deck = deck[1:]
        st.session_state.sorteadas = st.session_state.sorteadas + [carta_atual]

        novo_time = 2 if time_atual == 1 else 1
        novo_nome = nome_t2 if novo_time == 2 else nome_t1

        st.session_state.time_atual = novo_time
        st.session_state.carta_atual = None
        st.session_state.fase = "aguardando"
        st.session_state.turno_iniciado_em = _agora_ms()
        st.session_state.turno_contador += 1
        st.session_state.timer_rodando = False
        st.session_state.tempo_pausado_segundos = None

        fb_patch(sala_id, {
            "evento": None, "coord": None, "estado": "aguardando",
            "timer_iniciado_em": None,
            "time_atual": novo_time, "nome_turno": novo_nome,
            "deck_fb": st.session_state.deck,
            "sorteadas_fb": st.session_state.sorteadas,
            "estados_fb": estados,
        })
        st.rerun()


# =====================================================
# LAYOUT PRINCIPAL
# =====================================================
col_esq, col_dir = st.columns([2, 1], gap="large")


with col_esq:
    if jogo_acabou:
        vencedor, p1, p2, desc, cor_v = detectar_vencedor(estados, nome_t1, nome_t2)
        render_banner_vitoria(vencedor, p1, p2, desc, cor_v, nome_t1, nome_t2)

    st.markdown(
        f"<h2 style='margin-bottom:8px;'>🎯 Grade {tam}×{tam}</h2>",
        unsafe_allow_html=True,
    )
    st.caption("🔵 Time 1 · 🔴 Time 2 · ⚪ cinza = descartada")
    render_grid_estados(linhas, colunas, estados, nome_t1, nome_t2)
    st.divider()

    with st.expander(f"📜 Histórico ({len(st.session_state.sorteadas)} cartas)"):
        if st.session_state.sorteadas:
            for idx, c in enumerate(st.session_state.sorteadas, 1):
                estado = estados.get(c, 0)
                tag = {1: "🔵 Time 1", 2: "🔴 Time 2", 3: "❌ Descartada"}.get(estado, "?")
                st.markdown(f"`{idx}.` **{c}** — {tag}")
        else:
            st.caption("Nenhuma carta sorteada ainda.")


with col_dir:
    if jogo_acabou:
        vencedor_a, _, _, _, cor_a = detectar_vencedor(estados, nome_t1, nome_t2)
        if vencedor_a == "t1":
            label_time, emoji_time = nome_t1, "🏆🔵"
        elif vencedor_a == "t2":
            label_time, emoji_time = nome_t2, "🏆🔴"
        elif vencedor_a == "empate":
            label_time, emoji_time = "Empate", "🤝"
        else:
            label_time, emoji_time = "—", "🏁"
        st.markdown(
            f"""
            <div style="padding:24px 20px; background:{cor_a}15;
                        border:4px solid {cor_a}; border-radius:16px;
                        text-align:center; font-family:system-ui;
                        margin-bottom:16px; box-shadow: 0 3px 12px {cor_a}22;">
                <div style="font-size:13px;color:#64748b;
                            text-transform:uppercase;letter-spacing:3px;
                            font-weight:bold;">Fim de jogo</div>
                <div style="font-size:44px;font-weight:900;
                            color:{cor_a};margin-top:10px;
                            line-height:1.05;letter-spacing:1px;
                            text-shadow: 0 3px 14px {cor_a}22;">
                    {emoji_time} {label_time}</div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.markdown(
            f"""
            <div style="padding:24px 20px; background:{cor_turno}15;
                        border:4px solid {cor_turno}; border-radius:16px;
                        text-align:center; font-family:system-ui;
                        margin-bottom:16px; box-shadow: 0 3px 12px {cor_turno}22;">
                <div style="font-size:13px;color:#64748b;
                            text-transform:uppercase;letter-spacing:3px;
                            font-weight:bold;">Rodada do</div>
                <div style="font-size:44px;font-weight:900;
                            color:{cor_turno};margin-top:10px;
                            line-height:1.05;letter-spacing:1px;
                            text-shadow: 0 3px 14px {cor_turno}22;">
                    {nome_turno}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("""
        <div style="font-size:14px; color:#64748b; text-transform:uppercase;
                    letter-spacing:3px; text-align:center;
                    margin-bottom:10px; font-weight:bold;">
            ⏱️ Cronômetro
        </div>
    """, unsafe_allow_html=True)

    if jogo_acabou:
        st.info("🏁 Jogo encerrado.")
    elif st.session_state.timer_ativo:
        tempo_total = st.session_state.timer_minutos * 60 + st.session_state.timer_segundos
        if tempo_total > 0:
            rodando = st.session_state.timer_rodando
            pausado_seg = st.session_state.tempo_pausado_segundos
            restante_exibir = tempo_total if pausado_seg is None else pausado_seg

            if rodando:
                render_timer(
                    tempo_total=tempo_total,
                    iniciado_em_ms=st.session_state.turno_iniciado_em,
                    mudo=st.session_state.timer_mudo,
                    key=f"turno_{st.session_state.turno_contador}",
                )
            else:
                pct = (restante_exibir / tempo_total) * 100 if tempo_total > 0 else 0
                if pct > 50: cor = "#22c55e"
                elif pct > 20: cor = "#eab308"
                else: cor = "#dc2626"
                st.markdown(
                    f"""
                    <div style="padding:20px; border-radius:16px;
                                background:{cor}15; border:4px solid {cor};
                                text-align:center; font-family:system-ui;">
                        <div style="font-size:100px;font-weight:900;color:{cor};
                                    font-variant-numeric:tabular-nums;line-height:1;
                                    letter-spacing:3px;
                                    text-shadow: 0 3px 14px rgba(0,0,0,0.06);">
                            {restante_exibir // 60:02d}:{restante_exibir % 60:02d}</div>
                        <div style="font-size:13px;color:#64748b;
                                    text-transform:uppercase;letter-spacing:2px;
                                    margin-top:8px;">
                            {'Pausado' if pausado_seg is not None else 'Pronto para iniciar'}</div>
                        <div style="width:100%;height:14px;background:#e2e8f0;
                                    border-radius:7px;overflow:hidden;margin-top:16px;">
                            <div style="width:{pct}%;height:100%;background:{cor};"></div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

            st.write("")
            c_play, c_reset = st.columns(2)
            with c_play:
                if not rodando:
                    label = "▶️ Iniciar" if pausado_seg is None else "▶️ Retomar"
                    if st.button(label, type="primary", use_container_width=True, key="timer_play"):
                        if pausado_seg is None:
                            st.session_state.turno_iniciado_em = _agora_ms()
                        else:
                            segundos_decorridos = tempo_total - pausado_seg
                            st.session_state.turno_iniciado_em = (
                                _agora_ms() - (segundos_decorridos * 1000))
                        st.session_state.tempo_pausado_segundos = None
                        st.session_state.timer_rodando = True
                        st.session_state.turno_contador += 1
                        if sala_id:
                            fb_patch(sala_id, {
                                "timer_iniciado_em": st.session_state.turno_iniciado_em})
                        st.rerun()
                else:
                    if st.button("⏸️ Pausar", use_container_width=True, key="timer_pause"):
                        decorrido_ms = _agora_ms() - st.session_state.turno_iniciado_em
                        restante = max(0, int(tempo_total - decorrido_ms / 1000))
                        st.session_state.tempo_pausado_segundos = restante
                        st.session_state.timer_rodando = False
                        st.rerun()
            with c_reset:
                if st.button("🔄 Reiniciar", use_container_width=True, key="timer_reset"):
                    st.session_state.turno_iniciado_em = _agora_ms()
                    st.session_state.tempo_pausado_segundos = None
                    st.session_state.timer_rodando = False
                    st.session_state.turno_contador += 1
                    if sala_id:
                        fb_patch(sala_id, {"timer_iniciado_em": None})
                    st.rerun()

            if rodando:
                st.markdown("<div style='text-align:center;color:#22c55e;"
                    "font-weight:bold;padding-top:8px;font-size:15px;'>"
                    "🟢 Cronômetro rodando</div>", unsafe_allow_html=True)
            elif pausado_seg is not None:
                st.markdown("<div style='text-align:center;color:#eab308;"
                    "font-weight:bold;padding-top:8px;font-size:15px;'>"
                    "⏸️ Pausado</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div style='text-align:center;color:#64748b;"
                    f"font-weight:bold;padding-top:8px;font-size:15px;'>"
                    f"⚪ Aguardando início ({tempo_total // 60:02d}:{tempo_total % 60:02d})</div>",
                    unsafe_allow_html=True)
        else:
            st.info("Configure minutos ou segundos na barra lateral.")
    else:
        st.info("⏱️ Cronômetro desativado.")

    st.markdown("""
        <div style="font-size:14px; color:#64748b; text-transform:uppercase;
                    letter-spacing:3px; text-align:center;
                    margin-top:24px; margin-bottom:10px; font-weight:bold;">
            📊 Cartas Faltantes
        </div>
    """, unsafe_allow_html=True)

    faltantes = len(deck)
    pct_restante = (faltantes / total_cartas) * 100 if total_cartas > 0 else 0
    if faltantes == 0:
        cor_f, icone_f, label_f = "#22c55e", "🏁", "Baralho vazio"
    elif pct_restante > 50:
        cor_f, icone_f, label_f = "#22c55e", "🟢", "Ainda tem muito jogo"
    elif pct_restante > 20:
        cor_f, icone_f, label_f = "#eab308", "🟡", "Reta final chegando"
    else:
        cor_f, icone_f, label_f = "#dc2626", "🔴", "Últimas cartas!"

    sorteadas_qtd = len(st.session_state.sorteadas)
    st.markdown(
        f"""
        <div style="padding:24px 20px; border-radius:16px;
                    background:{cor_f}15; border:4px solid {cor_f};
                    text-align:center; font-family:system-ui;
                    box-shadow: 0 3px 12px {cor_f}22;">
            <div style="font-size:90px; font-weight:900; color:{cor_f};
                        line-height:1; font-variant-numeric:tabular-nums;
                        letter-spacing:2px; text-shadow: 0 3px 14px {cor_f}22;">
                {faltantes}</div>
            <div style="font-size:13px; color:#64748b;
                        text-transform:uppercase; letter-spacing:2px;
                        margin-top:8px; font-weight:bold;">
                {icone_f} cartas no baralho</div>
            <div style="width:100%; height:14px; background:#e2e8f0;
                        border-radius:7px; overflow:hidden; margin-top:16px;">
                <div style="width:{pct_restante}%; height:100%; background:{cor_f};"></div>
            </div>
            <div style="font-size:12px; color:#94a3b8; margin-top:10px;">
                {sorteadas_qtd} de {total_cartas} já sorteadas — {label_f}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
        <div style="font-size:14px; color:#64748b; text-transform:uppercase;
                    letter-spacing:3px; text-align:center;
                    margin-top:24px; margin-bottom:10px; font-weight:bold;">
            🎴 Ação da Rodada
        </div>
    """, unsafe_allow_html=True)

    if jogo_acabou:
        if st.button("🔄 Nova partida (mesma grade)", type="primary",
                     use_container_width=True, key="nova_partida_painel"):
            coords = gerar_coordenadas(tam)
            st.session_state.deck = coords
            st.session_state.estados = {c: 0 for c in coords}
            st.session_state.sorteadas = []
            st.session_state.time_atual = 1
            st.session_state.carta_atual = None
            st.session_state.fase = "aguardando"
            st.session_state.turno_iniciado_em = _agora_ms()
            st.session_state.turno_contador += 1
            st.session_state.timer_rodando = False
            st.session_state.tempo_pausado_segundos = None
            fb_patch(sala_id, {
                "coord": None, "evento": None, "estado": "aguardando",
                "timer_iniciado_em": None, "time_atual": 1,
                "nome_turno": nome_t1, "solicitar_sorteio": None,
                "baralho_vazio": False, "deck_fb": coords,
                "sorteadas_fb": [], "estados_fb": {c: 0 for c in coords},
            })
            st.rerun()

        if st.button("🎲 Nova grade (novo sorteio)", use_container_width=True,
                     key="nova_grade_painel"):
            st.session_state.linhas = None
            st.query_params.clear()
            st.rerun()

    elif fase == "aguardando":
        if st.button("🎴 SORTEAR CARTA SECRETA", type="primary",
                     use_container_width=True, key="sortear_desktop_painel"):
            if not deck:
                st.warning("Baralho vazio!")
            else:
                nova = deck[0]
                st.session_state.carta_atual = nova
                st.session_state.fase = "decidindo"
                st.session_state.timer_rodando = False
                st.session_state.tempo_pausado_segundos = None
                fb_patch(sala_id, {
                    "coord": nova, "estado": "sorteada", "evento": None,
                    "time_atual": time_atual, "nome_turno": nome_turno,
                    "timer_iniciado_em": None,
                })
                st.rerun()
        st.caption(f"👉 Rodada do **{nome_turno}** — clique acima ou aguarde "
                   "o tirador sortear pelo celular.")

    elif fase == "decidindo":
        st.markdown(
            f"""
            <div style="padding:20px; border-radius:12px;
                        background:{cor_turno}15;
                        border:3px dashed {cor_turno};
                        text-align:center; font-family:system-ui;">
                <div style="font-size:13px; color:#64748b;
                            text-transform:uppercase; letter-spacing:2px;">
                    ⏳ Carta em jogo</div>
                <div style="font-size:14px; color:#94a3b8; margin-top:8px;">
                    Aguardando resposta de {nome_turno} no celular...</div>
            </div>
            """, unsafe_allow_html=True)
        with st.expander("🆘 Emergência: usar botões do desktop"):
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("✅ Acertou (manual)", use_container_width=True,
                             key="manual_acerto_painel"):
                    fb_patch(sala_id, {"evento": "acertou"})
                    st.rerun()
            with col_b:
                if st.button("❌ Errou (manual)", use_container_width=True,
                             key="manual_erro_painel"):
                    fb_patch(sala_id, {"evento": "errou"})
                    st.rerun()


st.divider()
st.caption(
    "🎲 **Entre Linhas** — cada rodada, o tirador vê a carta secretamente no celular "
    "e decide **Acertou** ou **Errou**. Vence quem conquistar mais cartas."
)
