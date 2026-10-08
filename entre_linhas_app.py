"""
Entre Linhas — Desktop + QR code para o tirador
Layout em duas colunas com auto-scale, cartas faltantes e mensagem persistida.
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
        titulo = f"🏆 {nome_t1} VENCEU!"
        subtitulo = f"{p1} × {p2} cartas conquistadas"
        emoji = "🔵"
    elif vencedor == "t2":
        titulo = f"🏆 {nome_t2} VENCEU!"
        subtitulo = f"{p2} × {p1} cartas conquistadas"
        emoji = "🔴"
    elif vencedor == "empate":
        titulo = "🤝 EMPATE!"
        subtitulo = f"{p1} × {p2} cartas conquistadas"
        emoji = "⚪"
    else:
        titulo = "🏁 FIM DE JOGO"
        subtitulo = "Nenhuma carta foi conquistada"
        emoji = "🎲"

    st.markdown(
        f"""
        <div style="padding: 32px 24px;
                    background: linear-gradient(135deg, {cor}18, {cor}08);
                    border: 4px solid {cor}; border-radius: 20px;
                    text-align: center; font-family: system-ui;
                    margin: 16px 0 24px 0;
                    box-shadow: 0 8px 32px {cor}30;">
            <div style="font-size: 64px; line-height: 1; margin-bottom: 8px;">
                {emoji}
            </div>
            <div style="font-size: 38px; font-weight: bold;
                        color: {cor}; letter-spacing: 1px; line-height: 1.1;">
                {titulo}
            </div>
            <div style="font-size: 20px; color: #64748b;
                        margin-top: 12px; font-weight: 500;">
                {subtitulo}
            </div>
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
                    🔵 {nome_t1}
                </div>
                <div style="font-size: 42px; font-weight: bold;
                            color: {('#ffffff' if vencedor == 't1' else '#1e3c78')};
                            margin-top: 4px;">
                    {p1}
                </div>
                <div style="font-size: 12px;
                            color: {('#e0e7ff' if vencedor == 't1' else '#94a3b8')};
                            margin-top: 4px;">
                    {pct1:.0f}% das cartas
                </div>
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
                    🔴 {nome_t2}
                </div>
                <div style="font-size: 42px; font-weight: bold;
                            color: {('#ffffff' if vencedor == 't2' else '#dc2626')};
                            margin-top: 4px;">
                    {p2}
                </div>
                <div style="font-size: 12px;
                            color: {('#fee2e2' if vencedor == 't2' else '#94a3b8')};
                            margin-top: 4px;">
                    {pct2:.0f}% das cartas
                </div>
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
        .timer-wrap {{ padding: 20px 20px; border-radius: 16px;
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
        3: {"cell": 230, "coord": 65, "min_h": 100,
            "f_letter": 30, "f_word": 24, "f_coord": 48},
        4: {"cell": 200, "coord": 60, "min_h": 90,
            "f_letter": 28, "f_word": 22, "f_coord": 42},
        5: {"cell": 180, "coord": 55, "min_h": 82,
            "f_letter": 26, "f_word": 20, "f_coord": 38},
        6: {"cell": 160, "coord": 50, "min_h": 74,
            "f_letter": 24, "f_word": 18, "f_coord": 34},
    }
    sz = tamanhos.get(tam, tamanhos[5])
    cell = sz["cell"]
    coord_w = sz["coord"]
    min_h = sz["min_h"]
    f_letter = sz["f_letter"]
    f_word = sz["f_word"]
    f_coord = sz["f_coord"]

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

    .placar {{
        display: flex; gap: 24px; align-items: center;
        padding: 12px 16px; background: #f8fafc;
        border: 2px solid #cbd5e1; border-radius: 10px;
        margin-bottom: 14px; font-size: 16px;
        font-weight: 600; flex-wrap: wrap;
    }}
    .placar-time {{ display: flex; align-items: center; gap: 8px; }}
    .placar-dot {{
        width: 22px; height: 22px; border-radius: 4px;
        border: 2px solid #4a5568; flex-shrink: 0;
    }}
    .dot-b {{ background: #1e3c78; }}
    .dot-r {{ background: #dc2626; }}
    .dot-x {{ background: #94a3b8; }}
    .placar-nome {{
        font-size: 15px; max-width: 180px;
        overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    }}
    .placar-num {{
        font-size: 26px; min-width: 34px;
        text-align: center; font-weight: bold;
    }}
    .placar-num.blue {{ color: #1e3c78; }}
    .placar-num.red {{ color: #dc2626; }}
    .placar-num.grey {{ color: #64748b; }}

    .grid-wrap {{
        width: 100%;
        overflow: visible;
        position: relative;
    }}
    .grid-el {{
        display: grid;
        grid-template-columns:
            {coord_w}px
            {cell}px
            repeat({tam}, {cell}px);
        gap: 6px;
        width: fit-content;
        transform-origin: top left;
        transition: transform 0.15s ease-out;
    }}
    .cel {{
        border: 3px solid #4a5568;
        padding: 10px 6px;
        text-align: center;
        display: flex;
        align-items: center;
        justify-content: center;
        min-height: {min_h}px;
        font-size: {f_word}px;
        border-radius: 6px;
        user-select: none;
        overflow: hidden;
        word-break: break-word;
        line-height: 1.15;
    }}
    .cel-coord {{
        background: #f0f0f0;
        font-weight: bold;
        font-size: {f_letter}px;
        color: #323232;
    }}
    .cel-canto {{
        background: #e6e6e6;
        font-weight: bold;
        font-size: {f_letter + 2}px;
    }}
    .cel-col {{
        background: #c5e0b4;
        font-weight: bold;
        font-size: {f_word}px;
    }}
    .cel-lin {{
        background: #ffe699;
        font-weight: bold;
        font-size: {f_word}px;
    }}
    .cel-carta {{
        font-weight: 900;
        font-size: {f_coord}px;
        min-height: {min_h}px;
        transition: background 0.3s, color 0.3s;
        letter-spacing: 0.5px;
    }}
    .cel-carta.state-0 {{ background: #d9e1f2; color: #1e3c78; }}
    .cel-carta.state-1 {{ background: #1e3c78; color: #ffffff; }}
    .cel-carta.state-2 {{ background: #dc2626; color: #ffffff; }}
    .cel-carta.state-3 {{
        background: #cbd5e1; color: #64748b;
        text-decoration: line-through; opacity: 0.7;
    }}
</style></head><body>

<div class="placar">
    <div class="placar-time">
        <span class="placar-dot dot-b"></span>
        <span class="placar-nome">{nome_time_1}</span>
        <span class="placar-num blue">{p1}</span>
    </div>
    <div class="placar-time">
        <span class="placar-dot dot-r"></span>
        <span class="placar-nome">{nome_time_2}</span>
        <span class="placar-num red">{p2}</span>
    </div>
    <div class="placar-time">
        <span class="placar-dot dot-x"></span>
        <span class="placar-nome">Descartadas</span>
        <span class="placar-num grey">{desc}</span>
    </div>
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
    altura = alturas.get(tam, 1150)
    components.html(html, height=altura, scrolling=False)


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
    nome_turno = dados.get("nome_turno", "Time")
    cor_turno = "#1e3c78" if dados.get("time_atual", 1) == 1 else "#dc2626"

    st.markdown(
        f"""
        <div style="padding:10px 16px; background:{cor_turno}15;
                    border-left:6px solid {cor_turno}; border-radius:8px;
                    font-family:system-ui; margin-bottom:16px;">
            <div style="font-size:11px;color:#64748b;
                        text-transform:uppercase;letter-spacing:1px;">
                Vez de
            </div>
            <div style="font-size:20px;font-weight:bold;color:{cor_turno};">
                {nome_turno}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not coord or estado == "aguardando":
        st.info("⏳ Aguardando o dispositivo principal sortear uma carta...")
        st.caption("Não feche esta página.")
        return

    if evento:
        st.success("✅ Decisão registrada!")
        st.markdown(
            f"""
            <div style="text-align:center; padding:40px 20px;
                        background:#f1f5f9; border-radius:12px;
                        font-family:system-ui;">
                <div style="font-size:14px;color:#64748b;">Você marcou</div>
                <div style="font-size:32px;font-weight:bold;
                            color:{'#22c55e' if evento == 'acertou' else '#dc2626'};
                            margin-top:12px;">
                    {'✅ ACERTOU' if evento == 'acertou' else '❌ ERROU'}
                </div>
                <div style="font-size:13px;color:#94a3b8;margin-top:16px;">
                    Aguardando próxima rodada...
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        f"""
        <div style="text-align:center; padding:30px 20px;
                    background:#d9e1f2; border:4px solid #1e3c78;
                    border-radius:16px; margin-bottom:20px;
                    font-family:system-ui;">
            <div style="font-size:13px;color:#64748b;
                        text-transform:uppercase;letter-spacing:1.5px;">
                Sua carta secreta
            </div>
            <div style="font-size:120px;font-weight:bold;color:#1e3c78;
                        letter-spacing:8px; line-height:1; margin:16px 0;">
                {coord}
            </div>
            <div style="font-size:12px;color:#94a3b8;">
                Dê uma dica e aguarde seu time responder
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
            mm = restante_seg // 60
            ss = restante_seg % 60
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
                            text-align:center; font-family:system-ui;
                            margin:14px 0;">
                    <div style="font-size:12px;color:#64748b;
                                text-transform:uppercase;letter-spacing:1px;">
                        {icone} Cronômetro {texto_estado}
                    </div>
                    <div style="font-size:42px;font-weight:bold;color:{cor};
                                font-variant-numeric:tabular-nums;line-height:1;
                                margin-top:6px;">
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
    st.caption("Clique assim que o time der a resposta.")


# =====================================================
# INICIALIZAÇÃO
# =====================================================
defaults = {
    "linhas": None, "colunas": None, "tamanho_atual": None,
    "nome_time_1": "Time 1", "nome_time_2": "Time 2",
    "deck": [], "sorteadas": [], "estados": {},
    "time_atual": 1, "carta_atual": None,
    "fase": "aguardando",
    "sala_id": None,
    "timer_ativo": False, "timer_minutos": 2, "timer_segundos": 0,
    "timer_mudo": False, "turno_iniciado_em": _agora_ms(), "turno_contador": 0,
    "timer_rodando": False,
    "tempo_pausado_segundos": None,
    "msg_sucesso": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


qp = st.query_params
sala_qp = qp.get("sala")
role_qp = qp.get("role")

if role_qp == "tirador" and sala_qp:
    st.title("🎴 Controle do Tirador")
    render_mobile(sala_qp.upper())
    st.stop()


st.title("🎲 Entre Linhas — Jogo Principal")
st.caption("Biblioteca com **343 palavras**. Sorteie a grade, exiba o QR code e jogue com 2 times.")

# ---- Mensagem de sucesso persistida (após rerun do sorteio) ----
if st.session_state.msg_sucesso:
    st.success(st.session_state.msg_sucesso)
    st.session_state.msg_sucesso = None


# =====================================================
# SIDEBAR
# =====================================================
with st.sidebar:
    st.header("⚙️ Configurações")
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
        st.caption("▶️ Inicie pelo **celular do tirador** ou pelos botões ao lado.")

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
        st.rerun()

    # ---- QR CODE DO TIRADOR (dentro da sidebar) ----
    if st.session_state.sala_id:
        st.divider()
        st.subheader("📱 QR do Tirador")
        url_mobile = f"{APP_URL}?sala={st.session_state.sala_id}&role=tirador"
        st.image(gerar_qr_code(url_mobile), use_column_width=True)
        st.caption(f"Sala: **{st.session_state.sala_id}**")
        st.caption("O tirador escaneia uma vez com o celular.")


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
            "coord": None,
            "estado": "aguardando",
            "evento": None,
            "time_atual": 1,
            "nome_turno": st.session_state.nome_time_1,
            "criada_em": _agora_ms(),
            "timer_ativo": st.session_state.timer_ativo and tempo_total_seg > 0,
            "tempo_total": tempo_total_seg,
            "timer_iniciado_em": None,
        })

        # ---- Mensagem persistida + rerun para o QR aparecer ----
        st.session_state.msg_sucesso = f"✅ Grade {tamanho}×{tamanho} pronta! Sala {sala_id}."
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
if sala_id and fase == "decidindo":
    st_autorefresh(interval=1500, key="desktop_poll")
    dados_fb = fb_get(sala_id)
    evento = dados_fb.get("evento")
    timer_fb = dados_fb.get("timer_iniciado_em")

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
        st.session_state.time_atual = 2 if time_atual == 1 else 1
        st.session_state.carta_atual = None
        st.session_state.fase = "aguardando"
        st.session_state.turno_iniciado_em = _agora_ms()
        st.session_state.turno_contador += 1
        st.session_state.timer_rodando = False
        st.session_state.tempo_pausado_segundos = None
        fb_patch(sala_id, {
            "evento": None, "coord": None,
            "estado": "aguardando",
            "timer_iniciado_em": None,
        })
        st.rerun()

    elif evento == "errou":
        estados[carta_atual] = 3
        st.session_state.estados = estados
        st.session_state.deck = deck[1:]
        st.session_state.sorteadas = st.session_state.sorteadas + [carta_atual]
        st.session_state.time_atual = 2 if time_atual == 1 else 1
        st.session_state.carta_atual = None
        st.session_state.fase = "aguardando"
        st.session_state.turno_iniciado_em = _agora_ms()
        st.session_state.turno_contador += 1
        st.session_state.timer_rodando = False
        st.session_state.tempo_pausado_segundos = None
        fb_patch(sala_id, {
            "evento": None, "coord": None,
            "estado": "aguardando",
            "timer_iniciado_em": None,
        })
        st.rerun()


# =====================================================
# LAYOUT PRINCIPAL
# =====================================================
col_esq, col_dir = st.columns([2, 1], gap="large")


# =====================================================
# COLUNA ESQUERDA — GRADE
# =====================================================
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

    if jogo_acabou:
        col_n1, col_n2 = st.columns(2)
        with col_n1:
            if st.button("🔄 Nova partida (mesma grade)",
                         type="primary", use_container_width=True):
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
                    "coord": None, "evento": None,
                    "estado": "aguardando",
                    "timer_iniciado_em": None,
                    "time_atual": 1,
                    "nome_turno": nome_t1,
                })
                st.rerun()
        with col_n2:
            if st.button("🎲 Nova grade (novo sorteio)",
                         use_container_width=True):
                st.session_state.linhas = None
                st.rerun()

    elif fase == "aguardando":
        st.caption("👉 O tirador escaneia o QR code na barra lateral. "
                   "Depois clique abaixo para sortear a carta.")
        if st.button("🎴 Sortear carta secreta", type="primary", use_container_width=True):
            if not deck:
                st.warning("Baralho vazio!")
            else:
                nova = deck[0]
                st.session_state.carta_atual = nova
                st.session_state.fase = "decidindo"
                st.session_state.timer_rodando = False
                st.session_state.tempo_pausado_segundos = None
                fb_patch(sala_id, {
                    "coord": nova,
                    "estado": "sorteada",
                    "evento": None,
                    "time_atual": time_atual,
                    "nome_turno": nome_turno,
                    "timer_iniciado_em": None,
                })
                st.rerun()

    elif fase == "decidindo":
        st.info("⏳ **Carta em jogo** — aguardando o tirador responder no celular…")
        st.caption("🔄 Atualizando automaticamente a cada 1.5s.")
        with st.expander("🆘 Emergência: usar botões do desktop"):
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("✅ Acertou (manual)", use_container_width=True):
                    fb_patch(sala_id, {"evento": "acertou"})
                    st.rerun()
            with col_b:
                if st.button("❌ Errou (manual)", use_container_width=True):
                    fb_patch(sala_id, {"evento": "errou"})
                    st.rerun()

    with st.expander(f"📜 Histórico ({len(st.session_state.sorteadas)} cartas)"):
        if st.session_state.sorteadas:
            for idx, c in enumerate(st.session_state.sorteadas, 1):
                estado = estados.get(c, 0)
                tag = {1: "🔵 Time 1", 2: "🔴 Time 2", 3: "❌ Descartada"}.get(estado, "?")
                st.markdown(f"`{idx}.` **{c}** — {tag}")
        else:
            st.caption("Nenhuma carta sorteada ainda.")


# =====================================================
# COLUNA DIREITA
# =====================================================
with col_dir:

    # ---------- QUADRANTE A ----------
    if jogo_acabou:
        vencedor_a, _, _, _, cor_a = detectar_vencedor(estados, nome_t1, nome_t2)
        if vencedor_a == "t1":
            label_time = nome_t1
            emoji_time = "🏆🔵"
        elif vencedor_a == "t2":
            label_time = nome_t2
            emoji_time = "🏆🔴"
        elif vencedor_a == "empate":
            label_time = "Empate"
            emoji_time = "🤝"
        else:
            label_time = "—"
            emoji_time = "🏁"

        st.markdown(
            f"""
            <div style="padding:24px 20px; background:{cor_a}15;
                        border:4px solid {cor_a}; border-radius:16px;
                        text-align:center; font-family:system-ui;
                        margin-bottom:16px; box-shadow: 0 3px 12px {cor_a}22;">
                <div style="font-size:13px;color:#64748b;
                            text-transform:uppercase;letter-spacing:3px;
                            font-weight:bold;">
                    Fim de jogo
                </div>
                <div style="font-size:44px;font-weight:900;
                            color:{cor_a};margin-top:10px;
                            line-height:1.05;letter-spacing:1px;
                            text-shadow: 0 3px 14px {cor_a}22;">
                    {emoji_time} {label_time}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""
            <div style="padding:24px 20px; background:{cor_turno}15;
                        border:4px solid {cor_turno}; border-radius:16px;
                        text-align:center; font-family:system-ui;
                        margin-bottom:16px; box-shadow: 0 3px 12px {cor_turno}22;">
                <div style="font-size:13px;color:#64748b;
                            text-transform:uppercase;letter-spacing:3px;
                            font-weight:bold;">
                    Rodada do
                </div>
                <div style="font-size:44px;font-weight:900;
                            color:{cor_turno};margin-top:10px;
                            line-height:1.05;letter-spacing:1px;
                            text-shadow: 0 3px 14px {cor_turno}22;">
                    {nome_turno}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ---------- QUADRANTE B ----------
    st.markdown(
        """
        <div style="font-size:14px; color:#64748b;
                    text-transform:uppercase; letter-spacing:3px;
                    text-align:center; margin-bottom:10px; font-weight:bold;">
            ⏱️ Cronômetro
        </div>
        """,
        unsafe_allow_html=True,
    )

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
                if pct > 50:
                    cor = "#22c55e"
                elif pct > 20:
                    cor = "#eab308"
                else:
                    cor = "#dc2626"
                st.markdown(
                    f"""
                    <div style="padding:20px 20px; border-radius:16px;
                                background:{cor}15; border:4px solid {cor};
                                text-align:center; font-family:system-ui;">
                        <div style="font-size:100px;font-weight:900;color:{cor};
                                    font-variant-numeric:tabular-nums;line-height:1;
                                    letter-spacing:3px;
                                    text-shadow: 0 3px 14px rgba(0,0,0,0.06);">
                            {restante_exibir // 60:02d}:{restante_exibir % 60:02d}
                        </div>
                        <div style="font-size:13px;color:#64748b;
                                    text-transform:uppercase;letter-spacing:2px;
                                    margin-top:8px;">
                            {'Pausado' if pausado_seg is not None else 'Pronto para iniciar'}
                        </div>
                        <div style="width:100%;height:14px;background:#e2e8f0;
                                    border-radius:7px;overflow:hidden;margin-top:16px;">
                            <div style="width:{pct}%;height:100%;background:{cor};"></div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

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
                                _agora_ms() - (segundos_decorridos * 1000)
                            )
                        st.session_state.tempo_pausado_segundos = None
                        st.session_state.timer_rodando = True
                        st.session_state.turno_contador += 1
                        if sala_id:
                            fb_patch(sala_id, {
                                "timer_iniciado_em": st.session_state.turno_iniciado_em
                            })
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
                st.markdown(
                    "<div style='text-align:center;color:#22c55e;"
                    "font-weight:bold;padding-top:8px;font-size:15px;'>"
                    "🟢 Cronômetro rodando</div>",
                    unsafe_allow_html=True,
                )
            elif pausado_seg is not None:
                st.markdown(
                    "<div style='text-align:center;color:#eab308;"
                    "font-weight:bold;padding-top:8px;font-size:15px;'>"
                    "⏸️ Pausado</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"<div style='text-align:center;color:#64748b;"
                    f"font-weight:bold;padding-top:8px;font-size:15px;'>"
                    f"⚪ Aguardando início ({tempo_total // 60:02d}:{tempo_total % 60:02d})</div>",
                    unsafe_allow_html=True,
                )
        else:
            st.info("Configure minutos ou segundos na barra lateral.")
    else:
        st.info("⏱️ Cronômetro desativado.")

    # ---------- QUADRANTE C: CARTAS FALTANTES ----------
    st.markdown(
        """
        <div style="font-size:14px; color:#64748b;
                    text-transform:uppercase; letter-spacing:3px;
                    text-align:center; margin-top:24px; margin-bottom:10px;
                    font-weight:bold;">
            📊 Cartas Faltantes
        </div>
        """,
        unsafe_allow_html=True,
    )

    faltantes = len(deck)
    pct_restante = (faltantes / total_cartas) * 100 if total_cartas > 0 else 0

    if faltantes == 0:
        cor_f = "#22c55e"
        icone_f = "🏁"
        label_f = "Baralho vazio"
    elif pct_restante > 50:
        cor_f = "#22c55e"
        icone_f = "🟢"
        label_f = "Ainda tem muito jogo"
    elif pct_restante > 20:
        cor_f = "#eab308"
        icone_f = "🟡"
        label_f = "Reta final chegando"
    else:
        cor_f = "#dc2626"
        icone_f = "🔴"
        label_f = "Últimas cartas!"

    sorteadas_qtd = len(st.session_state.sorteadas)

    st.markdown(
        f"""
        <div style="padding:24px 20px; border-radius:16px;
                    background:{cor_f}15; border:4px solid {cor_f};
                    text-align:center; font-family:system-ui;
                    box-shadow: 0 3px 12px {cor_f}22;">
            <div style="font-size:90px; font-weight:900; color:{cor_f};
                        line-height:1; font-variant-numeric:tabular-nums;
                        letter-spacing:2px;
                        text-shadow: 0 3px 14px {cor_f}22;">
                {faltantes}
            </div>
            <div style="font-size:13px; color:#64748b;
                        text-transform:uppercase; letter-spacing:2px;
                        margin-top:8px; font-weight:bold;">
                {icone_f} cartas no baralho
            </div>
            <div style="width:100%; height:14px; background:#e2e8f0;
                        border-radius:7px; overflow:hidden; margin-top:16px;">
                <div style="width:{pct_restante}%; height:100%; background:{cor_f};"></div>
            </div>
            <div style="font-size:12px; color:#94a3b8; margin-top:10px;">
                {sorteadas_qtd} de {total_cartas} já sorteadas — {label_f}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =====================================================
# RODAPÉ
# =====================================================
st.divider()
st.caption(
    "🎲 **Entre Linhas** — cada rodada, o tirador vê a carta secretamente no celular "
    "e decide **Acertou** ou **Errou**. Vence quem conquistar mais cartas."
)
