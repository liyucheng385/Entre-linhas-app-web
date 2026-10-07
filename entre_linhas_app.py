"""
Entre Linhas — Desktop + QR code para o tirador
- Dispositivo principal: grade + controle de rodadas
- Celular do tirador: mostra coord + Acertou/Errou
- Comunicação via Firebase Realtime Database
"""

import io
import time
import random
import string
import hashlib
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

FIREBASE_URL = st.secrets.get("FIREBASE_URL", "")
APP_URL = st.secrets.get("APP_URL", "http://localhost:8501")


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
# FIREBASE — API REST
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
# FUNÇÕES UTILITÁRIAS
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
# TIMER
# =====================================================
def render_timer(tempo_total, iniciado_em_ms, mudo=False, key="timer"):
    if tempo_total <= 0:
        return
    mudo_js = "true" if mudo else "false"
    html = f"""
    <!DOCTYPE html><html><head><style>
        body {{ margin: 0; font-family: system-ui, sans-serif; }}
        .timer-wrap {{ padding: 16px 20px; border-radius: 12px;
            background: #22c55e15; border: 2px solid #22c55e;
            text-align: center; transition: background 0.3s, border-color 0.3s; }}
        .timer-tempo {{ font-size: 56px; font-weight: bold;
            font-variant-numeric: tabular-nums; line-height: 1;
            margin-bottom: 6px; color: #22c55e; transition: color 0.3s; }}
        .timer-label {{ font-size: 12px; color: #64748b;
            text-transform: uppercase; letter-spacing: 1.5px; }}
        .timer-bar {{ width: 100%; height: 10px; background: #e2e8f0;
            border-radius: 5px; overflow: hidden; margin-top: 14px; }}
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
    components.html(html, height=190, scrolling=False)


# =====================================================
# GRID (somente leitura)
# =====================================================
def render_grid_estados(linhas, colunas, estados, nome_time_1, nome_time_2):
    tam = len(linhas)
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
    body {{ margin: 0; font-family: system-ui, sans-serif; }}
    .placar {{ display: flex; gap: 24px; align-items: center;
        padding: 12px 16px; background: #f8fafc; border: 2px solid #cbd5e1;
        border-radius: 10px; margin-bottom: 14px; font-size: 16px;
        font-weight: 600; flex-wrap: wrap; }}
    .placar-time {{ display: flex; align-items: center; gap: 8px; }}
    .placar-dot {{ width: 22px; height: 22px; border-radius: 4px;
        border: 2px solid #4a5568; flex-shrink: 0; }}
    .dot-b {{ background: #1e3c78; }} .dot-r {{ background: #dc2626; }}
    .dot-x {{ background: #94a3b8; }}
    .placar-nome {{ font-size: 15px; max-width: 180px; overflow: hidden;
        text-overflow: ellipsis; white-space: nowrap; }}
    .placar-num {{ font-size: 26px; min-width: 34px; text-align: center; font-weight: bold; }}
    .placar-num.blue {{ color: #1e3c78; }} .placar-num.red {{ color: #dc2626; }}
    .placar-num.grey {{ color: #64748b; }}
    .grid-el {{ display: grid;
        grid-template-columns: 45px 130px repeat({tam}, 130px); gap: 4px; }}
    .cel {{ border: 2px solid #4a5568; padding: 12px 8px; text-align: center;
        display: flex; align-items: center; justify-content: center;
        min-height: 45px; font-size: 13px; border-radius: 4px; user-select: none; }}
    .cel-coord {{ background: #f0f0f0; font-weight: bold; font-size: 16px; color: #323232; }}
    .cel-canto {{ background: #e6e6e6; font-weight: bold; font-size: 20px; }}
    .cel-col {{ background: #c5e0b4; font-weight: bold; }}
    .cel-lin {{ background: #ffe699; font-weight: bold; }}
    .cel-carta {{ font-weight: bold; font-size: 22px; min-height: 55px;
        transition: background 0.3s, color 0.3s; }}
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
<div class="grid-el">{header1}{header2}{rows}</div>
</body></html>
"""
    components.html(html, height=int(180 + tam * 105), scrolling=False)


# =====================================================
# INTERFACE MOBILE (celular do tirador)
# =====================================================
def render_mobile(sala_id):
    """Interface simplificada que o celular do tirador vê."""
    st.markdown("""
    <style>
        #MainMenu, header, footer { visibility: hidden; }
        .block-container { padding-top: 1rem; padding-bottom: 1rem; max-width: 500px; }
    </style>
    """, unsafe_allow_html=True)

    if not FIREBASE_URL:
        st.error("⚠️ Firebase não configurado. Configure o FIREBASE_URL em Secrets.")
        return

    # Polling via autorefresh
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

    # Cabeçalho
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

    # Estado: aguardando o desktop sortear
    if not coord or estado == "aguardando":
        st.info("⏳ Aguardando o dispositivo principal sortear uma carta...")
        st.caption("Não feche esta página.")
        return

    # Estado: carta sorteada, aguardando decisão
    if evento:
        # Já foi decidido — aguardando próxima
        st.success("✅ Decisão registrada!")
        st.markdown(
            f"""
            <div style="text-align:center; padding:40px 20px;
                        background:#f1f5f9; border-radius:12px;
                        font-family:system-ui;">
                <div style="font-size:14px;color:#64748b;">
                    Você marcou
                </div>
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

    # Estado normal: mostra coord e botões
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
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# =====================================================
# ROTEADOR: mobile ou desktop?
# =====================================================
qp = st.query_params
sala_qp = qp.get("sala")
role_qp = qp.get("role")

# ---------- MODO MOBILE ----------
if role_qp == "tirador" and sala_qp:
    st.title("🎴 Controle do Tirador")
    render_mobile(sala_qp.upper())
    st.stop()


# ---------- MODO DESKTOP ----------
st.title("🎲 Entre Linhas — Jogo Principal")
st.caption("Biblioteca com **343 palavras**. Sorteie a grade, exiba o QR code e jogue com 2 times.")


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
        cm, cr = st.columns(2)
        with cm:
            icone = "🔇 Som off" if st.session_state.timer_mudo else "🔊 Som on"
            if st.button(icone, use_container_width=True):
                st.session_state.timer_mudo = not st.session_state.timer_mudo
                st.rerun()
        with cr:
            if st.button("🔄 Resetar", use_container_width=True):
                st.session_state.turno_iniciado_em = _agora_ms()
                st.session_state.turno_contador += 1
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
        st.rerun()


# =====================================================
# LÓGICA DE SORTEIO DA GRADE
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

        # Reset local
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

        # Estado inicial no Firebase
        fb_put(sala_id, {
            "coord": None,
            "estado": "aguardando",
            "evento": None,
            "time_atual": 1,
            "nome_turno": st.session_state.nome_time_1,
            "criada_em": _agora_ms(),
        })

        st.success(f"✅ Grade **{tamanho}×{tamanho}** pronta! Sala **{sala_id}**.")


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


# =====================================================
# POLLING: verifica se o celular respondeu
# =====================================================
if sala_id and fase == "decidindo":
    st_autorefresh(interval=1500, key="desktop_poll")
    dados_fb = fb_get(sala_id)
    evento = dados_fb.get("evento")

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
        fb_patch(sala_id, {"evento": None, "coord": None, "estado": "aguardando"})
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
        fb_patch(sala_id, {"evento": None, "coord": None, "estado": "aguardando"})
        st.rerun()


# =====================================================
# QR CODE (para o tirador)
# =====================================================
if sala_id:
    with st.expander("📱 QR code do tirador (clique para expandir)", expanded=True):
        col_qr, col_info = st.columns([1, 2])

        with col_qr:
            url_mobile = f"{APP_URL}?sala={sala_id}&role=tirador"
            st.image(gerar_qr_code(url_mobile), width=220)

        with col_info:
            st.markdown(f"### Sala: `{sala_id}`")
            st.markdown(
                "**Como usar:**\n"
                "1. O tirador da rodada escaneia o QR code com o celular.\n"
                "2. A tela do celular mostrará **apenas a coordenada** e os botões "
                "**✅ Acertou** / **❌ Errou**.\n"
                "3. Quando o time responder, o tirador clica no botão correspondente.\n"
                "4. O jogo atualiza automaticamente aqui."
            )
            st.caption(f"Link direto: `{url_mobile}`")


# =====================================================
# INDICADORES
# =====================================================
col1, col2, col3 = st.columns(3)
col1.metric("Cartas restantes", len(deck))
col2.metric("Sorteadas", len(st.session_state.sorteadas))
col3.metric("Total", total_cartas)
if total_cartas > 0:
    st.progress(len(st.session_state.sorteadas) / total_cartas)


# =====================================================
# TIMER
# =====================================================
if st.session_state.timer_ativo:
    tempo_total = st.session_state.timer_minutos * 60 + st.session_state.timer_segundos
    if tempo_total > 0:
        render_timer(
            tempo_total=tempo_total,
            iniciado_em_ms=st.session_state.turno_iniciado_em,
            mudo=st.session_state.timer_mudo,
            key=f"turno_{st.session_state.turno_contador}",
        )


# =====================================================
# PAINEL DE RODADA
# =====================================================
st.divider()

if not deck and carta_atual is None:
    st.success(f"🏁 **Fim de jogo!** Todas as {total_cartas} cartas foram sorteadas.")
    st.balloons()
else:
    st.markdown(
        f"""
        <div style="padding:14px 18px; background:{cor_turno}15;
                    border-left:6px solid {cor_turno}; border-radius:10px;
                    font-family:system-ui; margin-bottom:14px;">
            <div style="font-size:12px;color:#64748b;
                        text-transform:uppercase;letter-spacing:1.5px;">
                Rodada do
            </div>
            <div style="font-size:24px;font-weight:bold;color:{cor_turno};">
                {nome_turno}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if fase == "aguardando":
        st.caption("👉 O tirador pega o celular e escaneia o QR code acima. "
                   "Depois clique abaixo para sortear a carta.")
        if st.button("🎴 Sortear carta secreta", type="primary", use_container_width=True):
            if not deck:
                st.warning("Baralho vazio!")
            else:
                nova = deck[0]
                st.session_state.carta_atual = nova
                st.session_state.fase = "decidindo"
                fb_patch(sala_id, {
                    "coord": nova,
                    "estado": "sorteada",
                    "evento": None,
                    "time_atual": time_atual,
                    "nome_turno": nome_turno,
                })
                st.rerun()

    elif fase == "decidindo":
        st.markdown(
            f"""
            <div style="padding:24px; background:#f1f5f9; border-radius:12px;
                        text-align:center; font-family:system-ui;">
                <div style="font-size:14px;color:#64748b;">
                    Carta em jogo — aguardando o tirador responder no celular
                </div>
                <div style="font-size:44px;font-weight:bold;color:#94a3b8;
                            margin-top:12px;">❓</div>
                <div style="font-size:13px;color:#94a3b8;margin-top:12px;">
                    🔄 Atualizando automaticamente a cada 1.5s...
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Botão de fallback (se o celular der problema)
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


# =====================================================
# GRADE
# =====================================================
st.divider()
st.subheader(f"🎯 Grade {tam}×{tam}")
st.caption("🔵 Time 1 · 🔴 Time 2 · ⚪ cinza = descartada")

render_grid_estados(linhas, colunas, estados, nome_t1, nome_t2)


# =====================================================
# HISTÓRICO
# =====================================================
with st.expander(f"📜 Histórico ({len(st.session_state.sorteadas)} cartas)"):
    if st.session_state.sorteadas:
        for idx, c in enumerate(st.session_state.sorteadas, 1):
            estado = estados.get(c, 0)
            tag = {1: "🔵 Time 1", 2: "🔴 Time 2", 3: "❌ Descartada"}.get(estado, "?")
            st.markdown(f"`{idx}.` **{c}** — {tag}")
    else:
        st.caption("Nenhuma carta sorteada ainda.")


# =====================================================
# RODAPÉ
# =====================================================
st.divider()
st.caption(
    "🎲 **Entre Linhas — Desktop + Celular do tirador** — "
    "cada rodada, o tirador vê a carta secretamente no celular e decide Acertou/Errou."
)
