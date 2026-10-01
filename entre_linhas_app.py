"""
Entre Linhas — App Multiplayer com Supabase
- Dois celulares, dois times
- Código de sala + QR code
- Timer configurável com som de contagem
- Cartas secretas por turno
- Contador de cartas restantes
"""

import io
import random
import string
from datetime import datetime, timezone

import streamlit as st
import streamlit.components.v1 as components
from streamlit_autorefresh import st_autorefresh
from supabase import create_client
import qrcode


# =====================================================
# CONFIGURAÇÃO DA PÁGINA
# =====================================================
st.set_page_config(
    page_title="Entre Linhas — Multiplayer",
    page_icon="🎲",
    layout="wide",
)

SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
APP_URL = st.secrets.get("APP_URL", "http://localhost:8501")


@st.cache_resource
def get_supabase():
    return create_client(SUPABASE_URL, SUPABASE_KEY)


supabase = get_supabase()


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
# UTILITÁRIOS
# =====================================================
def gerar_codigo_sala():
    chars = string.ascii_uppercase + string.digits
    return "".join(random.choices(chars, k=6))


def criar_sala(tamanho, nome_time_1, nome_time_2, tempo_limite_segundos=0):
    codigo = gerar_codigo_sala()
    sorteadas = random.sample(PALAVRAS, tamanho * 2)
    linhas = sorteadas[:tamanho]
    colunas = sorteadas[tamanho:]

    coords = [f"{chr(64 + j)}{i}" for i in range(1, tamanho + 1) for j in range(1, tamanho + 1)]
    random.shuffle(coords)

    dados = {
        "codigo": codigo,
        "tamanho": tamanho,
        "linhas": linhas,
        "colunas": colunas,
        "deck": coords,
        "sorteadas": [],
        "turno": 1,
        "nome_time_1": nome_time_1,
        "nome_time_2": nome_time_2,
        "fase": "aguardando",
        "tempo_limite_segundos": tempo_limite_segundos,
        "turno_iniciado_em": datetime.now(timezone.utc).isoformat(),
    }
    supabase.table("salas").insert(dados).execute()
    return codigo


def buscar_sala(codigo):
    resp = supabase.table("salas").select("*").eq("codigo", codigo).execute()
    return resp.data[0] if resp.data else None


def atualizar_sala(codigo, campos):
    campos["atualizada_em"] = datetime.now(timezone.utc).isoformat()
    supabase.table("salas").update(campos).eq("codigo", codigo).execute()


def registrar_entrada(codigo, time):
    campo = f"jogador_{time}_presente"
    supabase.table("salas").update({campo: True}).eq("codigo", codigo).execute()


def registrar_carta_sorteada(codigo, coord, time):
    supabase.table("cartas_sorteadas").insert({
        "sala_codigo": codigo,
        "coord": coord,
        "time": time,
    }).execute()


def gerar_qr_code(url):
    qr = qrcode.QRCode(version=1, box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


def formatar_tempo(segundos):
    if segundos < 0:
        segundos = 0
    return f"{segundos // 60:02d}:{segundos % 60:02d}"


def calcular_tempo_restante(sala):
    total = sala.get("tempo_limite_segundos", 0) or 0
    if total == 0:
        return None, 0, 0
    iniciado = sala.get("turno_iniciado_em")
    if not iniciado:
        return total, total, 100
    try:
        inicio = datetime.fromisoformat(iniciado.replace("Z", "+00:00"))
        agora = datetime.now(timezone.utc)
        decorrido = int((agora - inicio).total_seconds())
        restante = max(0, total - decorrido)
        pct = (restante / total) * 100 if total else 0
        return restante, total, pct
    except Exception:
        return total, total, 100


# =====================================================
# COMPONENTES VISUAIS
# =====================================================
def render_cronometro(sala, cor_turno, nome_turno):
    restante, total, pct = calcular_tempo_restante(sala)
    if total == 0:
        return

    if pct > 50:
        cor, icone = "#22c55e", "🟢"
    elif pct > 20:
        cor, icone = "#eab308", "🟡"
    else:
        cor, icone = "#dc2626", "🔴"

    esgotado = restante == 0

    html = f"""
    <div style="
        padding:14px 18px; background:{cor}15;
        border:2px solid {cor}; border-radius:12px;
        margin:10px 0 16px 0; font-family:system-ui;">
        <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
            <div style="font-size:13px;color:#64748b;">{icone} Tempo de <b>{nome_turno}</b></div>
            <div style="font-size:32px; font-weight:bold; color:{cor};
                        font-variant-numeric: tabular-nums; line-height:1;">
                {formatar_tempo(restante)}
            </div>
        </div>
        <div style="width:100%; height:8px; background:#e2e8f0; border-radius:4px; overflow:hidden;">
            <div style="width:{pct}%; height:100%; background:{cor}; transition: width 0.5s linear;"></div>
        </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

    if esgotado:
        st.warning(f"⏰ **Tempo esgotado!** {nome_turno} deve concluir o turno imediatamente.")


def render_som_contagem(sala, mudo=False):
    total = sala.get("tempo_limite_segundos", 0) or 0
    if total == 0 or mudo:
        return
    iniciado = sala.get("turno_iniciado_em")
    if not iniciado:
        return

    html = f"""
    <!DOCTYPE html><html><head><meta charset="utf-8"></head><body>
    <script>
    (function() {{
        const totalSegundos = {total};
        const inicioMs = new Date("{iniciado}").getTime();
        const beepKey = "beeps_" + inicioMs;
        let beepedSeconds;
        try {{ beepedSeconds = JSON.parse(sessionStorage.getItem(beepKey) || "[]"); }}
        catch(e) {{ beepedSeconds = []; }}

        let ctx = null;
        function getCtx() {{
            if (!ctx) {{ try {{ ctx = new (window.AudioContext || window.webkitAudioContext)(); }}
                        catch(e) {{ return null; }} }}
            if (ctx && ctx.state === 'suspended') ctx.resume();
            return ctx;
        }}
        function beep(freq, dur, vol) {{
            const c = getCtx(); if (!c) return;
            try {{
                const osc = c.createOscillator();
                const gain = c.createGain();
                osc.type = 'sine';
                osc.frequency.value = freq;
                gain.gain.setValueAtTime(vol, c.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.001, c.currentTime + dur);
                osc.connect(gain); gain.connect(c.destination);
                osc.start(c.currentTime); osc.stop(c.currentTime + dur);
            }} catch(e) {{}}
        }}
        function tocarSom(restante) {{
            if (restante <= 0) beep(220, 0.8, 0.30);
            else if (restante <= 3) beep(880, 0.15, 0.25);
            else if (restante <= 10) beep(440, 0.10, 0.15);
        }}
        function tick() {{
            const decorrido = (Date.now() - inicioMs) / 1000;
            const restante = Math.ceil(totalSegundos - decorrido);
            if (restante <= 0) {{
                if (!beepedSeconds.includes(0)) {{
                    tocarSom(0); beepedSeconds.push(0);
                    try {{ sessionStorage.setItem(beepKey, JSON.stringify(beepedSeconds)); }} catch(e) {{}}
                }}
                return;
            }}
            if (restante <= 10 && !beepedSeconds.includes(restante)) {{
                tocarSom(restante); beepedSeconds.push(restante);
                try {{ sessionStorage.setItem(beepKey, JSON.stringify(beepedSeconds)); }} catch(e) {{}}
            }}
            setTimeout(tick, 250);
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
        setTimeout(tick, 300);
    }})();
    </script></body></html>
    """
    components.html(html, height=0, scrolling=False)


# =====================================================
# ESTADO DA SESSÃO
# =====================================================
if "sala_codigo" not in st.session_state:
    st.session_state.sala_codigo = None
if "meu_time" not in st.session_state:
    st.session_state.meu_time = None
if "som_mudo" not in st.session_state:
    st.session_state.som_mudo = False


# Parâmetro ?sala=XXXXXX da URL (QR code)
qp = st.query_params
if "sala" in qp and st.session_state.sala_codigo is None:
    st.session_state.sala_codigo = qp["sala"].upper()


# =====================================================
# TELA 1 — HOME
# =====================================================
def tela_home():
    st.title("🎲 Entre Linhas — Multiplayer")
    st.caption("Jogue com dois times em celulares separados.")

    tab_criar, tab_entrar = st.tabs(["🆕 Criar sala", "🔑 Entrar em sala"])

    with tab_criar:
        col1, col2 = st.columns(2)
        with col1:
            tamanho = st.select_slider("Tamanho da grade", options=[3, 4, 5, 6], value=5)
        with col2:
            nome_t1 = st.text_input("🔵 Nome do Time 1", value="Time 1", max_chars=20)
            nome_t2 = st.text_input("🔴 Nome do Time 2", value="Time 2", max_chars=20)

        st.markdown("##### ⏱️ Tempo por turno (opcional)")
        col_m, col_s, col_off = st.columns([2, 2, 3])
        with col_m:
            minutos = st.number_input("Minutos", min_value=0, max_value=60, value=2, step=1)
        with col_s:
            segundos = st.number_input("Segundos", min_value=0, max_value=59, value=0, step=5)
        with col_off:
            st.write(""); st.write("")
            sem_timer = st.checkbox("🚫 Sem cronômetro", value=False)

        tempo_total = 0 if sem_timer else (minutos * 60 + segundos)

        if tempo_total > 0:
            st.caption(f"⏱️ Cada turno terá **{minutos:02d}:{segundos:02d}**.")
        elif not sem_timer:
            st.caption("⚠️ Tempo zero. Selecione minutos/segundos ou marque **Sem cronômetro**.")
        else:
            st.caption("🚫 Cronômetro desativado — sem limite de tempo.")

        if st.button("🎲 Criar sala", type="primary", use_container_width=True):
            if not sem_timer and tempo_total == 0:
                st.error("Configure um tempo válido ou marque 'Sem cronômetro'.")
            else:
                codigo = criar_sala(tamanho, nome_t1 or "Time 1", nome_t2 or "Time 2", tempo_total)
                st.session_state.sala_codigo = codigo
                st.rerun()

    with tab_entrar:
        codigo_input = st.text_input(
            "Código da sala (6 caracteres)", max_chars=6, placeholder="Ex.: ABC123"
        ).upper().strip()
        if st.button("🔑 Entrar", type="primary", use_container_width=True):
            if not codigo_input:
                st.error("Digite o código da sala.")
            elif not buscar_sala(codigo_input):
                st.error("Sala não encontrada.")
            else:
                st.session_state.sala_codigo = codigo_input
                st.rerun()


# =====================================================
# TELA 2 — LOBBY
# =====================================================
def tela_lobby(sala):
    st.title(f"🎯 Sala **{sala['codigo']}**")

    if st.session_state.meu_time is None:
        col_qr, col_escolha = st.columns([1, 1])

        with col_qr:
            st.subheader("📱 Compartilhe o QR")
            url_join = f"{APP_URL}?sala={sala['codigo']}"
            st.image(gerar_qr_code(url_join), caption=f"Ou entre com o código **{sala['codigo']}**")
            st.code(sala["codigo"], language=None)

        with col_escolha:
            st.subheader("👥 Escolha seu time")
            pode_1 = not sala.get("jogador_1_presente")
            pode_2 = not sala.get("jogador_2_presente")

            if st.button(f"🔵 Entrar como {sala['nome_time_1']}",
                         disabled=not pode_1, use_container_width=True, type="primary"):
                registrar_entrada(sala["codigo"], 1)
                st.session_state.meu_time = 1
                st.rerun()

            if st.button(f"🔴 Entrar como {sala['nome_time_2']}",
                         disabled=not pode_2, use_container_width=True, type="primary"):
                registrar_entrada(sala["codigo"], 2)
                st.session_state.meu_time = 2
                st.rerun()

            st.caption("Cada time deve entrar em um dispositivo diferente.")
    else:
        st.success(f"✅ Você entrou como **{sala[f'nome_time_{st.session_state.meu_time}']}**")

        pronto_1 = sala.get("jogador_1_presente")
        pronto_2 = sala.get("jogador_2_presente")

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"🔵 **{sala['nome_time_1']}** — " + ("✅ Conectado" if pronto_1 else "⏳ Aguardando..."))
        with col2:
            st.markdown(f"🔴 **{sala['nome_time_2']}** — " + ("✅ Conectado" if pronto_2 else "⏳ Aguardando..."))

        if pronto_1 and pronto_2:
            st.success("🎉 Todos conectados! Iniciando jogo...")
            if st.button("▶️ Começar partida", type="primary", use_container_width=True):
                atualizar_sala(sala["codigo"], {
                    "fase": "jogo",
                    "turno": 1,
                    "turno_iniciado_em": datetime.now(timezone.utc).isoformat(),
                })
                st.rerun()
        else:
            st.info("Aguardando o outro time entrar. A página atualiza automaticamente.")
            st_autorefresh(interval=3000, key="lobby_refresh")


# =====================================================
# TELA 3 — JOGO
# =====================================================
def tela_jogo(sala):
    tam = sala["tamanho"]
    total_cartas = tam * tam
    deck = sala["deck"]
    sorteadas = sala["sorteadas"]
    restantes = len(deck)
    turno = sala["turno"]
    meu_time = st.session_state.meu_time

    nome_t1 = sala["nome_time_1"]
    nome_t2 = sala["nome_time_2"]
    nome_turno = nome_t1 if turno == 1 else nome_t2
    nome_eu = nome_t1 if meu_time == 1 else nome_t2
    cor_turno = "#1e3c78" if turno == 1 else "#dc2626"

    # Cabeçalho
    st.markdown(
        f"""
        <div style="padding:10px 16px; background:{cor_turno}15;
                    border-left:6px solid {cor_turno}; border-radius:8px;
                    margin-bottom:12px; font-family:system-ui;">
            <div style="font-size:12px;color:#64748b;">VOCÊ É</div>
            <div style="font-size:18px;font-weight:bold;">{nome_eu}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Botão de mudo
    if sala.get("tempo_limite_segundos", 0) > 0:
        col_mudo, _ = st.columns([1, 5])
        with col_mudo:
            icone = "🔇 Som off" if st.session_state.som_mudo else "🔊 Som on"
            if st.button(icone, use_container_width=True, key="btn_mudo"):
                st.session_state.som_mudo = not st.session_state.som_mudo
                st.rerun()

    # Contadores
    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Cartas restantes", restantes)
    col_b.metric("Já sorteadas", len(sorteadas))
    col_c.metric("Total", total_cartas)

    st.progress(len(sorteadas) / total_cartas if total_cartas else 0)

    # Cronômetro + Som
    render_cronometro(sala, cor_turno, nome_turno)
    render_som_contagem(sala, mudo=st.session_state.som_mudo)

    st.divider()

    # Painel de turno
    if turno == meu_time:
        st.markdown(f"<h3 style='color:{cor_turno};'>🎴 Sua vez, {nome_eu}!</h3>",
                    unsafe_allow_html=True)

        if sala["fase"] == "aguardando":
            st.caption("👉 Peça para o adversário olhar para o outro lado antes de clicar.")

            if st.button("🎴 Sortear minha carta", type="primary", use_container_width=True):
                if not deck:
                    st.warning("Baralho vazio!")
                else:
                    nova = deck[0]
                    atualizar_sala(sala["codigo"], {
                        "deck": deck[1:],
                        "carta_atual": nova,
                        "fase": "mostrando",
                        "turno_iniciado_em": datetime.now(timezone.utc).isoformat(),
                    })
                    st.rerun()

            if sala.get("tempo_limite_segundos", 0) > 0:
                if st.button("🔄 Reiniciar cronômetro", use_container_width=False):
                    atualizar_sala(sala["codigo"], {
                        "turno_iniciado_em": datetime.now(timezone.utc).isoformat(),
                    })
                    st.rerun()

        elif sala["fase"] == "mostrando":
            st.info("👀 Sua carta está visível. Memorize antes de fechar.")
            coord = sala["carta_atual"]

            st.markdown(
                f"""
                <div style="text-align:center; padding:40px; background:#d9e1f2;
                            border:4px solid #1e3c78; border-radius:16px;
                            margin:20px 0; font-family:system-ui;">
                    <div style="font-size:14px;color:#64748b;margin-bottom:8px;">Sua carta secreta</div>
                    <div style="font-size:120px;font-weight:bold;color:#1e3c78;letter-spacing:8px;">
                        {coord}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ Memorizei", type="primary", use_container_width=True):
                    atualizar_sala(sala["codigo"], {"fase": "jogando"})
                    st.rerun()
            with col2:
                if st.button("🔁 Trocar carta", use_container_width=True):
                    novo_deck = deck + [sala["carta_atual"]]
                    random.shuffle(novo_deck)
                    atualizar_sala(sala["codigo"], {
                        "deck": novo_deck[1:],
                        "carta_atual": novo_deck[0],
                    })
                    st.rerun()

        elif sala["fase"] == "jogando":
            st.markdown(
                """
                <div style="padding:20px;background:#f1f5f9;border-radius:12px;
                            text-align:center;font-family:system-ui;">
                    <div style="font-size:14px;color:#64748b;">Sua carta (já memorizada)</div>
                    <div style="font-size:48px;font-weight:bold;color:#64748b;">❓</div>
                    <div style="font-size:13px;color:#94a3b8;margin-top:8px;">
                        Dê a dica e aguarde o adversário adivinhar
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if st.button("✅ Concluir turno", type="primary", use_container_width=True):
                registrar_carta_sorteada(sala["codigo"], sala["carta_atual"], turno)
                atualizar_sala(sala["codigo"], {
                    "sorteadas": sorteadas + [sala["carta_atual"]],
                    "carta_atual": None,
                    "fase": "aguardando",
                    "turno": 2 if turno == 1 else 1,
                    "turno_iniciado_em": datetime.now(timezone.utc).isoformat(),
                })
                st.rerun()

    else:
        st.markdown(
            f"""
            <div style="padding:40px 20px; text-align:center;
                        background:#f8fafc; border:2px dashed #cbd5e1;
                        border-radius:12px; font-family:system-ui;">
                <div style="font-size:48px;">⏳</div>
                <div style="font-size:18px;font-weight:bold;color:#64748b;margin-top:12px;">
                    Aguardando o turno de {nome_turno}
                </div>
                <div style="font-size:14px;color:#94a3b8;margin-top:8px;">
                    Olhe para o outro lado enquanto {nome_turno} vê a carta
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption("🔄 A tela atualiza automaticamente a cada 2 segundos.")

    # Fim de jogo
    if restantes == 0 and sala["carta_atual"] is None:
        st.success(f"🏁 **Fim de jogo!** Todas as {total_cartas} cartas saíram.")
        st.balloons()

        if st.button("🔄 Nova partida (mesma sala)", use_container_width=True):
            coords = [f"{chr(64+j)}{i}" for i in range(1, tam+1) for j in range(1, tam+1)]
            random.shuffle(coords)
            atualizar_sala(sala["codigo"], {
                "deck": coords,
                "sorteadas": [],
                "carta_atual": None,
                "fase": "aguardando",
                "turno": 1,
                "turno_iniciado_em": datetime.now(timezone.utc).isoformat(),
            })
            st.rerun()

    # Histórico
    with st.expander(f"📜 Histórico de cartas sorteadas ({len(sorteadas)})"):
        if sorteadas:
            for idx, c in enumerate(sorteadas, 1):
                st.markdown(f"`{idx}.` **{c}**")
        else:
            st.caption("Nenhuma carta sorteada ainda.")

    # Sair
    st.divider()
    if st.button("🚪 Sair da sala", use_container_width=False):
        st.session_state.sala_codigo = None
        st.session_state.meu_time = None
        st.query_params.clear()
        st.rerun()

    st_autorefresh(interval=2000, key="jogo_refresh")


# =====================================================
# ROTEADOR PRINCIPAL
# =====================================================
if st.session_state.sala_codigo is None:
    tela_home()
else:
    sala = buscar_sala(st.session_state.sala_codigo)

    if not sala:
        st.error("❌ Sala não encontrada. Talvez tenha sido encerrada.")
        if st.button("🏠 Voltar para o início"):
            st.session_state.sala_codigo = None
            st.session_state.meu_time = None
            st.query_params.clear()
            st.rerun()
    elif sala["fase"] == "aguardando" and st.session_state.meu_time is None:
        tela_lobby(sala)
    elif sala["fase"] == "aguardando":
        tela_lobby(sala)
    else:
        tela_jogo(sala)
