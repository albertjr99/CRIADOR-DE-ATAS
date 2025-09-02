# Atas.py - Versão Simplificada para Deploy
from fastapi import FastAPI, Request, Form, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.responses import Response
from jinja2 import Environment, DictLoader, select_autoescape
from datetime import date, timedelta, datetime
from markupsafe import Markup
import io, sqlite3, os, json, logging

# Configurações
APP_TITLE = "Sistema de Atas — Comitê de Investimentos"
VERSION = "2.0.0"

PARTICIPANTES = [
    "Albert Iglésia Correa dos Santos Júnior",
    "Lucas José das Neves Rodrigues", 
    "Mariana Schneider Viana",
    "Shirlene Pires Mesquita",
    "Tatiana Gasparini Silva Stelzer",
]
CARGO = "Membro do Comitê de Investimentos"

PARTICIPANTES_MULHERES = {
    "Shirlene Pires Mesquita",
    "Mariana Schneider Viana", 
    "Tatiana Gasparini Silva Stelzer",
}

DB_PATH = os.getenv("DATABASE_URL", "./atas.db")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")

TOPICOS = [
    ("CHINA", "CHINA"),
    ("ESTADOS UNIDOS", "ESTADOS UNIDOS"),
    ("EUROPA", "EUROPA"),
    ("CENÁRIO POLÍTICO BRASILEIRO", "CENÁRIO POLÍTICO BRASILEIRO"),
    ("CENÁRIO ECONÔMICO BRASILEIRO", "CENÁRIO ECONÔMICO BRASILEIRO"),
]

# Funções auxiliares
def ptbr_date(d):
    return d.strftime("%d/%m/%Y")

def mes_ano_pt(d):
    nomes = {1:"janeiro",2:"fevereiro",3:"março",4:"abril",5:"maio",6:"junho",
             7:"julho",8:"agosto",9:"setembro",10:"outubro",11:"novembro",12:"dezembro"}
    return f"{nomes[d.month]} de {d.year}"

def menos_dois_meses(d):
    m = d.month - 2
    y = d.year
    if m <= 0:
        m += 12
        y -= 1
    dia = min(d.day, 28)
    return date(y, m, dia)

def primeira_quinta(ano, mes):
    d = date(ano, mes, 1)
    offset = (3 - d.weekday()) % 7
    return d + timedelta(days=offset)

# Database
def db_connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def ensure_schema():
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("""
        CREATE TABLE IF NOT EXISTS meeting (
            id INTEGER PRIMARY KEY,
            numero INTEGER NOT NULL,
            ano INTEGER NOT NULL,
            data TEXT NOT NULL,
            hora TEXT NOT NULL,
            local TEXT NOT NULL,
            lavrador TEXT DEFAULT ''
        );
    """)
    
    cur.execute("""
        CREATE TABLE IF NOT EXISTS scenario (
            id INTEGER PRIMARY KEY,
            participant TEXT NOT NULL,
            text TEXT NOT NULL DEFAULT '',
            topic TEXT NOT NULL DEFAULT '',
            UNIQUE(participant)
        );
    """)
    
    cur.execute("""
        CREATE TABLE IF NOT EXISTS system_config (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
    """)

    # Dados iniciais
    cur.execute("SELECT COUNT(1) FROM meeting")
    if cur.fetchone()[0] == 0:
        hoje = date.today()
        pq = primeira_quinta(hoje.year, hoje.month)
        cur.execute(
            "INSERT INTO meeting (numero, ano, data, hora, local, lavrador) VALUES (?, ?, ?, ?, ?, ?)",
            (4, hoje.year, pq.isoformat(), "14:00", "Sala nº 408 do 4º andar do IPAJM", "")
        )
    
    cur.execute("SELECT COUNT(1) FROM scenario")
    if cur.fetchone()[0] == 0:
        for p in PARTICIPANTES:
            cur.execute("INSERT OR IGNORE INTO scenario (participant, text, topic) VALUES (?, '', '')", (p,))
    
    # Configurações padrão
    defaults = [
        ("item2_default", "Não houve realocações de recursos desde a última reunião até a presente data."),
        ("assuntos_default", "– Assuntos gerais discutidos e/ou Eventos:")
    ]
    
    for key, value in defaults:
        cur.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES (?, ?)", (key, value))
    
    conn.commit()
    conn.close()

# Estado global simples
STATE = {
    "item2": "Não houve realocações de recursos desde a última reunião até a presente data.",
    "assuntos": "– Assuntos gerais discutidos e/ou Eventos:",
    "resumo": {"rentab": "", "difpp": "", "posicao": "abaixo", "risco": ""}
}

# WebSocket Manager
class ConnectionManager:
    def __init__(self):
        self.connections = set()

    async def connect(self, websocket):
        await websocket.accept()
        self.connections.add(websocket)

    def disconnect(self, websocket):
        self.connections.discard(websocket)

manager = ConnectionManager()

# Templates HTML simples
TEMPLATES = {
    "base.html": """<!doctype html>
<html lang="pt-br">
<head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <title>{{ title }}</title>
    <script src="https://unpkg.com/htmx.org@1.9.12"></script>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-50">
    <header class="bg-white shadow border-b">
        <div class="max-w-6xl mx-auto px-4 py-4 flex items-center justify-between">
            <h1 class="text-xl font-bold">{{ title }}</h1>
            <div class="flex space-x-3">
                <button hx-get="/preview/modal" hx-target="body" hx-swap="beforeend" 
                        class="bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded">Prévia</button>
                <a href="/export/docx" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded">DOCX</a>
            </div>
        </div>
    </header>
    <main class="max-w-6xl mx-auto px-4 py-8">
        {% block content %}{% endblock %}
    </main>
</body>
</html>""",

    "index.html": """{% extends 'base.html' %}
{% block content %}
<div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
    <div class="space-y-6">
        <div class="bg-white rounded-lg shadow p-6">
            <h2 class="text-lg font-semibold mb-4">Configurações da Reunião</h2>
            <div id="session-block">{% include 'partials/session.html' %}</div>
        </div>

        <div class="bg-white rounded-lg shadow p-6">
            <h2 class="text-lg font-semibold mb-4">Item 01 — Cenário por Analista</h2>
            <div id="item1-form" class="mb-4">{% include 'partials/item1_form.html' %}</div>
            <div id="status-list">{% include 'partials/status.html' %}</div>
        </div>

        <div class="bg-white rounded-lg shadow p-6">
            <h2 class="text-lg font-semibold mb-4">Itens 02 e 04</h2>
            <form hx-post="/preview/update" class="space-y-4">
                <div>
                    <label class="block text-sm font-medium mb-2">Item 02 — Movimentações Financeiras</label>
                    <textarea name="item2" rows="3" class="w-full px-3 py-2 border rounded">{{ item2 }}</textarea>
                </div>
                <div>
                    <label class="block text-sm font-medium mb-2">Item 04 — Assuntos Gerais</label>
                    <textarea name="assuntos" rows="3" class="w-full px-3 py-2 border rounded">{{ assuntos }}</textarea>
                </div>
                <button type="submit" class="bg-green-600 text-white px-4 py-2 rounded">Salvar</button>
            </form>
        </div>
    </div>

    <div class="space-y-6">
        <div class="bg-white rounded-lg shadow p-6">
            <h2 class="text-lg font-semibold mb-4">Item 03 — Parâmetros</h2>
            <form hx-post="/resumo/update" class="space-y-4">
                <div>
                    <label class="block text-sm font-medium mb-2">Rentabilidade (%)</label>
                    <input type="text" name="rentab" class="w-full px-3 py-2 border rounded"/>
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block text-sm font-medium mb-2">Diferença</label>
                        <input type="text" name="difpp" class="w-full px-3 py-2 border rounded"/>
                    </div>
                    <div>
                        <label class="block text-sm font-medium mb-2">Posição</label>
                        <select name="posicao" class="w-full px-3 py-2 border rounded">
                            <option value="abaixo">Abaixo</option>
                            <option value="acima">Acima</option>
                        </select>
                    </div>
                </div>
                <div>
                    <label class="block text-sm font-medium mb-2">Risco (%)</label>
                    <input type="text" name="risco" class="w-full px-3 py-2 border rounded"/>
                </div>
                <button type="submit" class="w-full bg-orange-600 text-white px-4 py-2 rounded">Salvar</button>
            </form>
        </div>
    </div>
</div>
{% endblock %}""",

    "partials/session.html": """<form hx-post="/meeting/update" hx-target="#session-block" hx-swap="outerHTML" class="space-y-4">
    <div class="grid grid-cols-2 md:grid-cols-4 gap-3">
        <div>
            <label class="block text-sm mb-1">Número</label>
            <input type="number" name="numero" value="{{ meeting.numero }}" class="w-full px-3 py-2 border rounded"/>
        </div>
        <div>
            <label class="block text-sm mb-1">Ano</label>
            <input type="number" name="ano" value="{{ meeting.ano }}" class="w-full px-3 py-2 border rounded"/>
        </div>
        <div>
            <label class="block text-sm mb-1">Data</label>
            <input type="date" name="data" value="{{ meeting.data }}" class="w-full px-3 py-2 border rounded"/>
        </div>
        <div>
            <label class="block text-sm mb-1">Hora</label>
            <input type="time" name="hora" value="{{ meeting.hora }}" class="w-full px-3 py-2 border rounded"/>
        </div>
    </div>
    <div>
        <label class="block text-sm mb-1">Local</label>
        <input type="text" name="local" value="{{ meeting.local }}" class="w-full px-3 py-2 border rounded"/>
    </div>
    <div>
        <label class="block text-sm mb-1">Responsável</label>
        <select name="lavrador" class="w-full px-3 py-2 border rounded">
            <option value="">Selecionar</option>
            {% for p in participantes %}
            <option value="{{p}}" {{ 'selected' if meeting.lavrador==p else '' }}>{{p}}</option>
            {% endfor %}
        </select>
    </div>
    <button type="submit" class="bg-blue-600 text-white px-4 py-2 rounded">Salvar</button>
</form>""",

    "partials/item1_form.html": """<form hx-post="/scenario/save" hx-target="#status-list" hx-swap="outerHTML" class="space-y-4">
    <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
        <div>
            <label class="block text-sm mb-1">Analista</label>
            <select name="participant" class="w-full px-3 py-2 border rounded">
                {% for p in participantes %}
                <option value="{{p}}" {{ 'selected' if p==participant else '' }}>{{p}}</option>
                {% endfor %}
            </select>
        </div>
        <div>
            <label class="block text-sm mb-1">Tema</label>
            <select name="topic" class="w-full px-3 py-2 border rounded">
                {% for val,label in topicos %}
                <option value="{{val}}" {{ 'selected' if val==topic else '' }}>{{label}}</option>
                {% endfor %}
            </select>
        </div>
        <div class="flex items-end">
            <button type="submit" class="w-full bg-purple-600 text-white px-3 py-2 rounded">Salvar</button>
        </div>
    </div>
    <div>
        <label class="block text-sm mb-1">Texto do Cenário</label>
        <textarea name="text" rows="4" class="w-full px-3 py-2 border rounded">{{ text }}</textarea>
    </div>
</form>""",

    "partials/status.html": """<div id="status-list">
    <h3 class="font-medium mb-3">Status dos Participantes</h3>
    <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
        {% for p in participantes %}
        {% set data = scenarios.get(p, {}) %}
        {% set tem_texto = data.get('text','').strip()|length > 0 %}
        {% set tem_tema = data.get('topic','').strip()|length > 0 %}
        {% set completo = tem_texto and tem_tema %}
        
        <div class="bg-gray-50 rounded p-3 border-l-4 {{ 'border-green-500' if completo else 'border-yellow-400' if tem_texto or tem_tema else 'border-gray-300' }}">
            <div class="flex items-center justify-between">
                <div>
                    <h4 class="text-sm font-medium">{{ p }}</h4>
                    {% if data.get('topic') %}
                    <p class="text-xs text-gray-600">{{ data.get('topic') }}</p>
                    {% endif %}
                    <span class="text-xs px-2 py-1 rounded {{ 'bg-green-100 text-green-800' if completo else 'bg-yellow-100 text-yellow-800' if tem_texto or tem_tema else 'bg-gray-100 text-gray-800' }}">
                        {{ 'Completo' if completo else 'Parcial' if tem_texto or tem_tema else 'Pendente' }}
                    </span>
                </div>
                <button class="text-xs bg-white border px-2 py-1 rounded"
                        hx-get="/scenario/form?participant={{p|urlencode}}"
                        hx-target="#item1-form" hx-swap="outerHTML">
                    Editar
                </button>
            </div>
        </div>
        {% endfor %}
    </div>
</div>""",

    "partials/preview_modal.html": """<div id="preview-modal" class="fixed inset-0 z-50 bg-black bg-opacity-50 flex items-center justify-center p-4">
    <div class="bg-white rounded-lg w-full max-w-5xl max-h-[90vh] flex flex-col">
        <div class="flex items-center justify-between p-4 border-b">
            <h3 class="text-lg font-semibold">Prévia da Ata</h3>
            <button onclick="document.getElementById('preview-modal')?.remove()">×</button>
        </div>
        <div class="flex-1 overflow-y-auto p-6">
            <div class="prose max-w-none" style="font-family:Calibri,Arial;font-size:11pt;color:#000;text-align:justify">
                {{ ata_html | safe }}
            </div>
        </div>
        <div class="p-4 border-t flex justify-end space-x-2">
            <a href="/export/docx" class="bg-blue-600 text-white px-4 py-2 rounded">DOCX</a>
        </div>
    </div>
</div>"""
}

# Jinja2
env = Environment(loader=DictLoader(TEMPLATES), autoescape=select_autoescape(["html","xml"]))

def render(name, **ctx):
    return HTMLResponse(env.get_template(name).render(**ctx))

def nl2br(value):
    if not value:
        return Markup("")
    return Markup(value.replace("\n", "<br/>"))

env.filters["nl2br"] = nl2br

# FastAPI App
app = FastAPI(title=APP_TITLE, version=VERSION)

@app.on_event("startup")
async def startup():
    ensure_schema()

# WebSocket
@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        manager.disconnect(websocket)

# Rotas principais
@app.get("/", response_class=HTMLResponse)
async def index():
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    if meeting_row:
        meeting = dict(meeting_row)
    else:
        hoje = date.today()
        meeting = {
            "numero": 1, 
            "ano": hoje.year, 
            "data": hoje.isoformat(), 
            "hora": "14:00", 
            "local": "Sala nº 408 do 4º andar do IPAJM", 
            "lavrador": ""
        }
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()
    
    return render("index.html", 
                  title=APP_TITLE,
                  participantes=PARTICIPANTES,
                  topicos=TOPICOS,
                  meeting=meeting,
                  scenarios=scenarios,
                  participant=PARTICIPANTES[0],
                  topic=TOPICOS[0][0],
                  text="",
                  item2=STATE["item2"],
                  assuntos=STATE["assuntos"])

@app.post("/meeting/update")
async def meeting_update(
    numero: int = Form(...), 
    ano: int = Form(...), 
    data: str = Form(...),
    hora: str = Form(...), 
    local: str = Form(...), 
    lavrador: str = Form("")
):
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT id FROM meeting ORDER BY id DESC LIMIT 1")
    row = cur.fetchone()
    meeting_id = row[0] if row else None
    
    if meeting_id:
        cur.execute("UPDATE meeting SET numero=?, ano=?, data=?, hora=?, local=?, lavrador=? WHERE id=?", 
                    (numero, ano, data, hora, local, lavrador, meeting_id))
    else:
        cur.execute("INSERT INTO meeting (numero, ano, data, hora, local, lavrador) VALUES (?, ?, ?, ?, ?, ?)", 
                    (numero, ano, data, hora, local, lavrador))
        meeting_id = cur.lastrowid
    
    conn.commit()
    cur.execute("SELECT * FROM meeting WHERE id=?", (meeting_id,))
    meeting = dict(cur.fetchone())
    conn.close()
    
    html = env.get_template("partials/session.html").render(meeting=meeting, participantes=PARTICIPANTES)
    return HTMLResponse(f'<div id="session-block">{html}</div>')

@app.get("/scenario/form")
async def scenario_form(participant: str):
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("SELECT participant, text, topic FROM scenario WHERE participant=?", (participant,))
    row = cur.fetchone()
    conn.close()
    
    text = row["text"] if row else ""
    topic = row["topic"] if row else TOPICOS[0][0]
    
    html = env.get_template("partials/item1_form.html").render(
        participantes=PARTICIPANTES, topicos=TOPICOS, participant=participant, topic=topic, text=text)
    return HTMLResponse(html)

@app.post("/scenario/save")
async def scenario_save(participant: str = Form(...), topic: str = Form(...), text: str = Form(...)):
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("INSERT OR REPLACE INTO scenario (participant, text, topic) VALUES (?, ?, ?)", 
                (participant, text, topic))
    conn.commit()
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()
    
    html = env.get_template("partials/status.html").render(participantes=PARTICIPANTES, scenarios=scenarios)
    return HTMLResponse(html)

@app.post("/preview/update")
async def preview_update(item2: str = Form(...), assuntos: str = Form(...)):
    STATE["item2"] = item2
    STATE["assuntos"] = assuntos
    return Response(status_code=204)

@app.post("/resumo/update")
async def resumo_update(
    rentab: str = Form(""), 
    difpp: str = Form(""), 
    posicao: str = Form("abaixo"), 
    risco: str = Form("")
):
    STATE["resumo"] = {"rentab": rentab, "difpp": difpp, "posicao": posicao, "risco": risco}
    return Response(status_code=204)

# Funções de geração da ata
def _prefixo_genero(nome):
    return "A Sra." if nome in PARTICIPANTES_MULHERES else "O Sr."

def numero_sessao_fmt(numero, ano):
    return f"{numero:03d}/{ano}"

def item1_single_paragraph(meeting, scenarios):
    d = date.fromisoformat(meeting["data"])
    
    ordinais = {
        1: "primeiro", 2: "segundo", 3: "terceiro", 4: "quarto", 5: "quinto",
        6: "sexto", 7: "sétimo", 8: "oitavo", 9: "nono", 10: "décimo"
    }

    intro = f"No {ordinais.get(d.day, str(d.day))} dia do mês de {mes_ano_pt(d).split(' de ')[0]} do ano de {d.year}, às {meeting['hora']} horas, na {meeting['local']}, ocorreu a {meeting['numero']}ª Reunião Ordinária dos Membros do Comitê de Investimentos. "

    partes = [intro]
    for nome in PARTICIPANTES:
        row = scenarios.get(nome, {"text": "", "topic": ""})
        texto = row.get("text", "").strip()
        tema = row.get("topic", "").strip()
        if texto:
            partes.append(f"<strong>{_prefixo_genero(nome)} {nome}</strong> falando sobre {tema}, {texto} ")

    return "<p>" + "".join(partes).strip() + "</p>"

def item3_html(meeting):
    d = date.fromisoformat(meeting["data"])
    d2 = menos_dois_meses(d)
    mes_ano = mes_ano_pt(d2)
    
    # Dividir texto longo em partes menores
    parte1 = "O Comitê de Investimentos, buscando transmitir maior transparência em relação às análises dos investimentos do Instituto e, em consequência, aderindo às normas do Pró-Gestão, elabora o Relatório de Análise de Investimentos IPAJM."
    
    parte2 = "Este relatório já foi encaminhado à SCO – Subgerência de Contabilidade e Orçamento, para posterior envio para análise do Conselho Fiscal do IPAJM."
    
    parte3 = f"Segue abaixo um resumo relativo aos itens abordados no Relatório supracitado de {mes_ano}:"
    
    p0 = parte1 + " " + parte2 + " " + parte3
    
    resumo = STATE["resumo"]
    rentab = resumo["rentab"]
    difpp = resumo["difpp"] 
    pos = resumo["posicao"]
    risco = resumo["risco"]
    
    p1 = f"1) Acompanhamento da rentabilidade - A rentabilidade consolidada dos investimentos do Fundo Previdenciário em {mes_ano} foi de {rentab}, ficando {difpp} p.p. {pos} da meta atuarial."
    
    p2 = f"2) Avaliação de risco da carteira - O grau de variação nas rentabilidades está coerente com o grau de risco assumido, em {risco}."
    
    p3 = f"3) Execução da Política de Investimentos – As movimentações financeiras realizadas no mês de {mes_ano} estão de acordo com as deliberações estabelecidas com a Diretoria de Investimentos e com a legislação vigente."
    
    p4_parte1 = f"4) Aderência a Política de Investimentos - Os recursos investidos, abrangendo a carteira consolidada, que representa o patrimônio total do RPPS sob gestão, estão aderentes à Política de Investimentos de {meeting['ano']}, respeitando o estabelecido na legislação em vigor e dentro dos percentuais definidos."
    
    p4_parte2 = "Considerando que as taxas ainda são negociadas acima da meta atuarial, seguimos com a estratégia de alcançar o alvo definido de 60% de alocação em Títulos Públicos."
    
    p4 = p4_parte1 + " " + p4_parte2
    
    paragrafos = [p0, p1, p2, p3, p4]
    return "\n".join([f"<p><strong>{p}</strong></p>" if i==0 else f"<p>{p}</p>" for i, p in enumerate(paragrafos)])

def ata_html_full(meeting, scenarios):
    d = date.fromisoformat(meeting["data"])
    numero_fmt = numero_sessao_fmt(meeting["numero"], meeting["ano"])
    presencas_html = "<br/>".join([f"<strong>{p}</strong> - {CARGO};" for p in PARTICIPANTES])
    
    blocos = [
        f"<p><strong>Sessão Ordinária nº {numero_fmt}</strong></p>",
        f"<p><strong>Data:</strong> {ptbr_date(d)}.<br/><strong>Hora:</strong> {meeting['hora']}h.<br/><strong>Local:</strong> {meeting['local']}.</p>",
        "<p><strong>Presenças:</strong></p>",
        f"<p>{presencas_html}</p>",
        "<p><strong>Ordem do Dia:</strong></p>",
        "<p>1. <strong>Cenário Político e Econômico Interno</strong> e <strong>Cenário Econômico Externo (EUA, Europa e China)</strong>;<br/>2. <strong>Movimentações e Aplicações financeiras</strong>;<br/>3. <strong>Acompanhamento dos Recursos Investidos</strong>;<br/>4. <strong>Assuntos Gerais</strong>.</p>",
        "<p><strong>Item 01 – Cenário Político e Econômico Interno e Cenário Econômico Externo (EUA, Europa e China):</strong></p>",
        item1_single_paragraph(meeting, scenarios),
        "<p><strong>Item 02 – Movimentações e Aplicações financeiras</strong></p>",
        f"<p>{STATE['item2']}</p>",
        "<p><strong>Item 03 – Acompanhamento dos Recursos Investidos:</strong></p>",
        item3_html(meeting),
        "<p><strong>Item 04 – Assuntos Gerais</strong></p>",
        f"<p>{STATE['assuntos']}</p>",
        f"<p>Nada mais havendo a tratar, foi encerrada a reunião e eu, <strong>{meeting.get('lavrador') or '___________________________________'}</strong>, lavrei a presente Ata, assinada pelos membros presentes do Comitê de Investimentos.</p>",
        f"<p><strong>{PARTICIPANTES[0]}</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>{PARTICIPANTES[1]}</strong><br/>{CARGO}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{CARGO}</p>",
        f"<p><strong>{PARTICIPANTES[2]}</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>{PARTICIPANTES[3]}</strong><br/>{CARGO}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{CARGO}</p>",
        f"<p><strong>{PARTICIPANTES[4]}</strong><br/>{CARGO}</p>",
    ]
    
    header = """<div style="width:100%; text-align:center; margin-bottom:20px;"><div style="font-family: Calibri; font-size:11pt; font-weight:700; color:#000;">GOVERNO DO ESTADO DO ESPÍRITO SANTO<br/>INSTITUTO DE PREVIDÊNCIA DOS<br/>SERVIDORES DO ESTADO DO ESPÍRITO SANTO</div><div style="border-bottom:1px solid #000; margin:10px 0; text-align:center;">IPAJM</div></div>"""
    
    return header + "\n".join(blocos)

@app.get("/preview/modal")
async def preview_modal():
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    if meeting_row:
        meeting = dict(meeting_row)
    else:
        hoje = date.today()
        meeting = {"numero": 1, "ano": hoje.year, "data": hoje.isoformat(), "hora": "14:00", "local": "Sala nº 408", "lavrador": ""}
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()
    
    ata_html = ata_html_full(meeting, scenarios)
    
    html = env.get_template("partials/preview_modal.html").render(ata_html=ata_html)
    return HTMLResponse(html)

# Exportação DOCX simples
try:
    from docx import Document as DocxDocument
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    @app.get("/export/docx")
    async def export_docx():
        conn = db_connect()
        cur = conn.cursor()
        
        cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
        meeting_row = cur.fetchone()
        if meeting_row:
            meeting = dict(meeting_row)
        else:
            hoje = date.today()
            meeting = {"numero": 1, "ano": hoje.year, "data": hoje.isoformat(), "hora": "14:00", "local": "Sala nº 408", "lavrador": ""}
        
        cur.execute("SELECT participant, text, topic FROM scenario")
        scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
        conn.close()

        d = date.fromisoformat(meeting["data"])
        numero_fmt = numero_sessao_fmt(meeting["numero"], meeting["ano"])

        doc = DocxDocument()
        style = doc.styles["Normal"]
        style.font.name = "Calibri"
        style.font.size = Pt(11)

        # Título
        p = doc.add_paragraph()
        run = p.add_run(f"Sessão Ordinária nº {numero_fmt}")
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(11)

        # Data/Hora/Local
        p = doc.add_paragraph()
        p.add_run(f"Data: {ptbr_date(d)} • Hora: {meeting['hora']}h • Local: {meeting['local']}")

        # Salvar
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        
        filename = f"Ata_{meeting['numero']:03d}-{meeting['ano']}.docx"
        return StreamingResponse(
            buf, 
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
except ImportError:
    @app.get("/export/docx") 
    async def export_docx():
        return JSONResponse({"error": "DOCX export not available"}, status_code=503)

# Health check
@app.get("/health")
async def health():
    return JSONResponse({"status": "ok", "version": VERSION})

@app.get("/api/status")
async def api_status():
    return JSONResponse({"status": "online", "version": VERSION})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
