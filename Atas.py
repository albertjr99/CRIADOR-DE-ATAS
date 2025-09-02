# Atas.py — Versão 2.0 Corrigida para Deploy
from fastapi import FastAPI, Request, Form, HTTPException, BackgroundTasks, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.responses import Response
from jinja2 import Environment, DictLoader, select_autoescape
from datetime import date, timedelta, datetime
from markupsafe import Markup
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io, sqlite3, os, json, asyncio, logging
from typing import Dict, Set, Optional, List
from pydantic import BaseModel, Field

# Exportadores
from docx import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from xhtml2pdf import pisa

# Configurações básicas
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

# Mulheres (para "A Sra.")
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

# Espaçamento (pt)
HEADER_GAP = 12
TITLE_GAP = 6  
SECTION_GAP = 12

# Models para validação
class MeetingUpdate(BaseModel):
    numero: int = Field(ge=1, le=9999)
    ano: int = Field(ge=2020, le=2050)
    data: str
    hora: str
    local: str = Field(min_length=1, max_length=200)
    lavrador: str = ""

class ScenarioUpdate(BaseModel):
    participant: str = Field(min_length=1)
    topic: str = Field(min_length=1) 
    text: str = Field(max_length=5000)

# Gerenciador de conexões WebSocket simplificado
class ConnectionManager:
    def __init__(self):
        self.active_connections: Dict[str, Set[WebSocket]] = {}
        self.user_sessions: Dict[WebSocket, str] = {}

    async def connect(self, websocket: WebSocket, user_id: str):
        await websocket.accept()
        if user_id not in self.active_connections:
            self.active_connections[user_id] = set()
        self.active_connections[user_id].add(websocket)
        self.user_sessions[websocket] = user_id

    def disconnect(self, websocket: WebSocket):
        if websocket in self.user_sessions:
            user_id = self.user_sessions[websocket]
            self.active_connections[user_id].discard(websocket)
            if not self.active_connections[user_id]:
                del self.active_connections[user_id]
            del self.user_sessions[websocket]

    async def broadcast_update(self, message: dict, exclude_user: str = None):
        message["timestamp"] = datetime.now().isoformat()
        for user_id, connections in self.active_connections.items():
            if user_id != exclude_user:
                for connection in connections:
                    try:
                        await connection.send_text(json.dumps(message))
                    except:
                        pass

manager = ConnectionManager()

# --------------------- Funções auxiliares
def ptbr_date(d: date) -> str:
    return d.strftime("%d/%m/%Y")

def mes_ano_pt(d: date) -> str:
    nomes = {1:"janeiro",2:"fevereiro",3:"março",4:"abril",5:"maio",6:"junho",
             7:"julho",8:"agosto",9:"setembro",10:"outubro",11:"novembro",12:"dezembro"}
    return f"{nomes[d.month]} de {d.year}"

def menos_dois_meses(d: date) -> date:
    m = d.month - 2; y = d.year
    if m <= 0: m += 12; y -= 1
    dia = min(d.day, 28)
    return date(y, m, dia)

def primeira_quinta(ano: int, mes: int) -> date:
    d = date(ano, mes, 1)
    offset = (3 - d.weekday()) % 7
    return d + timedelta(days=offset)

# --------------------- Database
def db_connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def ensure_schema():
    conn = db_connect()
    cur = conn.cursor()
    
    # Tabela principal de reuniões
    cur.execute("""
        CREATE TABLE IF NOT EXISTS meeting (
            id INTEGER PRIMARY KEY,
            numero INTEGER NOT NULL,
            ano INTEGER NOT NULL,
            data TEXT NOT NULL,
            hora TEXT NOT NULL,
            local TEXT NOT NULL,
            lavrador TEXT DEFAULT '',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
    """)
    
    # Tabela de cenários
    cur.execute("""
        CREATE TABLE IF NOT EXISTS scenario (
            id INTEGER PRIMARY KEY,
            participant TEXT NOT NULL,
            text TEXT NOT NULL DEFAULT '',
            topic TEXT NOT NULL DEFAULT '',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(participant)
        );
    """)
    
    # Tabela de configurações do sistema
    cur.execute("""
        CREATE TABLE IF NOT EXISTS system_config (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
    """)

    # Dados iniciais
    cur.execute("SELECT COUNT(1) FROM meeting")
    if cur.fetchone()[0] == 0:
        hoje = date.today()
        pq = primeira_quinta(hoje.year, hoje.month)
        cur.execute(
            "INSERT INTO meeting (numero, ano, data, hora, local, lavrador) VALUES (?, ?, ?, ?, ?, ?)",
            (4, hoje.year, pq.isoformat(), "14:00", "Sala nº 408 do 4º andar do IPAJM", ""),
        )
    
    cur.execute("SELECT COUNT(1) FROM scenario")
    if cur.fetchone()[0] == 0:
        for p in PARTICIPANTES:
            cur.execute("INSERT OR IGNORE INTO scenario (participant, text, topic) VALUES (?, '', '')", (p,))
    
    # Configurações padrão
    default_configs = [
        ("item2_default", "Não houve realocações de recursos desde a última reunião até a presente data."),
        ("assuntos_default", "– Assuntos gerais discutidos e/ou Eventos:"),
    ]
    
    for key, value in default_configs:
        cur.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES (?, ?)", (key, value))
    
    conn.commit()
    conn.close()

# Estado da aplicação simplificado
class AppState:
    def __init__(self):
        self.config = self.load_config()
        
    def load_config(self):
        try:
            conn = db_connect()
            cur = conn.cursor()
            cur.execute("SELECT key, value FROM system_config")
            config = dict(cur.fetchall())
            conn.close()
            return config
        except:
            return {
                "item2_default": "Não houve realocações de recursos desde a última reunião até a presente data.",
                "assuntos_default": "– Assuntos gerais discutidos e/ou Eventos:",
            }

state = AppState()

# Templates HTML básicos para funcionar
TEMPLATES = {
    "base.html": """<!doctype html>
<html lang="pt-br">
<head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <title>{{ title }}</title>
    <script src="https://unpkg.com/htmx.org@1.9.12"></script>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body { font-family: Inter, system-ui, -apple-system, sans-serif; }
        .prose.ata-preview { font-family: Calibri, Arial, sans-serif !important; font-size: 11pt !important; color: #000 !important; text-align: justify; line-height: 1.35; white-space: pre-wrap; }
    </style>
</head>
<body class="bg-gray-50 text-gray-900">
    <header class="bg-white shadow-sm border-b">
        <div class="max-w-6xl mx-auto px-4 py-4 flex items-center justify-between">
            <h1 class="text-xl font-bold text-gray-900">{{ title }}</h1>
            <div class="flex space-x-3">
                <button hx-get="/preview/modal" hx-target="body" hx-swap="beforeend" class="bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-lg">Prévia</button>
                <a href="/export/docx" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg">Exportar DOCX</a>
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
<div class="grid grid-cols-1 lg:grid-cols-2 gap-8">
    <div class="space-y-8">
        <!-- Configurações da Reunião -->
        <div class="bg-white rounded-xl shadow border p-6">
            <h2 class="text-lg font-semibold mb-4">Configurações da Reunião</h2>
            <div id="session-block">{% include 'partials/session.html' with context %}</div>
        </div>

        <!-- Item 01 - Cenários -->
        <div class="bg-white rounded-xl shadow border p-6">
            <h2 class="text-lg font-semibold mb-4">Item 01 — Cenário por Analista</h2>
            <div id="item1-form" class="mb-6">{% include 'partials/item1_form.html' with context %}</div>
            <div id="status-list">{% include 'partials/status.html' with context %}</div>
        </div>

        <!-- Itens 02 e 04 -->
        <div class="bg-white rounded-xl shadow border p-6">
            <h2 class="text-lg font-semibold mb-4">Itens 02 e 04 — Texto Livre</h2>
            <form hx-post="/preview/update" hx-target="this" hx-swap="none" class="space-y-4">
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-2">Item 02 — Movimentações e Aplicações Financeiras</label>
                    <textarea name="item2" rows="4" class="w-full px-3 py-2 border rounded-lg">{{ item2 }}</textarea>
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-2">Item 04 — Assuntos Gerais</label>
                    <textarea name="assuntos" rows="5" class="w-full px-3 py-2 border rounded-lg">{{ assuntos }}</textarea>
                </div>
                <button type="submit" class="bg-green-600 hover:bg-green-700 text-white px-6 py-2 rounded-lg">Salvar Itens 02 e 04</button>
            </form>
        </div>
    </div>

    <!-- Coluna lateral -->
    <div class="space-y-8">
        <!-- Item 03 - Parâmetros -->
        <div class="bg-white rounded-xl shadow border p-6">
            <h2 class="text-lg font-semibold mb-4">Item 03 — Parâmetros</h2>
            <form hx-post="/resumo/update" hx-target="this" hx-swap="none" class="space-y-4">
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-2">Rentabilidade (%)</label>
                    <input type="text" name="rentab" placeholder="ex: 1,23%" class="w-full px-3 py-2 border rounded-lg"/>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-2">Diferença (p.p.)</label>
                        <input type="text" name="difpp" placeholder="ex: 0,15" class="w-full px-3 py-2 border rounded-lg"/>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-2">Posição</label>
                        <select name="posicao" class="w-full px-3 py-2 border rounded-lg">
                            <option value="abaixo">Abaixo</option>
                            <option value="acima">Acima</option>
                        </select>
                    </div>
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-2">Risco (%)</label>
                    <input type="text" name="risco" placeholder="ex: 7,5%" class="w-full px-3 py-2 border rounded-lg"/>
                </div>
                <button type="submit" class="w-full bg-orange-600 hover:bg-orange-700 text-white px-4 py-2 rounded-lg">Salvar Parâmetros</button>
            </form>
        </div>
    </div>
</div>
{% endblock %}""",

    "partials/session.html": """<form hx-post="/meeting/update" hx-target="#session-block" hx-swap="outerHTML" class="space-y-4">
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
        <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Número</label>
            <input type="number" name="numero" value="{{ meeting.numero }}" min="1" class="w-full px-3 py-2 border rounded-lg"/>
        </div>
        <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Ano</label>
            <input type="number" name="ano" value="{{ meeting.ano }}" min="2020" class="w-full px-3 py-2 border rounded-lg"/>
        </div>
        <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Data</label>
            <input type="date" name="data" value="{{ meeting.data }}" class="w-full px-3 py-2 border rounded-lg"/>
        </div>
        <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Hora</label>
            <input type="time" name="hora" value="{{ meeting.hora }}" class="w-full px-3 py-2 border rounded-lg"/>
        </div>
    </div>
    <div>
        <label class="block text-sm font-medium text-gray-700 mb-1">Local</label>
        <input type="text" name="local" value="{{ meeting.local }}" class="w-full px-3 py-2 border rounded-lg"/>
    </div>
    <div>
        <label class="block text-sm font-medium text-gray-700 mb-1">Responsável pela ata</label>
        <select name="lavrador" class="w-full px-3 py-2 border rounded-lg">
            <option value="">Selecionar</option>
            {% for p in participantes %}
            <option value="{{p}}" {{ 'selected' if meeting.lavrador==p else '' }}>{{p}}</option>
            {% endfor %}
        </select>
    </div>
    <button type="submit" class="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-lg">Salvar Configurações</button>
</form>""",

    "partials/item1_form.html": """<form hx-post="/scenario/save" hx-target="#status-list" hx-swap="outerHTML" class="space-y-4">
    <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Analista</label>
            <select name="participant" class="w-full px-3 py-2 border rounded-lg">
                {% for p in participantes %}
                <option value="{{p}}" {{ 'selected' if p==participant else '' }}>{{p}}</option>
                {% endfor %}
            </select>
        </div>
        <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Tema</label>
            <select name="topic" class="w-full px-3 py-2 border rounded-lg">
                {% for val,label in topicos %}
                <option value="{{val}}" {{ 'selected' if val==topic else '' }}>{{label}}</option>
                {% endfor %}
            </select>
        </div>
        <div class="flex items-end">
            <button type="submit" class="w-full bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-lg">Salvar</button>
        </div>
    </div>
    <div>
        <label class="block text-sm font-medium text-gray-700 mb-1">Conteúdo do Cenário</label>
        <textarea name="text" rows="6" placeholder="Cole o texto do cenário..." class="w-full px-3 py-2 border rounded-lg">{{ text }}</textarea>
    </div>
</form>""",

    "partials/status.html": """<div id="status-list" class="space-y-3">
    <h3 class="text-md font-medium text-gray-900">Status dos Participantes</h3>
    <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        {% for p in participantes %}
        {% set scenario_data = scenarios.get(p, {}) %}
        {% set has_text = scenario_data.get('text','').strip()|length > 0 %}
        {% set has_topic = scenario_data.get('topic','').strip()|length > 0 %}
        {% set is_complete = has_text and has_topic %}
        
        <div class="bg-gray-50 rounded-lg p-3 border-l-4 {{ 'border-green-500' if is_complete else 'border-yellow-500' if has_text or has_topic else 'border-gray-300' }}">
            <div class="flex items-center justify-between">
                <div>
                    <h4 class="font-medium text-gray-900 text-sm">{{ p }}</h4>
                    {% if scenario_data.get('topic') %}
                    <p class="text-xs text-gray-600">{{ scenario_data.get('topic') }}</p>
                    {% endif %}
                    <span class="inline-flex items-center px-2 py-1 rounded-full text-xs {{ 'bg-green-100 text-green-800' if is_complete else 'bg-yellow-100 text-yellow-800' if has_text or has_topic else 'bg-gray-100 text-gray-800' }}">
                        {{ 'Completo' if is_complete else 'Parcial' if has_text or has_topic else 'Pendente' }}
                    </span>
                </div>
                <button class="text-xs bg-white border px-2 py-1 rounded hover:bg-gray-50"
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
            <button class="text-gray-400 hover:text-gray-600" onclick="document.getElementById('preview-modal')?.remove()">✕</button>
        </div>
        <div class="flex-1 overflow-y-auto p-6">
            <div class="prose ata-preview max-w-none">{{ ata_html | safe }}</div>
        </div>
        <div class="p-4 border-t flex justify-end space-x-2">
            <a href="/export/pdf" class="bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg">PDF</a>
            <a href="/export/docx" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg">DOCX</a>
        </div>
    </div>
</div>"""
}

# Configurar Jinja2
env = Environment(loader=DictLoader(TEMPLATES), autoescape=select_autoescape(["html","xml"]))

def render(name: str, **ctx) -> HTMLResponse:
    ctx['version'] = VERSION
    return HTMLResponse(env.get_template(name).render(**ctx))

def nl2br(value: str) -> Markup:
    if not value: return Markup("")
    return Markup(value.replace("\n", "<br/>"))

env.filters["nl2br"] = nl2br

# Inicialização da aplicação
app = FastAPI(title=APP_TITLE, version=VERSION)

# Servir arquivos estáticos se existirem
if os.path.exists(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.on_event("startup")
async def on_startup():
    os.makedirs(STATIC_DIR, exist_ok=True)
    ensure_schema()

# WebSocket simplificado
@app.websocket("/ws/{user_id}")
async def websocket_endpoint(websocket: WebSocket, user_id: str):
    try:
        await manager.connect(websocket, user_id)
        while True:
            data = await websocket.receive_text()
            message = json.loads(data)
            await manager.broadcast_update(message, exclude_user=user_id)
    except WebSocketDisconnect:
        manager.disconnect(websocket)

# --------------------- Rotas principais
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    meeting = dict(meeting_row) if meeting_row else {"numero": 1, "ano": date.today().year, "data": date.today().isoformat(), "hora": "14:00", "local": "Sala nº 408 do 4º andar do IPAJM", "lavrador": ""}
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()
    
    item2 = state.config.get("current_item2", state.config.get("item2_default", "Não houve realocações de recursos desde a última reunião até a presente data."))
    assuntos = state.config.get("current_assuntos", state.config.get("assuntos_default", "– Assuntos gerais discutidos e/ou Eventos:"))
    
    return render("index.html", 
                  title=APP_TITLE,
                  participantes=PARTICIPANTES,
                  topicos=TOPICOS,
                  meeting=meeting,
                  scenarios=scenarios,
                  participant=PARTICIPANTES[0],
                  topic=TOPICOS[0][0],
                  text="",
                  item2=item2,
                  assuntos=assuntos)

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
        cur.execute("UPDATE meeting SET numero=?, ano=?, data=?, hora=?, local=?, lavrador=?, updated_at=CURRENT_TIMESTAMP WHERE id=?", 
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
    
    cur.execute("INSERT OR REPLACE INTO scenario (participant, text, topic, updated_at) VALUES (?, ?, ?, CURRENT_TIMESTAMP)", 
                (participant, text, topic))
    conn.commit()
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()
    
    html = env.get_template("partials/status.html").render(participantes=PARTICIPANTES, scenarios=scenarios)
    return HTMLResponse(f'<div id="status-list" class="space-y-3">{html}</div>')

@app.post("/preview/update")
async def preview_update(item2: str = Form(...), assuntos: str = Form(...)):
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("INSERT OR REPLACE INTO system_config (key, value, updated_at) VALUES ('current_item2', ?, CURRENT_TIMESTAMP)", (item2,))
    cur.execute("INSERT OR REPLACE INTO system_config (key, value, updated_at) VALUES ('current_assuntos', ?, CURRENT_TIMESTAMP)", (assuntos,))
    conn.commit()
    conn.close()
    
    state.config['current_item2'] = item2
    state.config['current_assuntos'] = assuntos
    
    return Response(status_code=204)

@app.post("/resumo/update")
async def resumo_update(rentab: str = Form(""), difpp: str = Form(""), posicao: str = Form("abaixo"), risco: str = Form("")):
    conn = db_connect()
    cur = conn.cursor()
    
    resumo_config = {'resumo_rentab': rentab, 'resumo_difpp': difpp, 'resumo_posicao': posicao, 'resumo_risco': risco}
    
    for key, value in resumo_config.items():
        cur.execute("INSERT OR REPLACE INTO system_config (key, value, updated_at) VALUES (?, ?, CURRENT_TIMESTAMP)", (key, value))
    
    conn.commit()
    conn.close()
    state.config.update(resumo_config)
    
    return Response(status_code=204)

# Funções para geração da ata
def _prefixo_genero(nome: str) -> str:
    return "A Sra." if nome in PARTICIPANTES_MULHERES else "O Sr."

def numero_sessao_fmt(numero: int, ano: int) -> str:
    return f"{numero:03d}/{ano}"

def item1_single_paragraph(meeting: dict, scenarios: dict) -> str:
    d = date.fromisoformat(meeting["data"])
    
    ORDINAIS = {
        1: "primeiro", 2: "segundo", 3: "terceiro", 4: "quarto", 5: "quinto",
        6: "sexto", 7: "sétimo", 8: "oitavo", 9: "nono", 10: "décimo",
        11: "décimo primeiro", 12: "décimo segundo", 13: "décimo terceiro", 
        14: "décimo quarto", 15: "décimo quinto", 16: "décimo sexto",
        17: "décimo sétimo", 18: "décimo oitavo", 19: "décimo nono", 
        20: "vigésimo", 21: "vigésimo primeiro", 22: "vigésimo segundo",
        23: "vigésimo terceiro", 24: "vigésimo quarto", 25: "vigésimo quinto",
        26: "vigésimo sexto", 27: "vigésimo sétimo", 28: "vigésimo oitavo",
        29: "vigésimo nono", 30: "trigésimo", 31: "trigésimo primeiro"
    }

    intro = (
        f"No {ORDINAIS.get(d.day, str(d.day))} dia do mês de {mes_ano_pt(d).split(' de ')[0]} "
        f"do ano de {d.year}, às {meeting['hora']} horas, na {meeting['local']}, "
        f"ocorreu a {meeting['numero']}ª Reunião Ordinária dos Membros do Comitê de Investimentos. "
    )

    partes = [intro]
    for nome in PARTICIPANTES:
        row = scenarios.get(nome, {"text": "", "topic": ""})
        texto = (row.get("text") or "").strip()
        tema = (row.get("topic") or "").strip()
        if texto:
            partes.append(
                f"<strong>{_prefixo_genero(nome)} {nome}</strong> falando sobre {tema}, {texto} "
            )

    return "<p>" + "".join(partes).strip() + "</p>"

def item3_html(meeting: dict) -> str:
    d = date.fromisoformat(meeting["data"])
    d2 = menos_dois_meses(d)
    mes_ano = mes_ano_pt(d2)
    
    p0 = (
        "O Comitê de Investimentos, buscando transmitir maior transparência em relação às análises dos investimentos do Instituto e, em consequência, "
        "aderindo às normas do Pró-Gestão, elabora o "Relatório de Análise de Investimentos IPAJM". "
        "Este relatório já foi encaminhado à SCO – Subgerência de Contabilidade e Orçamento, para posterior envio para análise do Conselho Fiscal do IPAJM. "
        f"Segue abaixo um resumo relativo aos itens abordados no Relatório supracitado de {mes_ano}:"
    )
    
    rentab = state.config.get("resumo_rentab", "")
    difpp = state.config.get("resumo_difpp", "")
    pos = state.config.get("resumo_posicao", "abaixo")
    risco = state.config.get("resumo_risco", "")
    
    p1 = f"1) Acompanhamento da rentabilidade -  A rentabilidade consolidada dos investimentos do Fundo Previdenciário em {mes_ano} foi de {rentab}, ficando {difpp} p.p. {pos} da meta atuarial."
    p2 = f"2) Avaliação de risco da carteira - O grau de variação nas rentabilidades está coerente com o grau de risco assumido, em {risco}."
    p3 = f"3) Execução da Política de Investimentos – As movimentações financeiras realizadas no mês de {mes_ano} estão de acordo com as deliberações estabelecidas com a Diretoria de Investimentos e com a legislação vigente."
    p4 = (
        f"4) Aderência a Política de Investimentos - Os recursos investidos, abrangendo a carteira consolidada, que representa o patrimônio total do RPPS sob gestão, estão aderentes à Política de Investimentos de {meeting['ano']}, respeitando o estabelecido na legislação em vigor e dentro dos percentuais definidos. "
        "Considerando que as taxas ainda são negociadas acima da meta atuarial, seguimos com a estratégia de alcançar o alvo definido de 60% de alocação em Títulos Públicos."
    )
    
    return "\n".join([f"<p><strong>{t}</strong></p>" if i==0 else f"<p>{t}</p>" for i,t in enumerate([p0,p1,p2,p3,p4])])

def ata_html_full(meeting: dict, scenarios: dict) -> str:
    d = date.fromisoformat(meeting["data"])
    numero_fmt = numero_sessao_fmt(meeting["numero"], meeting["ano"])
    presencas_html = "<br/>".join([f"<strong>{p}</strong> - {CARGO};" for p in PARTICIPANTES])
    
    item2 = state.config.get("current_item2", state.config.get("item2_default", "Não houve realocações de recursos desde a última reunião até a presente data."))
    assuntos = state.config.get("current_assuntos", state.config.get("assuntos_default", "– Assuntos gerais discutidos e/ou Eventos:"))
    
    blocos = [
        f"<p><strong>Sessão Ordinária nº {numero_fmt}</strong></p>",
        f"<p><strong>Data:</strong> {ptbr_date(d)}.<br/><strong>Hora:</strong> {meeting['hora']}h.<br/><strong>Local:</strong> {meeting['local']}.</p>",
        "<p><strong>Presenças:</strong></p>",
        f"<p>{presencas_html}</p>",
        "<p><strong>Ordem do Dia:</strong></p>",
        "<p>1. <strong>Cenário Político e Econômico Interno</strong> e <strong>Cenário Econômico Externo (EUA, Europa e China)</strong>;<br/>"
        "2. <strong>Movimentações e Aplicações financeiras</strong>;<br/>"
        "3. <strong>Acompanhamento dos Recursos Investidos</strong>;<br/>"
        "4. <strong>Assuntos Gerais</strong>.</p>",
        "<p><strong>Item 01 – Cenário Político e Econômico Interno e Cenário Econômico Externo (EUA, Europa e China):</strong></p>",
        item1_single_paragraph(meeting, scenarios),
        "<p><strong>Item 02 – Movimentações e Aplicações financeiras</strong></p>",
        f"<p>{item2}</p>",
        "<p><strong>Item 03 – Acompanhamento dos Recursos Investidos:</strong></p>",
        item3_html(meeting),
        "<p><strong>Item 04 – Assuntos Gerais</strong></p>",
        f"<p>{assuntos}</p>",
        (lambda lav: f"<p>Nada mais havendo a tratar, foi encerrada a reunião e eu, "
                     f"{f'<strong>{lav}</strong>' if lav else '___________________________________'}, "
                     "lavrei a presente Ata, assinada pelos membros presentes do Comitê de Investimentos.</p>")((meeting.get('lavrador') or '').strip()),
        f"<p><strong>{PARTICIPANTES[0]}</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>{PARTICIPANTES[1]}</strong><br/>{CARGO}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{CARGO}</p>",
        f"<p><strong>{PARTICIPANTES[2]}</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>{PARTICIPANTES[3]}</strong><br/>{CARGO}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{CARGO}</p>",
        f"<p><strong>{PARTICIPANTES[4]}</strong><br/>{CARGO}</p>",
    ]
    
    # Cabeçalho HTML/PDF
    header_html = """
    <div style="width:100%; text-align:center; margin-bottom:20px;">
      <div style="font-family: Calibri; font-size:11pt; font-weight:700; color:#000;">
        GOVERNO DO ESTADO DO ESPÍRITO SANTO<br/>
        INSTITUTO DE PREVIDÊNCIA DOS<br/>
        SERVIDORES DO ESTADO DO ESPÍRITO SANTO
      </div>
      <div style="border-bottom:1px solid #000; margin:10px 0; text-align:center;">IPAJM</div>
    </div>
    """
    return header_html + "\n".join(blocos)

@app.get("/preview/modal")
async def preview_modal():
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    meeting = dict(meeting_row) if meeting_row else {"numero": 1, "ano": date.today().year, "data": date.today().isoformat(), "hora": "14:00", "local": "Sala nº 408", "lavrador": ""}
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()
    
    ata_html = ata_html_full(meeting, scenarios)
    
    html = env.get_template("partials/preview_modal.html").render(ata_html=ata_html)
    return HTMLResponse(html)

# --------------------- Exportações
@app.get("/export/pdf")
async def export_pdf():
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    meeting = dict(meeting_row) if meeting_row else {"numero": 1, "ano": date.today().year, "data": date.today().isoformat(), "hora": "14:00", "local": "Sala nº 408", "lavrador": ""}
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()
    
    html_body = ata_html_full(meeting, scenarios)
    html = f"""<html><head><meta charset='utf-8'>
      <style>
        body {{ font-family: Calibri, Arial, sans-serif; font-size: 11pt; color:#000; line-height: 1.4; margin: 20px; }}
        .content {{ white-space: normal; text-align: justify; }}
        p {{ margin: 6pt 0; }}
        strong {{ font-weight: 700; }}
      </style>
    </head><body><div class="content">{html_body}</div></body></html>"""
    
    pdf_buf = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html), dest=pdf_buf)
    pdf_buf.seek(0)
    
    filename = f"Ata_{meeting.get('numero', 1):03d}-{meeting.get('ano', date.today().year)}.pdf"
    return StreamingResponse(pdf_buf, media_type="application/pdf", headers={"Content-Disposition": f"attachment; filename={filename}"})

@app.get("/export/docx")
async def export_docx():
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    meeting = dict(meeting_row) if meeting_row else {"numero": 1, "ano": date.today().year, "data": date.today().isoformat(), "hora": "14:00", "local": "Sala nº 408", "lavrador": ""}
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()

    d = date.fromisoformat(meeting["data"]) if meeting.get("data") else date.today()
    numero_fmt = numero_sessao_fmt(meeting.get("numero", 1), meeting.get("ano", d.year))

    doc = DocxDocument()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # Adicionar título
    p = doc.add_paragraph()
    run = p.add_run(f"Sessão Ordinária nº {numero_fmt}")
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Adicionar outros conteúdos básicos
    p = doc.add_paragraph()
    p.add_run(f"Data: {ptbr_date(d)} • Hora: {meeting.get('hora', '14:00')}h • Local: {meeting.get('local', 'A definir')}")

    # Salvar e enviar
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    
    filename = f"Ata_{meeting.get('numero', 1):03d}-{meeting.get('ano', d.year)}.docx"
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", headers={"Content-Disposition": f"attachment; filename={filename}"})

# Health check para deploy
@app.get("/health")
async def health_check():
    return JSONResponse({"status": "healthy", "version": VERSION, "timestamp": datetime.now().isoformat()})

# API Status  
@app.get("/api/status")
async def api_status():
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM scenario WHERE text != ''")
    completed = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM scenario")
    total = cur.fetchone()[0]
    conn.close()
    
    return JSONResponse({
        "status": "online",
        "version": VERSION,
        "scenarios_completed": completed,
        "total_scenarios": total,
        "completion_rate": (completed / total) * 100 if total > 0 else 0
    })

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
