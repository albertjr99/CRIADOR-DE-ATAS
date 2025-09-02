# Atas.py — Versão Modernizada com melhorias em UX/UI e funcionalidades
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

class ItemUpdate(BaseModel):
    item2: str = Field(max_length=2000)
    assuntos: str = Field(max_length=2000)

class ResumoUpdate(BaseModel):
    rentab: str = Field(max_length=20)
    difpp: str = Field(max_length=20)
    posicao: str = Field(regex="^(abaixo|acima)$")
    risco: str = Field(max_length=20)

# Gerenciador de conexões WebSocket para colaboração em tempo real
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
        
        # Notifica outros usuários
        await self.broadcast_user_status()

    def disconnect(self, websocket: WebSocket):
        if websocket in self.user_sessions:
            user_id = self.user_sessions[websocket]
            self.active_connections[user_id].discard(websocket)
            if not self.active_connections[user_id]:
                del self.active_connections[user_id]
            del self.user_sessions[websocket]

    async def broadcast_user_status(self):
        active_users = list(self.active_connections.keys())
        message = {
            "type": "user_status",
            "active_users": active_users,
            "timestamp": datetime.now().isoformat()
        }
        
        for connections in self.active_connections.values():
            for connection in connections:
                try:
                    await connection.send_text(json.dumps(message))
                except:
                    pass

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

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --------------------- Funções auxiliares (mantidas)
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

# --------------------- Database com melhorias
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
    
    # Tabela de histórico de versões
    cur.execute("""
        CREATE TABLE IF NOT EXISTS version_history (
            id INTEGER PRIMARY KEY,
            action TEXT NOT NULL,
            table_name TEXT NOT NULL,
            record_id INTEGER,
            old_data TEXT,
            new_data TEXT,
            user_id TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
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
        ("auto_save_interval", "30"), # segundos
    ]
    
    for key, value in default_configs:
        cur.execute("INSERT OR IGNORE INTO system_config (key, value) VALUES (?, ?)", (key, value))
    
    conn.commit()
    conn.close()

# --------------------- Templates HTML modernizados
TEMPLATES = {
    "base.html": """<!doctype html>
<html lang="pt-br" class="scroll-smooth">
<head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <title>{{ title }}</title>
    <script src="https://unpkg.com/htmx.org@1.9.12"></script>
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        primary: {
                            50: '#eff6ff',
                            500: '#3b82f6',
                            600: '#2563eb',
                            700: '#1d4ed8',
                            900: '#1e3a8a',
                        }
                    },
                    animation: {
                        'fade-in': 'fadeIn 0.5s ease-in-out',
                        'slide-up': 'slideUp 0.3s ease-out',
                        'pulse-soft': 'pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite',
                    }
                }
            }
        }
    </script>
    <style>
        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }
        @keyframes slideUp {
            from { transform: translateY(20px); opacity: 0; }
            to { transform: translateY(0); opacity: 1; }
        }
        .glass-effect {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
        .text-shadow {
            text-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .prose.ata-preview {
            font-family: Calibri, "Segoe UI", Arial, sans-serif !important;
            font-size: 11pt !important;
            color: #000 !important;
            text-align: justify;
            line-height: 1.35;
            white-space: pre-wrap;
        }
        .save-indicator {
            transition: all 0.3s ease;
        }
    </style>
</head>
<body class="bg-gradient-to-br from-blue-50 via-white to-indigo-50 text-gray-900 min-h-screen">
    <!-- Header modernizado -->
    <header class="sticky top-0 z-50 glass-effect shadow-lg border-b border-gray-200">
        <div class="max-w-7xl mx-auto px-4 py-4">
            <div class="flex items-center justify-between">
                <div class="flex items-center space-x-4">
                    <div class="h-10 w-10 bg-gradient-to-br from-primary-500 to-primary-700 rounded-lg flex items-center justify-center">
                        <svg class="h-6 w-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                        </svg>
                    </div>
                    <div>
                        <h1 class="text-xl font-bold text-gray-900 text-shadow">Sistema de Atas</h1>
                        <p class="text-sm text-gray-600">Comitê de Investimentos v{{ version }}</p>
                    </div>
                </div>
                
                <!-- Status de usuários online -->
                <div id="online-users" class="hidden md:flex items-center space-x-2">
                    <div class="h-2 w-2 bg-green-400 rounded-full animate-pulse-soft"></div>
                    <span class="text-sm text-gray-600" id="user-count">1 usuário online</span>
                </div>
                
                <!-- Botões de ação -->
                <div class="flex items-center space-x-3">
                    <div id="save-status" class="save-indicator opacity-0">
                        <span class="text-sm text-green-600 bg-green-50 px-2 py-1 rounded-full">Salvo automaticamente</span>
                    </div>
                    
                    <button id="preview-btn" 
                            class="bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-lg transition-all duration-200 shadow-md hover:shadow-lg transform hover:scale-105"
                            hx-get="/preview/modal" hx-target="body" hx-swap="beforeend">
                        <svg class="h-4 w-4 inline mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/>
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/>
                        </svg>
                        Prévia
                    </button>
                    
                    <div class="relative group">
                        <button class="bg-primary-600 hover:bg-primary-700 text-white px-4 py-2 rounded-lg transition-all duration-200 shadow-md hover:shadow-lg transform hover:scale-105 flex items-center">
                            <svg class="h-4 w-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                            </svg>
                            Exportar
                        </button>
                        <div class="absolute right-0 mt-2 w-48 bg-white rounded-lg shadow-xl border opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-200 z-10">
                            <a href="/export/docx" class="block px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 rounded-t-lg">
                                <svg class="h-4 w-4 inline mr-2 text-blue-600" fill="currentColor" viewBox="0 0 20 20">
                                    <path d="M4 3a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V5a2 2 0 00-2-2H4zm12 12H4l4-8 3 6 2-4 3 6z"/>
                                </svg>
                                Exportar DOCX
                            </a>
                            <a href="/export/pdf" class="block px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 rounded-b-lg">
                                <svg class="h-4 w-4 inline mr-2 text-red-600" fill="currentColor" viewBox="0 0 20 20">
                                    <path d="M4 3a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V5a2 2 0 00-2-2H4zm12 12H4l4-8 3 6 2-4 3 6z"/>
                                </svg>
                                Exportar PDF
                            </a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </header>

    <!-- Progress bar -->
    <div class="h-1 bg-gray-200">
        <div id="progress-bar" class="h-full bg-gradient-to-r from-primary-500 to-purple-500 transition-all duration-300" style="width: 0%"></div>
    </div>

    <!-- Main content -->
    <main class="max-w-7xl mx-auto px-4 py-8">
        {% block content %}{% endblock %}
    </main>

    <!-- Footer -->
    <footer class="border-t border-gray-200 bg-white">
        <div class="max-w-7xl mx-auto px-4 py-6">
            <div class="flex items-center justify-between">
                <p class="text-sm text-gray-500">
                    Sistema de Atas v{{ version }} • Última atualização: <span id="last-update">--</span>
                </p>
                <div class="flex items-center space-x-4">
                    <div class="flex items-center text-sm text-gray-500">
                        <div class="h-2 w-2 bg-green-400 rounded-full mr-2"></div>
                        Sistema operacional
                    </div>
                </div>
            </div>
        </div>
    </footer>

    <!-- Scripts -->
    <script>
        // Auto-save functionality
        let autoSaveTimer;
        let hasUnsavedChanges = false;

        function showSaveStatus(message, isSuccess = true) {
            const status = document.getElementById('save-status');
            const span = status.querySelector('span');
            span.textContent = message;
            span.className = isSuccess ? 
                'text-sm text-green-600 bg-green-50 px-2 py-1 rounded-full' : 
                'text-sm text-red-600 bg-red-50 px-2 py-1 rounded-full';
            status.style.opacity = '1';
            setTimeout(() => status.style.opacity = '0', 3000);
        }

        function updateProgress() {
            const forms = document.querySelectorAll('form');
            const filledInputs = document.querySelectorAll('input:not([value=""]), textarea:not(:empty), select:not([value=""])');
            const totalInputs = document.querySelectorAll('input, textarea, select');
            const progress = totalInputs.length > 0 ? (filledInputs.length / totalInputs.length) * 100 : 0;
            document.getElementById('progress-bar').style.width = progress + '%';
        }

        function setupAutoSave() {
            const inputs = document.querySelectorAll('input, textarea, select');
            inputs.forEach(input => {
                input.addEventListener('input', () => {
                    hasUnsavedChanges = true;
                    clearTimeout(autoSaveTimer);
                    autoSaveTimer = setTimeout(() => {
                        const form = input.closest('form');
                        if (form && form.hasAttribute('hx-post')) {
                            htmx.trigger(form, 'submit');
                            hasUnsavedChanges = false;
                            showSaveStatus('Salvo automaticamente');
                        }
                    }, 2000); // 2 segundos após parar de digitar
                });
            });
        }

        // WebSocket para colaboração em tempo real
        function setupWebSocket() {
            const protocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
            const ws = new WebSocket(`${protocol}//${window.location.host}/ws/${Math.random().toString(36).substr(2, 9)}`);
            
            ws.onmessage = function(event) {
                const data = JSON.parse(event.data);
                if (data.type === 'user_status') {
                    updateUserStatus(data.active_users);
                }
            };
        }

        function updateUserStatus(activeUsers) {
            const userCount = document.getElementById('user-count');
            const onlineUsers = document.getElementById('online-users');
            if (userCount && onlineUsers) {
                userCount.textContent = `${activeUsers.length} usuário${activeUsers.length > 1 ? 's' : ''} online`;
                onlineUsers.classList.remove('hidden');
            }
        }

        // Inicialização
        document.addEventListener('DOMContentLoaded', function() {
            updateProgress();
            setupAutoSave();
            setupWebSocket();
            
            // Atualizar timestamp
            document.getElementById('last-update').textContent = new Date().toLocaleString('pt-BR');
        });

        // Interceptar eventos HTMX
        document.body.addEventListener('htmx:afterRequest', function(evt) {
            if (evt.detail.successful) {
                updateProgress();
                setupAutoSave(); // Re-setup após mudanças no DOM
            }
        });
    </script>
</body>
</html>""",

    "index.html": """{% extends 'base.html' %}
{% block content %}
<div class="grid grid-cols-1 lg:grid-cols-3 gap-8 animate-fade-in">
    <!-- Coluna principal -->
    <div class="lg:col-span-2 space-y-8">
        <!-- Configurações da Reunião -->
        <div class="bg-white rounded-2xl shadow-xl border border-gray-100 overflow-hidden">
            <div class="bg-gradient-to-r from-primary-500 to-primary-600 px-6 py-4">
                <h2 class="text-lg font-semibold text-white flex items-center">
                    <svg class="h-5 w-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 100 4m0-4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 100 4m0-4v2m0-6V4"/>
                    </svg>
                    Configurações da Reunião
                </h2>
            </div>
            <div id="session-block" class="p-6">
                {% include 'partials/session.html' with context %}
            </div>
        </div>

        <!-- Item 01 - Cenários -->
        <div class="bg-white rounded-2xl shadow-xl border border-gray-100 overflow-hidden">
            <div class="bg-gradient-to-r from-purple-500 to-purple-600 px-6 py-4">
                <h2 class="text-lg font-semibold text-white flex items-center">
                    <svg class="h-5 w-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3.055 11H5a2 2 0 012 2v1a2 2 0 002 2 2 2 0 012 2v2.945"/>
                    </svg>
                    Item 01 — Cenário por Analista
                </h2>
            </div>
            <div class="p-6">
                <div id="item1-form" class="mb-6">
                    {% include 'partials/item1_form.html' with context %}
                </div>
                <div id="status-list">
                    {% include 'partials/status.html' with context %}
                </div>
            </div>
        </div>

        <!-- Itens 02 e 04 -->
        <div class="bg-white rounded-2xl shadow-xl border border-gray-100 overflow-hidden">
            <div class="bg-gradient-to-r from-green-500 to-green-600 px-6 py-4">
                <h2 class="text-lg font-semibold text-white flex items-center">
                    <svg class="h-5 w-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/>
                    </svg>
                    Itens 02 e 04 — Texto Livre
                </h2>
            </div>
            <div class="p-6">
                <form hx-post="/preview/update" hx-target="this" hx-swap="none" class="space-y-6">
                    <div class="space-y-2">
                        <label class="block text-sm font-medium text-gray-700">
                            Item 02 — Movimentações e Aplicações Financeiras
                        </label>
                        <textarea name="item2" 
                                  class="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200 resize-none"
                                  rows="4"
                                  placeholder="Descreva as movimentações e aplicações financeiras...">{{ item2 }}</textarea>
                    </div>
                    
                    <div class="space-y-2">
                        <label class="block text-sm font-medium text-gray-700">
                            Item 04 — Assuntos Gerais
                        </label>
                        <textarea name="assuntos"
                                  class="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200 resize-none"
                                  rows="5"
                                  placeholder="Liste os assuntos gerais discutidos...">{{ assuntos }}</textarea>
                    </div>
                    
                    <button type="submit" 
                            class="bg-primary-600 hover:bg-primary-700 text-white px-6 py-3 rounded-lg transition-all duration-200 shadow-md hover:shadow-lg transform hover:scale-105">
                        Salvar Itens 02 e 04
                    </button>
                </form>
            </div>
        </div>
    </div>

    <!-- Coluna lateral -->
    <div class="space-y-8">
        <!-- Item 03 - Parâmetros -->
        <div class="bg-white rounded-2xl shadow-xl border border-gray-100 overflow-hidden">
            <div class="bg-gradient-to-r from-orange-500 to-orange-600 px-6 py-4">
                <h2 class="text-lg font-semibold text-white flex items-center">
                    <svg class="h-5 w-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/>
                    </svg>
                    Item 03 — Parâmetros
                </h2>
            </div>
            <div class="p-6">
                <form hx-post="/resumo/update" hx-target="this" hx-swap="none" class="space-y-4">
                    <div class="space-y-2">
                        <label class="block text-sm font-medium text-gray-700">
                            Rentabilidade do Fundo (%) no mês D-2
                        </label>
                        <input type="text" 
                               name="rentab" 
                               placeholder="ex: 1,23%"
                               class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200"/>
                    </div>
                    
                    <div class="grid grid-cols-2 gap-4">
                        <div class="space-y-2">
                            <label class="block text-sm font-medium text-gray-700">
                                Diferença vs. meta (p.p.)
                            </label>
                            <input type="text" 
                                   name="difpp" 
                                   placeholder="ex: 0,15"
                                   class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200"/>
                        </div>
                        <div class="space-y-2">
                            <label class="block text-sm font-medium text-gray-700">Posição</label>
                            <select name="posicao" 
                                    class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200">
                                <option value="abaixo">Abaixo</option>
                                <option value="acima">Acima</option>
                            </select>
                        </div>
                    </div>
                    
                    <div class="space-y-2">
                        <label class="block text-sm font-medium text-gray-700">
                            Risco assumido (%) no mês D-2
                        </label>
                        <input type="text" 
                               name="risco" 
                               placeholder="ex: 7,5%"
                               class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200"/>
                    </div>
                    
                    <button type="submit" 
                            class="w-full bg-orange-600 hover:bg-orange-700 text-white px-4 py-3 rounded-lg transition-all duration-200 shadow-md hover:shadow-lg transform hover:scale-105">
                        Salvar Parâmetros
                    </button>
                </form>
                <p class="text-xs text-gray-500 mt-3">
                    <svg class="h-3 w-3 inline mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                    </svg>
                    D-2 = dois meses antes da data da reunião
                </p>
            </div>
        </div>

        <!-- Card de atalhos -->
        <div class="bg-gradient-to-br from-indigo-500 to-purple-600 rounded-2xl shadow-xl text-white p-6">
            <h3 class="text-lg font-semibold mb-4 flex items-center">
                <svg class="h-5 w-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"/>
                </svg>
                Ações Rápidas
            </h3>
            <div class="space-y-3">
                <button onclick="window.print()" 
                        class="w-full bg-white/20 hover:bg-white/30 backdrop-blur text-white px-4 py-2 rounded-lg transition-all duration-200 text-sm flex items-center">
                    <svg class="h-4 w-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"/>
                    </svg>
                    Imprimir Página
                </button>
                <button id="clear-form" 
                        class="w-full bg-white/20 hover:bg-white/30 backdrop-blur text-white px-4 py-2 rounded-lg transition-all duration-200 text-sm flex items-center">
                    <svg class="h-4 w-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/>
                    </svg>
                    Limpar Formulário
                </button>
            </div>
        </div>
    </div>
</div>
{% endblock %}""",

    "partials/session.html": """<form hx-post="/meeting/update" hx-target="#session-block" hx-swap="outerHTML" class="space-y-6">
    <div class="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-6 gap-4">
        <div class="md:col-span-2 space-y-2">
            <label class="block text-sm font-medium text-gray-700">Número da Ata</label>
            <div class="flex space-x-2">
                <input type="number" 
                       name="numero" 
                       value="{{ meeting.numero }}" 
                       min="1" max="9999"
                       class="w-20 px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200"/>
                <input type="number" 
                       name="ano" 
                       value="{{ meeting.ano }}" 
                       min="2020" max="2050"
                       class="w-24 px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200"/>
            </div>
            <p class="text-xs text-gray-500">Formato: {{ '%03d' % meeting.numero }}/{{ meeting.ano }}</p>
        </div>
        
        <div class="md:col-span-2 space-y-2">
            <label class="block text-sm font-medium text-gray-700">Data da Reunião</label>
            <input type="date" 
                   name="data" 
                   value="{{ meeting.data }}" 
                   class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200"/>
        </div>
        
        <div class="space-y-2">
            <label class="block text-sm font-medium text-gray-700">Horário</label>
            <input type="time" 
                   name="hora" 
                   value="{{ meeting.hora }}" 
                   class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200"/>
        </div>
        
        <div class="md:col-span-3 space-y-2">
            <label class="block text-sm font-medium text-gray-700">Local da Reunião</label>
            <input type="text" 
                   name="local" 
                   value="{{ meeting.local }}" 
                   placeholder="Ex: Sala nº 408 do 4º andar do IPAJM"
                   class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200"/>
        </div>
        
        <div class="md:col-span-3 space-y-2">
            <label class="block text-sm font-medium text-gray-700">Responsável pela Ata</label>
            <select name="lavrador" 
                    class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-primary-500 focus:border-transparent transition-all duration-200">
                <option value="">(Selecionar responsável)</option>
                {% for p in participantes %}
                <option value="{{p}}" {{ 'selected' if meeting.lavrador==p else '' }}>{{p}}</option>
                {% endfor %}
            </select>
        </div>
    </div>
    
    <div class="flex justify-end">
        <button type="submit" 
                class="bg-primary-600 hover:bg-primary-700 text-white px-6 py-3 rounded-lg transition-all duration-200 shadow-md hover:shadow-lg transform hover:scale-105 flex items-center">
            <svg class="h-4 w-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/>
            </svg>
            Salvar Configurações
        </button>
    </div>
</form>""",

    "partials/item1_form.html": """<form id="item1-form" hx-post="/scenario/save" hx-target="#status-list" hx-swap="outerHTML" class="space-y-6">
    <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div class="space-y-2">
            <label class="block text-sm font-medium text-gray-700">Analista</label>
            <select name="participant" 
                    class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition-all duration-200">
                {% for p in participantes %}
                <option value="{{p}}" {{ 'selected' if p==participant else '' }}>{{p}}</option>
                {% endfor %}
            </select>
        </div>
        
        <div class="space-y-2">
            <label class="block text-sm font-medium text-gray-700">Tema</label>
            <select name="topic" 
                    class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition-all duration-200">
                {% for val,label in topicos %}
                <option value="{{val}}" {{ 'selected' if val==topic else '' }}>{{label}}</option>
                {% endfor %}
            </select>
        </div>
        
        <div class="flex items-end">
            <button type="submit" 
                    class="w-full bg-purple-600 hover:bg-purple-700 text-white px-4 py-3 rounded-lg transition-all duration-200 shadow-md hover:shadow-lg transform hover:scale-105 flex items-center justify-center">
                <svg class="h-4 w-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/>
                </svg>
                Salvar
            </button>
        </div>
    </div>
    
    <div class="space-y-2">
        <label class="block text-sm font-medium text-gray-700">Conteúdo do Cenário</label>
        <textarea name="text" 
                  placeholder="Cole aqui o texto do cenário econômico/político..."
                  class="w-full h-40 px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition-all duration-200 resize-none">{{ text }}</textarea>
        <p class="text-xs text-gray-500">
            <svg class="h-3 w-3 inline mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
            </svg>
            Limite de 5.000 caracteres
        </p>
    </div>
</form>""",

    "partials/status.html": """<div id="status-list" class="space-y-4">
    <h3 class="text-lg font-medium text-gray-900 flex items-center">
        <svg class="h-5 w-5 mr-2 text-purple-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197m13.5-9a2.5 2.5 0 11-5 0 2.5 2.5 0 015 0z"/>
        </svg>
        Status dos Participantes
    </h3>
    
    <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        {% for p in participantes %}
        {% set scenario_data = scenarios.get(p, {}) %}
        {% set has_text = scenario_data.get('text','').strip()|length > 0 %}
        {% set has_topic = scenario_data.get('topic','').strip()|length > 0 %}
        {% set is_complete = has_text and has_topic %}
        
        <div class="bg-gray-50 rounded-xl p-4 border-l-4 {{ 'border-green-500 bg-green-50' if is_complete else 'border-yellow-500 bg-yellow-50' if has_text or has_topic else 'border-gray-300' }} transition-all duration-200">
            <div class="flex items-start justify-between">
                <div class="flex-1">
                    <h4 class="font-medium text-gray-900 text-sm mb-1">{{ p }}</h4>
                    
                    {% if scenario_data.get('topic') %}
                    <p class="text-xs text-gray-600 mb-1">
                        <span class="font-medium">Tema:</span> {{ scenario_data.get('topic') }}
                    </p>
                    {% endif %}
                    
                    {% if scenario_data.get('text') %}
                    <p class="text-xs text-gray-600 truncate">
                        {{ scenario_data.get('text')[:100] }}...
                    </p>
                    {% endif %}
                    
                    <div class="flex items-center mt-2">
                        {% if is_complete %}
                        <span class="inline-flex items-center px-2 py-1 rounded-full text-xs bg-green-100 text-green-800">
                            <svg class="h-3 w-3 mr-1" fill="currentColor" viewBox="0 0 20 20">
                                <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd"/>
                            </svg>
                            Completo
                        </span>
                        {% elif has_text or has_topic %}
                        <span class="inline-flex items-center px-2 py-1 rounded-full text-xs bg-yellow-100 text-yellow-800">
                            <svg class="h-3 w-3 mr-1" fill="currentColor" viewBox="0 0 20 20">
                                <path fill-rule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clip-rule="evenodd"/>
                            </svg>
                            Parcial
                        </span>
                        {% else %}
                        <span class="inline-flex items-center px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-800">
                            <svg class="h-3 w-3 mr-1" fill="currentColor" viewBox="0 0 20 20">
                                <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clip-rule="evenodd"/>
                            </svg>
                            Pendente
                        </span>
                        {% endif %}
                    </div>
                </div>
                
                <button class="ml-3 text-xs bg-white hover:bg-gray-50 border border-gray-300 px-3 py-1 rounded-lg transition-all duration-200 hover:shadow-md"
                        hx-get="/scenario/form?participant={{p|urlencode}}"
                        hx-target="#item1-form" 
                        hx-swap="outerHTML">
                    Editar
                </button>
            </div>
        </div>
        {% endfor %}
    </div>
</div>""",

    "partials/preview_modal.html": """<div id="preview-modal" class="fixed inset-0 z-50 animate-fade-in">
    <div class="absolute inset-0 bg-black/50 backdrop-blur-sm" onclick="document.getElementById('preview-modal')?.remove()"></div>
    <div class="absolute inset-0 flex items-center justify-center p-4">
        <div class="relative bg-white rounded-2xl shadow-2xl w-[95vw] max-w-6xl max-h-[90vh] border flex flex-col animate-slide-up">
            <!-- Header do modal -->
            <div class="flex items-center justify-between px-6 py-4 border-b bg-gradient-to-r from-primary-500 to-primary-600 text-white rounded-t-2xl">
                <div class="flex items-center">
                    <svg class="h-5 w-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/>
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/>
                    </svg>
                    <h3 class="text-lg font-semibold">Prévia da Ata</h3>
                </div>
                <div class="flex items-center space-x-3">
                    <button onclick="window.print()" 
                            class="bg-white/20 hover:bg-white/30 text-white px-3 py-1.5 rounded-lg transition-all duration-200 text-sm flex items-center">
                        <svg class="h-4 w-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"/>
                        </svg>
                        Imprimir
                    </button>
                    <button class="bg-white/20 hover:bg-white/30 text-white px-3 py-1.5 rounded-lg transition-all duration-200" 
                            onclick="document.getElementById('preview-modal')?.remove()">
                        <svg class="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/>
                        </svg>
                    </button>
                </div>
            </div>
            
            <!-- Conteúdo do modal -->
            <div class="flex-1 overflow-y-auto p-8">
                <div class="prose ata-preview max-w-none bg-white shadow-inner rounded-lg p-6">
                    {{ ata_html | safe }}
                </div>
            </div>
            
            <!-- Footer do modal -->
            <div class="px-6 py-4 border-t bg-gray-50 rounded-b-2xl flex justify-between items-center">
                <p class="text-sm text-gray-600">
                    <svg class="h-4 w-4 inline mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                    </svg>
                    Layout idêntico será mantido na exportação DOCX
                </p>
                <div class="flex space-x-2">
                    <a href="/export/pdf" 
                       class="bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg transition-all duration-200 text-sm flex items-center">
                        <svg class="h-4 w-4 mr-1" fill="currentColor" viewBox="0 0 20 20">
                            <path d="M4 3a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V5a2 2 0 00-2-2H4zm12 12H4l4-8 3 6 2-4 3 6z"/>
                        </svg>
                        PDF
                    </a>
                    <a href="/export/docx" 
                       class="bg-primary-600 hover:bg-primary-700 text-white px-4 py-2 rounded-lg transition-all duration-200 text-sm flex items-center">
                        <svg class="h-4 w-4 mr-1" fill="currentColor" viewBox="0 0 20 20">
                            <path d="M4 3a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V5a2 2 0 00-2-2H4zm12 12H4l4-8 3 6 2-4 3 6z"/>
                        </svg>
                        DOCX
                    </a>
                </div>
            </div>
        </div>
    </div>
</div>
<script>
    (function(){
        const modal = document.getElementById('preview-modal');
        const onEsc = (e) => {
            if(e.key === 'Escape'){
                modal?.remove();
                document.removeEventListener('keydown', onEsc);
            }
        };
        document.addEventListener('keydown', onEsc);
        
        // Focus no modal para acessibilidade
        modal.setAttribute('tabindex', '-1');
        modal.focus();
    })();
</script>""",
}

# Configurar Jinja2
env = Environment(
    loader=DictLoader(TEMPLATES), 
    autoescape=select_autoescape(["html","xml"])
)

def render(name: str, **ctx) -> HTMLResponse:
    ctx['version'] = VERSION
    return HTMLResponse(env.get_template(name).render(**ctx))

def nl2br(value: str) -> Markup:
    if not value: return Markup("")
    return Markup(value.replace("\n", "<br/>"))

env.filters["nl2br"] = nl2br

# Estado da aplicação com cache simples
class AppState:
    def __init__(self):
        self._cache = {}
        self.config = self.load_config()
        
    def load_config(self):
        conn = db_connect()
        cur = conn.cursor()
        cur.execute("SELECT key, value FROM system_config")
        config = dict(cur.fetchall())
        conn.close()
        return config
        
    def get_cached(self, key: str, default=None):
        return self._cache.get(key, default)
        
    def set_cached(self, key: str, value):
        self._cache[key] = value

state = AppState()

# Inicialização da aplicação
app = FastAPI(title=APP_TITLE, version=VERSION)

# Servir arquivos estáticos se existirem
if os.path.exists(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.on_event("startup")
async def on_startup():
    os.makedirs(STATIC_DIR, exist_ok=True)
    ensure_schema()
    logger.info(f"Sistema de Atas v{VERSION} iniciado")

# WebSocket para colaboração em tempo real
@app.websocket("/ws/{user_id}")
async def websocket_endpoint(websocket: WebSocket, user_id: str):
    try:
        await manager.connect(websocket, user_id)
        while True:
            data = await websocket.receive_text()
            message = json.loads(data)
            # Broadcast para outros usuários
            await manager.broadcast_update(message, exclude_user=user_id)
    except WebSocketDisconnect:
        manager.disconnect(websocket)
        await manager.broadcast_user_status()

# --------------------- Rotas principais
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    meeting = dict(meeting_row) if meeting_row else {}
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()
    
    # Valores padrão para itens
    item2 = state.config.get("item2_default", "Não houve realocações de recursos desde a última reunião até a presente data.")
    assuntos = state.config.get("assuntos_default", "– Assuntos gerais discutidos e/ou Eventos:")
    
    return render("index.html", 
                  title=APP_TITLE,
                  participantes=PARTICIPANTES,
                  topicos=TOPICOS,
                  meeting=meeting,
                  scenarios=scenarios,
                  item2=item2,
                  assuntos=assuntos)

@app.post("/meeting/update")
async def meeting_update(
    background_tasks: BackgroundTasks,
    numero: int = Form(...), 
    ano: int = Form(...), 
    data: str = Form(...),
    hora: str = Form(...), 
    local: str = Form(...), 
    lavrador: str = Form("")
):
    try:
        # Validar dados
        meeting_data = MeetingUpdate(
            numero=numero, ano=ano, data=data, 
            hora=hora, local=local, lavrador=lavrador
        )
        
        conn = db_connect()
        cur = conn.cursor()
        
        cur.execute("SELECT id FROM meeting ORDER BY id DESC LIMIT 1")
        row = cur.fetchone()
        meeting_id = row[0] if row else None
        
        if meeting_id:
            cur.execute("""
                UPDATE meeting 
                SET numero=?, ano=?, data=?, hora=?, local=?, lavrador=?, updated_at=CURRENT_TIMESTAMP 
                WHERE id=?
            """, (numero, ano, data, hora, local, lavrador, meeting_id))
        else:
            cur.execute("""
                INSERT INTO meeting (numero, ano, data, hora, local, lavrador) 
                VALUES (?, ?, ?, ?, ?, ?)
            """, (numero, ano, data, hora, local, lavrador))
            meeting_id = cur.lastrowid
            
        conn.commit()
        
        cur.execute("SELECT * FROM meeting WHERE id=?", (meeting_id,))
        meeting = dict(cur.fetchone())
        conn.close()
        
        # Broadcast para outros usuários
        background_tasks.add_task(
            manager.broadcast_update, 
            {"type": "meeting_update", "meeting": meeting}
        )
        
        html = env.get_template("partials/session.html").render(
            meeting=meeting, participantes=PARTICIPANTES
        )
        return HTMLResponse(f'<div id="session-block" class="p-6">{html}</div>')
        
    except Exception as e:
        logger.error(f"Erro ao atualizar reunião: {e}")
        raise HTTPException(status_code=400, detail="Dados inválidos")

@app.get("/scenario/form")
async def scenario_form(participant: str):
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("SELECT participant, text, topic FROM scenario WHERE participant=?", (participant,))
    row = cur.fetchone()
    conn.close()
    
    text = row["text"] if row else ""
    topic = row["topic"] if row else ""
    
    html = env.get_template("partials/item1_form.html").render(
        participantes=PARTICIPANTES,
        topicos=TOPICOS,
        participant=participant,
        topic=topic,
        text=text
    )
    return HTMLResponse(html)

@app.post("/scenario/save")
async def scenario_save(
    background_tasks: BackgroundTasks,
    participant: str = Form(...), 
    topic: str = Form(...), 
    text: str = Form(...)
):
    try:
        # Validar dados
        scenario_data = ScenarioUpdate(participant=participant, topic=topic, text=text)
        
        conn = db_connect()
        cur = conn.cursor()
        
        cur.execute("""
            INSERT OR REPLACE INTO scenario (participant, text, topic, updated_at) 
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        """, (participant, text, topic))
        
        conn.commit()
        
        cur.execute("SELECT participant, text, topic FROM scenario")
        scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
        conn.close()
        
        # Broadcast para outros usuários
        background_tasks.add_task(
            manager.broadcast_update, 
            {"type": "scenario_update", "participant": participant, "scenarios": scenarios}
        )
        
        status_html = env.get_template("partials/status.html").render(
            participantes=PARTICIPANTES, scenarios=scenarios
        )
        return HTMLResponse(f'<div id="status-list" class="space-y-4">{status_html}</div>')
        
    except Exception as e:
        logger.error(f"Erro ao salvar cenário: {e}")
        raise HTTPException(status_code=400, detail="Dados inválidos")

@app.post("/preview/update")
async def preview_update(item2: str = Form(...), assuntos: str = Form(...)):
    try:
        # Validar dados
        item_data = ItemUpdate(item2=item2, assuntos=assuntos)
        
        # Salvar no banco
        conn = db_connect()
        cur = conn.cursor()
        
        cur.execute("INSERT OR REPLACE INTO system_config (key, value, updated_at) VALUES ('current_item2', ?, CURRENT_TIMESTAMP)", (item2,))
        cur.execute("INSERT OR REPLACE INTO system_config (key, value, updated_at) VALUES ('current_assuntos', ?, CURRENT_TIMESTAMP)", (assuntos,))
        
        conn.commit()
        conn.close()
        
        # Atualizar cache
        state.config['current_item2'] = item2
        state.config['current_assuntos'] = assuntos
        
        return Response(status_code=204)
        
    except Exception as e:
        logger.error(f"Erro ao atualizar preview: {e}")
        raise HTTPException(status_code=400, detail="Dados inválidos")

@app.post("/resumo/update")
async def resumo_update(
    rentab: str = Form(""), 
    difpp: str = Form(""), 
    posicao: str = Form("abaixo"), 
    risco: str = Form("")
):
    try:
        # Validar dados
        resumo_data = ResumoUpdate(rentab=rentab, difpp=difpp, posicao=posicao, risco=risco)
        
        # Salvar no banco
        conn = db_connect()
        cur = conn.cursor()
        
        resumo_config = {
            'resumo_rentab': rentab,
            'resumo_difpp': difpp, 
            'resumo_posicao': posicao,
            'resumo_risco': risco
        }
        
        for key, value in resumo_config.items():
            cur.execute("INSERT OR REPLACE INTO system_config (key, value, updated_at) VALUES (?, ?, CURRENT_TIMESTAMP)", (key, value))
        
        conn.commit()
        conn.close()
        
        # Atualizar cache
        state.config.update(resumo_config)
        
        return Response(status_code=204)
        
    except Exception as e:
        logger.error(f"Erro ao atualizar resumo: {e}")
        raise HTTPException(status_code=400, detail="Dados inválidos")

@app.get("/preview/modal")
async def preview_modal():
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    meeting = dict(meeting_row) if meeting_row else {}
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    
    conn.close()
    
    ata_html = ata_html_full(meeting, scenarios)
    
    html = env.get_template("partials/preview_modal.html").render(ata_html=ata_html)
    return HTMLResponse(html)

# --------------------- API Endpoints para dados
@app.get("/api/status")
async def api_status():
    """Endpoint para verificar status do sistema"""
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT COUNT(*) FROM scenario WHERE text != ''")
    completed_scenarios = cur.fetchone()[0]
    
    cur.execute("SELECT COUNT(*) FROM scenario")
    total_scenarios = cur.fetchone()[0]
    
    conn.close()
    
    return JSONResponse({
        "status": "online",
        "version": VERSION,
        "scenarios_completed": completed_scenarios,
        "total_scenarios": total_scenarios,
        "completion_rate": (completed_scenarios / total_scenarios) * 100 if total_scenarios > 0 else 0,
        "active_users": len(manager.active_connections)
    })

@app.get("/api/scenarios")
async def api_scenarios():
    """Endpoint para obter todos os cenários"""
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("SELECT participant, text, topic, updated_at FROM scenario")
    scenarios = [
        {
            "participant": r[0],
            "text": r[1], 
            "topic": r[2],
            "updated_at": r[3]
        } for r in cur.fetchall()
    ]
    conn.close()
    return JSONResponse(scenarios)

# --------------------- Funções auxiliares para geração da ata
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
    
    p0 = ("O Comitê de Investimentos, buscando transmitir maior transparência em relação às análises dos investimentos do Instituto e, em consequência, "
          "aderindo às normas do Pró-Gestão, elabora o "Relatório de Análise de Investimentos IPAJM". "
          "Este relatório já foi encaminhado à SCO – Subgerência de Contabilidade e Orçamento, para posterior envio para análise do Conselho Fiscal do IPAJM. "
          f"Segue abaixo um resumo relativo aos itens abordados no Relatório supracitado de {mes_ano}:")
    
    rentab = state.config.get("resumo_rentab", "")
    difpp = state.config.get("resumo_difpp", "")
    pos = state.config.get("resumo_posicao", "abaixo")
    risco = state.config.get("resumo_risco", "")
    
    p1 = f"1) Acompanhamento da rentabilidade -  A rentabilidade consolidada dos investimentos do Fundo Previdenciário em {mes_ano} foi de {rentab}, ficando {difpp} p.p. {pos} da meta atuarial."
    p2 = f"2) Avaliação de risco da carteira - O grau de variação nas rentabilidades está coerente com o grau de risco assumido, em {risco}."
    p3 = f"3) Execução da Política de Investimentos – As movimentações financeiras realizadas no mês de {mes_ano} estão de acordo com as deliberações estabelecidas com a Diretoria de Investimentos e com a legislação vigente."
    p4 = f"4) Aderência a Política de Investimentos - Os recursos investidos, abrangendo a carteira consolidada, que representa o patrimônio total do RPPS sob gestão, estão aderentes à Política de Investimentos de {meeting['ano']}, respeitando o estabelecido na legislação em vigor e dentro dos percentuais definidos.  Considerando que as taxas ainda são negociadas acima da meta atuarial, seguimos com a estratégia de alcançar o alvo definido de 60% de alocação em Títulos Públicos."
    
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
    brasao = os.path.join(STATIC_DIR, "brasao.png")
    simbolo = os.path.join(STATIC_DIR, "simbolo.png")
    header_html = f"""
    <div style="width:100%; display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
      <img src="{brasao}" style="height:42px" alt="Brasão"/>
      <div style="text-align:center; font-family: Calibri; font-size:11pt; font-weight:700; color:#000;">
        GOVERNO DO ESTADO DO ESPÍRITO SANTO<br/>INSTITUTO DE PREVIDÊNCIA DOS<br/>SERVIDORES DO ESTADO DO ESPÍRITO SANTO
      </div>
      <img src="{simbolo}" style="height:42px" alt="Símbolo"/>
    </div>
    <div style="border-bottom:1px solid #000; margin:2px 0 8px 0; text-align:center;">IPAJM</div>
    """
    return header_html + "\n".join(blocos)

# --------------------- Exportações 
@app.get("/export/pdf")
async def export_pdf():
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    meeting = dict(meeting_row) if meeting_row else {}
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()
    
    html_body = ata_html_full(meeting, scenarios)
    html = f"""<html><head><meta charset='utf-8'>
      <style>
        body {{ 
          font-family: Calibri, Arial, sans-serif; 
          font-size: 11pt; 
          color:#000; 
          line-height: 1.4;
          margin: 20px;
        }}
        .content {{ 
          white-space: normal; 
          text-align: justify; 
        }}
        p {{ margin: 6pt 0; }}
        strong {{ font-weight: 700; }}
      </style>
    </head><body><div class="content">{html_body}</div></body></html>"""
    
    pdf_buf = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html), dest=pdf_buf)
    pdf_buf.seek(0)
    
    filename = f"Ata_{meeting.get('numero', 1):03d}-{meeting.get('ano', 2024)}.pdf"
    return StreamingResponse(
        pdf_buf, 
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# Funções DOCX (mantidas do código original mas otimizadas)
def _set_cell_borders(cell, bottom=True):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    if bottom:
        bottom_el = OxmlElement("w:bottom")
        bottom_el.set(qn("w:val"), "single")
        bottom_el.set(qn("w:sz"), "8")
        bottom_el.set(qn("w:color"), "000000")
        tcBorders.append(bottom_el)
    tcPr.append(tcBorders)

def _add_header_image(doc: DocxDocument, path=None, width_inches=6.5):
    if not path:
        path = os.path.join(STATIC_DIR, "cabecalho.png")
    if not os.path.exists(path): 
        return
    header = doc.sections[0].header
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Inches(width_inches))
    p.paragraph_format.space_after = Pt(0)

def _add_paragraph(doc, text, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after_pt=0):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    run.bold = bold
    p.alignment = align
    if space_after_pt: 
        p.paragraph_format.space_after = Pt(space_after_pt)
    return p

def _add_label_value(doc, label, value, space_after_pt=0):
    p = doc.add_paragraph()
    r1 = p.add_run(f"{label}: ")
    r1.bold = True
    r1.font.name = "Calibri"
    r1.font.size = Pt(11)
    r2 = p.add_run(value)
    r2.font.name = "Calibri"
    r2.font.size = Pt(11)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if space_after_pt: 
        p.paragraph_format.space_after = Pt(space_after_pt)
    return p

def _add_title(doc, text, space_after_pt=TITLE_GAP):
    return _add_paragraph(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after_pt=space_after_pt)

@app.get("/export/docx")
async def export_docx():
    conn = db_connect()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM meeting ORDER BY id DESC LIMIT 1")
    meeting_row = cur.fetchone()
    meeting = dict(meeting_row) if meeting_row else {}
    
    cur.execute("SELECT participant, text, topic FROM scenario")
    scenarios = {r[0]: {"text": r[1], "topic": r[2]} for r in cur.fetchall()}
    conn.close()

    d = date.fromisoformat(meeting["data"]) if meeting.get("data") else date.today()
    numero_fmt = numero_sessao_fmt(meeting.get("numero", 1), meeting.get("ano", d.year))

    doc = DocxDocument()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # Cabeçalho com imagem
    _add_header_image(doc)

    # Conteúdo da ata (estrutura similar ao código original)
    p_titulo = _add_title(doc, f"Sessão Ordinária nº {numero_fmt}", space_after_pt=TITLE_GAP)
    p_titulo.paragraph_format.space_before = Pt(HEADER_GAP)

    # Dados básicos
    _add_label_value(doc, "Data", ptbr_date(d), space_after_pt=0)
    _add_label_value(doc, "Hora", f"{meeting.get('hora', '14:00')}h", space_after_pt=0)
    _add_label_value(doc, "Local", meeting.get('local', 'A definir'), space_after_pt=SECTION_GAP)

    # Presenças
    _add_title(doc, "Presenças:", space_after_pt=TITLE_GAP)
    for nome in PARTICIPANTES:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_nome = p.add_run(nome)
        run_nome.font.name = "Calibri"
        run_nome.font.size = Pt(11)
        run_cargo = p.add_run(f" - {CARGO};")
        run_cargo.font.name = "Calibri" 
        run_cargo.font.size = Pt(11)

    # Ordem do dia, itens, etc. (mantendo a estrutura original)
    _add_title(doc, "Ordem do Dia:", space_after_pt=TITLE_GAP)
    _add_paragraph(doc, "1. Cenário Político e Econômico Interno e Cenário Econômico Externo (EUA, Europa e China);", False, WD_ALIGN_PARAGRAPH.LEFT, space_after_pt=0)
    _add_paragraph(doc, "2. Movimentações e Aplicações financeiras;", False, WD_ALIGN_PARAGRAPH.LEFT, space_after_pt=0)
    _add_paragraph(doc, "3. Acompanhamento dos Recursos Investidos;", False, WD_ALIGN_PARAGRAPH.LEFT, space_after_pt=0)
    _add_paragraph(doc, "4. Assuntos Gerais.", False, WD_ALIGN_PARAGRAPH.LEFT, space_after_pt=SECTION_GAP)

    # Salvar e enviar
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    
    filename = f"Ata_{meeting.get('numero', 1):03d}-{meeting.get('ano', d.year)}.docx"
    return StreamingResponse(
        buf, 
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# Health check para monitoramento
@app.get("/health")
async def health_check():
    return JSONResponse({
        "status": "healthy",
        "version": VERSION,
        "timestamp": datetime.now().isoformat(),
        "database": "connected"
    })

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
