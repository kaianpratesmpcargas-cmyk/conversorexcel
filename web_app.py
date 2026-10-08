"""
MP CARGAS — Conversor SSW + Distribuição por Base
Servidor local. O envio de WhatsApp é feito pelo bot existente através de HTTP.
"""
import io
import json
import os
import sys
import uuid
from urllib.parse import quote
from datetime import datetime

import requests
from flask import Flask, request, jsonify, send_file, render_template_string

from parser_ssw import parse_ssw_table
from excel_export import export_to_excel

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
def get_data_file(filename):
    if os.getenv("VERCEL"):
        target = os.path.join("/tmp", filename)
        if not os.path.exists(target):
            source = os.path.join(BASE_DIR, filename)
            if os.path.exists(source):
                try:
                    import shutil
                    shutil.copy2(source, target)
                except Exception:
                    pass
        return target
    return os.path.join(BASE_DIR, filename)

MOTORISTAS_FILE = get_data_file("motoristas.json")
FILA_FILE = get_data_file("fila_tabelas.json")
WHATSAPP_API_URL = os.getenv("WHATSAPP_API_URL", "").strip()
WHATSAPP_API_TOKEN = os.getenv("WHATSAPP_API_TOKEN", "").strip()

app = Flask(__name__)


def load_json(path, default):
    try:
        if os.getenv("VERCEL") and not os.path.exists(path):
            base_path = os.path.join(BASE_DIR, os.path.basename(path))
            if os.path.exists(base_path):
                path = base_path
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return default


def save_json(path, data):
    try:
        temp = path + ".tmp"
        with open(temp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(temp, path)
    except OSError:
        tmp_target = os.path.join("/tmp", os.path.basename(path))
        temp = tmp_target + ".tmp"
        with open(temp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(temp, tmp_target)


def get_motoristas():
    return load_json(MOTORISTAS_FILE, {})


def get_fila():
    return load_json(FILA_FILE, [])


def whatsapp_configurado():
    return bool(WHATSAPP_API_URL)


def format_whatsapp_message(item):
    """Monta uma mensagem curta para WhatsApp contendo somente as NFs atrasadas."""
    records = item.get("records", [])

    notas = []
    for r in records:
        nf = str(r.get("N FISCAL") or "").strip()
        atraso = r.get("DIAS DE ATRASO")

        # Somente notas que realmente possuem atraso.
        try:
            dias = float(str(atraso).replace(",", ".")) if atraso not in (None, "") else 0
        except (TypeError, ValueError):
            dias = 0

        if nf and dias > 0:
            notas.append(nf)

    # Remove duplicadas mantendo a ordem da tabela.
    notas = list(dict.fromkeys(notas))

    lines = ["*NOTAS ATRASADAS:*"]
    if notas:
        lines.extend(f"{i}. {nf}" for i, nf in enumerate(notas, 1))
    else:
        lines.append("Nenhuma nota atrasada.")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------
HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MP CARGAS — SSW + WhatsApp</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
:root { --mp-yellow:#FFD100; --mp-black:#111111; --mp-white:#fff; --line:#e5e7eb; }
body { background:#f5f6f8; color:#111; font-family:Segoe UI,Arial,sans-serif; }
.navbar-mp { background:var(--mp-black); color:#fff; border-bottom:4px solid var(--mp-yellow); }
.brand { font-weight:800; letter-spacing:.4px; }
.card { border:0; border-radius:12px; box-shadow:0 3px 14px rgba(0,0,0,.07); }
.btn-mp { background:var(--mp-yellow); border:1px solid #d8b300; color:#111; font-weight:800; }
.btn-mp:hover { background:#e8bd00; color:#111; }
.btn-dark-mp { background:#111; color:#fff; font-weight:700; }
.btn-dark-mp:hover { background:#222; color:#fff; }
.badge-base { background:#111; color:#FFD100; }
.table thead th { background:#111; color:#fff; white-space:nowrap; }
.table-responsive { max-height:420px; overflow:auto; }
.status { font-size:.78rem; font-weight:800; padding:.35rem .55rem; border-radius:20px; }
.status-pronta { background:#e7f7ed; color:#18733a; }
.status-enviando { background:#fff4cc; color:#785b00; }
.status-enviada { background:#e7f7ed; color:#18733a; }
.status-erro { background:#fde8e8; color:#a11; }
.status-pendente { background:#eef2f7; color:#455; }
.metric { border-left:4px solid #FFD100; }
textarea { min-height:250px; resize:vertical; }
.small-muted { color:#6b7280; font-size:.88rem; }
</style>
</head>
<body>
<nav class="navbar navbar-mp px-3 py-3">
  <div class="container-fluid">
    <div class="brand fs-5">MP CARGAS <span class="fw-normal text-white-50">| Conversor SSW + Distribuição WhatsApp</span></div>
    <div class="small">Servidor local</div>
  </div>
</nav>

<main class="container-fluid px-3 px-lg-4 py-4">
  <div class="row g-4">
    <div class="col-xl-7">
      <div class="card p-4">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <div>
            <h4 class="mb-1 fw-bold">Nova tabela</h4>
            <div class="small-muted">Cole a tabela do SSW e escolha a base. A base NÃO é descoberta automaticamente.</div>
          </div>
          <span class="badge badge-base rounded-pill px-3 py-2">1 tabela por vez</span>
        </div>

        <label class="form-label fw-bold">Base</label>
        <select id="base" class="form-select form-select-lg mb-3"></select>
        <div id="destino-info" class="alert alert-light border mb-3">Selecione uma base.</div>

        <label class="form-label fw-bold">Tabela copiada do SSW</label>
        <textarea id="raw-text" class="form-control font-monospace" placeholder="Cole aqui a tabela exatamente como vem do SSW..."></textarea>

        <div class="d-flex gap-2 mt-3 flex-wrap">
          <button id="btn-paste" class="btn btn-outline-dark fw-bold">Colar</button>
          <button id="btn-save" class="btn btn-mp px-4">PROCESSAR E SALVAR NA FILA</button>
          <button id="btn-clear" class="btn btn-outline-secondary">Limpar</button>
        </div>
        <div id="error" class="alert alert-danger mt-3 d-none"></div>
      </div>

      <div class="card mt-4 p-4">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <div>
            <h4 class="mb-1 fw-bold">Fila de envio</h4>
            <div class="small-muted">Cada linha representa uma tabela/base independente.</div>
          </div>
          <div id="queue-count" class="fw-bold">0 tabelas</div>
        </div>
        <div class="table-responsive">
          <table class="table align-middle mb-0">
            <thead><tr><th>Base</th><th>Motorista</th><th>WhatsApp</th><th>CT-RCs</th><th>Status</th><th></th></tr></thead>
            <tbody id="queue-body"></tbody>
          </table>
        </div>
        <div class="d-flex justify-content-between align-items-center mt-3 flex-wrap gap-2">
          <button id="btn-send-all" class="btn btn-lg btn-success fw-bold px-4" onclick="abrirTodosWhatsApp()">📱 ENVIAR TABELAS NO WHATSAPP</button>
          <div id="whatsapp-status" class="small-muted">Cada linha também possui seu próprio botão para abrir o WhatsApp do motorista.</div>
        </div>
        <div id="send-result" class="mt-3"></div>
      </div>
    </div>

    <div class="col-xl-5">
      <div class="row g-3 mb-4">
        <div class="col-6"><div class="card metric p-3"><div class="small-muted">Tabelas na fila</div><div id="m-tabelas" class="fs-3 fw-bold">0</div></div></div>
        <div class="col-6"><div class="card metric p-3"><div class="small-muted">CT-RCs na fila</div><div id="m-ctrc" class="fs-3 fw-bold">0</div></div></div>
      </div>
      <div class="card p-4">
        <h5 class="fw-bold">Cadastro das bases</h5>
        <div class="small-muted mb-3">Esses dados já estão gravados no projeto.</div>
        <div id="drivers"></div>
      </div>
      <div class="card p-4 mt-4">
        <h5 class="fw-bold">Como funciona</h5>
        <ol class="mb-0 ps-3">
          <li class="mb-2">Cole uma tabela.</li>
          <li class="mb-2">Escolha a base.</li>
          <li class="mb-2">Clique em <b>Processar e Salvar</b>.</li>
          <li class="mb-2">Repita para as demais bases.</li>
          <li>Clique em <b>Enviar Todas</b>.</li>
        </ol>
      </div>
    </div>
  </div>
</main>

<script>
let bases = {};

async function api(url, options={}) {
  const r = await fetch(url, options);
  const d = await r.json();
  if (!r.ok) throw new Error(d.error || 'Erro na operação.');
  return d;
}

function esc(v) {
  return String(v ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
}

async function carregar() {
  const d = await api('/api/config');
  bases = d.motoristas;
  const select = document.getElementById('base');
  select.innerHTML = Object.keys(bases).sort().map(b => `<option value="${esc(b)}">${esc(b)}</option>`).join('');
  renderInfo();
  renderDrivers();
  renderFila(d.fila);
}

function renderInfo() {
  const b = document.getElementById('base').value;
  const x = bases[b];
  document.getElementById('destino-info').innerHTML = x
    ? `<b>${esc(b)}</b> → ${esc(x.motorista)} → <b>${esc(x.whatsapp)}</b>`
    : 'Selecione uma base.';
}

document.getElementById('base').addEventListener('change', renderInfo);

document.getElementById('btn-paste').addEventListener('click', async () => {
  try { document.getElementById('raw-text').value = await navigator.clipboard.readText(); }
  catch(e) { alert('Use Ctrl+V dentro da caixa de texto.'); }
});

document.getElementById('btn-clear').addEventListener('click', () => document.getElementById('raw-text').value='');

document.getElementById('btn-save').addEventListener('click', async () => {
  const raw = document.getElementById('raw-text').value.trim();
  const base = document.getElementById('base').value;
  const err = document.getElementById('error'); err.classList.add('d-none');
  if (!raw) { err.textContent='Cole uma tabela antes de salvar.'; err.classList.remove('d-none'); return; }
  try {
    await api('/api/queue', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({base, raw_text:raw})});
    document.getElementById('raw-text').value='';
    await refreshFila();
  } catch(e) { err.textContent=e.message; err.classList.remove('d-none'); }
});

async function refreshFila() { const d=await api('/api/config'); renderFila(d.fila); }

function renderFila(fila) {
  document.getElementById('queue-count').textContent = `${fila.length} tabela${fila.length===1?'':'s'}`;
  document.getElementById('m-tabelas').textContent = fila.length;
  document.getElementById('m-ctrc').textContent = fila.reduce((s,x)=>s+(x.records||[]).length,0);
  const body=document.getElementById('queue-body');
  body.innerHTML='';
  fila.forEach(x=>{
    const cls='status-'+String(x.status||'pendente').toLowerCase().replace(' ','-');
    body.innerHTML += `<tr>
      <td><b>${esc(x.base)}</b></td><td>${esc(x.motorista)}</td><td>${esc(x.whatsapp)}</td>
      <td>${(x.records||[]).length}</td><td><span class="status ${cls}">${esc(x.status)}</span></td>
      <td class="d-flex gap-1">
        <button class="btn btn-success fw-bold" onclick="openWhatsApp('${x.id}')">📱 ENVIAR WHATSAPP</button>
        <button class="btn btn-sm btn-outline-danger" onclick="removeItem('${x.id}')">Excluir</button>
      </td>
    </tr>`;
  });
}

async function removeItem(id) {
  if(!confirm('Excluir esta tabela da fila?')) return;
  try { await api('/api/queue/'+id,{method:'DELETE'}); await refreshFila(); } catch(e) { alert(e.message); }
}

function renderDrivers() {
  const el=document.getElementById('drivers'); el.innerHTML='';
  Object.keys(bases).sort().forEach(b=>{
    const x=bases[b]; el.innerHTML += `<div class="border-bottom py-2"><b>${esc(b)}</b> — ${esc(x.motorista)}<br><span class="small-muted">${esc(x.whatsapp)}</span></div>`;
  });
}

async function openWhatsApp(id) {
  // Abre a nova aba imediatamente no clique, antes do fetch, para o navegador não considerar popup.
  const novaAba = window.open('about:blank', '_blank');
  if (!novaAba) {
    alert('O navegador bloqueou a nova aba. Permita pop-ups para este aplicativo.');
    return;
  }
  try {
    const d = await api('/api/whatsapp/link/' + encodeURIComponent(id));
    novaAba.location.href = d.url;
    novaAba.focus();
  } catch(e) {
    novaAba.close();
    alert(e.message);
  }
}

async function abrirTodosWhatsApp() {
  try {
    const d = await api('/api/config');
    if (!d.fila.length) { alert('Não há tabelas na fila.'); return; }
    alert('As tabelas serão abertas uma por vez. No WhatsApp, confira a conversa e clique em Enviar.');
    for (const item of d.fila) {
      const link = await api('/api/whatsapp/link/' + encodeURIComponent(item.id));
      window.open(link.url, '_blank');
      await new Promise(r => setTimeout(r, 700));
    }
  } catch(e) { alert(e.message); }
}

carregar();
</script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/api/config')
def config():
    return jsonify({'motoristas': get_motoristas(), 'fila': get_fila()})


@app.route('/api/whatsapp/status')
def whatsapp_status():
    return jsonify({'configurado': whatsapp_configurado(), 'url_configurada': bool(WHATSAPP_API_URL)})


@app.route('/api/process', methods=['POST'])
def process():
    data = request.get_json() or {}
    raw_text = data.get('raw_text', '')
    records, err = parse_ssw_table(raw_text)
    if err:
        return jsonify({'error': err}), 400
    total = len(records)
    metrics = {
        'total': total,
        'atrasados': sum(1 for r in records if r.get('SITUAÇÃO') == 'ATRASADO'),
        'no_prazo': sum(1 for r in records if r.get('SITUAÇÃO') == 'NO PRAZO'),
        'sem_info': sum(1 for r in records if r.get('SITUAÇÃO') == 'SEM INFORMAÇÃO'),
        'peso_total': sum(r.get('PESO', 0.0) or 0.0 for r in records),
        'frete_total': sum(r.get('FRETE', 0.0) or 0.0 for r in records)
    }
    return jsonify({'records': records, 'metrics': metrics})


@app.route('/api/queue', methods=['POST'])
def queue_table():
    data = request.get_json() or {}
    base = str(data.get('base', '')).strip().upper()
    raw_text = str(data.get('raw_text', '')).strip()
    if not base or not raw_text:
        return jsonify({'error': 'Informe a base e cole a tabela.'}), 400

    motoristas = get_motoristas()
    if base not in motoristas:
        return jsonify({'error': f'Base {base} não está cadastrada.'}), 400

    records, err = parse_ssw_table(raw_text)
    if err:
        return jsonify({'error': err}), 400
    if not records:
        return jsonify({'error': 'Nenhum registro foi encontrado na tabela.'}), 400

    info = motoristas[base]
    fila = get_fila()
    item = {
        'id': uuid.uuid4().hex,
        'created_at': datetime.now().isoformat(timespec='seconds'),
        'base': base,
        'motorista': info['motorista'],
        'whatsapp': info['whatsapp'],
        'records': records,
        'status': 'PENDENTE',
        'error': None,
        'sent_at': None
    }
    fila.append(item)
    save_json(FILA_FILE, fila)
    return jsonify({'ok': True, 'item': item})


@app.route('/api/queue/<item_id>', methods=['DELETE'])
def delete_queue_item(item_id):
    fila = get_fila()
    nova = [x for x in fila if x.get('id') != item_id]
    if len(nova) == len(fila):
        return jsonify({'error': 'Tabela não encontrada.'}), 404
    save_json(FILA_FILE, nova)
    return jsonify({'ok': True})


@app.route('/api/whatsapp/link/<item_id>')
def whatsapp_link(item_id):
    fila = get_fila()
    item = next((x for x in fila if x.get('id') == item_id), None)
    if not item:
        return jsonify({'error': 'Tabela não encontrada.'}), 404

    numero = ''.join(ch for ch in str(item.get('whatsapp', '')) if ch.isdigit())
    if not numero:
        return jsonify({'error': 'Número do motorista não cadastrado.'}), 400

    # Garante DDI 55 para os números cadastrados sem DDI.
    if len(numero) in (10, 11):
        numero = '55' + numero

    mensagem = format_whatsapp_message(item)
    url = f"https://wa.me/{numero}?text={quote(mensagem, safe='')}"
    return jsonify({'ok': True, 'url': url, 'numero': numero})


def send_to_whatsapp(item):
    if not WHATSAPP_API_URL:
        return False, 'WHATSAPP_API_URL não configurada.'

    payload = {
        'to': item['whatsapp'],
        'message': format_whatsapp_message(item),
        'base': item['base'],
        'motorista': item['motorista'],
        'queue_id': item['id'],
        'idempotency_key': f"mp-cargas-{item['id']}"
    }
    headers = {'Content-Type': 'application/json'}
    if WHATSAPP_API_TOKEN:
        headers['Authorization'] = f'Bearer {WHATSAPP_API_TOKEN}'

    try:
        r = requests.post(WHATSAPP_API_URL, json=payload, headers=headers, timeout=20)
        if 200 <= r.status_code < 300:
            return True, None
        return False, f'Bot retornou HTTP {r.status_code}: {r.text[:300]}'
    except requests.RequestException as exc:
        return False, f'Falha de conexão com o bot: {exc}'


@app.route('/api/send-all', methods=['POST'])
def send_all():
    fila = get_fila()
    if not fila:
        return jsonify({'error': 'A fila está vazia.'}), 400
    if not WHATSAPP_API_URL:
        return jsonify({'error': 'O projeto está pronto, mas a integração do seu bot ainda precisa ser ligada. Configure WHATSAPP_API_URL apontando para o endpoint de envio do bot.'}), 400

    sent = 0
    failed = 0
    for item in fila:
        if item.get('status') == 'ENVIADA':
            continue
        item['status'] = 'ENVIANDO'
        ok, error = send_to_whatsapp(item)
        if ok:
            item['status'] = 'ENVIADA'
            item['sent_at'] = datetime.now().isoformat(timespec='seconds')
            item['error'] = None
            sent += 1
        else:
            item['status'] = 'ERRO'
            item['error'] = error
            failed += 1
        save_json(FILA_FILE, fila)

    msg = f'{sent} tabela(s) enviada(s).'
    if failed:
        msg += f' {failed} tabela(s) com erro ficaram marcadas para revisão.'
    return jsonify({'ok': True, 'sent': sent, 'failed': failed, 'message': msg})


@app.route('/api/export', methods=['POST'])
def export_file():
    data = request.get_json() or {}
    records = data.get('records', [])
    if not records:
        return jsonify({'error': 'Nenhum registro para exportar.'}), 400
    import tempfile
    temp_path = os.path.join(tempfile.gettempdir(), f'_temp_relatorio_{os.getpid()}.xlsx')
    try:
        export_to_excel(records, temp_path)
        with open(temp_path, 'rb') as f:
            file_data = io.BytesIO(f.read())
        today = datetime.now().strftime('%Y-%m-%d')
        return send_file(file_data, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'MP_CARGAS_SSW_{today}.xlsx')
    finally:
        if os.path.exists(temp_path):
            try: os.remove(temp_path)
            except OSError: pass


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=False)
