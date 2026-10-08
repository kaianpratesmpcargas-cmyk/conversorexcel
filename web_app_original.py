"""
Servidor Web Local (Localhost) — MP CARGAS Conversor SSW
Permite acessar o sistema pelo navegador (http://localhost:5000).
"""
import io
import re
import os
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string
import pandas as pd
from parser_ssw import parse_ssw_table
from excel_export import export_to_excel

app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MP CARGAS — Conversor SSW</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" rel="stylesheet">
    <style>
        body { background-color: #f8fafc; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .navbar-brand { font-weight: 800; letter-spacing: 0.5px; }
        .card-stat { border-radius: 10px; border: none; box-shadow: 0 2px 6px rgba(0,0,0,0.06); transition: transform 0.2s; }
        .card-stat:hover { transform: translateY(-2px); }
        .bg-atrasado { background-color: #ffebee !important; color: #c62828 !important; }
        .bg-noprazo { background-color: #e8f5e9 !important; color: #2e7d32 !important; }
        .bg-seminfo { background-color: #f5f5f5 !important; color: #616161 !important; }
        .badge-atrasado { background-color: #fde8e8; color: #9b1c1c; border: 1px solid #f8b4b4; }
        .badge-noprazo { background-color: #def7ec; color: #03543f; border: 1px solid #84e1bc; }
        .badge-seminfo { background-color: #f3f4f6; color: #374151; border: 1px solid #e5e7eb; }
        .table-responsive { max-height: 480px; overflow-y: auto; }
        .table thead th { position: sticky; top: 0; background: #1a365d; color: white; z-index: 10; font-size: 0.85rem; }
    </style>
</head>
<body>
    <nav class="navbar navbar-dark" style="background-color: #1a365d;">
        <div class="container-fluid px-4 py-2">
            <span class="navbar-brand mb-0 h1"><i class="fa-solid fa-truck-fast me-2"></i>MP CARGAS <small class="text-white-50 fs-6 ms-2">| Conversor SSW → Excel</small></span>
            <span class="text-light small"><i class="fa-solid fa-circle text-success me-1"></i> Servidor Local Ativo</span>
        </div>
    </nav>

    <div class="container-fluid px-4 py-4">
        <!-- ÁREA DE ENTRADA -->
        <div id="input-section" class="card shadow-sm border-0 mb-4">
            <div class="card-body p-4">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h5 class="card-title fw-bold text-dark mb-0"><i class="fa-solid fa-paste text-primary me-2"></i>Cole a tabela copiada do SSW</h5>
                    <button class="btn btn-outline-primary btn-sm" id="btn-paste-clipboard">
                        <i class="fa-regular fa-clipboard me-1"></i> Colar da Área de Transferência
                    </button>
                </div>
                <div class="form-group mb-3">
                    <textarea id="raw-text" class="form-control font-monospace" rows="9" placeholder="Cole aqui os dados copiados do SSW (| CTRC | Emissão | ... |)"></textarea>
                </div>
                <div class="d-flex gap-2">
                    <button id="btn-process" class="btn btn-primary px-4 fw-bold" style="background-color: #1a365d; border-color: #1a365d;">
                        <i class="fa-solid fa-play me-1"></i> PROCESSAR TABELA
                    </button>
                    <button id="btn-clear" class="btn btn-outline-secondary px-3">
                        <i class="fa-solid fa-eraser me-1"></i> Limpar
                    </button>
                </div>
                <div id="error-alert" class="alert alert-danger mt-3 d-none" role="alert"></div>
            </div>
        </div>

        <!-- ÁREA DE RESULTADOS -->
        <div id="result-section" class="d-none">
            <!-- CARDS DE MÉTRICAS -->
            <div class="row g-3 mb-4">
                <div class="col-md-2 col-6">
                    <div class="card card-stat p-3 bg-white">
                        <div class="text-muted small fw-bold">TOTAL DE CT-RCs</div>
                        <div class="fs-3 fw-bold text-dark mt-1" id="m-total">0</div>
                    </div>
                </div>
                <div class="col-md-2 col-6">
                    <div class="card card-stat p-3 bg-atrasado">
                        <div class="small fw-bold">ATRASADOS</div>
                        <div class="fs-3 fw-bold mt-1" id="m-atrasados">0</div>
                    </div>
                </div>
                <div class="col-md-2 col-6">
                    <div class="card card-stat p-3 bg-noprazo">
                        <div class="small fw-bold">NO PRAZO</div>
                        <div class="fs-3 fw-bold mt-1" id="m-noprazo">0</div>
                    </div>
                </div>
                <div class="col-md-2 col-6">
                    <div class="card card-stat p-3 bg-seminfo">
                        <div class="small fw-bold">SEM INFORMAÇÃO</div>
                        <div class="fs-3 fw-bold mt-1" id="m-seminfo">0</div>
                    </div>
                </div>
                <div class="col-md-2 col-6">
                    <div class="card card-stat p-3 bg-white">
                        <div class="text-muted small fw-bold">PESO TOTAL</div>
                        <div class="fs-4 fw-bold text-primary mt-1" id="m-peso">0,00 KG</div>
                    </div>
                </div>
                <div class="col-md-2 col-6">
                    <div class="card card-stat p-3 bg-white">
                        <div class="text-muted small fw-bold">FRETE TOTAL</div>
                        <div class="fs-4 fw-bold text-primary mt-1" id="m-frete">R$ 0,00</div>
                    </div>
                </div>
            </div>

            <!-- CONTROLES, FILTROS E EXPORTAÇÃO -->
            <div class="card shadow-sm border-0 mb-4">
                <div class="card-body p-3">
                    <div class="row g-2 align-items-center">
                        <div class="col-md-3">
                            <div class="input-group">
                                <span class="input-group-text bg-white"><i class="fa-solid fa-magnifying-glass text-muted"></i></span>
                                <input type="text" id="filtro-busca" class="form-control" placeholder="Buscar CTRC, NF, Destino...">
                            </div>
                        </div>
                        <div class="col-md-2">
                            <select id="filtro-situacao" class="form-select">
                                <option value="TODOS">Situação: TODAS</option>
                                <option value="ATRASADO">ATRASADO</option>
                                <option value="NO PRAZO">NO PRAZO</option>
                                <option value="SEM INFORMAÇÃO">SEM INFORMAÇÃO</option>
                            </select>
                        </div>
                        <div class="col-md-2">
                            <select id="filtro-destino" class="form-select">
                                <option value="TODOS">Destino: TODOS</option>
                            </select>
                        </div>
                        <div class="col-md-5 d-flex justify-content-end gap-2">
                            <button id="btn-export-excel" class="btn btn-success fw-bold">
                                <i class="fa-solid fa-file-excel me-1"></i> EXPORTAR EXCEL (.xlsx)
                            </button>
                            <button id="btn-new-import" class="btn btn-secondary fw-bold">
                                <i class="fa-solid fa-rotate-left me-1"></i> Nova Importação
                            </button>
                        </div>
                    </div>
                </div>
            </div>

            <!-- TABELA DE RESULTADOS -->
            <div class="card shadow-sm border-0">
                <div class="card-body p-0">
                    <div class="table-responsive">
                        <table class="table table-hover table-striped align-middle mb-0" id="tabela-dados">
                            <thead>
                                <tr>
                                    <th>CTRC</th>
                                    <th>DATA DE EMISSÃO</th>
                                    <th>NOTA FISCAL</th>
                                    <th class="text-center">DIAS DE ATRASO</th>
                                    <th class="text-center">SITUAÇÃO</th>
                                </tr>
                            </thead>
                            <tbody id="tabela-corpo"></tbody>
                        </table>
                    </div>
                    <div class="p-3 bg-light border-top small text-muted d-flex justify-content-between">
                        <span id="registros-contador">Exibindo 0 registros</span>
                        <span>MP CARGAS — Sistema Local</span>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        let allRecords = [];

        // Colar da Área de Transferência
        document.getElementById('btn-paste-clipboard').addEventListener('click', async () => {
            try {
                const text = await navigator.clipboard.readText();
                document.getElementById('raw-text').value = text;
            } catch (err) {
                alert("Para colar, utilize o atalho Ctrl + V diretamente dentro da caixa de texto.");
            }
        });

        // Limpar
        document.getElementById('btn-clear').addEventListener('click', () => {
            document.getElementById('raw-text').value = '';
            document.getElementById('error-alert').classList.add('d-none');
        });

        // Processar
        document.getElementById('btn-process').addEventListener('click', async () => {
            const rawText = document.getElementById('raw-text').value.trim();
            const errBox = document.getElementById('error-alert');
            errBox.classList.add('d-none');

            if (!rawText) {
                errBox.textContent = "Por favor, cole a tabela antes de processar.";
                errBox.classList.remove('d-none');
                return;
            }

            try {
                const res = await fetch('/api/process', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ raw_text: rawText })
                });

                const data = await res.json();
                if (!res.ok || data.error) {
                    errBox.textContent = data.error || "Erro ao processar dados.";
                    errBox.classList.remove('d-none');
                    return;
                }

                allRecords = data.records;
                renderDashboard(data.metrics, allRecords);
                populateDestinos(allRecords);
                applyFilters();

                document.getElementById('result-section').classList.remove('d-none');
                document.getElementById('result-section').scrollIntoView({ behavior: 'smooth' });
            } catch (err) {
                errBox.textContent = "Falha ao conectar com o serviço: " + err.message;
                errBox.classList.remove('d-none');
            }
        });

        // Nova Importação
        document.getElementById('btn-new-import').addEventListener('click', () => {
            document.getElementById('result-section').classList.add('d-none');
            window.scrollTo({ top: 0, behavior: 'smooth' });
        });

        // Exportar Excel
        document.getElementById('btn-export-excel').addEventListener('click', async () => {
            if (!allRecords.length) return;
            const res = await fetch('/api/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ records: allRecords })
            });

            if (res.ok) {
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                const d = new Date().toISOString().slice(0, 10);
                a.download = `MP_CARGAS_SSW_${d}.xlsx`;
                document.body.appendChild(a);
                a.click();
                a.remove();
            } else {
                alert("Erro ao gerar o arquivo Excel.");
            }
        });

        function renderDashboard(metrics, records) {
            document.getElementById('m-total').textContent = metrics.total.toLocaleString('pt-BR');
            document.getElementById('m-atrasados').textContent = metrics.atrasados.toLocaleString('pt-BR');
            document.getElementById('m-noprazo').textContent = metrics.no_prazo.toLocaleString('pt-BR');
            document.getElementById('m-seminfo').textContent = metrics.sem_info.toLocaleString('pt-BR');
            document.getElementById('m-peso').textContent = metrics.peso_total.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " KG";
            document.getElementById('m-frete').textContent = "R$ " + metrics.frete_total.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }

        function populateDestinos(records) {
            const select = document.getElementById('filtro-destino');
            select.innerHTML = '<option value="TODOS">Destino: TODOS</option>';
            const destinos = [...new Set(records.map(r => r['DESTINO']).filter(Boolean))].sort();
            destinos.forEach(d => {
                const opt = document.createElement('option');
                opt.value = d;
                opt.textContent = d;
                select.appendChild(opt);
            });
        }

        // Filtros
        document.getElementById('filtro-busca').addEventListener('input', applyFilters);
        document.getElementById('filtro-situacao').addEventListener('change', applyFilters);
        document.getElementById('filtro-destino').addEventListener('change', applyFilters);

        function applyFilters() {
            const query = document.getElementById('filtro-busca').value.toLowerCase().trim();
            const sit = document.getElementById('filtro-situacao').value;
            const dest = document.getElementById('filtro-destino').value;

            const filtered = allRecords.filter(r => {
                if (sit !== 'TODOS' && r['SITUAÇÃO'] !== sit) return false;
                if (dest !== 'TODOS' && r['DESTINO'] !== dest) return false;
                if (query) {
                    const rowText = `${r['CTRC']} ${r['N FISCAL']} ${r['REMETENTE']} ${r['DESTINATÁRIO']} ${r['DESTINO']}`.toLowerCase();
                    if (!rowText.includes(query)) return false;
                }
                return true;
            });

            const tbody = document.getElementById('tabela-corpo');
            tbody.innerHTML = '';

            filtered.forEach(r => {
                const tr = document.createElement('tr');
                const sitClass = r['SITUAÇÃO'] === 'ATRASADO' ? 'badge-atrasado' : (r['SITUAÇÃO'] === 'NO PRAZO' ? 'badge-noprazo' : 'badge-seminfo');
                const diasAtraso = r['DIAS DE ATRASO'] !== null && r['DIAS DE ATRASO'] !== undefined ? r['DIAS DE ATRASO'] : '-';
                
                tr.innerHTML = `
                    <td class="fw-bold text-nowrap">${r['CTRC'] || ''}</td>
                    <td class="text-nowrap">${r['DATA EMISSÃO'] || ''}</td>
                    <td class="fw-bold">${r['N FISCAL'] || ''}</td>
                    <td class="text-center fw-bold fs-6">${diasAtraso}</td>
                    <td class="text-center"><span class="badge ${sitClass} px-2 py-1">${r['SITUAÇÃO']}</span></td>
                `;
                tbody.appendChild(tr);
            });

            document.getElementById('registros-contador').textContent = `Exibindo ${filtered.length} de ${allRecords.length} registros`;
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/process', methods=['POST'])
def process():
    data = request.get_json() or {}
    raw_text = data.get('raw_text', '')
    records, err = parse_ssw_table(raw_text)

    if err:
        return jsonify({'error': err}), 400

    # Indicadores
    total = len(records)
    atrasados = sum(1 for r in records if r.get('SITUAÇÃO') == 'ATRASADO')
    no_prazo = sum(1 for r in records if r.get('SITUAÇÃO') == 'NO PRAZO')
    sem_info = sum(1 for r in records if r.get('SITUAÇÃO') == 'SEM INFORMAÇÃO')
    peso_total = sum(r.get('PESO', 0.0) or 0.0 for r in records)
    frete_total = sum(r.get('FRETE', 0.0) or 0.0 for r in records)

    metrics = {
        'total': total,
        'atrasados': atrasados,
        'no_prazo': no_prazo,
        'sem_info': sem_info,
        'peso_total': peso_total,
        'frete_total': frete_total
    }

    return jsonify({'records': records, 'metrics': metrics})

@app.route('/api/export', methods=['POST'])
def export_file():
    data = request.get_json() or {}
    records = data.get('records', [])
    if not records:
        return jsonify({'error': 'Nenhum registro para exportar.'}), 400

    # Exportar para arquivo temporário (compatível com ambiente serverless como Vercel)
    import tempfile
    temp_path = os.path.join(tempfile.gettempdir(), f"_temp_relatorio_{os.getpid()}.xlsx")
    try:
        export_to_excel(records, temp_path)
        with open(temp_path, "rb") as f:
            file_data = io.BytesIO(f.read())
        
        today = datetime.now().strftime("%Y-%m-%d")
        return send_file(
            file_data,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"MP_CARGAS_SSW_{today}.xlsx"
        )
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

if __name__ == '__main__':
    # Roda no localhost porta 5000
    app.run(host='127.0.0.1', port=5000, debug=False)
