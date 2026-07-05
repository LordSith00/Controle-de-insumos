<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Conecta Cargo | Gestão de Insumos</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="style.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
<body>

<header class="main-header">
    <div class="container d-flex justify-content-between align-items-center">
        <img src="https://conectacargo.com.br/wp-content/uploads/2021/05/logo-conecta-cargo.png" style="height: 50px;" alt="Conecta Cargo">
        <div class="header-right d-flex align-items-center gap-3">
            <span class="badge-unidade"><i class="fas fa-map-marker-alt"></i> Unidade: <span id="unidadeLabel">Não informado</span></span>
            <button class="btn-user" type="button"><i class="fas fa-user-circle"></i> Sair</button>
        </div>
    </div>
</header>

<section class="hero-banner">
    <div class="container text-center">
        <h1>Painel de Controle de Insumos</h1>
        <p>Implementação da Metodologia <strong>Just In Time</strong> para otimização de custos</p>
    </div>
</section>

<div class="container">

    <!-- FONTE DE DADOS -->
    <div class="content-box shadow-sm mb-4" id="dataSourceBox">
        <div class="d-flex justify-content-between align-items-start flex-wrap gap-3">
            <div>
                <h4 class="mb-1"><i class="fas fa-file-excel"></i> Fonte de Dados</h4>
                <p class="mb-0" id="fileStatusText" aria-live="polite">Nenhuma planilha carregada ainda. Envie um arquivo .xlsx para preencher o painel.</p>
            </div>
            <button class="btn btn-conecta btn-sm" id="exportBtn" type="button">
                <i class="fas fa-file-export"></i> Baixar planilha (dados atuais)
            </button>
        </div>

        <div id="dropZone" class="drop-zone mt-3 text-center" role="button" tabindex="0" aria-label="Enviar planilha Excel">
            <i class="fas fa-cloud-upload-alt fa-2x mb-2" aria-hidden="true"></i>
            <p class="mb-2">Arraste o arquivo .xlsx aqui, ou</p>
            <label for="fileInput" class="btn btn-conecta btn-sm mb-0">Selecionar planilha</label>
            <input type="file" id="fileInput" accept=".xlsx,.xls" hidden>
            <div id="uploadSpinner" class="mt-2 d-none">
                <i class="fas fa-spinner fa-spin"></i> Lendo planilha...
            </div>
        </div>
    </div>

    <h3 class="section-title"><i class="fas fa-traffic-light"></i> Controle Visual de Estoque (Kanban) </h3>
    <div class="row g-4 mb-5" id="kanbanRow">
        <div class="col-md-4">
            <div class="kanban-col status-red">
                <div class="col-label d-flex justify-content-between">
                    <span>Nível Crítico / Compra Imediata</span>
                    <span class="col-count" id="count-critico">0</span>
                </div>
                <div id="col-critico" class="kanban-cards"></div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="kanban-col status-yellow">
                <div class="col-label d-flex justify-content-between">
                    <span>Atenção / Reposição Próxima</span>
                    <span class="col-count" id="count-atencao">0</span>
                </div>
                <div id="col-atencao" class="kanban-cards"></div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="kanban-col status-green">
                <div class="col-label d-flex justify-content-between">
                    <span>Estoque Abastecido</span>
                    <span class="col-count" id="count-abastecido">0</span>
                </div>
                <div id="col-abastecido" class="kanban-cards"></div>
            </div>
        </div>
    </div>

    <div class="row">
        <div class="col-lg-8">
            <div class="content-box shadow-sm">
                <h4><i class="fas fa-chart-bar"></i> Ranking de Giro de Materiais </h4>
                <div id="chartEmptyState" class="text-center text-muted py-5">
                    <i class="fas fa-chart-bar fa-2x mb-2" aria-hidden="true"></i>
                    <p class="mb-0">O gráfico aparece aqui depois que uma planilha for carregada.</p>
                </div>
                <canvas id="graficoEstoque" class="d-none"></canvas>
            </div>
        </div>
        <div class="col-lg-4">
            <div class="content-box shadow-sm">
                <h4><i class="fas fa-exchange-alt"></i> Movimentação Rápida</h4>
                <div class="form-group mb-3">
                    <label for="insumoSelect">Insumo</label>
                    <select class="form-select" id="insumoSelect" disabled>
                        <option value="">Carregue a planilha primeiro</option>
                    </select>
                </div>
                <div class="form-group mb-2">
                    <label for="qtdInput">Quantidade</label>
                    <input type="number" class="form-control" id="qtdInput" placeholder="0" min="1" disabled>
                    <div class="form-text text-danger d-none" id="qtdError">Informe uma quantidade válida.</div>
                </div>
                <button class="btn btn-conecta w-100 mb-2" id="btnEntrada" type="button" disabled>Registrar Entrada</button>
                <button class="btn btn-outline-secondary w-100" id="btnSaida" type="button" disabled>Registrar Saída (Baixa)</button>
            </div>
        </div>
    </div>

</div>

<footer class="footer">

    <div class="card shadow-sm p-4 mt-4">
        <div class="d-flex justify-content-between align-items-center mb-3">
            <h4 class="mb-0"><i class="fas fa-history me-2"></i>Histórico Diário </h4>
            <button class="btn btn-success btn-sm" id="exportBtn2" type="button">
                <i class="fas fa-file-export"></i> Exportar Relatório
            </button>
        </div>
        <div class="table-responsive">
            <table class="table table-hover mb-0">
                <thead class="table-light">
                    <tr>
                        <th scope="col">Data</th>
                        <th scope="col">Tipo</th>
                        <th scope="col">Insumo</th>
                        <th scope="col" class="text-end">Quantidade</th>
                    </tr>
                </thead>
                <tbody id="historicoBody">
                    <tr><td colspan="4" class="text-center text-muted py-4">Sem dados carregados.</td></tr>
                </tbody>
            </table>
        </div>
    </div>

    <div class="container text-center mt-4">
        <img src="https://conectacargo.com.br/wp-content/uploads/2023/06/Logo-Conecta-branco.png" style="width:130px;" class="mb-3" alt="Conecta Cargo">
        <p>Sistema Integrado de Gestão de Armazenagem </p>
        <a href="https://conectacargo.com.br" target="_blank" rel="noopener">Acesse o portal oficial</a>
    </div>
</footer>

<!-- TOASTS -->
<div id="toastContainer" aria-live="assertive" aria-atomic="true"></div>

<script>
/* =========================================================
   ESTADO GLOBAL
   ========================================================= */
let estoqueData = [];     // [{ insumo, quantidade, unidade, status, tag, unidadeOperacional }]
let historicoData = [];   // [{ data, tipo, insumo, quantidade }]
let unidadeAtual = '';
let chartInstance = null;
let currentFileName = '';

/* =========================================================
   TOAST / FEEDBACK
   ========================================================= */
function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const icons = { info: 'fa-circle-info', success: 'fa-circle-check', error: 'fa-circle-exclamation' };
    const el = document.createElement('div');
    el.className = `app-toast app-toast-${type}`;
    el.innerHTML = `<i class="fas ${icons[type] || icons.info}"></i><span>${message}</span>`;
    container.appendChild(el);
    requestAnimationFrame(() => el.classList.add('show'));
    setTimeout(() => {
        el.classList.remove('show');
        setTimeout(() => el.remove(), 200);
    }, 3500);
}

/* =========================================================
   HELPERS
   ========================================================= */
function normalizeKey(str) {
    return str.toString().trim().toLowerCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function getField(row, aliases) {
    const keys = Object.keys(row);
    for (const alias of aliases) {
        const found = keys.find(k => normalizeKey(k) === normalizeKey(alias));
        if (found !== undefined && row[found] !== '') return row[found];
    }
    return undefined;
}

function normalizeStatus(status) {
    const s = normalizeKey(status || '');
    if (s.includes('crit')) return 'critico';
    if (s.includes('aten')) return 'atencao';
    if (s.includes('abast') || s.includes('ok') || s.includes('normal')) return 'abastecido';
    return 'atencao';
}

function statusLabel(status) {
    return { critico: 'Crítico', atencao: 'Atenção', abastecido: 'Abastecido' }[status] || 'Atenção';
}

function formatDateBR(value) {
    if (value instanceof Date && !isNaN(value)) return value.toLocaleDateString('pt-BR');
    if (typeof value === 'number') {
        const d = XLSX.SSF.parse_date_code(value);
        if (d) return `${String(d.d).padStart(2, '0')}/${String(d.m).padStart(2, '0')}/${d.y}`;
    }
    return value || '';
}

function setControlsEnabled(enabled) {
    ['insumoSelect', 'qtdInput', 'btnEntrada', 'btnSaida'].forEach(id => {
        document.getElementById(id).disabled = !enabled;
    });
}

/* =========================================================
   LEITURA DA PLANILHA
   ========================================================= */
const ALLOWED_EXTENSIONS = ['.xlsx', '.xls'];

function isValidSpreadsheet(file) {
    const name = file.name.toLowerCase();
    return ALLOWED_EXTENSIONS.some(ext => name.endsWith(ext));
}

function handleWorkbookFile(file) {
    if (!isValidSpreadsheet(file)) {
        showToast('Formato não suportado. Envie um arquivo .xlsx ou .xls.', 'error');
        return;
    }

    currentFileName = file.name;
    document.getElementById('uploadSpinner').classList.remove('d-none');

    const reader = new FileReader();
    reader.onload = (evt) => {
        try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            processWorkbook(workbook);
            showToast('Planilha carregada com sucesso.', 'success');
        } catch (err) {
            console.error(err);
            showToast('Não foi possível ler essa planilha. Verifique o arquivo e tente novamente.', 'error');
        } finally {
            document.getElementById('uploadSpinner').classList.add('d-none');
        }
    };
    reader.onerror = () => {
        document.getElementById('uploadSpinner').classList.add('d-none');
        showToast('Erro ao ler o arquivo.', 'error');
    };
    reader.readAsArrayBuffer(file);
}

function findHeaderRowIndex(rows, mustContain) {
    for (let i = 0; i < rows.length; i++) {
        const rowNormalized = rows[i].map(c => normalizeKey(c));
        if (mustContain.some(alias => rowNormalized.includes(normalizeKey(alias)))) {
            return i;
        }
    }
    return -1;
}

function sheetToJsonSmart(sheet, headerAliases) {
    if (!sheet) return [];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
    if (!rows.length) return [];

    const headerRowIndex = findHeaderRowIndex(rows, headerAliases);
    if (headerRowIndex === -1) {
        // não achou cabeçalho reconhecível: usa o comportamento padrão (primeira linha como cabeçalho)
        return XLSX.utils.sheet_to_json(sheet, { defval: '' });
    }

    const headers = rows[headerRowIndex];
    return rows.slice(headerRowIndex + 1).map(row => {
        const obj = {};
        headers.forEach((h, i) => { if (h !== '') obj[h] = row[i] !== undefined ? row[i] : ''; });
        return obj;
    });
}

function processWorkbook(workbook) {
    const sheetNames = workbook.SheetNames;
    const estoqueSheetName = sheetNames.find(n => normalizeKey(n).includes('estoque')) || sheetNames[0];
    const historicoSheetName = sheetNames.find(n => normalizeKey(n).includes('historic')) || sheetNames[1];

    const estoqueRaw = sheetToJsonSmart(workbook.Sheets[estoqueSheetName], ['Insumo', 'Item', 'Material']);
    const historicoRaw = historicoSheetName
        ? sheetToJsonSmart(workbook.Sheets[historicoSheetName], ['Data'])
        : [];

    estoqueData = estoqueRaw.map(row => ({
        insumo: (getField(row, ['Insumo', 'Item', 'Material']) || '').toString().trim(),
        quantidade: Number(getField(row, ['Quantidade', 'Qtd', 'Estoque'])) || 0,
        unidade: getField(row, ['Unidade', 'UN', 'Un']) || 'un.',
        status: normalizeStatus(getField(row, ['Status', 'Situacao', 'Nivel'])),
        tag: getField(row, ['Tag', 'Observacao', 'Prioridade']) || '',
        unidadeOperacional: getField(row, ['Unidade Operacional', 'Filial', 'Unidade de Negocio']) || ''
    })).filter(r => r.insumo);

    historicoData = historicoRaw.map(row => ({
        data: formatDateBR(getField(row, ['Data'])),
        tipo: getField(row, ['Tipo', 'Categoria']) || '',
        insumo: (getField(row, ['Insumo', 'Item', 'Material']) || '').toString().trim(),
        quantidade: Number(getField(row, ['Quantidade', 'Qtd'])) || 0
    }));

    if (!estoqueData.length) {
        const headersFound = estoqueRaw.length ? Object.keys(estoqueRaw[0]) : [];
        console.warn('Aba usada como Estoque:', estoqueSheetName);
        console.warn('Colunas encontradas nessa aba:', headersFound);
        console.warn('Linhas lidas (antes do filtro por Insumo):', estoqueRaw);

        if (!headersFound.length) {
            showToast(`A aba "${estoqueSheetName}" parece estar vazia.`, 'error');
        } else if (!headersFound.some(h => normalizeKey(h) === 'insumo' || normalizeKey(h) === 'item' || normalizeKey(h) === 'material')) {
            showToast(`Não encontrei uma coluna "Insumo" na aba "${estoqueSheetName}". Colunas encontradas: ${headersFound.join(', ')}.`, 'error');
        } else {
            showToast(`A aba "${estoqueSheetName}" tem as colunas certas, mas todas as linhas de Insumo vieram vazias. Verifique se o cabeçalho está mesmo na primeira linha da planilha.`, 'error');
        }
    }

    unidadeAtual = (estoqueData.find(i => i.unidadeOperacional) || {}).unidadeOperacional || '';
    document.getElementById('unidadeLabel').textContent = unidadeAtual || 'Não informado';

    document.getElementById('fileStatusText').innerHTML =
        `Planilha carregada: <strong>${currentFileName}</strong> — ${estoqueData.length} insumo(s), ${historicoData.length} movimentação(ões).`;

    setControlsEnabled(estoqueData.length > 0);
    renderAll();
}

/* =========================================================
   RENDERIZAÇÃO
   ========================================================= */
function renderAll() {
    renderKanban();
    renderChart();
    renderTabela();
    renderSelect();
}

function renderKanban() {
    const cols = { critico: 'col-critico', atencao: 'col-atencao', abastecido: 'col-abastecido' };
    Object.values(cols).forEach(id => document.getElementById(id).innerHTML = '');
    const counts = { critico: 0, atencao: 0, abastecido: 0 };

    if (!estoqueData.length) {
        Object.values(cols).forEach(id => {
            document.getElementById(id).innerHTML = '<p class="text-muted small mb-0">Sem dados.</p>';
        });
    } else {
        estoqueData.forEach(item => {
            counts[item.status] = (counts[item.status] || 0) + 1;
            const container = document.getElementById(cols[item.status] || cols.atencao);
            const iconClass = item.status === 'critico'
                ? 'fa-exclamation-circle text-danger'
                : (item.status === 'abastecido' ? 'fa-check-circle text-success' : '');

            const card = document.createElement('div');
            card.className = 'insumo-card';
            card.innerHTML = `
                ${item.tag ? `<div class="card-tag">${item.tag}</div>` : ''}
                <h6>${item.insumo}</h6>
                <div class="d-flex justify-content-between align-items-center">
                    <span class="qtd-val">${item.quantidade} ${item.unidade}</span>
                    ${iconClass ? `<i class="fas ${iconClass}"></i>` : ''}
                </div>`;
            container.appendChild(card);
        });
    }

    Object.entries(counts).forEach(([status, n]) => {
        document.getElementById(`count-${status}`).textContent = n;
    });
}

function renderChart() {
    const canvas = document.getElementById('graficoEstoque');
    const emptyState = document.getElementById('chartEmptyState');

    if (!estoqueData.length) {
        canvas.classList.add('d-none');
        emptyState.classList.remove('d-none');
        if (chartInstance) { chartInstance.destroy(); chartInstance = null; }
        return;
    }

    canvas.classList.remove('d-none');
    emptyState.classList.add('d-none');

    const colorMap = { critico: '#d32f2f', atencao: '#fbc02d', abastecido: '#388e3c' };
    const labels = estoqueData.map(i => i.insumo);
    const data = estoqueData.map(i => i.quantidade);
    const colors = estoqueData.map(i => colorMap[i.status] || '#888');

    if (chartInstance) chartInstance.destroy();
    chartInstance = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: { labels, datasets: [{ label: 'Nível de Estoque', data, backgroundColor: colors, borderRadius: 8 }] },
        options: { responsive: true, plugins: { legend: { display: false } } }
    });
}

function renderTabela() {
    const body = document.getElementById('historicoBody');
    body.innerHTML = '';
    if (!historicoData.length) {
        body.innerHTML = '<tr><td colspan="4" class="text-center text-muted py-4">Sem dados carregados.</td></tr>';
        return;
    }
    [...historicoData].reverse().forEach(mov => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${mov.data}</td><td>${mov.tipo}</td><td>${mov.insumo}</td><td class="text-end">${mov.quantidade}</td>`;
        body.appendChild(tr);
    });
}

function renderSelect() {
    const select = document.getElementById('insumoSelect');
    select.innerHTML = '';
    if (!estoqueData.length) {
        select.innerHTML = '<option value="">Carregue a planilha primeiro</option>';
        return;
    }
    estoqueData.forEach((item, idx) => {
        const opt = document.createElement('option');
        opt.value = idx;
        opt.textContent = `${item.insumo} (${item.quantidade} ${item.unidade})`;
        select.appendChild(opt);
    });
}

/* =========================================================
   MOVIMENTAÇÃO (ENTRADA / SAÍDA)
   ========================================================= */
function registrarMovimento(tipoMovimento) {
    const select = document.getElementById('insumoSelect');
    const qtdInput = document.getElementById('qtdInput');
    const qtdError = document.getElementById('qtdError');
    const idx = select.value;
    const qtd = Number(qtdInput.value);

    const valido = idx !== '' && estoqueData[idx] && qtd > 0 && Number.isFinite(qtd);
    qtdError.classList.toggle('d-none', valido);
    if (!valido) return;

    const item = estoqueData[idx];
    if (tipoMovimento === 'Saída' && qtd > item.quantidade) {
        showToast(`Quantidade maior que o estoque atual de "${item.insumo}" (${item.quantidade} ${item.unidade}).`, 'error');
        return;
    }

    item.quantidade = tipoMovimento === 'Entrada' ? item.quantidade + qtd : item.quantidade - qtd;

    historicoData.push({
        data: new Date().toLocaleDateString('pt-BR'),
        tipo: tipoMovimento,
        insumo: item.insumo,
        quantidade: qtd
    });

    qtdInput.value = '';
    renderAll();
    // manter o mesmo insumo selecionado após o re-render
    document.getElementById('insumoSelect').value = idx;
    showToast(`${tipoMovimento} registrada para "${item.insumo}".`, 'success');
}

document.getElementById('btnEntrada').addEventListener('click', () => registrarMovimento('Entrada'));
document.getElementById('btnSaida').addEventListener('click', () => registrarMovimento('Saída'));
document.getElementById('qtdInput').addEventListener('input', () => {
    document.getElementById('qtdError').classList.add('d-none');
});

/* =========================================================
   UPLOAD (CLIQUE E ARRASTAR/SOLTAR)
   ========================================================= */
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');

fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleWorkbookFile(e.target.files[0]);
    fileInput.value = ''; // permite reenviar o mesmo arquivo depois
});

['dragenter', 'dragover'].forEach(evt =>
    dropZone.addEventListener(evt, (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); })
);
['dragleave', 'drop'].forEach(evt =>
    dropZone.addEventListener(evt, (e) => { e.preventDefault(); dropZone.classList.remove('drag-over'); })
);
dropZone.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files[0];
    if (file) handleWorkbookFile(file);
});
dropZone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
});

/* =========================================================
   EXPORTAR PLANILHA (SEMPRE COM OS DADOS ATUAIS DO SITE)
   ========================================================= */
function exportarPlanilhaAtual() {
    const wb = XLSX.utils.book_new();

    const estoqueExport = estoqueData.length
        ? estoqueData.map(i => ({
            Insumo: i.insumo,
            Quantidade: i.quantidade,
            Unidade: i.unidade,
            Status: statusLabel(i.status),
            Tag: i.tag,
            'Unidade Operacional': i.unidadeOperacional
        }))
        : [{ Insumo: '', Quantidade: '', Unidade: '', Status: '', Tag: '', 'Unidade Operacional': '' }];

    const historicoExport = historicoData.length
        ? historicoData.map(m => ({ Data: m.data, Tipo: m.tipo, Insumo: m.insumo, Quantidade: m.quantidade }))
        : [{ Data: '', Tipo: '', Insumo: '', Quantidade: '' }];

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(estoqueExport), 'Estoque');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(historicoExport), 'Historico');

    const baseName = currentFileName ? currentFileName.replace(/\.xlsx?$/i, '') : 'conecta_cargo_estoque';
    XLSX.writeFile(wb, `${baseName}.xlsx`);
    showToast('Planilha exportada com os dados atuais do painel.', 'success');
}

document.getElementById('exportBtn').addEventListener('click', exportarPlanilhaAtual);
document.getElementById('exportBtn2').addEventListener('click', exportarPlanilhaAtual);
</script>

<style>
:root {
    --qa-danger: #d32f2f;
    --qa-warning: #fbc02d;
    --qa-success: #388e3c;
}

.drop-zone {
    border: 2px dashed #c9c9c9;
    border-radius: 12px;
    padding: 24px;
    cursor: pointer;
    transition: background-color .15s ease, border-color .15s ease;
}
.drop-zone:hover,
.drop-zone:focus-visible {
    border-color: #388e3c;
    outline: none;
}
.drop-zone.drag-over {
    background-color: #f4f9f4;
    border-color: #388e3c;
}

.col-count {
    font-size: .75rem;
    font-weight: 600;
    background: rgba(0,0,0,.06);
    border-radius: 999px;
    padding: 1px 9px;
    line-height: 1.5;
}

button:disabled,
select:disabled,
input:disabled {
    cursor: not-allowed;
    opacity: .55;
}

.btn:focus-visible,
.form-select:focus-visible,
.form-control:focus-visible,
a:focus-visible {
    outline: 2px solid #388e3c;
    outline-offset: 2px;
}

#toastContainer {
    position: fixed;
    top: 16px;
    right: 16px;
    z-index: 1080;
    display: flex;
    flex-direction: column;
    gap: 8px;
    max-width: 320px;
}
.app-toast {
    display: flex;
    align-items: center;
    gap: 8px;
    background: #fff;
    border-left: 4px solid #888;
    border-radius: 8px;
    padding: 10px 14px;
    font-size: .9rem;
    box-shadow: 0 4px 14px rgba(0,0,0,.15);
    opacity: 0;
    transform: translateX(12px);
    transition: opacity .2s ease, transform .2s ease;
}
.app-toast.show { opacity: 1; transform: translateX(0); }
.app-toast-success { border-left-color: var(--qa-success); }
.app-toast-error { border-left-color: var(--qa-danger); }
.app-toast-info { border-left-color: #0d6efd; }
</style>

</body>
</html>
