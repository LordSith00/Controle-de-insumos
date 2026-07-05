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

<!-- TELA DE LOGIN -->
<div id="loginScreen" class="login-overlay">
    <form id="loginForm" class="login-box" autocomplete="off">
        <img src="https://conectacargo.com.br/wp-content/uploads/2021/05/logo-conecta-cargo.png" alt="Conecta Cargo" class="login-logo">
        <h3>Acesso ao Painel</h3>
        <p class="text-muted mb-4">Insira suas credenciais para continuar</p>

        <div class="form-group mb-3 text-start">
            <label for="loginUser">Usuário</label>
            <input type="text" id="loginUser" class="form-control" required autocomplete="username">
        </div>
        <div class="form-group mb-2 text-start">
            <label for="loginPass">Senha</label>
            <input type="password" id="loginPass" class="form-control" required autocomplete="current-password">
        </div>
        <div id="loginError" class="text-danger small mb-3 d-none">
            <i class="fas fa-circle-exclamation"></i> Usuário ou senha incorretos.
        </div>
        <button type="submit" class="btn btn-conecta w-100">Entrar</button>
    </form>
</div>

<!-- CONTEÚDO DO PAINEL (liberado após login) -->
<div id="appRoot" class="d-none">

<header class="main-header">
    <div class="container d-flex justify-content-between align-items-center">
        <img src="https://conectacargo.com.br/wp-content/uploads/2023/06/Logo-Conecta-branco.png" style="height: 42px;" alt="Conecta Cargo">
        <div class="header-right d-flex align-items-center gap-3">
            <span class="badge-unidade"><i class="fas fa-map-marker-alt"></i> Unidade: <span id="unidadeLabel">Não informado</span></span>
            <span class="badge-unidade"><i class="fas fa-user"></i> <span id="loggedUserLabel">—</span></span>
            <button class="btn-user" type="button" id="btnLogout"><i class="fas fa-user-circle"></i> Sair</button>
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

    <!-- ABAS -->
    <ul class="nav app-tabs mb-4" id="mainTabs" role="tablist">
        <li class="nav-item">
            <button class="nav-link active" type="button" data-tab="painel">
                <i class="fas fa-table-columns"></i> Painel
            </button>
        </li>
        <li class="nav-item">
            <button class="nav-link" type="button" data-tab="historico">
                <i class="fas fa-history"></i> Histórico de Alterações
            </button>
        </li>
    </ul>

    <!-- ABA: PAINEL -->
    <div class="tab-pane" id="tab-painel">
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

    <!-- ABA: HISTÓRICO DE ALTERAÇÕES -->
    <div class="tab-pane d-none" id="tab-historico">
        <div class="content-box shadow-sm">
            <h4 class="mb-3"><i class="fas fa-history me-2"></i>Histórico de Alterações</h4>
            <div class="table-responsive">
                <table class="table table-hover mb-0">
                    <thead class="table-light">
                        <tr>
                            <th scope="col">Data</th>
                            <th scope="col">Tipo</th>
                            <th scope="col">Insumo</th>
                            <th scope="col" class="text-end">Quantidade</th>
                            <th scope="col">Usuário</th>
                        </tr>
                    </thead>
                    <tbody id="historicoBody">
                        <tr><td colspan="5" class="text-center text-muted py-4">Sem dados carregados.</td></tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- FONTE DE DADOS -->
    <div class="content-box shadow-sm mt-4" id="dataSourceBox">
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

</div>

<footer class="footer">
    <div class="container text-center mt-4">
        <img src="https://conectacargo.com.br/wp-content/uploads/2023/06/Logo-Conecta-branco.png" style="width:130px;" class="mb-3" alt="Conecta Cargo">
        <p>Sistema Integrado de Gestão de Armazenagem </p>
        <a href="https://conectacargo.com.br" target="_blank" rel="noopener">Acesse o portal oficial</a>
    </div>
</footer>

</div> <!-- fim #appRoot -->

<!-- POPUP DE ITENS CRÍTICOS -->
<div id="criticalModalOverlay" class="modal-overlay d-none">
    <div class="modal-box">
        <div class="modal-icon"><i class="fas fa-triangle-exclamation"></i></div>
        <h4>Atenção — Itens Críticos</h4>
        <p class="text-muted">Estes insumos precisam de reposição imediata:</p>
        <ul id="criticalList" class="critical-list"></ul>
        <button id="closeCriticalModal" class="btn btn-conecta w-100 mt-2">Entendi</button>
    </div>
</div>

<!-- TOASTS -->
<div id="toastContainer" aria-live="assertive" aria-atomic="true"></div>

<script>
/* =========================================================
   LOGIN (tela de acesso)
   ========================================================= */
/* =========================================================
   USUÁRIOS VÁLIDOS
   Para adicionar, remover ou alterar um usuário, edite as
   linhas abaixo — formato: 'usuario': 'senha'.
   ========================================================= */
const VALID_CREDENTIALS = {
    'bruno.melo': 'bruno.melo002',
    'user.1':  'user1',
    'user.2':  'user2',
    'user.3':  'user3',
    'user.4':  'user4',
    'user.5':  'user5',
    'user.6':  'user6',
    'user.7':  'user7',
    'user.8':  'user8',
    'user.9':  'user9',
    'user.10': 'user10'
};
let loggedInUser = '';

document.getElementById('loginForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const user = document.getElementById('loginUser').value.trim();
    const pass = document.getElementById('loginPass').value;
    const errorEl = document.getElementById('loginError');

    if (VALID_CREDENTIALS[user] !== undefined && VALID_CREDENTIALS[user] === pass) {
        loggedInUser = user;
        document.getElementById('loggedUserLabel').textContent = user;
        document.getElementById('loginScreen').classList.add('login-hide');
        document.getElementById('appRoot').classList.remove('d-none');
        setTimeout(() => document.getElementById('loginScreen').remove(), 300);
    } else {
        errorEl.classList.remove('d-none');
        document.getElementById('loginPass').value = '';
        document.getElementById('loginPass').focus();
    }
});

document.getElementById('btnLogout').addEventListener('click', () => {
    location.reload();
});

/* =========================================================
   ABAS (Painel / Histórico de Alterações)
   ========================================================= */
document.querySelectorAll('#mainTabs .nav-link').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('#mainTabs .nav-link').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        document.querySelectorAll('.tab-pane').forEach(p => p.classList.add('d-none'));
        document.getElementById(`tab-${btn.dataset.tab}`).classList.remove('d-none');
    });
});

/* =========================================================
   POPUP DE ITENS CRÍTICOS
   ========================================================= */
function showCriticalPopup() {
    const criticos = estoqueData.filter(i => computeStatus(i.quantidade) === 'critico');
    if (!criticos.length) return;

    const list = document.getElementById('criticalList');
    list.innerHTML = criticos
        .map(i => `<li><strong>${i.quantidade} ${i.unidade}</strong> de <strong>${i.insumo}</strong> em estado crítico</li>`)
        .join('');

    document.getElementById('criticalModalOverlay').classList.remove('d-none');
}

document.getElementById('closeCriticalModal').addEventListener('click', () => {
    document.getElementById('criticalModalOverlay').classList.add('d-none');
});

/* =========================================================
   ESTADO GLOBAL
   ========================================================= */
let estoqueData = [];     // [{ insumo, quantidade, unidade, tag, unidadeOperacional }]
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

function computeStatus(quantidade) {
    if (quantidade > 100) return 'abastecido';
    if (quantidade > 30) return 'atencao';
    return 'critico';
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
            showCriticalPopup();
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
        tag: getField(row, ['Tag', 'Observacao', 'Prioridade']) || '',
        unidadeOperacional: getField(row, ['Unidade Operacional', 'Filial', 'Unidade de Negocio']) || ''
    })).filter(r => r.insumo);

    historicoData = historicoRaw.map(row => ({
        data: formatDateBR(getField(row, ['Data'])),
        tipo: getField(row, ['Tipo', 'Categoria']) || '',
        insumo: (getField(row, ['Insumo', 'Item', 'Material']) || '').toString().trim(),
        quantidade: Number(getField(row, ['Quantidade', 'Qtd'])) || 0,
        usuario: getField(row, ['Usuario', 'Responsavel', 'Alterado por']) || 'Importado da planilha'
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
        const counters = { critico: 0, atencao: 0, abastecido: 0 };
        estoqueData.forEach(item => {
            const status = computeStatus(item.quantidade);
            counts[status] = (counts[status] || 0) + 1;
            const container = document.getElementById(cols[status] || cols.atencao);
            const iconMap = {
                critico: 'fa-exclamation-circle status-icon-critico',
                atencao: 'fa-triangle-exclamation status-icon-atencao',
                abastecido: 'fa-check-circle status-icon-abastecido'
            };
            const delay = (counters[status]++) * 60;

            const card = document.createElement('div');
            card.className = `insumo-card card-status-${status}`;
            card.style.animationDelay = `${delay}ms`;
            card.innerHTML = `
                ${item.tag ? `<div class="card-tag">${item.tag}</div>` : ''}
                <h6>${item.insumo}</h6>
                <div class="d-flex justify-content-between align-items-center">
                    <span class="qtd-val">${item.quantidade} ${item.unidade}</span>
                    <i class="fas ${iconMap[status]}" aria-hidden="true"></i>
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
    const colors = estoqueData.map(i => colorMap[computeStatus(i.quantidade)] || '#888');

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
        body.innerHTML = '<tr><td colspan="5" class="text-center text-muted py-4">Sem dados carregados.</td></tr>';
        return;
    }
    [...historicoData].reverse().forEach(mov => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${mov.data}</td><td>${mov.tipo}</td><td>${mov.insumo}</td><td class="text-end">${mov.quantidade}</td><td>${mov.usuario || '—'}</td>`;
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
        quantidade: qtd,
        usuario: loggedInUser || 'Desconhecido'
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
            Status: statusLabel(computeStatus(i.quantidade)),
            Tag: i.tag,
            'Unidade Operacional': i.unidadeOperacional
        }))
        : [{ Insumo: '', Quantidade: '', Unidade: '', Status: '', Tag: '', 'Unidade Operacional': '' }];

    const historicoExport = historicoData.length
        ? historicoData.map(m => ({ Data: m.data, Tipo: m.tipo, Insumo: m.insumo, Quantidade: m.quantidade, Usuario: m.usuario || '' }))
        : [{ Data: '', Tipo: '', Insumo: '', Quantidade: '', Usuario: '' }];

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(estoqueExport), 'Estoque');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(historicoExport), 'Historico');

    const baseName = currentFileName ? currentFileName.replace(/\.xlsx?$/i, '') : 'conecta_cargo_estoque';
    XLSX.writeFile(wb, `${baseName}.xlsx`);
    showToast('Planilha exportada com os dados atuais do painel.', 'success');
}

document.getElementById('exportBtn').addEventListener('click', exportarPlanilhaAtual);
</script>

<style>
:root {
    --qa-danger: #d32f2f;
    --qa-warning: #fbc02d;
    --qa-success: #388e3c;
    --brand-navy: #000080;
}

/* =========================================================
   IDENTIDADE VISUAL — cabeçalho e rodapé azul-marinho
   ========================================================= */
.main-header {
    background: var(--brand-navy) !important;
    box-shadow: 0 2px 10px rgba(0, 0, 0, .15);
}
.main-header .badge-unidade {
    color: #fff;
    background: rgba(255, 255, 255, .12);
    border-radius: 999px;
    padding: 4px 12px;
    font-size: .82rem;
}
.main-header .btn-user {
    color: #fff;
    background: rgba(255, 255, 255, .12);
    border: 1px solid rgba(255, 255, 255, .25);
    border-radius: 8px;
    padding: 6px 14px;
    transition: background .15s ease;
}
.main-header .btn-user:hover { background: rgba(255, 255, 255, .22); }

body .hero-banner,
section.hero-banner {
    background: var(--brand-navy) !important;
    background-image: none !important;
    margin-top: 0 !important;
    padding: 28px 0 36px !important;
}
body .hero-banner h1,
section.hero-banner h1 { color: #fff !important; }
body .hero-banner p,
section.hero-banner p { color: #d7dbf5 !important; }
body .hero-banner strong,
section.hero-banner strong { color: #fff !important; }

.footer {
    background: var(--brand-navy) !important;
    color: #f1f2f6;
    padding: 32px 0 24px;
}
.footer p { color: #cfd2e6; }
.footer a { color: #fff; text-decoration: underline; }
.footer a:hover { color: #d9e0ff; }

/* =========================================================
   ABAS (Painel / Histórico de Alterações)
   ========================================================= */
.app-tabs {
    border-bottom: 1px solid #e5e7eb;
    gap: 4px;
}
.app-tabs .nav-link {
    background: none;
    border: none;
    padding: 10px 18px;
    font-weight: 600;
    font-size: .92rem;
    color: #6b7280;
    border-bottom: 3px solid transparent;
    transition: color .15s ease, border-color .15s ease;
}
.app-tabs .nav-link:hover { color: var(--brand-navy); }
.app-tabs .nav-link.active {
    color: var(--brand-navy);
    border-bottom-color: var(--brand-navy);
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

/* =========================================================
   KANBAN — bordas, cores por status, ícones e animações
   ========================================================= */
.kanban-col {
    background: #fff;
    border: 1px solid #e7e7ea;
    border-radius: 16px;
    padding: 18px 18px 8px;
    min-height: 240px;
    box-shadow: 0 2px 10px rgba(15, 23, 42, .05);
    transition: box-shadow .25s ease, transform .25s ease;
    border-top: 4px solid #cbd5e1;
}
.kanban-col:hover {
    box-shadow: 0 10px 24px rgba(15, 23, 42, .09);
    transform: translateY(-2px);
}
.kanban-col.status-red    { border-top-color: var(--qa-danger); }
.kanban-col.status-yellow { border-top-color: var(--qa-warning); }
.kanban-col.status-green  { border-top-color: var(--qa-success); }

.col-label {
    font-weight: 600;
    font-size: .92rem;
    letter-spacing: .01em;
    color: #374151;
    margin-bottom: 14px;
    padding-bottom: 10px;
    border-bottom: 1px solid #eef0f2;
}
.col-count {
    background: rgba(15, 23, 42, .06);
}
.kanban-col.status-red .col-count    { background: rgba(211, 47, 47, .12); color: var(--qa-danger); }
.kanban-col.status-yellow .col-count { background: rgba(251, 192, 2, .18); color: #8a6100; }
.kanban-col.status-green .col-count  { background: rgba(56, 142, 60, .12); color: var(--qa-success); }

.insumo-card {
    background: #fff;
    border: 1px solid #ececef;
    border-left: 4px solid #cbd5e1;
    border-radius: 12px;
    padding: 13px 15px;
    margin-bottom: 12px;
    box-shadow: 0 1px 3px rgba(15, 23, 42, .04);
    transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
    animation: cardIn .35s ease both;
}
.insumo-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 18px rgba(15, 23, 42, .1);
    border-color: #dcdfe4;
}
.insumo-card h6 { margin: 2px 0 8px; font-weight: 600; color: #1f2937; }
.qtd-val { font-weight: 600; color: #111827; }

.card-status-critico    { border-left-color: var(--qa-danger); }
.card-status-atencao    { border-left-color: var(--qa-warning); }
.card-status-abastecido { border-left-color: var(--qa-success); }

.status-icon-critico    { color: var(--qa-danger); animation: pulseIcon 1.8s ease-in-out infinite; }
.status-icon-atencao    { color: #b8860b; }
.status-icon-abastecido { color: var(--qa-success); }

.card-tag {
    display: inline-block;
    font-size: .68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .04em;
    color: #6b7280;
    background: #f3f4f6;
    border-radius: 999px;
    padding: 2px 9px;
    margin-bottom: 6px;
}

@keyframes cardIn {
    from { opacity: 0; transform: translateY(8px); }
    to   { opacity: 1; transform: translateY(0); }
}
@keyframes pulseIcon {
    0%, 100% { opacity: 1; transform: scale(1); }
    50%      { opacity: .55; transform: scale(1.15); }
}
@media (prefers-reduced-motion: reduce) {
    .insumo-card, .status-icon-critico { animation: none !important; }
    .insumo-card:hover, .kanban-col:hover { transform: none !important; }
}

/* =========================================================
   OUTROS DETALHES DE POLIMENTO
   ========================================================= */
.content-box {
    transition: box-shadow .2s ease;
}
.btn-conecta, .btn-outline-secondary, .btn-success {
    transition: transform .12s ease, box-shadow .12s ease;
}
.btn-conecta:hover, .btn-success:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(56, 142, 60, .25);
}

/* =========================================================
   TELA DE LOGIN
   ========================================================= */
.login-overlay {
    position: fixed;
    inset: 0;
    z-index: 2000;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg, #f4f6f5 0%, #e7efe8 100%);
    padding: 20px;
    transition: opacity .3s ease;
}
.login-overlay.login-hide { opacity: 0; pointer-events: none; }

.login-box {
    background: #fff;
    border-radius: 18px;
    padding: 40px 36px;
    width: 100%;
    max-width: 380px;
    text-align: center;
    box-shadow: 0 20px 50px rgba(15, 23, 42, .12);
    animation: loginIn .4s ease both;
}
.login-logo { height: 42px; margin-bottom: 18px; }
.login-box h3 { font-weight: 700; color: #1f2937; margin-bottom: 4px; }
.login-box label { font-size: .85rem; font-weight: 600; color: #4b5563; margin-bottom: 4px; display: block; }
.login-box .form-control {
    border-radius: 10px;
    padding: 10px 12px;
}
.login-box .form-control:focus {
    border-color: var(--qa-success);
    box-shadow: 0 0 0 3px rgba(56, 142, 60, .15);
}

@keyframes loginIn {
    from { opacity: 0; transform: translateY(14px) scale(.98); }
    to   { opacity: 1; transform: translateY(0) scale(1); }
}

/* =========================================================
   POPUP DE ITENS CRÍTICOS
   ========================================================= */
.modal-overlay {
    position: fixed;
    inset: 0;
    z-index: 1900;
    background: rgba(15, 23, 42, .45);
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
    animation: overlayIn .2s ease both;
}
.modal-box {
    background: #fff;
    border-radius: 18px;
    padding: 32px 30px 26px;
    width: 100%;
    max-width: 420px;
    text-align: center;
    box-shadow: 0 24px 60px rgba(15, 23, 42, .25);
    animation: modalPop .3s ease both;
}
.modal-icon {
    width: 56px;
    height: 56px;
    margin: 0 auto 14px;
    border-radius: 50%;
    background: rgba(251, 192, 2, .15);
    color: #b8860b;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.5rem;
    animation: pulseIcon 1.8s ease-in-out infinite;
}
.modal-box h4 { font-weight: 700; color: #1f2937; margin-bottom: 4px; }
.critical-list {
    list-style: none;
    padding: 0;
    margin: 16px 0;
    text-align: left;
    max-height: 240px;
    overflow-y: auto;
}
.critical-list li {
    background: rgba(211, 47, 47, .06);
    border-left: 3px solid var(--qa-danger);
    border-radius: 8px;
    padding: 9px 12px;
    margin-bottom: 8px;
    font-size: .92rem;
    color: #333;
}

@keyframes overlayIn { from { opacity: 0; } to { opacity: 1; } }
@keyframes modalPop {
    from { opacity: 0; transform: translateY(10px) scale(.96); }
    to   { opacity: 1; transform: translateY(0) scale(1); }
}
@media (prefers-reduced-motion: reduce) {
    .login-box, .modal-box, .modal-icon { animation: none !important; }
}
</style>

</body>
</html>
