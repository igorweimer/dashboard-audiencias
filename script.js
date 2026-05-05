/* ============================================================
   FOREIGOR 2.0 — Lógica Principal
   Portal de Publicações Jurídicas — Gramado Parks
   ============================================================ */

// ─────────── CONFIG ───────────
const CONFIG = {
    // Google Apps Script Web App URL (será configurado após deploy)
    APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbya-JAKmKggVzjXgPBl_ypNZI9WUzubwtUePFMUw5Au4r9Oje8Zjc6UtdmgaTZk_-9J/exec',
    // Google Sheets CSV export (fallback para leitura direta)
    SHEETS_CSV_URL: 'https://docs.google.com/spreadsheets/d/1qwX3VE34d0C6lD68x00Jj09Un9Rnwh2qOninaOA9Bww/export?format=csv&gid=1744303682',
    SHEET_ID: '1qwX3VE34d0C6lD68x00Jj09Un9Rnwh2qOninaOA9Bww',
    SHEET_GID: '1744303682',
    PER_PAGE_DEFAULT: 25,
    REFRESH_INTERVAL: 60 * 1000, // 60 segundos — menos interrupções ao trabalhar
};

// ─────────── MAPEAMENTO DE TRIBUNAIS ───────────
// Formato CNJ: NNNNNNN-DD.AAAA.J.TR.OOOO
// J=8 → Justiça Estadual (Cível), J=5 → Justiça do Trabalho
const TRIBUNAL_MAP = {
    // Justiça Estadual (J=8)
    '8.01': { nome: 'TJAC - Acre', sistema: 'e-SAJ', url: 'https://esaj.tjac.jus.br/esaj/portal.do?servico=190090' },
    '8.02': { nome: 'TJAL - Alagoas', sistema: 'PJe', url: 'https://pje.tjal.jus.br/pje/login.seam' },
    '8.03': { nome: 'TJAP - Amapá', sistema: 'PJe', url: 'https://pje.tjap.jus.br/pje/login.seam' },
    '8.04': { nome: 'TJAM - Amazonas', sistema: 'e-SAJ', url: 'https://consultasaj.tjam.jus.br/cpopg/open.do' },
    '8.05': { nome: 'TJBA - Bahia', sistema: 'PJe', url: 'https://pje.tjba.jus.br/pje/login.seam' },
    '8.06': { nome: 'TJCE - Ceará', sistema: 'PJe', url: 'https://pje.tjce.jus.br/pje1grau/login.seam' },
    '8.07': { nome: 'TJDFT - Distrito Federal', sistema: 'PJe', url: 'https://pje.tjdft.jus.br/pje/login.seam' },
    '8.08': { nome: 'TJES - Espírito Santo', sistema: 'PJe', url: 'https://pje.tjes.jus.br/pje/login.seam' },
    '8.09': { nome: 'TJGO - Goiás', sistema: 'PJe', url: 'https://pje.tjgo.jus.br/pje/login.seam' },
    '8.10': { nome: 'TJMA - Maranhão', sistema: 'PJe', url: 'https://pje.tjma.jus.br/pje/login.seam' },
    '8.11': { nome: 'TJMT - Mato Grosso', sistema: 'PJe', url: 'https://pje.tjmt.jus.br/pje/login.seam' },
    '8.12': { nome: 'TJMS - Mato Grosso do Sul', sistema: 'e-SAJ', url: 'https://esaj.tjms.jus.br/esaj/portal.do?servico=190090' },
    '8.13': { nome: 'TJMG - Minas Gerais', sistema: 'PJe', url: 'https://pje.tjmg.jus.br/pje/login.seam' },
    '8.14': { nome: 'TJPA - Pará', sistema: 'PJe', url: 'https://pje.tjpa.jus.br/pje/login.seam' },
    '8.15': { nome: 'TJPB - Paraíba', sistema: 'PJe', url: 'https://pje.tjpb.jus.br/pje/login.seam' },
    '8.16': { nome: 'TJPR - Paraná', sistema: 'Projudi', url: 'https://projudi.tjpr.jus.br/projudi/' },
    '8.17': { nome: 'TJPE - Pernambuco', sistema: 'PJe', url: 'https://pje.tjpe.jus.br/pje/login.seam' },
    '8.18': { nome: 'TJPI - Piauí', sistema: 'PJe', url: 'https://pje.tjpi.jus.br/pje/login.seam' },
    '8.19': { nome: 'TJRJ - Rio de Janeiro', sistema: 'e-SAJ', url: 'https://www3.tjrj.jus.br/ejud/ConsultaProcesso.aspx' },
    '8.20': { nome: 'TJRN - Rio Grande do Norte', sistema: 'PJe', url: 'https://pje.tjrn.jus.br/pje/login.seam' },
    '8.21': { nome: 'TJRS - Rio Grande do Sul', sistema: 'eproc', url: 'https://eproc1g.tjrs.jus.br/eproc/' },
    '8.22': { nome: 'TJRO - Rondônia', sistema: 'PJe', url: 'https://pje.tjro.jus.br/pje/login.seam' },
    '8.23': { nome: 'TJRR - Roraima', sistema: 'PJe', url: 'https://pje.tjrr.jus.br/pje/login.seam' },
    '8.24': { nome: 'TJSC - Santa Catarina', sistema: 'eproc', url: 'https://eproc1g.tjsc.jus.br/eproc/' },
    '8.25': { nome: 'TJSE - Sergipe', sistema: 'PJe', url: 'https://pje.tjse.jus.br/pje/login.seam' },
    '8.26': { nome: 'TJSP - São Paulo', sistema: 'e-SAJ', url: 'https://esaj.tjsp.jus.br/esaj/portal.do?servico=190090' },
    '8.27': { nome: 'TJTO - Tocantins', sistema: 'PJe', url: 'https://pje.tjto.jus.br/pje/login.seam' },
    // Justiça do Trabalho (J=5)
    '5.01': { nome: 'TRT1 - Rio de Janeiro', sistema: 'PJe-JT', url: 'https://pje.trt1.jus.br/primeirograu/login.seam' },
    '5.02': { nome: 'TRT2 - São Paulo', sistema: 'PJe-JT', url: 'https://pje.trt2.jus.br/primeirograu/login.seam' },
    '5.03': { nome: 'TRT3 - Minas Gerais', sistema: 'PJe-JT', url: 'https://pje.trt3.jus.br/primeirograu/login.seam' },
    '5.04': { nome: 'TRT4 - Rio Grande do Sul', sistema: 'PJe-JT', url: 'https://pje.trt4.jus.br/primeirograu/login.seam' },
    '5.05': { nome: 'TRT5 - Bahia', sistema: 'PJe-JT', url: 'https://pje.trt5.jus.br/primeirograu/login.seam' },
    '5.06': { nome: 'TRT6 - Pernambuco', sistema: 'PJe-JT', url: 'https://pje.trt6.jus.br/primeirograu/login.seam' },
    '5.07': { nome: 'TRT7 - Ceará', sistema: 'PJe-JT', url: 'https://pje.trt7.jus.br/primeirograu/login.seam' },
    '5.08': { nome: 'TRT8 - Pará/Amapá', sistema: 'PJe-JT', url: 'https://pje.trt8.jus.br/primeirograu/login.seam' },
    '5.09': { nome: 'TRT9 - Paraná', sistema: 'PJe-JT', url: 'https://pje.trt9.jus.br/primeirograu/login.seam' },
    '5.10': { nome: 'TRT10 - DF/Tocantins', sistema: 'PJe-JT', url: 'https://pje.trt10.jus.br/primeirograu/login.seam' },
    '5.11': { nome: 'TRT11 - Amazonas/Roraima', sistema: 'PJe-JT', url: 'https://pje.trt11.jus.br/primeirograu/login.seam' },
    '5.12': { nome: 'TRT12 - Santa Catarina', sistema: 'PJe-JT', url: 'https://pje.trt12.jus.br/primeirograu/login.seam' },
    '5.13': { nome: 'TRT13 - Paraíba', sistema: 'PJe-JT', url: 'https://pje.trt13.jus.br/primeirograu/login.seam' },
    '5.14': { nome: 'TRT14 - Rondônia/Acre', sistema: 'PJe-JT', url: 'https://pje.trt14.jus.br/primeirograu/login.seam' },
    '5.15': { nome: 'TRT15 - Campinas', sistema: 'PJe-JT', url: 'https://pje.trt15.jus.br/primeirograu/login.seam' },
    '5.16': { nome: 'TRT16 - Maranhão', sistema: 'PJe-JT', url: 'https://pje.trt16.jus.br/primeirograu/login.seam' },
    '5.17': { nome: 'TRT17 - Espírito Santo', sistema: 'PJe-JT', url: 'https://pje.trt17.jus.br/primeirograu/login.seam' },
    '5.18': { nome: 'TRT18 - Goiás', sistema: 'PJe-JT', url: 'https://pje.trt18.jus.br/primeirograu/login.seam' },
    '5.20': { nome: 'TRT20 - Sergipe', sistema: 'PJe-JT', url: 'https://pje.trt20.jus.br/primeirograu/login.seam' },
    '5.21': { nome: 'TRT21 - Rio Grande do Norte', sistema: 'PJe-JT', url: 'https://pje.trt21.jus.br/primeirograu/login.seam' },
    '5.22': { nome: 'TRT22 - Piauí', sistema: 'PJe-JT', url: 'https://pje.trt22.jus.br/primeirograu/login.seam' },
    '5.23': { nome: 'TRT23 - Mato Grosso', sistema: 'PJe-JT', url: 'https://pje.trt23.jus.br/primeirograu/login.seam' },
    '5.24': { nome: 'TRT24 - Mato Grosso do Sul', sistema: 'PJe-JT', url: 'https://pje.trt24.jus.br/primeirograu/login.seam' },
};

// ─────────── STATE ───────────
let allPublications = [];
let filteredPublications = [];
let selectedPublications = new Set();
let currentPage = 1;
let perPage = CONFIG.PER_PAGE_DEFAULT;
let filters = {
    status: 'todos',
    area: 'todos',
    empreendimento: [],
    advogado: [],
    dateStart: null,
    dateEnd: null,
    cnjSearch: '',
};
let datepickerTarget = null;
let datepickerMonth = new Date().getMonth();
let datepickerYear = new Date().getFullYear();
let expandedUid = null;
let selectedRowUid = null;
let sortColumn = 'data';
let sortDirection = 'desc';
let statusCache = {};
// Known CNJs for "new" indicator
let knownCnjs = new Set();
try { knownCnjs = new Set(JSON.parse(localStorage.getItem('foreigor_known_cnjs')) || []); } catch(e) {}
// Delegation log — persists across page reloads (fallback when Sheets is slow)
let delegationLog = {};
try { delegationLog = JSON.parse(localStorage.getItem('foreigor_delegation_log')) || {}; } catch(e) { delegationLog = {}; }
// Optimistic updates: alterações locais que permanecem até a confirmação pelo servidor.
let optimisticUpdates = {};
try { optimisticUpdates = JSON.parse(localStorage.getItem('foreigor_optimistic_updates')) || {}; } catch(e) { optimisticUpdates = {}; }

// ─── LOCKED STATUSES — "Cimento e Concreto" ───
// Publicações tratadas ficam TRAVADAS aqui. Uma vez marcada como LIDO ou DESCONSIDERADO,
// a publicação nunca volta a ser NÃO LIDO automaticamente — só se o usuário desmarcar manualmente.
// Estrutura: { [_uid]: { status, cnj, lockedAt } }
let lockedStatuses = {};
try { lockedStatuses = JSON.parse(localStorage.getItem('foreigor_locked_statuses')) || {}; } catch(e) { lockedStatuses = {}; }

// Fila de Sincronização em Lote (Batch Sync)
let syncQueue = [];
try { syncQueue = JSON.parse(localStorage.getItem('foreigor_sync_queue')) || []; } catch(e) { syncQueue = []; }
let isSyncing = false;

function saveSyncQueue() {
    try { localStorage.setItem('foreigor_sync_queue', JSON.stringify(syncQueue)); } catch(e) {}
}

function saveLockedStatuses() {
    try { localStorage.setItem('foreigor_locked_statuses', JSON.stringify(lockedStatuses)); } catch(e) {}
}

// ─────────── INIT ───────────
document.addEventListener('DOMContentLoaded', () => {
    loadStatusCache();
    setupEventListeners();
    injectConfirmModal();
    fetchPublications();
    // Auto-refresh da planilha (Google Sheets é a única fonte da verdade)
    setInterval(fetchPublications, CONFIG.REFRESH_INTERVAL);
    // Processamento da fila de uploads em lote
    setInterval(processSyncQueue, 3000);
});

// ─── Detecta se o usuário está digitando em algum campo do accordion ───
function isUserEditing() {
    const el = document.activeElement;
    if (!el) return false;
    const tag = el.tagName.toLowerCase();
    return (tag === 'textarea' || tag === 'input') && !!el.closest('.detail-row, .delegation-box, .mbox-subject-row');
}

// ─── Modal de confirmação (caixinha no canto inferior) ───
let _confirmCallback = null;
function injectConfirmModal() {
    const div = document.createElement('div');
    div.id = 'confirm-read-modal';
    div.style.cssText = 'display:none;position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#1e293b;color:#f8fafc;padding:16px 24px;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,.4);z-index:9999;display:none;align-items:center;gap:16px;font-family:Inter,sans-serif;font-size:.9rem;min-width:320px;';
    div.innerHTML = `
        <span id="confirm-read-msg" style="flex:1">Tem certeza que quer marcar como <strong>Lida</strong>?</span>
        <button onclick="document.getElementById('confirm-read-modal').style.display='none';_confirmCallback=null;" style="background:#475569;color:#fff;border:none;padding:7px 14px;border-radius:7px;cursor:pointer;font-size:.85rem;">Cancelar</button>
        <button onclick="if(_confirmCallback){_confirmCallback();_confirmCallback=null;}document.getElementById('confirm-read-modal').style.display='none';" style="background:#10b981;color:#fff;border:none;padding:7px 14px;border-radius:7px;cursor:pointer;font-size:.85rem;font-weight:700;">✓ Sim</button>
    `;
    document.body.appendChild(div);
}
function askConfirmRead(msg, callback) {
    const modal = document.getElementById('confirm-read-modal');
    if (!modal) { callback(); return; }
    document.getElementById('confirm-read-msg').innerHTML = msg;
    _confirmCallback = callback;
    modal.style.display = 'flex';
}

// ─────────── FETCH DATA ───────────
async function fetchPublications() {
    // Se o usuário está digitando, atualiza os dados mas NÃO re-renderiza
    // para não apagar o que ele está escrevendo
    if (isUserEditing()) {
        try {
            const cacheBust = `?t=${Date.now()}`;
            const response = await fetch(CONFIG.APPS_SCRIPT_URL + cacheBust);
            const data = await response.json();
            allPublications = data.map(normalizePublication);
            detectNewPublications();
            setSyncStatus('ok', 'Conectado (modo edição)');
        } catch(e) { /* silencioso */ }
        return;
    }

    setSyncStatus('loading', 'Sincronizando...');

    try {
        if (CONFIG.APPS_SCRIPT_URL) {
            const cacheBust = `?t=${Date.now()}`;
            const response = await fetch(CONFIG.APPS_SCRIPT_URL + cacheBust);
            const data = await response.json();
            allPublications = data.map(normalizePublication);
        } else {
            allPublications = await fetchViaGoogleViz();
        }

        loadEmailQueueFromSheets();
        detectNewPublications();
        // Preservar a página atual durante auto-refresh (quando já havia dados carregados)
        const isAutoRefresh = filteredPublications.length > 0;
        applyFilters(isAutoRefresh);
        populateFilterDropdowns();
        updateCounters();
        setSyncStatus('ok', 'Conectado');
    } catch (err) {
        console.error('Erro ao buscar publicações:', err);
        setSyncStatus('error', 'Sem internet / ' + (err.message || 'Erro desconhecido'));
        if (allPublications.length === 0) {
            document.getElementById('pub-table-body').innerHTML = `
                <tr><td colspan="7" class="empty-state"><div class="empty-content">
                    <i data-lucide="wifi-off" style="width:40px;height:40px;color:var(--danger);"></i>
                    <p>Erro ao carregar publicações. Verifique sua conexão.</p>
                    <button onclick="fetchPublications()" class="action-btn btn-mark-read" style="margin-top:8px;">
                        <i data-lucide="refresh-cw"></i> Tentar novamente
                    </button>
                </div></td></tr>`;
            lucide.createIcons();
        }
    }
}


const DETECTION_DATA = {
    companies: [
        {"name": "GRAMADO PROMOÇÃO DE VENDAS S.A.", "cnpj": "25.381.865/0001-76"},
        {"name": "GRAMADO PROMOÇÃO DE VENDAS S.A.", "cnpj": "25.381.865/0009-23"},
        {"name": "GRAMADO PROMOÇÃO DE VENDAS S.A.", "cnpj": "25.381.865/0005-08"},
        {"name": "GRAMADO PROMOÇÃO DE VENDAS S.A.", "cnpj": "25.381.865/0004-19"},
        {"name": "GRAMADO PARKS INVESTIMENTOS E INTERMEDIACOES S.A - EM RECUPERACAO JUDICIAL", "cnpj": "00.369.161/0001-57"},
        {"name": "GRAMADO PARKS INVESTIMENTOS E INTERMEDIACOES LTDA - SCP (GRAMADO EXCLUSIVE RESORT)", "cnpj": "24.551.706/0001-00"},
        {"name": "GRAMADO PARKS INVESTIMENTOS E INTERMEDIACOES LTDA - SCP BUONA VITTA RESORT", "cnpj": "26.748.943/0001-90"},
        {"name": "BRASIL PARQUES TEMATICOS E DE DIVERSAO S/A - EM RECUPERACAO JUDICIAL", "cnpj": "37.233.270/0001-52"},
        {"name": "ARC RIO PARQUES TEMATICOS E DE DIVERSAO S.A.", "cnpj": "30.309.571/0001-73"},
        {"name": "GRAMADO TERMAS PARK PARQUES TEMATICOS LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "15.195.705/0001-89"},
        {"name": "GRAMADO TERMAS PARK PARQUES TEMATICOS LTDA SCP", "cnpj": "46.562.763/0001-27"},
        {"name": "SNOWLAND PARTICIPACOES E CONSULTORIA LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "13.820.324/0001-18"},
        {"name": "PARQUE DA NEVE GRAMADO - SCP", "cnpj": "23.859.210/0001-35"},
        {"name": "FOZ STAR PARQUES TEMATICOS E DE DIVERSAO LTDA. - EM RECUPERACAO JUDICIAL", "cnpj": "37.546.880/0001-06"},
        {"name": "FERRIS WHEEL - INVESTIMENTOS E PARTICIPACOES LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "32.522.523/0001-94"},
        {"name": "PARQUE AQUATICO CARNEIROS SPE LTDA", "cnpj": "35.830.898/0001-00"},
        {"name": "GRAMADO HYDROS INCORPORACOES - SPE LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "29.989.181/0001-02"},
        {"name": "GTR HOTEIS E RESORT LTDA", "cnpj": "16.966.397/0001-00"},
        {"name": "CARNEIROS RESORT INCORPORACOES SPE LTDA", "cnpj": "35.805.067/0001-88"},
        {"name": "PRIME FOZ INCORPORACOES SPE S/A", "cnpj": "30.870.334/0001-87"},
        {"name": "GP RESTAURANTE LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "31.010.847/0001-80"},
        {"name": "GP RESTAURANTE LTDA - FILIAL 1 - GER", "cnpj": "31.010.847/0002-61"},
        {"name": "GP RESTAURANTE LTDA - FILIAL 2 - GBV", "cnpj": "31.010.847/0003-42"},
        {"name": "GP RESTAURANTE LTDA - FILIAL 3 - GVI", "cnpj": "31.010.847/0004-23"},
        {"name": "GP RESTAURANTE LTDA", "cnpj": "31.010.847/0005-04"},
        {"name": "GP RESTAURANTE LTDA", "cnpj": "31.010.847/0006-95"},
        {"name": "MAGIC SNOWLAND OPERADORA TURISTICA LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "36.479.337/0001-70"},
        {"name": "MAGIC SNOWLAND OPERADORA TURISTICA LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "36.479.337/0002-51"},
        {"name": "GRAMADO PRIME ADMINISTRACAO HOTELEIRA LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "38.382.915/0001-81"},
        {"name": "GRAMADO PRIME ADMINISTRACAO HOTELEIRA LTDA - GBV", "cnpj": "38.382.915/0003-43"},
        {"name": "GRAMADO PRIME ADMINISTRACAO HOTELEIRA LTDA - GVI", "cnpj": "38.382.915/0004-24"},
        {"name": "GRAMADO PRIME ADMINISTRACAO HOTELEIRA LTDA - GER", "cnpj": "38.382.915/0002-62"},
        {"name": "GRAMADO PRIME ADMINISTRACAO HOTELEIRA LTDA - SCP I - GER", "cnpj": "42.428.612/0001-20"},
        {"name": "GRAMADO PRIME ADMINISTRACAO HOTELEIRA LTDA - SCP II - GBV", "cnpj": "42.474.929/0001-00"},
        {"name": "GRAMADO PRIME ADMINISTRACAO HOTELEIRA LTDA - SCP III - GVI", "cnpj": "42.714.088/0001-53"},
        {"name": "CONDOMINIO GRAMADO BUONA VITTA RESORT SPA", "cnpj": "42.383.685/0001-42"},
        {"name": "CONDOMINIO GRAMADO BV RESORT", "cnpj": "36.775.916/0001-60"},
        {"name": "CONDOMINIO GRAMADO EXCLUSIVE RESORT", "cnpj": "36.775.850/0001-09"},
        {"name": "GRAMADO MUSEU DO FESTIVAL DE CINEMA LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "22.584.232/0001-77"},
        {"name": "PRAIA DO FORTE PARQUES TEMATICOS E DE DIVERSAO LTDA", "cnpj": "44.678.210/0001-09"},
        {"name": "PRAIA DO FORTE CLUB LTDA", "cnpj": "44.908.269/0001-46"},
        {"name": "PRAIA DO FORTE RESORT LTDA", "cnpj": "44.611.928/0001-88"},
        {"name": "PRAIA DO FORTE VILLAS LTDA", "cnpj": "44.907.937/0001-10"},
        {"name": "GP VACATION CLUB LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "23.279.530/0001-16"},
        {"name": "GPK FIDELIDADE LTDA", "cnpj": "44.287.804/0001-99"},
        {"name": "GRAMADO BV RESORT INCORPORACOES SPE LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "23.448.583/0001-13"},
        {"name": "GRAMADO PARKS TURISMO LTDA", "cnpj": "41.742.472/0001-05"},
        {"name": "JARDIM CANELA INCORPORACOES LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "23.991.346/0001-02"},
        {"name": "LAGO-NEGRO RESTAURANTE LTDA - EM RECUPERACAO JUDICIAL", "cnpj": "13.747.277/0001-24"}
    ],
    lawyers: [
        {"name": "RACHEL BROCK", "oab": "49.636"},
        {"name": "PAULA RENATA MONTEIRO DE BRITO", "oab": "109.453"},
        {"name": "RENAN PERIM SIQUEIRA", "oab": "132.154A"},
        {"name": "GIOVANI MARCEL GONÇALVES DA SILVA", "oab": "104.372"}
    ]
};


function extractEntitiesFromTeor(teor, cnj) {
    let empFound = "NÃO IDENTIFICADO";
    let advFound = "NÃO IDENTIFICADO";
    if (!teor) return { empFound, advFound };

    const upperTeor = (String(teor).length > 3000 ? String(teor).substring(0, 3000) : String(teor)).toUpperCase();
    const cleanTeorForCNPJ = upperTeor.replace(/[.\-/]/g, '');

    // 1. Procurar Advogados (Pelo Nome ou OAB)
    for (const adv of DETECTION_DATA.lawyers) {
        const cleanOAB = adv.oab.replace(/[.\-/]/g, '');
        if (upperTeor.includes(adv.name) || cleanTeorForCNPJ.includes(cleanOAB)) {
            advFound = adv.name;
            break;
        }
    }

    // 2. Procurar Empreendimentos (Pelo Nome ou CNPJ)
    for (const emp of DETECTION_DATA.companies) {
        const cleanCNPJ = emp.cnpj.replace(/[.\-/]/g, '');
        if (upperTeor.includes(emp.name) || cleanTeorForCNPJ.includes(cleanCNPJ)) {
            empFound = emp.name;
            break;
        }
    }

    // Fallback: inteligente se nada foi encontrado da lista principal
    if (empFound === "NÃO IDENTIFICADO") {
        const match = upperTeor.match(/(?:EXECUTADO\(S\)|REQUERIDO\(S?\)|R[ÉE]U|DEVEDOR)\s*:\s*([^;]+?)(?=\s*(?:SENTENÇA|DECISÃO|DESPACHO|ATO|CITAÇÃO|TRATA-SE|CLASSE|\.|\n|$))/i);
        if (match && match[1]) {
            let possibleName = match[1].trim();
            if (possibleName.length > 5 && possibleName.length < 80) empFound = possibleName;
        }
    }
    
    return { empFound, advFound };
}

// Google Visualization API — suporta CORS nativamente
async function fetchViaGoogleViz() {
    const gid = CONFIG.SHEET_GID;
    const url = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:json&gid=${gid}`;

    const response = await fetch(url);
    const text = await response.text();

    const jsonStr = text.match(/google\.visualization\.Query\.setResponse\(([\s\S]*)\);?/);
    if (!jsonStr || !jsonStr[1]) {
        throw new Error('Formato de resposta inesperado do Google Sheets');
    }

    const data = JSON.parse(jsonStr[1]);
    const cols = data.table.cols;
    const rows = data.table.rows;
    const publications = [];

    // Mapeamento dinâmico de colunas
    const colIdx = {};
    cols.forEach((c, idx) => {
        let label = (c.label || '').trim();
        // Se a primeira coluna tiver o script do usuário acidentalmente, consideramos 'Data Disponibilização'
        if (idx === 0 && (label.startsWith('fetch') || label === '')) label = 'Data Disponibilização';
        colIdx[label] = idx;
    });

    const getVal = (rowCells, labelPossibilities) => {
        // Tenta match exato primeiro
        for (let label of labelPossibilities) {
            const idx = colIdx[label];
            if (idx !== undefined && rowCells[idx]) {
                const val = rowCells[idx].f || rowCells[idx].v || '';
                return typeof val === 'string' ? val.trim() : val;
            }
        }
        // Tenta match parcial (case insensitive)
        for (let label of labelPossibilities) {
            for (let cName in colIdx) {
                if (cName.toLowerCase().includes(label.toLowerCase()) && rowCells[colIdx[cName]]) {
                    const val = rowCells[colIdx[cName]].f || rowCells[colIdx[cName]].v || '';
                    return typeof val === 'string' ? val.trim() : val;
                }
            }
        }
        return '';
    };

    const getFormat = (rowCells, label) => {
        const idx = colIdx[label];
        if (idx !== undefined && rowCells[idx]) return rowCells[idx].f || '';
        return '';
    };

    for (const row of rows) {
        const cells = row.c;
        if (!cells) continue;

        let dataDisp = getVal(cells, ['Data', 'Data Disponibilização', 'Disponibilização', 'Disp']);
        // Fallback fixo para a Coluna B (índice 1) da planilha
        if (!dataDisp && cells[1]) {
            dataDisp = cells[1].f || cells[1].v || '';
        }
        
        let dataStr = '';
        try {
            if (typeof dataDisp === 'object' && dataDisp instanceof Date) {
                dataStr = `${String(dataDisp.getDate()).padStart(2,'0')}/${String(dataDisp.getMonth()+1).padStart(2,'0')}/${dataDisp.getFullYear()}`;
            } else if (typeof dataDisp === 'string' && dataDisp.match(/^\d{4}-\d{2}-\d{2}T/)) {
                const d = new Date(dataDisp);
                dataStr = `${String(d.getUTCDate()).padStart(2,'0')}/${String(d.getUTCMonth()+1).padStart(2,'0')}/${d.getUTCFullYear()}`;
            } else if (dataDisp) {
                dataStr = String(dataDisp).substring(0, 10);
            }
        } catch (e) {
            dataStr = String(dataDisp);
        }
        const cnj = getVal(cells, ['Número Processo (CNJ)', 'CNJ']);
        const tipoPub = getVal(cells, ['Tipo Publicação Tribunal', 'Tipo Publicação']);
        const idPubRaw = getVal(cells, ['Id Publicação']);
        const idPub = (idPubRaw && String(idPubRaw).trim()) ? String(idPubRaw).trim() : '';
        const orgao = getVal(cells, ['Orgão Expedidor', 'Órgão', 'Vara']);
        let textoPub = getVal(cells, ['Texto', 'Teor', 'Conteú']); // "Conteú" matches Conteúdo
        const idIntelivix = getVal(cells, ['Id_Intelivix']);
        const statusSpreadsheet = getVal(cells, ['Status']);
        const observacoesPlanilha = getVal(cells, ['Observações', 'Observacoes', 'Observacao']);
        
        // Se textoPub ainda vazio ou muito curto, pegar o maior texto disponível (fallback seguro para leitura da planilha)
        if (!textoPub || textoPub.length < 50) {
            let maxLen = 0;
            let longestStr = textoPub;
            for (let i = 0; i < cells.length; i++) {
                if (cells[i] && typeof cells[i].v === 'string' && !cells[i].v.includes('|')) {
                    if (cells[i].v.length > maxLen) {
                        maxLen = cells[i].v.length;
                        longestStr = cells[i].v;
                    }
                }
            }
            if (maxLen > 50) textoPub = longestStr;
        }

        // Se Empreendimento/Advogado foram exportados
        let empreendimento = getVal(cells, ['Empreendimento', 'Projeto', 'Cliente']);
        let advogado = getVal(cells, ['Advogado', 'Responsável', 'Profissional']);
        const entities = extractEntitiesFromTeor(textoPub, cnj);

        if (!empreendimento || empreendimento === 'Não Identificado') empreendimento = entities.empFound;
        if (!advogado || advogado === 'Não Identificado') advogado = entities.advFound;
        
        if (!cnj || !dataDisp) continue;

        const cnjStr = String(cnj).trim();
        if (!/\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}/.test(cnjStr)) continue;

        const cnjStr2 = String(cnj).trim();
        
        // UID DETERMINÍSTICO: idPub + CNJ + hash do teor
        // Isso garante que o UID seja idêntico entre reloads e sincronizações
        const _teorHash = String(textoPub || '').length;
        const _uid = `${idPub || 'noid'}_${cnjStr2}_${_teorHash}`;

        // Status precedence: Locked > Optimistic > Planilha
        let rawStatus = String(statusSpreadsheet).trim().toUpperCase();
        let finalStatus;

        // 1. LOCKED (cimento e concreto): status travado pelo usuário — imutável até desmarcar manualmente
        if (lockedStatuses[_uid]) {
            finalStatus = lockedStatuses[_uid].status;
        }
        // 2. OPTIMISTIC: mudança local ainda não confirmada pelo servidor
        else if (optimisticUpdates[_uid] || optimisticUpdates[cnjStr2]) {
            finalStatus = (optimisticUpdates[_uid] || optimisticUpdates[cnjStr2]).status;
        }
        // 3. PLANILHA: fonte da verdade remota
        else {
            if (rawStatus === 'LIDO' || rawStatus === 'LIDA') {
                finalStatus = 'LIDO';
                // Auto-travar: planilha confirmou LIDO → proteger localmente
                if (!lockedStatuses[_uid]) {
                    lockedStatuses[_uid] = { status: 'LIDO', cnj: cnjStr2, lockedAt: Date.now() };
                    saveLockedStatuses();
                }
            } else if (rawStatus === 'DESCONSIDERADO' || rawStatus === 'IGNORADO') {
                finalStatus = 'DESCONSIDERADO';
                // Auto-travar: planilha confirmou DESCONSIDERADO → proteger localmente
                if (!lockedStatuses[_uid]) {
                    lockedStatuses[_uid] = { status: 'DESCONSIDERADO', cnj: cnjStr2, lockedAt: Date.now() };
                    saveLockedStatuses();
                }
            } else if (rawStatus === 'NÃO LIDO' || rawStatus === 'NAO LIDO' || rawStatus === 'NÃO LIDA' || rawStatus === 'NAO LIDA') {
                finalStatus = 'NÃO LIDO';
            } else {
                finalStatus = 'NÃO LIDO';
            }
        }

        // Órgão Expedidor: usar da planilha ou detectar automaticamente pelo CNJ
        const detectedTribunal = detectTribunal(cnj);
        const orgaoFinal = orgao || (detectedTribunal ? detectedTribunal.nome : '');
        publications.push({
            _uid: _uid,
            dataDisponibilizacao: dataStr, // Alias para sort actions
            dataStr: dataStr,
            cnj: cnj,
            tipoPublicacao: tipoPub,
            idPublicacao: idPub,
            orgaoExpedidor: orgaoFinal,
            textoPublicacao: textoPub || 'Conteúdo não disponível',
            idIntelivix: idIntelivix,
            empreendimento: empreendimento || 'Não Identificado',
            advogado: advogado || 'Não Identificado',
            area: detectArea(cnj),
            tribunal: detectedTribunal,
            status: finalStatus,
            observacoes: observacoesPlanilha || '',
        });
    }

    return publications;
}

// ─────────── PARSE CSV ───────────
function parseCSV(csvText) {
    const lines = csvText.split('\n');
    const publications = [];

    // Find header line
    let headerIdx = 0;
    for (let i = 0; i < Math.min(lines.length, 5); i++) {
        if (lines[i].includes('Data Disponibilização') || lines[i].includes('Número Processo')) {
            headerIdx = i;
            break;
        }
    }

    // Parse records — handle multi-line CSV fields (quoted)
    const records = parseCSVRecords(lines.slice(headerIdx + 1).join('\n'));

    for (const fields of records) {
        if (fields.length < 6) continue;

        const dataDisp = fields[0]?.trim();
        const cnj = fields[1]?.trim();
        const tipoPub = fields[2]?.trim();
        const idPub = fields[3]?.trim();
        const orgao = fields[4]?.trim();
        // fields[5] = Número (not used)
        const textoPub = fields[6]?.trim() || '';
        const idIntelivix = fields[7]?.trim() || '';

        if (!cnj || !dataDisp) continue;

        const _uid = String(idPub || `${cnj}_${dataDisp}_${(textoPub||'').length}`);
        const pub = {
            _uid: _uid,
            dataDisponibilizacao: dataDisp,
            cnj: cnj,
            tipoPublicacao: tipoPub,
            idPublicacao: idPub,
            orgaoExpedidor: orgao,
            textoPublicacao: textoPub,
            idIntelivix: idIntelivix,
            area: detectArea(cnj),
            tribunal: detectTribunal(cnj),
            status: getStatus(cnj),
        };

        publications.push(pub);
    }

    return publications;
}

function parseCSVRecords(text) {
    const records = [];
    let currentRecord = [];
    let currentField = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const nextChar = text[i + 1];

        if (inQuotes) {
            if (char === '"' && nextChar === '"') {
                currentField += '"';
                i++; // skip next quote
            } else if (char === '"') {
                inQuotes = false;
            } else {
                currentField += char;
            }
        } else {
            if (char === '"') {
                inQuotes = true;
            } else if (char === ',') {
                currentRecord.push(currentField);
                currentField = '';
            } else if (char === '\r') {
                // skip \r
            } else if (char === '\n') {
                currentRecord.push(currentField);
                currentField = '';
                if (currentRecord.some(f => f.trim())) {
                    records.push(currentRecord);
                }
                currentRecord = [];
            } else {
                currentField += char;
            }
        }
    }

    // Last record
    if (currentField || currentRecord.length > 0) {
        currentRecord.push(currentField);
        if (currentRecord.some(f => f.trim())) {
            records.push(currentRecord);
        }
    }

    return records;
}

function normalizePublication(row) {
    const rawDate = row['Data Disponibilização'] || row.dataDisponibilizacao || '';
    const cnj = row['Número Processo (CNJ)'] || row.cnj || '';
    const textoPub = row['Texto Publicação'] || row.textoPublicacao || '';
    const orgaoRaw = row['Orgão Expedidor Publicação'] || row.orgaoExpedidor || '';

    // --- Format date to DD/MM/YYYY ---
    let dataStr = '';
    try {
        if (typeof rawDate === 'string' && rawDate.match(/^\d{4}-\d{2}-\d{2}/)) {
            const d = new Date(rawDate);
            dataStr = `${String(d.getUTCDate()).padStart(2,'0')}/${String(d.getUTCMonth()+1).padStart(2,'0')}/${d.getUTCFullYear()}`;
        } else if (typeof rawDate === 'string' && rawDate.match(/^\d{2}\/\d{2}\/\d{4}/)) {
            dataStr = rawDate.substring(0, 10); // Already DD/MM/YYYY
        } else if (rawDate) {
            dataStr = String(rawDate).substring(0, 10);
        }
    } catch(e) {
        dataStr = String(rawDate);
    }

    // --- NLP: Extract empreendimento and advogado from the publication text ---
    const entities = extractEntitiesFromTeor(textoPub, cnj);
    const empreendimento = entities.empFound;
    const advogado = entities.advFound;

    // --- Tribunal and Area detection ---
    const detectedTribunal = detectTribunal(cnj);
    const orgaoFinal = orgaoRaw || (detectedTribunal ? detectedTribunal.nome : '');

    // --- Status (Sheets-First — PLANILHA É A ÚNICA FONTE DA VERDADE) ---
    const _cnjNorm = String(cnj).trim();
    const idPubRaw = row['Id Publicação'] || row.idPublicacao || '';
    const _teorHash = String(textoPub || '').length;
    const _uid = `${idPubRaw || 'noid'}_${_cnjNorm}_${_teorHash}`;

    // --- Status (Prioridade: Locked > Optimistic > Planilha) ---
    let rawStatus = String(row.Status || row.status || '').trim().toUpperCase();
    let finalStatus;

    // 1. LOCKED (cimento e concreto): status travado pelo usuário — imutável até desmarcar manualmente
    if (lockedStatuses[_uid]) {
        finalStatus = lockedStatuses[_uid].status;
    }
    // 2. OPTIMISTIC: mudança local ainda não confirmada pelo servidor
    else if (optimisticUpdates[_uid] || optimisticUpdates[_cnjNorm]) {
        finalStatus = (optimisticUpdates[_uid] || optimisticUpdates[_cnjNorm]).status;
    }
    // 3. PLANILHA: fonte da verdade remota
    else {
        if (rawStatus === 'LIDO' || rawStatus === 'LIDA') {
            finalStatus = 'LIDO';
            // Auto-travar: planilha confirmou LIDO → proteger localmente
            if (!lockedStatuses[_uid]) {
                lockedStatuses[_uid] = { status: 'LIDO', cnj: _cnjNorm, lockedAt: Date.now() };
                saveLockedStatuses();
            }
        } else if (rawStatus === 'DESCONSIDERADO' || rawStatus === 'IGNORADO') {
            finalStatus = 'DESCONSIDERADO';
            // Auto-travar: planilha confirmou DESCONSIDERADO → proteger localmente
            if (!lockedStatuses[_uid]) {
                lockedStatuses[_uid] = { status: 'DESCONSIDERADO', cnj: _cnjNorm, lockedAt: Date.now() };
                saveLockedStatuses();
            }
        } else if (rawStatus === 'NÃO LIDO' || rawStatus === 'NAO LIDO' || rawStatus === 'NÃO LIDA' || rawStatus === 'NAO LIDA') {
            finalStatus = 'NÃO LIDO';
        } else {
            finalStatus = 'NÃO LIDO';
        }
    }
    return {
        _uid: _uid,
        dataDisponibilizacao: dataStr,
        dataStr: dataStr,
        cnj: cnj,
        tipoPublicacao: row['Tipo Publicação Tribunal'] || row.tipoPublicacao || '',
        idPublicacao: row['Id Publicação'] || row.idPublicacao || '',
        orgaoExpedidor: orgaoFinal,
        textoPublicacao: textoPub || 'Conteúdo não disponível',
        idIntelivix: row['Id_Intelivix'] || row.idIntelivix || '',
        empreendimento: empreendimento || 'Não Identificado',
        advogado: advogado || 'Não Identificado',
        area: detectArea(cnj),
        tribunal: detectedTribunal,
        status: finalStatus,
        observacoes: row['Observações'] || row['Observacoes'] || row.observacoes || '',
    };
}

// ─────────── CNJ HELPERS ───────────
// CNJ: NNNNNNN-DD.AAAA.J.TR.OOOO
function parseCNJ(cnj) {
    // Remove any extra chars
    const match = cnj.match(/(\d{7})-(\d{2})\.(\d{4})\.(\d)\.(\d{2})\.(\d{4})/);
    if (!match) return null;
    return {
        sequencial: match[1],
        digito: match[2],
        ano: match[3],
        justica: match[4],  // J
        tribunal: match[5], // TR
        origem: match[6],   // OOOO
    };
}

function detectArea(cnj) {
    const parsed = parseCNJ(cnj);
    if (!parsed) return 'desconhecido';
    if (parsed.justica === '5') return 'trabalhista';
    if (parsed.justica === '8') return 'civel';
    return 'outro';
}

function detectTribunal(cnj) {
    const parsed = parseCNJ(cnj);
    if (!parsed) return null;
    const key = `${parsed.justica}.${parsed.tribunal}`;
    let tribunal = TRIBUNAL_MAP[key] || { nome: `Tribunal ${parsed.tribunal}`, sistema: 'Desconhecido', url: '' };
    
    // Evita mutar o mapa global de tribunais
    tribunal = { ...tribunal };

    // Regra EPROC SP e RJ: Se o sequencial iniciar com '4'
    if ((key === '8.26' || key === '8.19') && parsed.sequencial.startsWith('4')) {
        tribunal.sistema = 'eproc';
        if (key === '8.26') { tribunal.url = 'https://eproc1g.tjsp.jus.br/eproc/'; tribunal.nome = 'TJSP - Eproc'; }
        if (key === '8.19') { tribunal.url = 'https://eproc1g.tjrj.jus.br/eproc/'; tribunal.nome = 'TJRJ - Eproc'; }
    }

    return tribunal;
}

// ─────────── STATUS CACHE (localStorage) ───────────
function getStatus(cnjOrUid) {
    return statusCache[cnjOrUid] || 'NÃO LIDO';
}

function setStatus(cnj, status, _uid, idPublicacao) {
    const optKey = _uid || cnj;

    // ── LOCKED STATUSES (cimento e concreto) ──
    if (status === 'LIDO' || status === 'DESCONSIDERADO') {
        // Travar: esse status nunca volta sozinho enquanto não desmarcar manualmente
        lockedStatuses[optKey] = { status: status, cnj: cnj, lockedAt: Date.now() };
        saveLockedStatuses();
    } else {
        // NÃO LIDO = usuário desmarcou manualmente → remover o travamento
        if (lockedStatuses[optKey]) {
            delete lockedStatuses[optKey];
            saveLockedStatuses();
        }
        // Também limpar optimistic para que o próximo fetch busque da planilha
        if (optimisticUpdates[optKey]) { delete optimisticUpdates[optKey]; saveOptimisticUpdates(); }
    }

    // 1. Optimistic: chave precisa para consistência visual imediata
    optimisticUpdates[optKey] = { status: status, timestamp: Date.now() };
    saveOptimisticUpdates();

    // 2. statusCache como fallback offline
    statusCache[cnj] = status;
    if (_uid) statusCache[_uid] = status;
    saveStatusCache();

    // 3. Fila de sync — inclui idPublicacao para o Apps Script achar a linha exata
    const existingIdx = syncQueue.findIndex(u => u._uid === optKey);
    if (existingIdx !== -1) {
        syncQueue[existingIdx].status = status;
    } else {
        syncQueue.push({ cnj: cnj, status: status, _uid: optKey, idPublicacao: idPublicacao || '', observacoes: '' });
    }

    saveSyncQueue(); // Persiste a fila
    processSyncQueue();
}

async function processSyncQueue() {
    if (isSyncing || syncQueue.length === 0 || !CONFIG.APPS_SCRIPT_URL) return;
    
    isSyncing = true;
    const batch = [...syncQueue]; // Cópia da fila atual
    
    try {
        console.log(`[Sync] Enviando lote de ${batch.length} atualizações...`);
        setSyncStatus('loading', `Salvando ${batch.length} status...`);
        
        const res = await fetch(CONFIG.APPS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({ action: 'updateStatusBatch', updates: batch }),
        });
        
        if (res.ok) {
            console.log(`[Sync] ✓ Lote de ${batch.length} salvo com sucesso na planilha.`);

            // Remove da fila apenas os itens enviados
            syncQueue = syncQueue.filter(qItem => !batch.some(bItem => bItem._uid === qItem._uid && bItem.status === qItem.status));
            saveSyncQueue();

            // Limpar optimisticUpdates pelos _uid confirmados
            // NOTA: NÃO limpamos lockedStatuses aqui — eles ficam até o usuário desmarcar manualmente.
            // O optimistic pode ser limpo pois a planilha já foi atualizada.
            batch.forEach(item => {
                const key = item._uid || item.cnj;
                // Só limpar o optimistic se o locked já protege esse item
                // (evita janela de tempo entre limpeza do optimistic e próximo fetch)
                if (lockedStatuses[key]) {
                    delete optimisticUpdates[key];
                }
            });
            saveOptimisticUpdates();

            setSyncStatus('ok', 'Conectado');
        } else {
            console.warn(`[Sync] ✗ Resposta não-OK da planilha no lote.`);
            setSyncStatus('error', 'Sem internet / Falha ao sincronizar');
        }
    } catch (err) {
        console.error(`[Sync] ✗ Erro ao salvar lote:`, err);
        setSyncStatus('error', 'Sem internet / Falha na conexão');
        // Mantém na fila para re-tentativa na próxima rodada do setInterval
    } finally {
        isSyncing = false;
    }
}

function loadStatusCache() {
    try {
        const cached = localStorage.getItem('foreigor_status_cache');
        if (cached) statusCache = JSON.parse(cached);
    } catch (e) {
        statusCache = {};
    }

    // ── MIGRAÇÃO ÚNICA: promover entradas antigas para lockedStatuses ──
    // Executa apenas uma vez por sessão para não repetir processamento.
    // Garante que publicações tratadas ANTES desta atualização permaneçam protegidas.
    if (!sessionStorage.getItem('foreigor_locked_migrated')) {
        let migrated = 0;
        const LOCK_STATUSES = ['LIDO', 'LIDA', 'DESCONSIDERADO', 'IGNORADO'];

        // 1. Migrar do statusCache (chaves antigas por CNJ ou _uid)
        for (const key in statusCache) {
            const st = String(statusCache[key] || '').trim().toUpperCase();
            if (LOCK_STATUSES.includes(st) && !lockedStatuses[key]) {
                const normalized = (st === 'LIDA') ? 'LIDO' : (st === 'IGNORADO') ? 'DESCONSIDERADO' : st;
                lockedStatuses[key] = { status: normalized, cnj: key, lockedAt: Date.now() };
                migrated++;
            }
        }

        // 2. Migrar do optimisticUpdates (chaves por _uid)
        for (const key in optimisticUpdates) {
            const entry = optimisticUpdates[key];
            const st = String(entry && entry.status || '').trim().toUpperCase();
            if (LOCK_STATUSES.includes(st) && !lockedStatuses[key]) {
                const normalized = (st === 'LIDA') ? 'LIDO' : (st === 'IGNORADO') ? 'DESCONSIDERADO' : st;
                lockedStatuses[key] = { status: normalized, cnj: key, lockedAt: Date.now() };
                migrated++;
            }
        }

        if (migrated > 0) {
            saveLockedStatuses();
            console.log(`[ForeIgor] Migração: ${migrated} status travados com sucesso.`);
        }
        sessionStorage.setItem('foreigor_locked_migrated', '1');
    }
}

function saveStatusCache() {
    try {
        localStorage.setItem('foreigor_status_cache', JSON.stringify(statusCache));
    } catch (e) {
        console.error('Erro ao salvar cache:', e);
    }
}

function saveOptimisticUpdates() {
    try {
        localStorage.setItem('foreigor_optimistic_updates', JSON.stringify(optimisticUpdates));
    } catch (e) {
        console.error('Erro ao salvar optimisticUpdates:', e);
    }
}

// ─────────── FILTERS ───────────
function populateFilterDropdowns() {
    const empreendimentos = new Set();
    const advogados = new Set();
    
    // 1. Pegar tudo que veio das publicações (dinâmico)
    allPublications.forEach(pub => {
        if (pub.empreendimento && pub.empreendimento !== 'Não Identificado') empreendimentos.add(pub.empreendimento);
        if (pub.advogado && pub.advogado !== 'Não Identificado') advogados.add(pub.advogado);
    });

    // 2. Adicionar os nomes da lista de detecção (garante que apareçam mesmo sem pubs)
    DETECTION_DATA.companies.forEach(c => empreendimentos.add(c.name));
    DETECTION_DATA.lawyers.forEach(l => advogados.add(l.name));

    const empSelect = document.getElementById('emp-dropdown');
    const advSelect = document.getElementById('adv-dropdown');
    
    if (empSelect) {
        const sortedEmp = Array.from(empreendimentos).sort();
        let html = sortedEmp.map(e => 
            `<label class="dropdown-item"><input type="checkbox" value="${escapeHtml(e)}" onchange="updateMultiFilter('empreendimento')"> ${escapeHtml(e)}</label>`
        ).join('');
        
        // Sempre oferecer "Não Identificado" no final
        html += `<hr style="margin:8px 0; border:0; border-top:1px dashed #cbd5e1;">`;
        html += `<label class="dropdown-item" style="color:#94a3b8;"><input type="checkbox" value="Não Identificado" onchange="updateMultiFilter('empreendimento')"> Não Identificado</label>`;
        
        empSelect.innerHTML = html;
        lucide.createIcons();
    }
    
    if (advSelect) {
        const sortedAdv = Array.from(advogados).sort();
        let advHtml = sortedAdv.map(a => 
            `<label class="dropdown-item"><input type="checkbox" value="${escapeHtml(a)}" onchange="updateMultiFilter('advogado')"> ${escapeHtml(a)}</label>`
        ).join('');

        advHtml += `<hr style="margin:8px 0; border:0; border-top:1px dashed #cbd5e1;">`;
        advHtml += `<label class="dropdown-item" style="color:#94a3b8;"><input type="checkbox" value="Não Identificado" onchange="updateMultiFilter('advogado')"> Não Identificado</label>`;
        
        advSelect.innerHTML = advHtml;
    }
}

// ─────────── AREA FILTER DROPDOWN ───────────
function setAreaFilter(area) {
    filters.area = area;
    // Update label
    const label = document.getElementById('area-btn-label');
    if (area === 'todos') label.textContent = 'Todas';
    else if (area === 'civel') label.textContent = 'Cível';
    else if (area === 'trabalhista') label.textContent = 'Trabalhista';
    // Update active state
    document.querySelectorAll('.area-opt').forEach(el => {
        el.classList.remove('active-area');
    });
    const clicked = document.querySelector(`.area-opt[onclick*="'${area}'"]`);
    if (clicked) clicked.classList.add('active-area');
    // Close dropdown
    const dd = document.getElementById('area-dropdown-container');
    if (dd) dd.classList.add('hidden');
    applyFilters();
}

function updateMultiFilter(type) {
    if (type === 'empreendimento') {
        const checkboxes = document.querySelectorAll('#emp-dropdown input:checked');
        filters.empreendimento = Array.from(checkboxes).map(c => c.value);
        const label = document.getElementById('emp-btn-label');
        label.textContent = filters.empreendimento.length ? `${filters.empreendimento.length} Selecionados` : 'Todos';
    } else if (type === 'advogado') {
        const checkboxes = document.querySelectorAll('#adv-dropdown input:checked');
        filters.advogado = Array.from(checkboxes).map(c => c.value);
        const label = document.getElementById('adv-btn-label');
        label.textContent = filters.advogado.length ? `${filters.advogado.length} Selecionados` : 'Todos';
    }
    applyFilters();
}

function updateGroupFilter(parentCb) {
    const container = parentCb.closest('.dropdown-children');
    if (!container) return;
    const children = container.querySelectorAll('input[type="checkbox"]');
    children.forEach(cb => { if (cb !== parentCb) cb.checked = parentCb.checked; });
}

function applyFilters(preservePage) {
    filteredPublications = allPublications.filter(pub => {
        // Status filter
        if (filters.status === 'lido' && pub.status !== 'LIDO') return false;
        if (filters.status === 'nao_lido' && pub.status !== 'NÃO LIDO') return false;
        if (filters.status === 'desconsiderado' && pub.status !== 'DESCONSIDERADO') return false;
        if (filters.status === 'novo' && !pub.isNew) return false;
        // Hide desconsiderado from normal views
        if (filters.status !== 'desconsiderado' && filters.status !== 'todos' && pub.status === 'DESCONSIDERADO') return false;
        if (filters.status === 'todos' && pub.status === 'DESCONSIDERADO') return false;

        // Area filter
        if (filters.area === 'civel' && pub.area !== 'civel') return false;
        if (filters.area === 'trabalhista' && pub.area !== 'trabalhista') return false;

        // Arrays filter (multiselect)
        if (filters.empreendimento.length > 0 && !filters.empreendimento.includes(pub.empreendimento)) return false;
        if (filters.advogado.length > 0 && !filters.advogado.includes(pub.advogado)) return false;

        // CNJ search filter
        if (filters.cnjSearch && filters.cnjSearch.length > 0) {
            if (!pub.cnj.includes(filters.cnjSearch)) return false;
        }

        // Date range filter
        if (filters.dateStart || filters.dateEnd) {
            const pubDate = parseDate(pub.dataDisponibilizacao);
            if (!pubDate) return false;

            if (filters.dateStart && pubDate < filters.dateStart) return false;
            if (filters.dateEnd) {
                const endOfDay = new Date(filters.dateEnd);
                endOfDay.setHours(23, 59, 59, 999);
                if (pubDate > endOfDay) return false;
            }
        }

        return true;
    });

    // Se as publicações não retornaram nada de texto (isso evita que fiquem todas agrupadas), vamos checar status 'lido' count
    updateCounters();

    // Aplicar ordenação dinâmica
    applySorting();

    // Só resetar para a página 1 quando os filtros mudam (NÃO quando marca lido/não lido)
    if (!preservePage) {
        currentPage = 1;
    } else {
        // Garantir que a página atual ainda é válida após a filtragem
        const isAll = perPage === 'all';
        const totalPages = isAll ? 1 : Math.ceil(filteredPublications.length / perPage);
        if (currentPage > totalPages && totalPages > 0) currentPage = totalPages;
    }
    renderTable();
}

// ─────────── SORTING ───────────
function applySorting() {
    filteredPublications.sort((a, b) => {
        let cmp = 0;
        if (sortColumn === 'data') {
            const da = parseDate(a.dataDisponibilizacao);
            const db = parseDate(b.dataDisponibilizacao);
            if (!da || !db) return 0;
            cmp = da - db;
        } else if (sortColumn === 'status') {
            // NÃO LIDO vem antes de LIDO em asc
            const sa = a.status === 'LIDO' ? 1 : 0;
            const sb = b.status === 'LIDO' ? 1 : 0;
            cmp = sa - sb;
        }
        return sortDirection === 'asc' ? cmp : -cmp;
    });
}

function sortBy(column) {
    if (sortColumn === column) {
        // Mesma coluna — inverte direção
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        sortColumn = column;
        sortDirection = column === 'data' ? 'desc' : 'asc'; // data começa desc, status começa asc
    }
    updateSortArrows();
    applySorting();
    renderTable();
}

function updateSortArrows() {
    const arrowData = document.getElementById('sort-arrow-data');
    const arrowStatus = document.getElementById('sort-arrow-status');
    
    arrowData.textContent = sortColumn === 'data' ? (sortDirection === 'asc' ? '▲' : '▼') : '';
    arrowStatus.textContent = sortColumn === 'status' ? (sortDirection === 'asc' ? '▲' : '▼') : '';
    
    document.getElementById('th-data').classList.toggle('th-sort-active', sortColumn === 'data');
    document.getElementById('th-status').classList.toggle('th-sort-active', sortColumn === 'status');
}

// ─────────── CARD CLICK FILTER ───────────
function filterByCard(statusFilter) {
    filters.status = statusFilter;
    // Sincronizar botões de filtro na barra
    document.querySelectorAll('#status-filter .fbtn').forEach(b => {
        b.classList.toggle('active', b.dataset.status === statusFilter);
    });
    // Highlight card ativo
    document.querySelectorAll('.card-clickable').forEach(c => c.classList.remove('card-active'));
    if (statusFilter === 'todos') document.getElementById('card-total').classList.add('card-active');
    else if (statusFilter === 'nao_lido') document.getElementById('card-unread').classList.add('card-active');
    else if (statusFilter === 'lido') document.getElementById('card-read').classList.add('card-active');
    applyFilters();
}

function parseDate(dateStr) {
    if (!dateStr) return null;
    // Format: DD/MM/YYYY
    const parts = dateStr.split('/');
    if (parts.length !== 3) return null;
    return new Date(parts[2], parts[1] - 1, parts[0]);
}

function formatDate(dateStr) {
    if (!dateStr) return '–';
    return dateStr; // Already in DD/MM/YYYY format
}

// ─────────── RENDER TABLE ───────────
function renderTable() {
    const tbody = document.getElementById('pub-table-body');
    const total = filteredPublications.length;

    if (total === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" class="empty-state">
                    <div class="empty-content">
                        <i data-lucide="inbox" style="width:40px;height:40px;"></i>
                        <p>Nenhuma publicação encontrada com os filtros atuais.</p>
                    </div>
                </td>
            </tr>`;
        updatePagination(0, 0, 0);
        lucide.createIcons();
        return;
    }

    const isAll = perPage === 'all';
    const itemsPerPage = isAll ? total : perPage;
    const totalPages = isAll ? 1 : Math.ceil(total / itemsPerPage);
    const startIdx = isAll ? 0 : (currentPage - 1) * itemsPerPage;
    const endIdx = isAll ? total : Math.min(startIdx + itemsPerPage, total);
    const pageItems = filteredPublications.slice(startIdx, endIdx);

    let html = '';
    for (const pub of pageItems) {
        const isRead = pub.status === 'LIDO';
        const isDesc = pub.status === 'DESCONSIDERADO';
        const rowClass = isRead ? 'row-read' : (isDesc ? 'row-desconsiderado' : 'row-unread');
        const isExpanded = expandedUid === pub._uid;
        const isChecked = selectedPublications.has(pub._uid);
        const newDot = pub.isNew ? '<span class="new-dot" title="Nova"></span>' : '';
        // Borda preta na linha selecionada (mesmo que a caixa esteja fechada)
        const isSelectedBorder = selectedRowUid === pub._uid;
        const expandedStyle = isSelectedBorder ? ' style="outline:2.5px solid #1e293b;outline-offset:-1px;"' : '';
        const selectedClass = isChecked ? ' selected' : '';
        html += `<tr class="pub-row ${rowClass}${selectedClass}"${expandedStyle} onclick="toggleDetailRow('${escapeHtml(pub._uid)}')">`;

        const areaChip = pub.area === 'civel'
            ? '<span class="area-chip chip-civel">Cível</span>'
            : pub.area === 'trabalhista'
                ? '<span class="area-chip chip-trab">Trabalhista</span>'
                : '<span class="area-chip" style="background:#f5f5f5;color:#999;">–</span>';

        const statusBadge = isRead
            ? '<span class="status-badge badge-read">Lida</span>'
            : isDesc
                ? '<span class="status-badge badge-desconsiderado">Ignorada</span>'
                : '<span class="status-badge badge-unread">Não Lida</span>';

        // Lida = check verdinho, Não lida = circulo vazio
        const toggleBtn = isRead
            ? `<button class="action-icon-btn btn-eye" style="color: #10b981; border-color: #10b981; background: #ecfdf5;" data-tooltip="Marcar como Não Lida" onclick="event.stopPropagation(); toggleStatus('${escapeHtml(pub._uid)}')"><i data-lucide="check-circle"></i></button>`
            : `<button class="action-icon-btn btn-eye-off" data-tooltip="Marcar como Lida" onclick="event.stopPropagation(); toggleStatus('${escapeHtml(pub._uid)}')"><i data-lucide="circle"></i></button>`;

        const chevronIcon = isExpanded ? 'chevron-up' : 'chevron-down';

        html += `
            <td class="cb-cell" onclick="event.stopPropagation();">
                <input type="checkbox" class="row-checkbox" value="${escapeHtml(pub._uid)}" 
                       ${isChecked ? 'checked' : ''} 
                       onclick="toggleRowSelection(event, '${escapeHtml(pub._uid)}')">
            </td>
            <td class="status-cell">
                <div style="display: flex; align-items: center; gap: 6px;">
                    ${newDot}${statusBadge}
                </div>
            </td>
            <td><strong>${escapeHtml(pub.dataStr)}</strong></td>
            <td>${areaChip}</td>
            <td class="cnj-cell" onclick="event.stopPropagation();" title="Clique para copiar">
                <span class="processo-number" style="cursor:pointer;" onclick="copyCNJ(event,'${escapeHtml(pub.cnj)}')">${escapeHtml(pub.cnj)}</span>
            </td>
            <td><span class="tipo-text">${escapeHtml(pub.tipoPublicacao)}</span></td>
            <td><span class="orgao-text" title="${escapeHtml(pub.orgaoExpedidor)}">${escapeHtml(pub.orgaoExpedidor)}</span></td>
            <td class="actions-cell">
                ${toggleBtn}
                <button class="action-icon-btn btn-view" data-tooltip="${isExpanded ? 'Fechar' : 'Ver Detalhes'}" onclick="event.stopPropagation(); toggleDetailRow('${escapeHtml(pub._uid)}')"><i data-lucide="${chevronIcon}"></i></button>
            </td>
        </tr>`;

        // Linha accordion (expandida inline)
        if (isExpanded) {
            html += renderAccordionRow(pub);
        }
    }

    tbody.innerHTML = html;
    updatePagination(startIdx + 1, endIdx, total);
    lucide.createIcons();
}

// ─────────── SELEÇÃO LOTE & RELATÓRIOS ───────────
function toggleSelectAll() {
    const cbAll = document.getElementById('select-all-cb');
    const isChecked = cbAll && cbAll.checked;
    
    if (isChecked) {
        const cbs = document.querySelectorAll('.row-checkbox');
        cbs.forEach(cb => {
            cb.checked = true;
            selectedPublications.add(cb.value);
            cb.closest('tr').classList.add('selected');
        });
    } else {
        const cbs = document.querySelectorAll('.row-checkbox');
        cbs.forEach(cb => {
            cb.checked = false;
            selectedPublications.delete(cb.value);
            cb.closest('tr').classList.remove('selected');
        });
    }
    updateBulkActionBar();
}

function toggleRowSelection(event, _uid) {
    event.stopPropagation();
    const cb = event.target;
    const isChecked = cb.checked;
    
    if (isChecked) {
        selectedPublications.add(_uid);
        cb.closest('tr').classList.add('selected');
    } else {
        selectedPublications.delete(_uid);
        cb.closest('tr').classList.remove('selected');
        const selectAllCb = document.getElementById('select-all-cb');
        if (selectAllCb) selectAllCb.checked = false;
    }
    updateBulkActionBar();
}

function updateBulkActionBar() {
    const bulkBar = document.getElementById('bulk-action-bar');
    const countEl = document.getElementById('bulk-count');
    if (!bulkBar || !countEl) return;
    
    if (selectedPublications.size > 0) {
        countEl.textContent = selectedPublications.size;
        bulkBar.style.display = 'flex';
        bulkBar.classList.remove('hidden');
    } else {
        bulkBar.style.display = 'none';
        bulkBar.classList.add('hidden');
        const selectAllCb = document.getElementById('select-all-cb');
        if (selectAllCb) selectAllCb.checked = false;
    }
}

function clearBulkSelection() {
    selectedPublications.clear();
    const cbs = document.querySelectorAll('.row-checkbox');
    cbs.forEach(cb => {
        cb.checked = false;
        cb.closest('tr').classList.remove('selected');
    });
    const selectAllCb = document.getElementById('select-all-cb');
    if (selectAllCb) selectAllCb.checked = false;
    updateBulkActionBar();
}

async function bulkMarkAs(isReadStatus) {
    if (selectedPublications.size === 0) return;
    
    const newStatus = isReadStatus ? 'LIDO' : 'NÃO LIDO';
    const uidArray = Array.from(selectedPublications);
    
    uidArray.forEach(_uid => {
        const pub = allPublications.find(p => p._uid === _uid);
        if (pub) {
            pub.status = newStatus;
            setStatus(pub.cnj, newStatus, pub._uid);
        }
    });
    
    clearBulkSelection();
    applyFilters();
}

function downloadReport() {
    if (filteredPublications.length === 0) {
        alert("Sem dados para exportar com os filtros atuais.");
        return;
    }

    const rows = filteredPublications.map(pub => ({
        'Data Disponibilização': pub.dataStr || '',
        'Processo (CNJ)': pub.cnj || '',
        'Status': pub.status || '',
        'Área': pub.area || '',
        'Tipo Publicação': pub.tipoPublicacao || '',
        'Órgão Expedidor': pub.orgaoExpedidor || '',
        'Empreendimento': pub.empreendimento || '',
        'Advogado': pub.advogado || '',
        'Teor Publicação': (pub.textoPublicacao || '').substring(0, 32000),
    }));

    if (typeof XLSX !== 'undefined') {
        const ws = XLSX.utils.json_to_sheet(rows);
        const colWidths = Object.keys(rows[0]).map(k => ({ wch: Math.max(k.length + 2, 18) }));
        ws['!cols'] = colWidths;
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Publicações');
        
        try {
            XLSX.writeFile(wb, `relatorio_foreigor_GP_${new Date().toISOString().slice(0,10)}.xlsx`);
        } catch (e) {
            console.error('Erro ao baixar Excel', e);
            alert("Erro ao gerar o arquivo Excel.");
        }
    } else {
        // Fallback CSV
        let csv = 'Data;Processo (CNJ);Status;Área;Órgão Expedidor;Empreendimento;Advogado;Teor\n';
        rows.forEach(r => {
            csv += Object.values(r).map(v => `"${String(v).replace(/"/g,'""')}"`).join(';') + '\n';
        });
        const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `relatorio_foreigor_GP_${new Date().toISOString().slice(0,10)}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 100);
    }
}

// ─────────── ACCORDION ROW (cascata inline — REDESENHADO) ───────────
function renderAccordionRow(pub) {
    const tribunal = pub.tribunal;
    const isRead = pub.status === 'LIDO';
    const toggleBtnLabel = isRead ? 'Marcar como Não Lida' : 'Marcar como Lida';
    const toggleBtnClass = isRead ? 'btn-mark-unread' : 'btn-mark-read';
    const toggleBtnIcon = isRead ? 'circle' : 'check-circle';

    // Se o tribunal não foi mapeado, fallback dinâmico
    const nomeTribunal = tribunal ? tribunal.nome : 'Tribunal não mapeado';
    const link = (tribunal && tribunal.url) ? tribunal.url : '#';
    const teor = pub.textoPublicacao || 'Conteúdo não disponível.';

    return `
    <tr class="detail-row" data-uid="${escapeHtml(pub._uid)}">
        <td colspan="7">
            <div class="cascade-container">
                <!-- Left Sidebar: Metadata -->
                <div class="cascade-sidebar">
                    <div class="cascade-cnj-header">
                        <label><i data-lucide="scale"></i> Processo Judicial</label>
                        <p>${escapeHtml(pub.cnj)}</p>
                    </div>

                    <div class="cascade-meta-list">
                        <div class="meta-row">
                            <label><i data-lucide="hash"></i> ID Publicação</label>
                            <span>${escapeHtml(pub.idPublicacao || 'N/A')}</span>
                        </div>
                        <div class="meta-row">
                            <label><i data-lucide="tag"></i> Tipo</label>
                            <span>${escapeHtml(pub.tipoPublicacao || 'N/A')}</span>
                        </div>
                        
                        <div class="meta-row" style="border-top: 1px solid var(--border-light); padding-top: 10px;">
                            <label><i data-lucide="briefcase"></i> Empreendimento</label>
                            <span style="font-weight: 700; color: var(--primary);">${escapeHtml(pub.empreendimento)}</span>
                        </div>
                        <div class="meta-row">
                            <label><i data-lucide="user"></i> Advogado</label>
                            <span>${escapeHtml(pub.advogado)}</span>
                        </div>
                        <div class="meta-row">
                            <label><i data-lucide="map-pin"></i> Órgão Expedidor</label>
                            <span>${escapeHtml(pub.orgaoExpedidor)}</span>
                        </div>
                    </div>

                    <div style="margin-top: auto; padding-top: 12px;">
                        <a href="${link}" target="_blank" class="tribunal-link-btn">
                            <div class="btn-shine"></div>
                            <i data-lucide="external-link"></i>
                            <span>ACESSAR PROCESSO NO ${pub.tribunal?.sistema || 'TRIBUNAL'}</span>
                        </a>
                    </div>
                </div>

                <!-- Right Main: Teor da Publicação -->
                <div class="cascade-main">
                    <div class="cascade-teor">
                        <label><i data-lucide="file-text"></i> Teor da Publicação</label>
                        <div class="teor-box">${escapeHtml(teor)}</div>
                    </div>
                    
                    <div class="cascade-action-row" style="margin-top: 15px; display: flex; gap: 10px; align-items: stretch; flex-wrap: wrap;">
                        <button class="action-btn ${toggleBtnClass}" onclick="event.stopPropagation(); toggleStatusInline('${escapeHtml(pub._uid)}')">
                            <i data-lucide="${toggleBtnIcon}"></i> ${toggleBtnLabel}
                        </button>
                        <button class="action-btn btn-desconsiderar" style="background-color: #e2e8f0; color: #475569;" onclick="event.stopPropagation(); markStatusInline('${escapeHtml(pub._uid)}', 'DESCONSIDERADO')">
                            <i data-lucide="trash-2"></i> Desconsiderar
                        </button>
                    </div>

                    <!-- CAIXA DE DELEGAÇÃO -->
                    <div class="delegation-box">
                        <label style="font-size: 0.75rem; font-weight: 700; color: var(--primary); margin-bottom: 8px; display: inline-flex; align-items: center; gap: 6px;"><i data-lucide="send" style="width:14px;height:14px;"></i> DELEGAR PARA COLEGA (E-MAIL)</label>
                        ${(function() {
                            const savedInfo = pub.observacoes && pub.observacoes.startsWith('Delegado para') ? pub.observacoes : (delegationLog[pub.cnj] || '');
                            if (!savedInfo) return '';
                            return `<div style="padding: 8px; background: #f0fdf4; border: 1px solid #16a34a; color: #166534; font-size: 0.8rem; border-radius: 4px; margin-bottom: 10px; display: flex; align-items: start; gap: 6px;"><i data-lucide="check-circle" style="width: 14px; height: 14px; margin-top: 2px;"></i> <span><strong>Já salvo:</strong> ${escapeHtml(savedInfo)}</span></div>`;
                        })()}
                        <div class="del-checkbox-list" style="display: flex; flex-direction: column; gap: 4px; margin-bottom: 10px; max-height: 200px; overflow-y: auto; padding: 8px; border: 1px solid #d1d5db; border-radius: 6px; background: #fff;">
                                                        <label class="del-email-opt"><input type="checkbox" value="processos.juridico@gramadoparks.com" data-name="Processos"> Processos</label>
                            <label class="del-email-opt"><input type="checkbox" value="renan.siqueira@gramadoparks.com" data-name="Renan"> Renan</label>
                            <label class="del-email-opt"><input type="checkbox" value="giovani.silva@gramadoparks.com" data-name="Giovani"> Giovani</label>
                            <label class="del-email-opt"><input type="checkbox" value="bruna.santos@gramadoparks.com" data-name="Bruna"> Bruna</label>
                            <label class="del-email-opt"><input type="checkbox" value="fernanda.meneses@gramadoparks.com" data-name="Fernanda"> Fernanda</label>
                            <label class="del-email-opt"><input type="checkbox" value="igor.weimer@gramadoparks.com" data-name="Igor"> Igor</label>
                            <label class="del-email-opt"><input type="checkbox" value="franciele.ribeiro@gramadoparks.com" data-name="Franciele"> Franciele</label>
                            <label class="del-email-opt"><input type="checkbox" value="rafaela.cardoso@gramadoparks.com" data-name="Rafaela"> Rafaela</label>
                            <label class="del-email-opt"><input type="checkbox" value="luciano.ferreira@gramadoparks.com" data-name="Luciano"> Luciano</label>
                        </div>
                        <textarea class="del-obs-text" placeholder="Ex: Para recorrer com urgência... Você também pode dar Ctrl+V aqui para colar prints." onpaste="handlePaste(event, '${escapeHtml(pub._uid)}')" style="width: 100%; height: 60px; padding: 8px; border-radius: 6px; border: 1px solid #d1d5db; resize: none; outline: none; font-size: 0.85rem; font-family: 'Inter', sans-serif;"></textarea>
                        
                        <!-- Upload de Anexos -->
                        <div class="del-upload-area" id="del-upload-${escapeHtml(pub._uid).replace(/[^a-zA-Z0-9]/g,'')}">
                            <div class="del-upload-header">
                                <span class="del-upload-label"><i data-lucide="paperclip"></i> Anexos (Prints / PDFs)</span>
                                <button class="del-upload-btn" onclick="event.stopPropagation(); document.getElementById('file-input-${escapeHtml(pub._uid).replace(/[^a-zA-Z0-9]/g,'')}').click()">
                                    <i data-lucide="upload"></i> Escolher Arquivo
                                </button>
                            </div>
                            <input type="file" id="file-input-${escapeHtml(pub._uid).replace(/[^a-zA-Z0-9]/g,'')}" 
                                   accept="image/*,.pdf" multiple 
                                   style="display:none;" 
                                   onchange="handleFileUpload(this, '${escapeHtml(pub._uid)}')">
                            <div class="del-file-preview" id="file-preview-${escapeHtml(pub._uid).replace(/[^a-zA-Z0-9]/g,'')}"></div>
                        </div>

                        <div style="display: flex; gap: 10px; margin-top: 10px;">
                            <button class="action-btn" style="background-color: var(--primary); color: white; border: none;" onclick="event.stopPropagation(); addToDelegationQueue(this, '${escapeHtml(pub._uid)}')">
                                <i data-lucide="plus-circle"></i> Adicionar à Fila de Disparos
                            </button>
                        </div>
                    </div>

                </div>
            </div>
        </td>
    </tr>`;
}

function toggleDetailRow(_uid) {
    selectedRowUid = _uid; // Define esta linha como a selecionada (borda)
    if (expandedUid === _uid) {
        expandedUid = null; // Fechar se já está aberto
    } else {
        expandedUid = _uid; // Abrir este
    }
    renderTable();
}

function updatePagination(from, to, total) {
    const info = total > 0 ? `${from} – ${to} de ${total}` : '–';

    document.getElementById('page-info').textContent = info;
    document.getElementById('page-info-bottom').textContent = info;

    const isAll = perPage === 'all';
    const totalPages = isAll ? 1 : Math.ceil(total / perPage);

    const prevDisabled = currentPage <= 1 || isAll;
    const nextDisabled = currentPage >= totalPages || isAll;

    document.getElementById('prev-page').disabled = prevDisabled;
    document.getElementById('next-page').disabled = nextDisabled;
    document.getElementById('prev-page-bottom').disabled = prevDisabled;
    document.getElementById('next-page-bottom').disabled = nextDisabled;

    // Update table count
    document.getElementById('table-count').textContent = `(${total} resultado${total !== 1 ? 's' : ''})`;
}

// ─────────── UPDATE COUNTERS ───────────
function updateCounters() {
    const total = filteredPublications.length;
    const unread = filteredPublications.filter(p => p.status !== 'LIDO' && p.status !== 'DESCONSIDERADO').length;
    const read = filteredPublications.filter(p => p.status === 'LIDO').length;
    // "Novas" é sempre calculado sobre TODAS as publicações, não só as filtradas
    const newCount = allPublications.filter(p => p.isNew).length;

    animateCounter('count-total', total);
    animateCounter('count-unread', unread);
    animateCounter('count-read', read);

    const countNewEl = document.getElementById('count-new');
    if (countNewEl) animateCounter('count-new', newCount);
}

function animateCounter(elementId, targetValue) {
    const el = document.getElementById(elementId);
    const current = parseInt(el.textContent) || 0;
    if (current === targetValue) return;

    el.classList.add('number-animated');
    el.textContent = targetValue;
    setTimeout(() => el.classList.remove('number-animated'), 400);
}

// ─────────── SYNC STATUS ───────────
function setSyncStatus(type, text) {
    const el = document.getElementById('sync-status');
    el.className = 'sync-status' + (type === 'error' ? ' error' : '');
    const iconName = type === 'ok' ? 'wifi' : type === 'loading' ? 'loader' : 'wifi-off';
    el.innerHTML = `<i data-lucide="${iconName}" class="sync-icon"></i><span>${text}</span>`;
    lucide.createIcons();
}

// ─────────── TRATAMENTO DE STATUS NO DOM ───────────






// ─────────── STATUS ACTIONS (TABELA & INLINE) ───────────

// Copia CNJ para a área de transferência
function copyCNJ(event, cnj) {
    event.stopPropagation();
    navigator.clipboard.writeText(cnj).then(() => {
        showToast(`CNJ copiado: ${cnj}`, 'success');
    }).catch(() => {
        // Fallback para navegadores sem clipboard API
        const ta = document.createElement('textarea');
        ta.value = cnj; ta.style.position='fixed'; ta.style.opacity='0';
        document.body.appendChild(ta); ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        showToast(`CNJ copiado!`, 'success');
    });
}

/**
 * toggleStatus - Alterna o status entre LIDO e NÃO LIDO (com confirmação ao marcar como Lida)
 */
function toggleStatus(_uid) {
    const pub = allPublications.find(p => p._uid === _uid);
    if (!pub) return;
    const next = pub.status === 'LIDO' ? 'NÃO LIDO' : 'LIDO';
    if (next === 'LIDO') {
        askConfirmRead('Tem certeza que quer marcar como <strong>Lida</strong>?', () => {
            pub.status = next;
            setStatus(pub.cnj, next, pub._uid, pub.idPublicacao);
            selectedRowUid = _uid;
            applyFilters(true);
        });
    } else {
        pub.status = next;
        setStatus(pub.cnj, next, pub._uid, pub.idPublicacao);
        selectedRowUid = _uid;
        applyFilters(true);
    }
}

/**
 * toggleStatusInline - Chamado para alternar status e fechar a linha
 */
function toggleStatusInline(_uid) {
    const pub = allPublications.find(p => p._uid === _uid);
    if (!pub) return;
    const next = pub.status === 'LIDO' ? 'NÃO LIDO' : 'LIDO';
    if (next === 'LIDO') {
        askConfirmRead('Tem certeza que quer marcar como <strong>Lida</strong>?', () => {
            pub.status = next;
            setStatus(pub.cnj, next, pub._uid, pub.idPublicacao);
            expandedUid = null;
            selectedRowUid = _uid;
            applyFilters(true);
        });
    } else {
        pub.status = next;
        setStatus(pub.cnj, next, pub._uid, pub.idPublicacao);
        expandedUid = null;
        selectedRowUid = _uid;
        applyFilters(true);
    }
}

/**
 * markStatusInline - Atualiza status específico e fecha linha
 */
function markStatusInline(_uid, newStatus) {
    const pub = allPublications.find(p => p._uid === _uid);
    if (!pub) return;
    if (newStatus === 'LIDO') {
        askConfirmRead('Tem certeza que quer marcar como <strong>Lida</strong>?', () => {
            pub.status = newStatus;
            setStatus(pub.cnj, newStatus, pub._uid, pub.idPublicacao);
            expandedUid = null;
            selectedRowUid = _uid;
            applyFilters(true);
        });
        return;
    }
    pub.status = newStatus;
    setStatus(pub.cnj, newStatus, pub._uid, pub.idPublicacao);
    expandedUid = null;
    selectedRowUid = _uid;
    applyFilters(true);
}



// ─────────── DETAIL MODAL (removido — agora usa accordion inline) ───────────
// As funções openDetail, renderDetailStatus e closeDetail foram
// substituídas por toggleDetailRow() e renderAccordionRow() acima.

// ─────────── DATEPICKER ───────────
function openDatepicker(target) {
    datepickerTarget = target;
    const overlay = document.getElementById('datepicker-overlay');
    overlay.classList.remove('hidden');
    renderDatepicker();
}

function closeDatepicker() {
    document.getElementById('datepicker-overlay').classList.add('hidden');
    datepickerTarget = null;
}

function renderDatepicker() {
    const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
        'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

    document.getElementById('dp-title').textContent = `${meses[datepickerMonth]} ${datepickerYear}`;

    const firstDay = new Date(datepickerYear, datepickerMonth, 1).getDay();
    const daysInMonth = new Date(datepickerYear, datepickerMonth + 1, 0).getDate();
    const today = new Date();

    let html = '';

    // Empty cells before first day
    for (let i = 0; i < firstDay; i++) {
        html += '<button class="dp-day dp-empty" disabled></button>';
    }

    for (let d = 1; d <= daysInMonth; d++) {
        const date = new Date(datepickerYear, datepickerMonth, d);
        const isToday = date.toDateString() === today.toDateString();

        // Check if selected
        const selectedDate = datepickerTarget === 'start' ? filters.dateStart : filters.dateEnd;
        const isSelected = selectedDate && date.toDateString() === selectedDate.toDateString();

        let cls = 'dp-day';
        if (isToday) cls += ' dp-today';
        if (isSelected) cls += ' dp-selected';

        html += `<button class="${cls}" onclick="selectDate(${d})">${d}</button>`;
    }

    document.getElementById('dp-grid').innerHTML = html;
}

function selectDate(day) {
    const date = new Date(datepickerYear, datepickerMonth, day);

    if (datepickerTarget === 'start') {
        filters.dateStart = date;
        document.getElementById('date-start-label').textContent = formatDateShort(date);
        document.getElementById('date-start-box').classList.add('has-date');
    } else {
        filters.dateEnd = date;
        document.getElementById('date-end-label').textContent = formatDateShort(date);
        document.getElementById('date-end-box').classList.add('has-date');
    }

    closeDatepicker();
    applyFilters();
}

function formatDateShort(date) {
    const dd = String(date.getDate()).padStart(2, '0');
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const yy = String(date.getFullYear());
    return `${dd}/${mm}/${yy}`;
}

function clearDates() {
    filters.dateStart = null;
    filters.dateEnd = null;
    document.getElementById('date-start-label').textContent = 'Data Inicial';
    document.getElementById('date-end-label').textContent = 'Data Final';
    document.getElementById('date-start-box').classList.remove('has-date');
    document.getElementById('date-end-box').classList.remove('has-date');
    applyFilters();
}

// ─────────── EVENT LISTENERS ───────────
function setupEventListeners() {
    // Status filter buttons
    document.querySelectorAll('#status-filter .fbtn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#status-filter .fbtn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            filters.status = btn.dataset.status;
            applyFilters();
        });
    });

    // Area filter is now a dropdown — handled by setAreaFilter()

    // Per page select
    document.getElementById('per-page-select').addEventListener('change', (e) => {
        const val = e.target.value;
        perPage = val === 'all' ? 'all' : parseInt(val);
        currentPage = 1;
        renderTable();
    });

    // Date pickers
    document.getElementById('date-start-box').addEventListener('click', () => openDatepicker('start'));
    document.getElementById('date-end-box').addEventListener('click', () => openDatepicker('end'));
    document.getElementById('clear-dates-btn').addEventListener('click', clearDates);

    // Datepicker navigation
    document.getElementById('dp-prev').addEventListener('click', () => {
        datepickerMonth--;
        if (datepickerMonth < 0) { datepickerMonth = 11; datepickerYear--; }
        renderDatepicker();
    });
    document.getElementById('dp-next').addEventListener('click', () => {
        datepickerMonth++;
        if (datepickerMonth > 11) { datepickerMonth = 0; datepickerYear++; }
        renderDatepicker();
    });
    document.getElementById('dp-cancel').addEventListener('click', closeDatepicker);
    document.getElementById('dp-today').addEventListener('click', () => {
        const today = new Date();
        datepickerMonth = today.getMonth();
        datepickerYear = today.getFullYear();
        selectDate(today.getDate());
    });

    // Close datepicker on overlay click
    document.getElementById('datepicker-overlay').addEventListener('click', (e) => {
        if (e.target.id === 'datepicker-overlay') closeDatepicker();
    });

    // Pagination
    document.getElementById('prev-page').addEventListener('click', () => { currentPage--; renderTable(); scrollToTable(); });
    document.getElementById('next-page').addEventListener('click', () => { currentPage++; renderTable(); scrollToTable(); });
    document.getElementById('prev-page-bottom').addEventListener('click', () => { currentPage--; renderTable(); scrollToTable(); });
    document.getElementById('next-page-bottom').addEventListener('click', () => { currentPage++; renderTable(); scrollToTable(); });

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            if (expandedUid) { expandedUid = null; renderTable(); }
            closeDatepicker();
        }
    });
}

function scrollToTable() {
    document.getElementById('table-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─────────── UTILITIES ───────────
function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str).replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function toggleDropdown(id) {
    const el = document.getElementById(id);
    if (!el) return;
    document.querySelectorAll('.dropdown-content').forEach(d => {
        if (d.id !== id) d.classList.add('hidden');
    });
    el.classList.toggle('hidden');
}
document.addEventListener('click', (e) => {
    if (!e.target.closest('.dropdown-wrapper')) {
        document.querySelectorAll('.dropdown-content').forEach(d => d.classList.add('hidden'));
    }
});

// Inicializar

// ─────────── MAILBOX & DELEGATION SYSTEM ───────────
let delegationQueue = {};
try {
    delegationQueue = JSON.parse(localStorage.getItem('foreigor_delegation_queue')) || {};
} catch(e) { console.error('Resetting queue', e); }

// ─── Sincronização da Fila de E-mails com a planilha ───
// A fila é salva em uma aba separada "FilaEmail" no Google Sheets.
// Isso garante que a Rachel veja a mesma fila em qualquer dispositivo.

async function loadEmailQueueFromSheets() {
    if (!CONFIG.APPS_SCRIPT_URL) return;
    try {
        const res = await fetch(CONFIG.APPS_SCRIPT_URL + `?action=getEmailQueue&t=${Date.now()}`);
        const data = await res.json();
        if (data && data.queue) {
            // Mesclar com fila local (arquivos base64 ficam apenas localmente)
            const sheetsQueue = data.queue;
            for (const key in sheetsQueue) {
                if (!delegationQueue[key]) {
                    delegationQueue[key] = sheetsQueue[key];
                } else {
                    // Mesclar itens: planilha é a verdade para metadados, local preserva arquivos
                    const localFiles = {};
                    delegationQueue[key].items.forEach(it => { localFiles[it.cnj] = it.files; });
                    delegationQueue[key] = sheetsQueue[key];
                    delegationQueue[key].items.forEach(it => {
                        if (localFiles[it.cnj]) it.files = localFiles[it.cnj];
                    });
                }
            }
            // Remover itens que a planilha não tem mais (foram limpos por outro usuário)
            for (const key in delegationQueue) {
                if (!sheetsQueue[key]) delete delegationQueue[key];
            }
            localStorage.setItem('foreigor_delegation_queue', JSON.stringify(
                JSON.parse(JSON.stringify(delegationQueue, (k, v) => k === 'files' ? undefined : v))
            ));
            updateDelegationBadge();
        }
    } catch (e) {
        console.warn('[Queue] Não foi possível carregar fila de e-mails da planilha:', e.message);
    }
}

async function saveEmailQueueToSheets() {
    if (!CONFIG.APPS_SCRIPT_URL) return;
    try {
        // Serializar sem os arquivos base64 (ficam apenas localmente)
        const queueClone = JSON.parse(JSON.stringify(delegationQueue, (k, v) => k === 'files' ? undefined : v));
        await fetch(CONFIG.APPS_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({ action: 'saveEmailQueue', queue: queueClone })
        });
        console.log('[Queue] Fila de e-mails sincronizada com a planilha.');
    } catch (e) {
        console.warn('[Queue] Não foi possível salvar fila na planilha:', e.message);
    }
}

// init() removido — inicialização via DOMContentLoaded

function saveDelegationQueue() {
    try {
        const queueClone = JSON.parse(JSON.stringify(delegationQueue));
        for (let key in queueClone) {
            queueClone[key].items.forEach(i => delete i.files); // Remove files (base64) so it doesn't break localStorage
        }
        localStorage.setItem('foreigor_delegation_queue', JSON.stringify(queueClone));
    } catch (e) {
        console.error('Erro ao salvar fila no localStorage:', e);
    }
    // Sincronizar com a planilha (fonte da verdade)
    saveEmailQueueToSheets();
    updateDelegationBadge();
}

// Handler para o evento de colar (Ctrl+V) na textarea de observações
function handlePaste(e, cnj) {
    if (e.clipboardData && e.clipboardData.items) {
        const items = Array.from(e.clipboardData.items);
        const filesToProcess = [];
        items.forEach(item => {
            if (item.kind === 'file' && (item.type.indexOf('image/') !== -1 || item.type === 'application/pdf')) {
                filesToProcess.push(item.getAsFile());
            }
        });
        if (filesToProcess.length > 0) {
            // Se tem arquivos (prints) vamos processar!
            handleFileUpload({ files: filesToProcess }, cnj);
        }
    }
}

// Armazenamento temporário de arquivos por CNJ (não vai no localStorage)
let tempFiles = {};

function addToDelegationQueue(btn, _uid) {
    const pub = allPublications.find(p => p._uid === _uid);
    if (!pub) return;

    // Expand accordion container to find elements correctly
    const container = btn.closest('.detail-row');
    const checkboxes = container.querySelectorAll('.del-email-opt input:checked');
    const obs = container.querySelector('.del-obs-text').value.trim();
    
    const files = tempFiles[_uid] || [];

    if (checkboxes.length === 0) {
        alert('Selecione ao menos um colega para delegar.');
        return;
    }

    const cbArray = Array.from(checkboxes);
    const addedNames = cbArray.map(cb => cb.dataset.name);

    // Se múltiplos destinatários, agrupar em UM único e-mail
    if (cbArray.length > 1) {
        const emails = cbArray.map(cb => cb.value);
        const groupKey = emails.sort().join(',');
        const groupName = addedNames.join(', ');

        if (!delegationQueue[groupKey]) {
            delegationQueue[groupKey] = { name: groupName, isGroup: true, items: [] };
        }

        const existing = delegationQueue[groupKey].items.findIndex(item => item._uid === _uid);
        if (existing !== -1) {
            delegationQueue[groupKey].items[existing].obs = obs;
            delegationQueue[groupKey].items[existing].files = files;
        } else {
            delegationQueue[groupKey].items.push({
                cnj: pub.cnj,
                _uid: _uid,
                orgao: pub.orgaoExpedidor,
                tipo: pub.tipoPublicacao,
                dataStr: pub.dataStr || pub.dataDisponibilizacao || '',
                obs: obs,
                files: files
            });
        }
    } else {
        // Destinatário único — comportamento original
        const email = cbArray[0].value;
        const name = cbArray[0].dataset.name;

        if (!delegationQueue[email]) {
            delegationQueue[email] = { name: name, items: [] };
        }

        const existing = delegationQueue[email].items.findIndex(item => item._uid === _uid);
        if (existing !== -1) {
            delegationQueue[email].items[existing].obs = obs;
            delegationQueue[email].items[existing].files = files;
        } else {
            delegationQueue[email].items.push({ 
                cnj: pub.cnj, 
                _uid: _uid,
                orgao: pub.orgaoExpedidor, 
                tipo: pub.tipoPublicacao,
                dataStr: pub.dataStr || pub.dataDisponibilizacao || '',
                obs: obs,
                files: files
            });
        }
    }

    // NÃO marcar como lida automaticamente — usuário decide quando marcar
    // Apenas salvar o log de quem foi delegado
    saveDelegationInfo(pub.cnj, addedNames, obs);

    saveDelegationQueue();
    renderTable();
    updateDelegationBadge();
    showToast(`Adicionado à fila de ${addedNames.join(', ')}`, 'success');
}

function updateDelegationBadge() {
    const badge = document.getElementById('mailbox-badge');
    if (!badge) return;
    let total = 0;
    Object.values(delegationQueue).forEach(v => total += v.items.length);
    badge.textContent = total;
    badge.style.display = total > 0 ? 'flex' : 'none';
}

// ─────────── FILE UPLOAD HANDLING ───────────
function handleFileUpload(input, _uid) {
    const files = Array.from(input.files);
    if (!files.length) return;
    
    if (!tempFiles[_uid]) tempFiles[_uid] = [];
    
    files.forEach(file => {
        const reader = new FileReader();
        reader.onload = function(e) {
            const fileData = {
                name: file.name,
                type: file.type,
                size: file.size,
                data: e.target.result // base64 data URL
            };
            tempFiles[_uid].push(fileData);
            renderFilePreview(_uid);
        };
        reader.readAsDataURL(file);
    });
}

function renderFilePreview(_uid) {
    // Usar o UID para ID HTML (sanitizado)
    const safeId = _uid.replace(/[^a-zA-Z0-9]/g, '');
    const container = document.getElementById('file-preview-' + safeId);
    if (!container) return;
    
    const files = tempFiles[_uid] || [];
    if (files.length === 0) {
        container.innerHTML = '';
        return;
    }
    
    container.innerHTML = files.map((f, idx) => {
        const isImage = f.type.startsWith('image/');
        const isPdf = f.type === 'application/pdf';
        const chipClass = isImage ? 'file-image' : isPdf ? 'file-pdf' : '';
        const icon = isImage ? '🖼️' : isPdf ? '📄' : '📎';
        const thumb = isImage ? `<img src="${f.data}" class="del-file-thumb" alt="${escapeHtml(f.name)}">` : '';
        
        return `<div class="del-file-chip ${chipClass}">
            ${thumb}
            ${icon} ${escapeHtml(f.name.length > 20 ? f.name.substring(0, 17) + '...' : f.name)}
            <button class="del-file-remove" onclick="event.stopPropagation(); removeFile('${escapeHtml(_uid)}', ${idx})" title="Remover">✕</button>
        </div>`;
    }).join('');
}

function removeFile(_uid, idx) {
    if (tempFiles[_uid]) {
        tempFiles[_uid].splice(idx, 1);
        renderFilePreview(_uid);
    }
}

/**
 * Adiciona a publicação selecionada à fila de disparos (Mailbox)
 * baseada nos colegas selecionados na cascata.
 */
function addToMailboxInline(_uid) {
    console.log(`[Mailbox] Iniciando adição para UID: ${_uid}`);
    
    const pub = allPublications.find(p => p._uid === _uid);
    if (!pub) {
        console.error(`[Mailbox] Publicação não encontrada para o UID: ${_uid}`);
        return;
    }

    // Buscar a linha específica pelo data-uid que acabamos de adicionar
    const row = document.querySelector(`.detail-row[data-uid="${_uid}"]`);
    if (!row) {
        console.error(`[Mailbox] Linha de detalhe não encontrada para o UID: ${_uid}`);
        alert('Erro interno: a visualização desta publicação não foi localizada.');
        return;
    }
    
    // Selecionar apenas os checkboxes desta linha
    const selectedCheckboxes = row.querySelectorAll('input[type="checkbox"]:checked');
    const selected = Array.from(selectedCheckboxes);

    if (selected.length === 0) {
        alert('Por favor, selecione ao menos um colega na lista acima.');
        return;
    }

    const obsField = document.getElementById(`obs-${_uid}`);
    const obs = obsField ? obsField.value.trim() : '';
    const filesAttached = tempFiles[_uid] || [];

    console.log(`[Mailbox] Delegando para ${selected.length} colegas.`);

    // Se múltiplos destinatários selecionados, agrupar em um único e-mail
    if (selected.length > 1) {
        const emails = selected.map(cb => cb.getAttribute('data-email'));
        const nomes = selected.map(cb => cb.getAttribute('data-nome'));
        const groupKey = emails.sort().join(',');
        const groupName = nomes.join(', ');
        
        if (!delegationQueue[groupKey]) {
            delegationQueue[groupKey] = { name: groupName, isGroup: true, items: [] };
        }
        
        const alreadyExists = delegationQueue[groupKey].items.some(it => it._uid === _uid);
        if (!alreadyExists) {
            delegationQueue[groupKey].items.push({
                cnj: pub.cnj,
                _uid: _uid,
                orgao: pub.orgaoExpedidor,
                tipo: pub.tipoPublicacao,
                dataStr: pub.dataStr || pub.dataDisponibilizacao || '',
                obs: obs,
                files: [...filesAttached]
            });
        }
        
        saveDelegationInfo(pub.cnj, nomes, obs);
    } else {
        // Destinatário único — comportamento original
        const cb = selected[0];
        const email = cb.getAttribute('data-email');
        const nome = cb.getAttribute('data-nome');
        
        if (!delegationQueue[email]) {
            delegationQueue[email] = { name: nome, items: [] };
        }
        
        const alreadyExists = delegationQueue[email].items.some(it => it._uid === _uid);
        if (!alreadyExists) {
            delegationQueue[email].items.push({
                cnj: pub.cnj,
                _uid: _uid,
                orgao: pub.orgaoExpedidor,
                tipo: pub.tipoPublicacao,
                dataStr: pub.dataStr || pub.dataDisponibilizacao || '',
                obs: obs,
                files: [...filesAttached]
            });
        }
        
        saveDelegationInfo(pub.cnj, [nome], obs);
    }

    // Limpar arquivos temporários desta publicação
    delete tempFiles[_uid];
    
    // Persistir
    saveDelegationQueue();
    
    // Feedback
    alert('Publicação adicionada à fila de disparos com sucesso!');
    updateDelegationBadge();
    
    // Fechar a linha
    expandedUid = null;
    renderTable();
}

// ─────────── REDESIGNED MAILBOX ───────────
function renderMailbox() {
    document.getElementById('mailbox-modal').classList.remove('hidden');
    const container = document.getElementById('mailbox-container');
    
    // Count total
    let totalItems = 0;
    Object.values(delegationQueue).forEach(v => totalItems += v.items.length);
    
    const countHeader = document.getElementById('mailbox-count-header');
    if (countHeader) countHeader.textContent = `${totalItems} itens`;
    
    // Hide send-all if empty
    const sendAllBtn = document.getElementById('send-all-btn');
    if (sendAllBtn) sendAllBtn.style.display = totalItems > 0 ? 'flex' : 'none';
    
    if (totalItems === 0) {
        container.innerHTML = `
        <div class="mbox-empty">
            <i data-lucide="inbox"></i>
            <p>Nenhuma publicação na fila de disparo.</p>
            <span style="font-size: 0.82rem; margin-top: 8px;">Adicione publicações expandindo uma linha na tabela e clicando em "Adicionar à Fila".</span>
        </div>`;
        lucide.createIcons();
        return;
    }
    
    let html = '';
    
    for (const [queueKey, data] of Object.entries(delegationQueue)) {
        if (!data.items || data.items.length === 0) continue;
        
        const actualEmail = getActualEmail(queueKey);
        const safeKey = queueKey.replace(/[^a-zA-Z0-9]/g, '_');
        const initials = data.name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
        const isCustomGroup = data.isCustomGroup || false;
        
        // Auto-generate subject from publication dates (not today)
        const defaultSubject = generateSubjectFromDates(data.items);
        const currentSubject = data.subject || defaultSubject;
        
        let itemsHtml = data.items.map((item, idx) => {
            const obsClass = item.obs ? 'mbox-item-obs' : 'mbox-item-obs mbox-item-obs-empty';
            const obsText = item.obs || 'Sem instrução — clique para editar';
            const dateTag = item.dataStr ? `<span class="mbox-item-date">${escapeHtml(item.dataStr)}</span>` : '';
            
            let attachHtml = '';
            if (item.files && item.files.length > 0) {
                attachHtml = `<div class="mbox-item-attachments">
                    ${item.files.map(f => {
                        const icon = f.type?.startsWith('image/') ? 'image' : 'file-text';
                        return `<span class="mbox-attachment-badge"><i data-lucide="${icon}"></i> ${escapeHtml(f.name?.substring(0, 15) || 'arquivo')}</span>`;
                    }).join('')}
                </div>`;
            }
            
            return `<div class="mbox-item">
                <div class="mbox-item-select">
                    <input type="checkbox" class="mbox-item-cb" data-key="${escapeHtml(queueKey)}" data-idx="${idx}">
                </div>
                <div class="mbox-item-content">
                    <div class="mbox-item-cnj">${escapeHtml(item.cnj)} ${dateTag}</div>
                    <div class="mbox-item-orgao">${escapeHtml(item.orgao || '')} ${item.tipo ? '— ' + escapeHtml(item.tipo) : ''}</div>
                    <div class="${obsClass}" contenteditable="true" 
                         onblur="updateItemObs('${escapeHtml(queueKey)}', ${idx}, this.textContent)"
                         title="Clique para editar">${escapeHtml(obsText)}</div>
                    ${attachHtml}
                </div>
                <button class="mbox-item-remove" onclick="removeQueueItem('${escapeHtml(queueKey)}', ${idx})" title="Remover">
                    <i data-lucide="trash-2"></i>
                </button>
            </div>`;
        }).join('');
        
        html += `
        <div class="mbox-recipient-card" id="mbox-card-${safeKey}">
            <div class="mbox-recipient-header">
                <div class="mbox-recipient-info">
                    <div class="mbox-recipient-avatar">${initials}</div>
                    <div>
                        <div class="mbox-recipient-name">${escapeHtml(data.name)}${isCustomGroup ? ' <span class="mbox-custom-tag">Título Personalizado</span>' : ''}</div>
                        <div class="mbox-recipient-email">${escapeHtml(actualEmail)} · ${data.items.length} publicação(ões)</div>
                    </div>
                </div>
                <div class="mbox-recipient-actions">
                    <button class="mbox-send-btn" onclick="flushMailbox('${escapeHtml(queueKey)}')" id="btn-flush-${safeKey}">
                        <i data-lucide="send"></i> Enviar E-mails
                    </button>
                </div>
            </div>
            
            <div class="mbox-subject-row">
                <div class="mbox-subject-label"><i data-lucide="mail" style="width:14px;height:14px;"></i> Assunto do E-mail:</div>
                <input type="text" class="mbox-subject-input" id="subject-${safeKey}" 
                       value="${escapeHtml(currentSubject)}" 
                       onchange="updateRecipientSubject('${escapeHtml(queueKey)}', this.value)" 
                       placeholder="Título do e-mail">
            </div>
            
            <div class="mbox-items-list">${itemsHtml}</div>
            
            <div class="mbox-split-bar">
                <button class="mbox-split-btn" onclick="splitSelectedItems('${escapeHtml(queueKey)}')" title="Separa itens selecionados em um novo e-mail com título diferente">
                    <i data-lucide="split"></i> Separar Selecionados com Outro Título
                </button>
            </div>
        </div>`;
    }
    
    container.innerHTML = html;
    lucide.createIcons();
}

function closeMailbox() {
    document.getElementById('mailbox-modal').classList.add('hidden');
}

function updateItemObs(email, idx, newObs) {
    if (delegationQueue[email] && delegationQueue[email].items[idx]) {
        delegationQueue[email].items[idx].obs = newObs.trim();
        saveDelegationQueue();
    }
}

function removeQueueItem(email, idx) {
    if (delegationQueue[email] && delegationQueue[email].items[idx]) {
        delegationQueue[email].items.splice(idx, 1);
        if (delegationQueue[email].items.length === 0) {
            delete delegationQueue[email];
        }
        saveDelegationQueue();
        renderMailbox();
    }
}

// ─────────── EMAIL FORMATTING & SENDING ───────────

// Extrai e-mail real de uma chave da fila (suporta grupos personalizados)
function getActualEmail(queueKey) {
    if (queueKey.includes('|custom|')) return queueKey.split('|custom|')[0];
    return queueKey;
}

// Gera título do e-mail baseado nas datas das publicações (não na data de hoje)
function generateSubjectFromDates(items) {
    const dates = [...new Set((items || []).map(i => i.dataStr).filter(Boolean))];
    if (dates.length === 0) {
        return `Intimações ${formatDateShort(new Date())}`;
    }
    const sorted = dates.sort((a, b) => {
        const pa = a.split('/'); const pb = b.split('/');
        return new Date(pa[2], pa[1]-1, pa[0]) - new Date(pb[2], pb[1]-1, pb[0]);
    });
    if (sorted.length === 1) return `Intimações ${sorted[0]}`;
    if (sorted.length === 2) return `Intimações ${sorted[0]} e ${sorted[1]}`;
    const last = sorted.pop();
    return `Intimações ${sorted.join(', ')} e ${last}`;
}

// Atualiza o assunto de um destinatário na fila
function updateRecipientSubject(queueKey, newSubject) {
    if (delegationQueue[queueKey]) {
        delegationQueue[queueKey].subject = newSubject.trim();
        saveDelegationQueue();
    }
}

// Separa itens selecionados em um novo grupo com título personalizado
function splitSelectedItems(queueKey) {
    const checkboxes = document.querySelectorAll(`.mbox-item-cb[data-key="${queueKey}"]:checked`);
    if (checkboxes.length === 0) {
        alert('Selecione ao menos um item (checkbox à esquerda) para separar.');
        return;
    }
    const newSubject = prompt('Digite o título do novo e-mail:', 'Intimações');
    if (!newSubject) return;
    
    const data = delegationQueue[queueKey];
    if (!data) return;
    
    const actualEmail = getActualEmail(queueKey);
    const indices = Array.from(checkboxes).map(cb => parseInt(cb.dataset.idx)).sort((a,b) => b - a);
    
    const newKey = `${actualEmail}|custom|${Date.now()}`;
    delegationQueue[newKey] = {
        name: data.name,
        subject: newSubject,
        isCustomGroup: true,
        items: []
    };
    
    indices.forEach(idx => {
        const item = data.items.splice(idx, 1)[0];
        delegationQueue[newKey].items.unshift(item);
    });
    
    if (data.items.length === 0) {
        delete delegationQueue[queueKey];
    }
    
    saveDelegationQueue();
    renderMailbox();
    showToast(`${indices.length} item(ns) separados com título "${newSubject}".`, 'success');
}

// Salva info de delegação no localStorage para exibição persistente
function saveDelegationInfo(cnj, names, obs) {
    const info = `Delegado para ${names.join(', ')}${obs ? ': ' + obs : ''}`;
    delegationLog[cnj] = info;
    try {
        localStorage.setItem('foreigor_delegation_log', JSON.stringify(delegationLog));
    } catch(e) { console.error('Erro ao salvar delegation log:', e); }
}

function getGreeting() {
    const hour = new Date().getHours();
    if (hour >= 0 && hour < 12) return 'Bom dia';
    if (hour >= 12 && hour < 18) return 'Boa tarde';
    return 'Boa noite';
}

function buildEmailBody(recipientNames, items, isGroup) {
    const greeting = getGreeting();
    let body = `<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">`;
    
    if (isGroup) {
        body += `<p>${greeting}, <strong>pessoal</strong>!</p><br>`;
    } else {
        body += `<p>${greeting}, <strong>${recipientNames}</strong>!</p><br>`;
    }
    
    items.forEach(item => {
        let obsText = item.obs && item.obs !== 'Sem instrução — clique para editar' ? ` - ${item.obs}` : '';
        body += `<p style="margin: 0 0 5px 0;">${item.cnj}${obsText}</p>`;
    });
    
    body += `<br><p>Atenciosamente,<br><strong>Rachel Brock</strong></p></div>`;
    
    return body;
}

function flushMailbox(queueKey) {
    if (!CONFIG.APPS_SCRIPT_URL) {
        showToast('Erro: APPS_SCRIPT_URL não configurado.', 'error');
        return;
    }

    const data = delegationQueue[queueKey];
    if (!data || data.items.length === 0) return;
    
    const actualEmail = getActualEmail(queueKey);
    const safeKey = queueKey.replace(/[^a-zA-Z0-9]/g, '_');
    
    // Pegar o assunto do input (permite edição em tempo real) ou do objeto
    const subjectInput = document.getElementById(`subject-${safeKey}`);
    const subject = subjectInput ? subjectInput.value.trim() : (data.subject || generateSubjectFromDates(data.items));
    
    const htmlBody = buildEmailBody(data.name, data.items, data.isGroup);
    
    // Preparar os anexos
    let attachments = [];
    data.items.forEach(item => {
        if (item.files && item.files.length > 0) {
            item.files.forEach(f => {
                attachments.push({
                    name: String(f.name),
                    mimeType: String(f.type),
                    content: String(f.data) // Base64
                });
            });
        }
    });

    const btnId = `btn-flush-${safeKey}`;
    const btn = document.getElementById(btnId);
    if (btn) btn.innerHTML = `<i data-lucide="loader" class="spin"></i> Enviando...`;

    showToast(`O e-mail para ${data.name} está sendo enviado...`, 'info');

    // Mapear apenas para log de texto
    const textLog = itemsToSimpleText(data.items);
    
    // Salvar info de delegação para cada CNJ enviado
    data.items.forEach(item => {
        saveDelegationInfo(item.cnj, [data.name], item.obs || '');
    });

    fetch(CONFIG.APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
            action: 'sendEmails',
            to: actualEmail,
            subject: subject,
            htmlBody: htmlBody,
            attachments: attachments
        })
    }).then(res => res.json())
      .then(resData => {
          if (resData.status === 'ok') {
              showToast(`✓ E-mail enviado para ${data.name}!`, 'success');
              
              // Registrar no histórico
              saveToHistory(actualEmail, data.name, data.items, subject);
              // Log na planilha
              logDelegationToSheets(actualEmail, data.name, data.items, textLog);
              
              delete delegationQueue[queueKey];
              saveDelegationQueue();
              renderMailbox();
          } else {
              throw new Error(resData.error || 'Erro desconhecido');
          }
      })
      .catch(err => {
          console.error(err);
          showToast(`Erro ao enviar e-mail: ${err.message}`, 'error');
          if (btn) btn.innerHTML = `<i data-lucide="send"></i> Enviar E-mails`;
          lucide.createIcons();
      });
}

function itemsToSimpleText(items) {
    return items.map(i => `${i.cnj} - ${i.obs || ''}`).join(' | ');
}

// Botao de Enviar Todos (pode ser iterativo ou uma super requisição)
function flushAllMailbox() {
    if (Object.keys(delegationQueue).length === 0) return;
    
    if (!confirm('Deseja enviar todos os e-mails da fila pendente de uma vez?')) return;
    
    showToast('Iniciando envio em lote...', 'info');
    
    // Itera enviando 1 a 1 para não explodir tempo limite do Apps Script
    const emails = Object.keys(delegationQueue);
    let delay = 0;
    
    emails.forEach((email) => {
        setTimeout(() => {
            flushMailbox(email);
        }, delay);
        delay += 1000; // 1 second headstart per email
    });
}

function downloadBase64File(dataUrl, filename) {
    const a = document.createElement('a');
    a.href = dataUrl;
    a.download = filename || 'anexo';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

// ─────────── LOG DELEGATION TO SHEETS ───────────
function logDelegationToSheets(email, name, items, emailText) {
    if (!CONFIG.APPS_SCRIPT_URL) return;
    
    const cnjs = items.map(i => i.cnj);
    
    fetch(CONFIG.APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
            action: 'logDelegation',
            cnjs: cnjs,
            destinatario: name,
            emailTexto: emailText.substring(0, 500)
        })
    }).catch(err => console.error('Erro ao logar delegação:', err));
}

// ─────────── HISTORY SYSTEM ───────────
let emailHistory = [];
try {
    emailHistory = JSON.parse(localStorage.getItem('foreigor_email_history')) || [];
} catch(e) { emailHistory = []; }

function saveToHistory(email, name, items, subject) {
    const now = new Date();
    const entry = {
        date: `${String(now.getDate()).padStart(2,'0')}/${String(now.getMonth()+1).padStart(2,'0')}/${now.getFullYear()}`,
        time: `${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`,
        timestamp: now.getTime(),
        to: name,
        email: email,
        subject: subject,
        items: items.map(i => ({ cnj: i.cnj, tipo: i.tipo, obs: i.obs })),
        attachmentCount: items.reduce((sum, i) => sum + (i.files?.length || 0), 0)
    };
    emailHistory.unshift(entry);
    // Manter apenas os últimos 200 registros
    if (emailHistory.length > 200) emailHistory = emailHistory.slice(0, 200);
    localStorage.setItem('foreigor_email_history', JSON.stringify(emailHistory));
}

function toggleHistoryPanel() {
    const panel = document.getElementById('history-panel');
    const overlay = document.getElementById('history-overlay');
    
    const isOpen = panel.classList.contains('open');
    
    if (isOpen) {
        panel.classList.remove('open');
        panel.classList.add('hidden');
        overlay.classList.remove('open');
        overlay.classList.add('hidden');
    } else {
        panel.classList.remove('hidden');
        overlay.classList.remove('hidden');
        // Force reflow before adding class for animation
        panel.offsetHeight;
        overlay.offsetHeight;
        panel.classList.add('open');
        overlay.classList.add('open');
        renderHistory();
    }
}

function renderHistory() {
    const container = document.getElementById('history-container');
    
    if (emailHistory.length === 0) {
        container.innerHTML = `
        <div class="hist-empty">
            <i data-lucide="clock"></i>
            <p>Nenhum envio registrado ainda.</p>
            <span style="font-size: 0.82rem;">O histórico aparecerá aqui quando você enviar e-mails pela fila de disparos.</span>
        </div>`;
        lucide.createIcons();
        return;
    }
    
    // Agrupar por data
    const grouped = {};
    emailHistory.forEach(entry => {
        if (!grouped[entry.date]) grouped[entry.date] = [];
        grouped[entry.date].push(entry);
    });
    
    let html = '';
    
    for (const [date, entries] of Object.entries(grouped)) {
        html += `
        <div class="hist-day-group">
            <div class="hist-day-header">
                <i data-lucide="calendar"></i>
                <span>${escapeHtml(date)}</span>
                <span style="font-size:0.72rem; color:var(--text-light); margin-left:auto;">${entries.length} envio(s)</span>
            </div>`;
        
        entries.forEach(entry => {
            const itemsList = entry.items.map(i => {
                let line = i.cnj;
                if (i.tipo) line += ` – ${i.tipo}`;
                return line;
            }).join('<br>');
            
            const attachInfo = entry.attachmentCount > 0 
                ? `<span style="color:#6d28d9; font-size:0.72rem;">📎 ${entry.attachmentCount} anexo(s)</span>` 
                : '';
            
            html += `
            <div class="hist-email-card">
                <div class="hist-email-subject">${escapeHtml(entry.subject)}</div>
                <div class="hist-email-to">
                    <i data-lucide="user"></i> ${escapeHtml(entry.to)} (${escapeHtml(entry.email)})
                </div>
                <div class="hist-email-items">${itemsList}</div>
                <div class="hist-email-time">${escapeHtml(entry.time)} ${attachInfo}</div>
            </div>`;
        });
        
        html += `</div>`;
    }
    
    container.innerHTML = html;
    lucide.createIcons();
}

// ─────────── TOAST NOTIFICATIONS ───────────
function showToast(message, type = 'info') {
    const existing = document.querySelector('.toast');
    if (existing) existing.remove();
    
    const toast = document.createElement('div');
    toast.className = `toast ${type === 'success' ? 'toast-success' : type === 'error' ? 'toast-error' : ''}`;
    
    const iconName = type === 'success' ? 'check-circle' : type === 'error' ? 'alert-circle' : 'info';
    toast.innerHTML = `<i data-lucide="${iconName}"></i> ${escapeHtml(message)}`;
    
    document.body.appendChild(toast);
    lucide.createIcons();
    
    setTimeout(() => {
        if (toast.parentNode) toast.remove();
    }, 3000);
}

// ─────────── NEW PUBLICATIONS DETECTION ───────────
// Estratégia: marca como "novo" toda publicação que seja do dia de hoje (dataDisponibilizacao)
// E que ainda não tenha sido LIDA (ou desconsiderada).
function detectNewPublications() {
    const today = new Date();
    // Reutilizando utilitário para pegar a data de hoje no formato DD/MM/YYYY
    const todayStr = formatDateShort(today);

    allPublications.forEach(pub => {
        if (pub.dataStr === todayStr && pub.status === 'NÃO LIDO') {
            pub.isNew = true;
        } else {
            pub.isNew = false;
        }
    });

    // Cleanup caso haja cache antigo no localStorage que não precise mais
    try { localStorage.removeItem('foreigor_known_cnjs'); } catch(e) {}
}

// ─────────── DROPDOWN SEARCH & CLEAR ───────────
function filterEmpDropdown() {
    const input = document.getElementById('emp-search-input');
    if (!input) return;
    const query = input.value.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
    const items = document.querySelectorAll('#emp-dropdown .dropdown-item, #emp-dropdown .dropdown-group');
    items.forEach(el => {
        const text = el.textContent.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
        el.style.display = text.includes(query) ? '' : 'none';
    });
}

function clearEmpFilter() {
    document.querySelectorAll('#emp-dropdown input:checked').forEach(cb => cb.checked = false);
    const searchInput = document.getElementById('emp-search-input');
    if (searchInput) searchInput.value = '';
    filterEmpDropdown();
    updateMultiFilter('empreendimento');
}

function clearAdvFilter() {
    document.querySelectorAll('#adv-dropdown input:checked').forEach(cb => cb.checked = false);
    updateMultiFilter('advogado');
}

// Initialize badge on load
updateDelegationBadge();

