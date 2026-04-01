// ==========================================
// ESTADO GLOBAL
// ==========================================
let allHearings = [];
let currentView = 'dia';
let geminiApiKey = localStorage.getItem('geminiApiKey') || '';
let selectedHearing = null;
let selectedChartPerson = null; // Pessoa selecionada no gráfico
let selectedChartRole = null;   // 'adv' ou 'prep'
let currentFilteredForChart = []; // Audiências filtradas pelo gráfico

// Equipe oficial da Gramado Parks
const ADV_WHITELIST = ['EDUARDO', 'GIOVANI', 'RENAN'];
const PREP_WHITELIST = ['BRUNA', 'IGOR', 'FRAN', 'RAFA', 'FERNANDA'];
const DOC9_LABEL = 'DOC9';
const SEM_RESP_LABEL = 'Sem Responsável';

// Função para validar se um texto é de fato um link de reunião
function isValidMeetingLink(text) {
    if (!text) return false;
    const lower = text.toLowerCase();
    
    // Ignorar avisos padrões de tribunal e códigos soltos
    const isWarning = lower.includes('verificar se o link') || 
                      lower.includes('ainda não disponível') ||
                      lower.includes('endereço poderá ser cadastrado') ||
                      lower.includes('sala virtual foi divulgado');
    
    if (isWarning) return false;

    // Se tiver http ou www, é link
    if (lower.includes('http') || lower.includes('www.')) return true;

    // Provedores comuns sem http prefixado
    const hosts = ['zoom.us', 'teams.microsoft', 'meet.google', 'webex.com', 'bit.ly', 'biturl', '.jus.br'];
    if (hosts.some(h => lower.includes(h))) return true;

    return false;
}

// ==========================================
// ELEMENTOS DOM
// ==========================================
const fileInput = document.getElementById('csv-file');
const tableBody = document.getElementById('table-body');
const totalAudienciasEl = document.getElementById('total-audiencias');
const viewTitleEl = document.getElementById('view-title');
const filterBtns = document.querySelectorAll('.filter-btn');
const configBtn = document.getElementById('config-ia-btn');
const pessoaFilter = document.getElementById('pessoa-filter');
const calPessoaFilter = document.getElementById('cal-pessoa-filter');

// Modal
const modal = document.getElementById('hearing-modal');
const closeBtn = document.getElementById('close-modal-btn');
const modalProcesso = document.getElementById('modal-processo');
const modalPartes = document.getElementById('modal-partes');
const modalData = document.getElementById('modal-data');
const modalAdvogado = document.getElementById('modal-advogado');
const modalPreposto = document.getElementById('modal-preposto');
const modalLink = document.getElementById('modal-link');
const existingSummarySection = document.getElementById('existing-summary-section');
const modalResumoTexto = document.getElementById('modal-resumo-texto');
const aiGeneratorSection = document.getElementById('ai-generator-section');
const aiResultSection = document.getElementById('ai-result-section');
const aiOutput = document.getElementById('ai-output');
const loadingOverlay = document.getElementById('loading-overlay');

// ==========================================
// CONFIGURAÇÃO DA CHAVE DA API (Gemini)
// ==========================================
configBtn.addEventListener('click', () => {
    const key = prompt("Cole aqui a sua Chave de API do Google Gemini (Gemini PRO):", geminiApiKey);
    if (key !== null) {
        geminiApiKey = key.trim();
        localStorage.setItem('geminiApiKey', geminiApiKey);
        if (geminiApiKey) alert("Chave IA configurada com sucesso!");
    }
});

// ==========================================
// CSV LOADING (Google Sheets)
// ==========================================
const PLANILHA_LINK = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTgROOudqi7NF6i3E_JS5lNdjgnfohAqGzVbCgHXPx1qcrGLKy2jstoAjzZRIp8bmKFdA_zU6R1f57_/pub?gid=438372030&single=true&output=csv";

const PROXIES = [
    "https://api.allorigins.win/raw?url=",
    "https://corsproxy.io/?",
    "https://api.codetabs.com/v1/proxy?quest="
];

async function tryFetchWithProxies(url) {
    for (const proxy of PROXIES) {
        try {
            const res = await fetch(proxy + encodeURIComponent(url), { signal: AbortSignal.timeout(8000) });
            if (res.ok) return await res.text();
        } catch (e) {
            console.warn("Proxy falhou, tentando próximo:", proxy, e.message);
        }
    }
    throw new Error("Todos os proxies falharam.");
}

async function loadData() {
    tableBody.innerHTML = `<tr><td colspan="6" class="empty-state"><div class="empty-content"><p style="color:var(--text-muted)">Conectando à Planilha...</p></div></td></tr>`;
    try {
        const csvText = await tryFetchWithProxies(PLANILHA_LINK);
        Papa.parse(csvText, {
            header: true,
            skipEmptyLines: true,
            complete: function (results) {
                allHearings = results.data;
                populatePessoaFilters();
                updateDashboard();
            },
            error: function (err) {
                console.error("Erro CSV:", err);
                showManualLoadOption();
            }
        });
    } catch (err) {
        console.error("Erro de conexão:", err);
        showManualLoadOption();
    }
}

function showManualLoadOption() {
    tableBody.innerHTML = `
        <tr><td colspan="6" class="empty-state">
            <div class="empty-content">
                <i data-lucide="wifi-off"></i>
                <p><strong>Não foi possível conectar à planilha automaticamente.</strong></p>
                <p style="font-size:0.85rem; margin-top: 8px; color: var(--text-muted)">
                    Clique no botão <strong>"Atualizar Planilha"</strong> acima para tentar novamente.
                </p>
            </div>
        </td></tr>`;
    lucide.createIcons();
}

loadData();

fileInput.addEventListener('click', function (e) {
    e.preventDefault();
    loadData();
});

// ==========================================
// POPULAR FILTROS DE PESSOA
// ==========================================
function populatePessoaFilters() {
    const allPeople = [
        ...ADV_WHITELIST.map(n => ({ name: n, role: 'Adv.' })),
        ...PREP_WHITELIST.map(n => ({ name: n, role: 'Prep.' }))
    ];

    [pessoaFilter, calPessoaFilter].forEach(sel => {
        sel.innerHTML = '<option value="todos">👥 Todos</option>';
        allPeople.forEach(p => {
            const opt = document.createElement('option');
            opt.value = p.name;
            opt.textContent = `${p.name} (${p.role})`;
            sel.appendChild(opt);
        });
    });

    pessoaFilter.addEventListener('change', () => updateDashboard());
    calPessoaFilter.addEventListener('change', () => renderCalendar(calendarYear, calendarMonth));
}

// ==========================================
// NAVEGAÇÃO SPA (ABAS)
// ==========================================
const navTabs = document.querySelectorAll('.nav-tab');
const agendaView = document.getElementById('agenda-view');
const metricsView = document.getElementById('metrics-view');

function switchToAgenda() {
    navTabs.forEach(t => t.classList.remove('active'));
    document.querySelector('[data-target="agenda-view"]').classList.add('active');
    agendaView.classList.remove('hidden');
    metricsView.classList.add('hidden');
}

navTabs.forEach(tab => {
    tab.addEventListener('click', () => {
        navTabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        const target = tab.getAttribute('data-target');
        if (target === 'agenda-view') {
            agendaView.classList.remove('hidden');
            metricsView.classList.add('hidden');
        } else {
            agendaView.classList.add('hidden');
            metricsView.classList.remove('hidden');
            renderMetrics();
        }
    });
});

filterBtns.forEach(btn => {
    if (btn.id === 'config-ia-btn') return;
    btn.addEventListener('click', () => {
        switchToAgenda();
        filterBtns.forEach(b => { if (b.id !== 'config-ia-btn') b.classList.remove('active'); });
        btn.classList.add('active');
        currentView = btn.getAttribute('data-view');

        if (currentView === 'dia') viewTitleEl.textContent = 'Hoje';
        if (currentView === 'semana') viewTitleEl.textContent = 'nesta Semana';
        if (currentView === 'mes') viewTitleEl.textContent = 'neste Mês';
        if (currentView === 'todas') viewTitleEl.textContent = '(Todas as Futuras)';

        const scheduleSection = document.querySelector('.schedule-section');
        const calendarSection = document.getElementById('calendar-section');
        const summaryCards = document.querySelector('.summary-cards');

        if (currentView === 'mes') {
            scheduleSection.classList.add('hidden');
            summaryCards.classList.add('hidden');
            calendarSection.classList.remove('hidden');
            renderCalendar(calendarYear, calendarMonth);
        } else {
            scheduleSection.classList.remove('hidden');
            summaryCards.classList.remove('hidden');
            calendarSection.classList.add('hidden');
            updateDashboard();
        }
    });
});

// ==========================================
// NORMALIZAÇÃO DE NOMES
// ==========================================
function normalizeTeamMember(name, whitelist) {
    if (!name || name.trim() === '') return SEM_RESP_LABEL;
    const upper = name.trim().toUpperCase();
    // Ignorar "Sim/Não" que aparecem em planilhas de teste
    if (upper === 'SIM' || upper === 'NÃO' || upper === 'NAO') return SEM_RESP_LABEL;
    for (const allowed of whitelist) {
        if (upper.includes(allowed)) return allowed;
    }
    return DOC9_LABEL; // Nome externo = DOC9
}

function isDoc9(hearing) {
    // DOC9 é determinado APENAS pela modalidade PRESENCIAL
    // Nomes externos em audiências virtuais NÃO geram DOC9
    const modalidade = (hearing['MODALIDADE'] || '').toUpperCase();
    return modalidade.includes('PRESENCIAL') && !modalidade.includes('SEMI');
}

function pessoaMatchesHearing(pessoaName, hearing) {
    if (pessoaName === 'todos') return true;
    const rawAdv = hearing['ADVOGADO/A'] || hearing['ADVOGADO'] || '';
    const rawPrep = hearing['PREPOSTO'] || '';
    const upper = pessoaName.toUpperCase();
    return rawAdv.toUpperCase().includes(upper) || rawPrep.toUpperCase().includes(upper);
}

// ==========================================
// UPDATE DASHBOARD (tabela)
// ==========================================
function updateDashboard() {
    if (allHearings.length === 0) return;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const pessoaSel = pessoaFilter ? pessoaFilter.value : 'todos';

    const filteredHearings = allHearings.filter(hearing => {
        const dataStr = hearing['DATA'] || hearing['Data'] || hearing['data'];
        if (!dataStr) return false;
        const parts = dataStr.split('/');
        if (parts.length !== 3) return false;
        const hearingDate = new Date(parts[2], parts[1] - 1, parts[0]);
        hearingDate.setHours(0, 0, 0, 0);

        let dateOk = false;
        if (currentView === 'dia') {
            dateOk = hearingDate.getTime() === today.getTime();
        } else if (currentView === 'semana') {
            const dayOfWeek = today.getDay();
            const startOfWeek = new Date(today);
            startOfWeek.setDate(today.getDate() - dayOfWeek);
            const endOfWeek = new Date(startOfWeek);
            endOfWeek.setDate(startOfWeek.getDate() + 6);
            dateOk = hearingDate >= startOfWeek && hearingDate <= endOfWeek;
        } else if (currentView === 'mes') {
            dateOk = hearingDate.getMonth() === today.getMonth() && hearingDate.getFullYear() === today.getFullYear();
        } else if (currentView === 'todas') {
            dateOk = hearingDate >= today;
        }

        if (!dateOk) return false;
        return pessoaMatchesHearing(pessoaSel, hearing);
    });

    totalAudienciasEl.textContent = filteredHearings.length;
    renderTable(filteredHearings);
}

// ==========================================
// RENDER TABLE
// ==========================================
function renderTable(hearings) {
    tableBody.innerHTML = '';

    if (hearings.length === 0) {
        tableBody.innerHTML = `<tr><td colspan="6" class="empty-state"><div class="empty-content"><i data-lucide="calendar-x"></i><p>Nenhuma audiência encontrada.</p></div></td></tr>`;
        lucide.createIcons();
        return;
    }

    hearings.forEach(h => {
        const data = h['DATA'] || '-';
        const hora = h['HORÁRIO'] || h['HORA'] || '-';
        const autor = h['AUTOR'] || '-';
        const reu = h['RÉU'] || h['REU'] || '-';
        const processo = h['N° DO PROCESSO'] || h['PROCESSO'] || '-';
        const advogado = h['ADVOGADO/A'] || h['ADVOGADO'] || '-';
        const preposto = h['PREPOSTO'] || '-';
        const vara = h['VARA'] || '-';
        const comarca = h['COMARCA'] || '-';
        const modalidade = h['MODALIDADE'] || '-';
        const tipoCompromisso = h['TIPO DE COMPROMISSO'] || h['TIPO'] || '';

        // Detectar se é DOC9
        const doc9 = isDoc9(h);

        // Audiências passadas
        let isPast = false;
        try {
            const dateParts = data.split('/');
            if (dateParts.length === 3) {
                const exactDate = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);
                if (hora && hora.includes(':')) {
                    const tp = hora.split(':');
                    exactDate.setHours(parseInt(tp[0]), parseInt(tp[1]), 0);
                } else {
                    exactDate.setHours(23, 59, 59);
                }
                if (exactDate < new Date()) isPast = true;
            }
        } catch (e) {}

        const tr = document.createElement('tr');
        if (isPast) tr.classList.add('past-hearing');
        if (doc9) tr.classList.add('doc9-row');

        tr.onclick = () => openHearingModal(h);

        // Badge do tipo de compromisso
        const tipoHtml = tipoCompromisso
            ? `<br><span class="chip chip-tipo" style="margin-top:4px;">${tipoCompromisso}</span>`
            : '';

        // Badge DOC9 na coluna de equipe
        const doc9Badge = doc9 ? `<span class="chip chip-doc9" style="margin-left:4px;">DOC9</span>` : '';

        tr.innerHTML = `
            <td>
                <strong>${data}</strong><br>
                <span style="color: var(--text-muted); font-size: 0.85rem;"><i data-lucide="clock" style="width:14px;height:14px;vertical-align:middle"></i> ${hora}</span>
            </td>
            <td>
                <span class="chip chip-gold">${autor}</span><br>
                <span style="font-size: 0.85rem">x ${reu}</span>
            </td>
            <td><strong style="font-size:0.9rem">${processo}</strong></td>
            <td>
                <span style="display:flex;align-items:center;justify-content:center;gap:5px;font-weight:600;flex-wrap:wrap;">
                    <i data-lucide="user" style="width:14px;height:14px;flex-shrink:0;"></i> ${advogado} (Adv.) ${doc9Badge}
                </span>
                <span style="display:flex;align-items:center;justify-content:center;gap:5px;font-size:0.85rem;color:var(--text-muted);margin-top:2px;">
                    <i data-lucide="contact" style="width:14px;height:14px;flex-shrink:0;"></i> ${preposto} (Prep.)
                </span>
            </td>
            <td>${vara}<br><span style="color: var(--text-muted); font-size: 0.85rem">${comarca}</span></td>
            <td>
                <span class="chip">${modalidade}</span>
                ${tipoHtml}
            </td>
        `;
        tableBody.appendChild(tr);
    });
    lucide.createIcons();
}

// ==========================================
// MODAL (Painel Inteligente)
// ==========================================
function openHearingModal(hearingInfo) {
    selectedHearing = hearingInfo;
    modalProcesso.textContent = hearingInfo['N° DO PROCESSO'] || '-';
    modalPartes.textContent = (hearingInfo['AUTOR'] || 'Autor') + ' vs ' + (hearingInfo['RÉU'] || 'Réu');
    modalData.textContent = (hearingInfo['DATA'] || '') + ' às ' + (hearingInfo['HORÁRIO'] || '');
    modalAdvogado.textContent = hearingInfo['ADVOGADO/A'] || hearingInfo['ADVOGADO'] || '-';
    modalPreposto.textContent = hearingInfo['PREPOSTO'] || '-';

    const obsOriginais = hearingInfo['OBSERVAÇÃO'] || hearingInfo['OBSERVACAO'] || 'Sem link cadastrado.';
    modalLink.innerHTML = obsOriginais.replace(/(https?:\/\/[^\s]+)/g, '<a href="$1" target="_blank" style="color:var(--primary-color);text-decoration:underline;word-break:break-all;">$1</a>');

    const obs = hearingInfo['RESUMO IA'] || hearingInfo['RESUMO_IA'] || hearingInfo['RESUMO'] || '';
    if (obs.length > 20) {
        existingSummarySection.classList.remove('hidden');
        aiGeneratorSection.classList.add('hidden');
        modalResumoTexto.innerHTML = obs;
    } else {
        existingSummarySection.classList.add('hidden');
        aiGeneratorSection.classList.remove('hidden');
        aiResultSection.classList.add('hidden');
        document.getElementById('pdf-inicial').value = '';
        document.getElementById('pdf-contestacao').value = '';
        document.getElementById('pdf-inicial-status').textContent = '';
        document.getElementById('pdf-contestacao-status').textContent = '';
    }

    modal.classList.remove('hidden');
}

closeBtn.onclick = () => modal.classList.add('hidden');

// ==========================================
// EXTRAÇÃO DE PDF
// ==========================================
async function extractTextFromPDF(fileInputEl, statusEl) {
    const file = fileInputEl.files[0];
    if (!file) return "";
    statusEl.textContent = "Lendo arquivo...";
    try {
        const fileReader = new FileReader();
        const arrayBuffer = await new Promise((resolve, reject) => {
            fileReader.onload = () => resolve(fileReader.result);
            fileReader.onerror = reject;
            fileReader.readAsArrayBuffer(file);
        });
        const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
        let finalString = "";
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            finalString += content.items.map(item => item.str).join(" ") + "\n";
        }
        statusEl.textContent = "Lido com sucesso!";
        return finalString;
    } catch (e) {
        statusEl.textContent = "Erro ao ler PDF!";
        console.error(e);
        return "";
    }
}

// ==========================================
// GERAÇÃO DE ROTEIRO COM IA
// ==========================================
document.getElementById('generate-ai-btn').addEventListener('click', async () => {
    if (!geminiApiKey) {
        alert("Por favor, configure sua chave da API do Gemini primeiro!");
        return;
    }

    loadingOverlay.classList.remove('hidden');

    const textoInicial = await extractTextFromPDF(document.getElementById('pdf-inicial'), document.getElementById('pdf-inicial-status'));
    const textoContestacao = await extractTextFromPDF(document.getElementById('pdf-contestacao'), document.getElementById('pdf-contestacao-status'));

    if (!textoInicial && !textoContestacao) {
        loadingOverlay.classList.add('hidden');
        alert("Por favor, anexe ao menos um arquivo PDF.");
        return;
    }

    const promptText = `Você é um assistente jurídico extremamente direto e focado.
    
Abaixo estão os textos extraídos dos PDFs da Petição Inicial e da nossa Contestação (Gramado Parks). 
Gere um resumo formatado EXATAMENTE como este modelo simples:

**O QUE O AUTOR ALEGOU E PEDE:**
- Alegou que [ponto principal 1]
- Reclamou de [ponto principal 2]
- Pediu [pedido principal e valores, se houver]

**NOSSA CONTESTAÇÃO (O QUE FALAMOS NA DEFESA):**
- Falamos que [argumento 1]
- Justificamos que [argumento 2]
- Provamos que [argumento 3, se houver]

TEXTO DA INICIAL:
${textoInicial}

TEXTO DA CONTESTAÇÃO:
${textoContestacao}`;

    try {
        const modelsRes = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${geminiApiKey}`);
        if (!modelsRes.ok) throw new Error("Sua chave de API é inválida ou expirou.");

        const modelsData = await modelsRes.json();
        const availableModels = modelsData.models.filter(m =>
            m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent")
        );

        if (availableModels.length === 0) throw new Error("Sua permissão do Google bloqueia o uso de ferramentas de IA.");

        let bestModel = availableModels.find(m => m.name.includes("2.5-flash")) ||
            availableModels.find(m => m.name.includes("2.0-flash")) ||
            availableModels.find(m => m.name.includes("1.5-flash")) ||
            availableModels.find(m => m.name.includes("pro")) ||
            availableModels[0];

        const modelUrlName = bestModel.name.replace("models/", "");
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelUrlName}:generateContent?key=${geminiApiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ contents: [{ parts: [{ text: promptText }] }] })
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.error ? data.error.message : "Desconhecido");

        const gerado = data.candidates[0].content.parts[0].text;
        loadingOverlay.classList.add('hidden');
        aiResultSection.classList.remove('hidden');
        aiOutput.innerHTML = gerado.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');

        document.getElementById('copy-summary-btn').onclick = () => {
            navigator.clipboard.writeText(gerado).then(() => {
                alert("Roteiro COPIADO! Cole na coluna 'OBSERVAÇÃO' da planilha.");
            });
        };
    } catch (e) {
        loadingOverlay.classList.add('hidden');
        alert("Erro ao falar com o Google: " + e.message);
    }
});

// ==========================================
// PAINEL GERENCIAL (Métricas e Gráficos)
// ==========================================
let advChartInstance = null;
let prepChartInstance = null;

const metricMonthSelect = document.getElementById('metric-month-select');
const metricTotal = document.getElementById('metric-total-audiencias');

if (metricMonthSelect) {
    metricMonthSelect.addEventListener('change', () => {
        clearChartSelection();
        drawCharts(metricMonthSelect.value);
    });
}

function renderMetrics() {
    if (allHearings.length === 0) return;

    const mesesSet = new Set();
    allHearings.forEach(h => {
        const dataStr = h['DATA'] || h['Data'] || h['data'];
        if (dataStr && dataStr.split('/').length === 3) {
            const parts = dataStr.split('/');
            mesesSet.add(`${parts[1]}/${parts[2]}`);
        }
    });

    const MONTHS_PT_SHORT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

    const mesesArray = Array.from(mesesSet).sort((a, b) => {
        const [mA, yA] = a.split('/');
        const [mB, yB] = b.split('/');
        return (new Date(yA, mA - 1)) - (new Date(yB, mB - 1));
    });

    metricMonthSelect.innerHTML = '<option value="todos">Todo o Período</option>';
    mesesArray.forEach(m => {
        const [mes, ano] = m.split('/');
        const option = document.createElement('option');
        option.value = m;
        option.textContent = `${MONTHS_PT_SHORT[parseInt(mes) - 1]}/${ano}`;
        metricMonthSelect.appendChild(option);
    });

    const hoje = new Date();
    const currentMonthStr = `${String(hoje.getMonth() + 1).padStart(2, '0')}/${hoje.getFullYear()}`;
    if (mesesArray.includes(currentMonthStr)) {
        metricMonthSelect.value = currentMonthStr;
    }

    clearChartSelection();
    drawCharts(metricMonthSelect.value);
}

// ==========================================
// DESENHAR GRÁFICOS
// ==========================================
function getHearingsForMonth(selectedMonth) {
    if (selectedMonth === "todos") return allHearings;
    return allHearings.filter(h => {
        const d = h['DATA'] || '';
        const parts = d.split('/');
        return parts.length === 3 && `${parts[1]}/${parts[2]}` === selectedMonth;
    });
}

function drawCharts(selectedMonth) {
    const mHearings = getHearingsForMonth(selectedMonth);
    metricTotal.textContent = mHearings.length;

    const advogadosCount = {};
    const prepostosCount = {};

    mHearings.forEach(h => {
        const rawAdv = h['ADVOGADO/A'] || h['ADVOGADO'] || h['Advogado'] || '';
        const rawPrep = h['PREPOSTO'] || h['Preposto'] || '';

        const adv = normalizeTeamMember(rawAdv, ADV_WHITELIST);
        const prep = normalizeTeamMember(rawPrep, PREP_WHITELIST);

        advogadosCount[adv] = (advogadosCount[adv] || 0) + 1;
        prepostosCount[prep] = (prepostosCount[prep] || 0) + 1;
    });

    // Garante que todos da equipe aparecem mesmo com 0
    ADV_WHITELIST.forEach(name => { if (!advogadosCount[name]) advogadosCount[name] = 0; });
    PREP_WHITELIST.forEach(name => { if (!prepostosCount[name]) prepostosCount[name] = 0; });

    const sortFn = (count) => (a, b) => {
        if (a === SEM_RESP_LABEL) return 1;
        if (b === SEM_RESP_LABEL) return -1;
        if (a === DOC9_LABEL) return 1;
        if (b === DOC9_LABEL) return -1;
        return count[b] - count[a];
    };

    const advLabels = Object.keys(advogadosCount).sort(sortFn(advogadosCount));
    const advData = advLabels.map(l => advogadosCount[l]);

    const prepLabels = Object.keys(prepostosCount).sort(sortFn(prepostosCount));
    const prepData = prepLabels.map(l => prepostosCount[l]);

    // Cores diferenciadas: equipe=verde/ouro, DOC9=roxo, Sem Resp=cinza
    const advColors = advLabels.map(l => {
        if (l === SEM_RESP_LABEL) return '#dce3e0';
        if (l === DOC9_LABEL) return '#c17fd0';
        return '#165b41';
    });
    const prepColors = prepLabels.map(l => {
        if (l === SEM_RESP_LABEL) return '#dce3e0';
        if (l === DOC9_LABEL) return '#c17fd0';
        return '#d4af37';
    });

    const makeChartOptions = (labels, role) => ({
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: ctx => ` ${ctx.raw} audiência(s)` } }
        },
        scales: {
            y: {
                beginAtZero: true,
                ticks: { stepSize: 1 },
                grid: { color: 'rgba(0,0,0,0.05)' }
            },
            x: {
                ticks: { font: { weight: '600', size: 13 }, color: '#2c3e38' },
                grid: { display: false }
            }
        },
        onClick: (evt, elements) => {
            if (elements.length > 0) {
                const idx = elements[0].index;
                const personName = labels[idx];
                selectChartPerson(personName, role, metricMonthSelect.value);
            }
        }
    });

    if (advChartInstance) advChartInstance.destroy();
    if (prepChartInstance) prepChartInstance.destroy();

    const advCtx = document.getElementById('advogadoChart').getContext('2d');
    advChartInstance = new Chart(advCtx, {
        type: 'bar',
        data: {
            labels: advLabels,
            datasets: [{ label: 'Audiências', data: advData, backgroundColor: advColors, borderRadius: 6 }]
        },
        options: makeChartOptions(advLabels, 'adv')
    });

    const prepCtx = document.getElementById('prepostoChart').getContext('2d');
    prepChartInstance = new Chart(prepCtx, {
        type: 'bar',
        data: {
            labels: prepLabels,
            datasets: [{ label: 'Audiências', data: prepData, backgroundColor: prepColors, borderRadius: 6 }]
        },
        options: makeChartOptions(prepLabels, 'prep')
    });
}

// ==========================================
// CLIQUE NO GRÁFICO → DETALHES DA PESSOA
// ==========================================
function selectChartPerson(personName, role, selectedMonth) {
    selectedChartPerson = personName;
    selectedChartRole = role;

    const mHearings = getHearingsForMonth(selectedMonth);

    // Filtrar audiências da pessoa selecionada
    const personHearings = mHearings.filter(h => {
        const rawAdv = h['ADVOGADO/A'] || h['ADVOGADO'] || '';
        const rawPrep = h['PREPOSTO'] || '';
        const advNorm = normalizeTeamMember(rawAdv, ADV_WHITELIST);
        const prepNorm = normalizeTeamMember(rawPrep, PREP_WHITELIST);
        if (role === 'adv') return advNorm === personName || (personName === DOC9_LABEL && advNorm === DOC9_LABEL) || (personName === SEM_RESP_LABEL && advNorm === SEM_RESP_LABEL);
        if (role === 'prep') return prepNorm === personName || (personName === DOC9_LABEL && prepNorm === DOC9_LABEL) || (personName === SEM_RESP_LABEL && prepNorm === SEM_RESP_LABEL);
        return false;
    });

    currentFilteredForChart = personHearings;

    // Atualizar visual dos gráficos: esmaece quem não foi clicado
    dimChartBars(role, personName);

    // Mostrar seção de detalhes
    const detailSection = document.getElementById('chart-detail-section');
    detailSection.classList.remove('hidden');
    detailSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

    const roleLabel = role === 'adv' ? 'Advogado(a)' : 'Preposto(a)';
    document.getElementById('detail-person-badge').textContent = personName;
    document.getElementById('detail-count-text').textContent = `${personHearings.length} audiência(s) como ${roleLabel}`;

    renderDetailTable(personHearings, role);
}

function dimChartBars(role, selectedName) {
    const applyDim = (chart, labels) => {
        const colors = chart.data.datasets[0].backgroundColor;
        const newColors = labels.map((l, i) => {
            const originalColor = colors[i];
            return l === selectedName ? originalColor : originalColor + '55'; // adiciona transparência
        });
        chart.data.datasets[0].backgroundColor = newColors;
        chart.update();
    };

    if (role === 'adv' && advChartInstance) {
        applyDim(advChartInstance, advChartInstance.data.labels);
    }
    if (role === 'prep' && prepChartInstance) {
        applyDim(prepChartInstance, prepChartInstance.data.labels);
    }
}

function renderDetailTable(hearings, role) {
    const tbody = document.getElementById('detail-table-body');
    tbody.innerHTML = '';

    if (hearings.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" class="empty-state"><div class="empty-content"><i data-lucide="calendar-x"></i><p>Nenhuma audiência encontrada.</p></div></td></tr>`;
        lucide.createIcons();
        return;
    }

    hearings.forEach(h => {
        const data = h['DATA'] || '-';
        const hora = h['HORÁRIO'] || h['HORA'] || '-';
        const autor = h['AUTOR'] || '-';
        const reu = h['RÉU'] || h['REU'] || '-';
        const processo = h['N° DO PROCESSO'] || h['PROCESSO'] || '-';
        const advogado = h['ADVOGADO/A'] || h['ADVOGADO'] || '-';
        const preposto = h['PREPOSTO'] || '-';
        const vara = h['VARA'] || '-';
        const comarca = h['COMARCA'] || '-';
        const modalidade = h['MODALIDADE'] || '-';
        const tipoCompromisso = h['TIPO DE COMPROMISSO'] || h['TIPO'] || '';

        const papel = role === 'adv'
            ? `<span class="chip" style="background:#e8f5ee;color:#165b41;">Advogado: ${advogado}</span>`
            : `<span class="chip" style="background:#fdf5e0;color:#8f6d0f;">Preposto: ${preposto}</span>`;

        const tipoHtml = tipoCompromisso ? `<br><span class="chip chip-tipo">${tipoCompromisso}</span>` : '';

        const tr = document.createElement('tr');
        tr.onclick = () => openHearingModal(h);
        tr.innerHTML = `
            <td><strong>${data}</strong><br><span style="color:var(--text-muted);font-size:0.85rem;">${hora}</span></td>
            <td><span class="chip chip-gold">${autor}</span><br><span style="font-size:0.85rem;">x ${reu}</span></td>
            <td><strong style="font-size:0.9rem">${processo}</strong></td>
            <td>${papel}</td>
            <td>${vara}<br><span style="color:var(--text-muted);font-size:0.85rem;">${comarca}</span></td>
            <td><span class="chip">${modalidade}</span>${tipoHtml}</td>
        `;
        tbody.appendChild(tr);
    });
    lucide.createIcons();
}

function clearChartSelection() {
    selectedChartPerson = null;
    selectedChartRole = null;
    currentFilteredForChart = [];
    document.getElementById('chart-detail-section').classList.add('hidden');

    // Restaurar cores originais dos gráficos
    const restoreColors = (chart, colorFn) => {
        if (!chart) return;
        chart.data.datasets[0].backgroundColor = chart.data.labels.map(colorFn);
        chart.update();
    };

    if (advChartInstance) {
        restoreColors(advChartInstance, l => {
            if (l === SEM_RESP_LABEL) return '#dce3e0';
            if (l === DOC9_LABEL) return '#c17fd0';
            return '#165b41';
        });
    }
    if (prepChartInstance) {
        restoreColors(prepChartInstance, l => {
            if (l === SEM_RESP_LABEL) return '#dce3e0';
            if (l === DOC9_LABEL) return '#c17fd0';
            return '#d4af37';
        });
    }
}

document.getElementById('clear-chart-selection-btn').addEventListener('click', () => {
    clearChartSelection();
    drawCharts(metricMonthSelect.value);
});

// ==========================================
// EXPORTAR EXCEL (SheetJS)
// ==========================================
function exportToExcel(hearings, filename) {
    if (!hearings || hearings.length === 0) {
        alert("Nenhuma audiência para exportar com os filtros atuais.");
        return;
    }

    const MONTHS_PT_SHORT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

    // Montar dados para a planilha
    const rows = hearings.map(h => ({
        'Data': h['DATA'] || '',
        'Hora': h['HORÁRIO'] || h['HORA'] || '',
        'Autor': h['AUTOR'] || '',
        'Réu': h['RÉU'] || h['REU'] || '',
        'Processo': h['N° DO PROCESSO'] || h['PROCESSO'] || '',
        'Advogado': h['ADVOGADO/A'] || h['ADVOGADO'] || '',
        'Preposto': h['PREPOSTO'] || '',
        'Vara': h['VARA'] || '',
        'Comarca': h['COMARCA'] || '',
        'Modalidade': h['MODALIDADE'] || '',
        'Tipo': h['TIPO DE COMPROMISSO'] || h['TIPO'] || '',
        'Observação': h['OBSERVAÇÃO'] || h['OBSERVACAO'] || '',
        'DOC9': isDoc9(h) ? 'Sim' : 'Não'
    }));

    const ws = XLSX.utils.json_to_sheet(rows);

    // Estilo da largura das colunas
    ws['!cols'] = [
        { wch: 12 }, { wch: 8 }, { wch: 25 }, { wch: 25 },
        { wch: 22 }, { wch: 18 }, { wch: 18 }, { wch: 25 },
        { wch: 20 }, { wch: 14 }, { wch: 25 }, { wch: 30 }, { wch: 6 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Audiências');
    XLSX.writeFile(wb, filename);
}

// Botão: Baixar tudo do mês
document.getElementById('download-excel-btn').addEventListener('click', () => {
    const selectedMonth = metricMonthSelect.value;
    const hearings = getHearingsForMonth(selectedMonth);
    const label = selectedMonth === 'todos' ? 'todos' : selectedMonth.replace('/', '-');
    exportToExcel(hearings, `audiencias_${label}.xlsx`);
});

// Botão: Baixar filtrado (com pessoa selecionada no gráfico)
document.getElementById('download-filtered-excel-btn').addEventListener('click', () => {
    const selectedMonth = metricMonthSelect.value;
    const label = selectedMonth === 'todos' ? 'todos' : selectedMonth.replace('/', '-');
    const personLabel = selectedChartPerson ? `_${selectedChartPerson.replace(' ', '_')}` : '';
    exportToExcel(currentFilteredForChart, `audiencias_${label}${personLabel}.xlsx`);
});

// ==========================================
// CALENDÁRIO MENSAL INTERATIVO
// ==========================================
const hojeDate = new Date();
let calendarMonth = hojeDate.getMonth();
let calendarYear = hojeDate.getFullYear();

const MONTHS_PT = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

// Feriados Nacionais e Regionais (Gramado/RS) para 2026
const FERIADOS_2026 = {
    '1/1': 'Confraternização Universal',
    '3/4': 'Sexta-feira Santa',
    '21/4': 'Tiradentes',
    '1/5': 'Dia do Trabalho',
    '26/5': 'N. Sra. de Caravaggio (RS)',
    '4/6': 'Corpus Christi',
    '7/9': 'Independência',
    '20/9': 'Rev. Farroupilha (RS)',
    '12/10': 'N. Sra. Aparecida',
    '2/11': 'Finados',
    '15/11': 'Proclamação da República',
    '20/11': 'Consciência Negra',
    '8/12': 'N. Sra. Imaculada Conceição',
    '25/12': 'Natal'
};

function renderCalendar(year, month) {
    calendarYear = year;
    calendarMonth = month;

    document.getElementById('cal-title').textContent = `${MONTHS_PT[month]} ${year}`;
    const grid = document.getElementById('calendar-grid');
    grid.innerHTML = '';

    const now = new Date();
    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();

    const calPessoaSel = calPessoaFilter ? calPessoaFilter.value : 'todos';

    // Indexar audiências do mês por dia
    const hearingsByDay = {};
    allHearings.forEach(h => {
        const dataStr = h['DATA'] || '';
        const parts = dataStr.split('/');
        if (parts.length === 3) {
            const d = parseInt(parts[0]);
            const m = parseInt(parts[1]) - 1;
            const y = parseInt(parts[2]);
            if (m === month && y === year) {
                if (!hearingsByDay[d]) hearingsByDay[d] = [];
                hearingsByDay[d].push(h);
            }
        }
    });

    // Células vazias antes do dia 1
    for (let i = 0; i < firstDay; i++) {
        const empty = document.createElement('div');
        empty.className = 'cal-day empty';
        grid.appendChild(empty);
    }

    // Dias do mês
    for (let day = 1; day <= daysInMonth; day++) {
        const cell = document.createElement('div');
        cell.className = 'cal-day';

        const isToday = day === now.getDate() && month === now.getMonth() && year === now.getFullYear();
        if (isToday) cell.classList.add('today');

        // Checar Feriado
        const feriadoKey = `${day}/${month + 1}`;
        if (FERIADOS_2026[feriadoKey]) {
            cell.classList.add('feriado-cell');
            const feriadoName = document.createElement('div');
            feriadoName.className = 'feriado-name';
            feriadoName.textContent = FERIADOS_2026[feriadoKey];
            cell.appendChild(feriadoName);
        }

        const numDiv = document.createElement('div');
        numDiv.className = 'cal-day-num';
        numDiv.textContent = day;
        cell.appendChild(numDiv);

        const dayHearings = hearingsByDay[day] || [];
        dayHearings.forEach(h => {
            const hora = h['HORÁRIO'] || h['HORA'] || '';
            const autor = h['AUTOR'] || 'Autor';
            const rawAdv = h['ADVOGADO/A'] || h['ADVOGADO'] || h['Advogado'] || '';
            const rawPrep = h['PREPOSTO'] || h['Preposto'] || '';
            const obs = h['OBSERVAÇÃO'] || h['OBSERVACAO'] || h['Observação'] || '';
            const tipoCompromisso = h['TIPO DE COMPROMISSO'] || h['TIPO'] || '';

            const advNorm = normalizeTeamMember(rawAdv, ADV_WHITELIST);
            const prepNorm = normalizeTeamMember(rawPrep, PREP_WHITELIST);

            const semAdv = !rawAdv.trim();
            const semPrep = !rawPrep.trim();
            const semObs = !obs.trim();
            const modalidade = (h['MODALIDADE'] || '').toUpperCase();
            const advogadoTxt = rawAdv.toUpperCase();
            const hasRealLink = isValidMeetingLink(obs);

            // === LÓGICA DE STATUS — 5 REGRAS FIXAS ===
            // A MODALIDADE é sempre o critério principal.
            let statusClass = '';

            if (advogadoTxt.includes('CANCELADA')) {
                // REGRA 1: Cancelada
                statusClass = 'status-cancelada';

            } else if (modalidade.includes('PRESENCIAL') && !modalidade.includes('SEMI')) {
                // REGRA 2: Presencial (DOC9)
                // Sem ADV e sem PREP = Pendente (Split)
                // Com ADV ou PREP = Confirmada (Magenta)
                statusClass = (semAdv && semPrep) ? 'status-doc9-split' : 'status-doc9';

            } else {
                // REGRA 3: Virtual, Semi-Presencial ou sem modalidade
                // Sem link real = Azul (Sem Link)
                // Com link real = Amarelo (Virtual)
                statusClass = hasRealLink ? 'status-virtual' : 'status-sem-link';
            }

            const dateTime = new Date(year, month, day);
            if (hora && hora.includes(':')) {
                const tp = hora.split(':');
                dateTime.setHours(parseInt(tp[0]), parseInt(tp[1]), 0);
            } else {
                dateTime.setHours(23, 59, 59);
            }
            const isPast = dateTime < now;

            const ev = document.createElement('div');
            ev.className = 'cal-event' + (statusClass ? ' ' + statusClass : '') + (isPast ? ' past-event' : '');

            // Filtro de pessoa no calendário
            const matchesPessoa = pessoaMatchesHearing(calPessoaSel, h);
            if (calPessoaSel !== 'todos' && !matchesPessoa) {
                ev.classList.add('filtered-out');
            }

            const tipoLabel = tipoCompromisso ? ` [${tipoCompromisso}]` : '';
            ev.title = `${hora} | ${autor}${tipoLabel} | Adv: ${rawAdv || 'N/A'} | Prep: ${rawPrep || 'N/A'} | ${statusClass === 'status-doc9' ? 'DOC9' : ''}`;
            ev.textContent = `${hora ? hora + ' ' : ''}${autor}`;
            ev.onclick = (e) => {
                e.stopPropagation();
                openHearingModal(h);
            };
            cell.appendChild(ev);
        });

        grid.appendChild(cell);
    }

    lucide.createIcons();
}

// Botões de navegar (Mês anterior / Próximo)
document.getElementById('cal-prev').addEventListener('click', () => {
    calendarMonth--;
    if (calendarMonth < 0) { calendarMonth = 11; calendarYear--; }
    renderCalendar(calendarYear, calendarMonth);
});

document.getElementById('cal-next').addEventListener('click', () => {
    calendarMonth++;
    if (calendarMonth > 11) { calendarMonth = 0; calendarYear++; }
    renderCalendar(calendarYear, calendarMonth);
});
