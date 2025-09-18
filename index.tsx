// --- DECLARAÇÕES GLOBAIS ---
declare var XLSX: any;
declare var jspdf: any;
declare var html2canvas: any;
declare const firebase: any;

// Resolve ReferenceError para bibliotecas CDN
const Chart = (window as any).Chart;
const ChartDataLabels = (window as any).ChartDataLabels;


// --- CONFIGURAÇÃO E INICIALIZAÇÃO DO FIREBASE ---
const firebaseConfig = {
    apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
    authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
    projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
    storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
    messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
    appId: import.meta.env.VITE_FIREBASE_APP_ID
};
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();


// --- TYPE DEFINITIONS ---
interface QuotaData {
    __id: number;
    [key: string]: string | number | null;
    'DATA REGISTRO DI': string | null;
    'REGISTRATION_TYPE'?: 'QUOTA' | 'INTEGRAL' | null;
    'QUOTE': string | null;
    'VALOR USD': string | number | null;
    'PO ': string | null;
    'PO': string | null;
    'PROJECT': string | null;
    'LI NUMBER': string | null;
    'QTD VEÍCULOS': string | number | null;
}

// --- TRANSLATIONS ---
const translations = {
    'pt-BR': {
        pageTitle: 'Dashboard de Controle de Quotas de Importação',
        headerTitle: 'DASHBOARD DE CONTROLE DE QUOTAS',
        promptToUpload: 'Carregue a planilha de prioridade de registro para começar',
        exportPDF: 'Exportar PDF',
        exportCSV: 'Exportar CSV',
        uploadSheet: 'Carregar Planilha',
        quotaEV: 'QUOTA EV',
        quotaPHEV: 'QUOTA PHEV',
        totalUSD: 'Total (USD)',
        usedUSD: 'Utilizado (USD)',
        usedVehicles: 'Utilizado (Veículos)',
        inOrderUSD: 'Em Pedido (USD)',
        inOrderVehicles: 'Em Pedido (Veículos)',
        balanceUSD: 'SALDO (USD)',
        summary: 'Resumo Geral',
        totalOrders: 'Total de Pedidos na Planilha',
        usedOrders: 'Pedidos Utilizados (DI Registrada)',
        pendingOrders: 'Pedidos em Aberto / Pendentes',
        integralSummary: 'Resumo - Reg. Integral',
        integralOrders: 'Total Pedidos (Integral)',
        integralTotalValue: 'Valor Total (USD)',
        integralTotalVehicles: 'Total Veículos (Integral)',
        quotaVisualization: 'Visualização das Quotas (USD)',
        filterByLIPO: 'Filtrar por LI ou PO...',
        pendingOrdersList: 'Pedidos em Aberto / Pendentes',
        usedOrdersList: 'Pedidos Utilizados (DI Registrada)',
        integralOrdersList: 'Pedidos com Registro Integral',
        waitingForFile: 'Aguardando arquivo...',
        selectFileToStart: 'Selecione a planilha de prioridade de registro para iniciar.',
        toastLoaded: 'Dashboard de quotas carregado!',
        toastNoSheet: 'Nenhuma planilha válida (ex: "Sheet1" ou "Planilha1") foi encontrada.',
        toastEmptySheet: 'A planilha está vazia.',
        toastProcessError: 'Erro ao processar arquivo.',
        toastRegisterSuccess: 'Pedido registrado com sucesso!',
        toastRegisterError: 'Erro: Pedido não encontrado.',
        toastCancelSuccess: 'Registro cancelado com sucesso.',
        toastNoDataExport: 'Não há dados para exportar.',
        toastExportCsvError: 'Erro ao exportar CSV.',
        toastExportCsvSuccess: 'Dados exportados para CSV!',
        toastExportPdfError: 'Erro ao gerar PDF.',
        statusRegistered: (date: string) => `DI Registrada em: ${date}`,
        statusIntegral: (date: string) => `Reg. Integral em: ${date}`,
        statusPending: 'Pendente',
        noPO: 'Sem PO',
        noProject: 'Sem Projeto',
        notAvailable: 'N/A',
        noQuota: 'SEM COTA',
        vehicles: 'Veículos',
        registerDI: 'Registrar DI',
        registerIntegral: 'Reg. Integral',
        cancelRegister: 'Cancelar Registro',
        chartUsed: 'Utilizado',
        chartBalance: 'Saldo',
        chartEVQuota: 'Quota EV',
        chartPHEVQuota: 'Quota PHEV',
        loadingProcess: 'Processando...',
        loadingGenerate: 'Gerando...',
        loadingExport: 'Exportando...',
        lastUpdate: (sheetName: string, date: string) => `Dados de "${sheetName}" | Sincronizado em: ${date}`,
        noItemsFound: 'Nenhum item encontrado.',
    },
    'zh-CN': {
        // ... (traduções em Chinês)
    }
};

// --- UI ELEMENTS MAPPING ---
const UIElements = {
    fileUpload: document.getElementById('file-upload') as HTMLInputElement,
    exportPdfBtn: document.getElementById('export-pdf-btn') as HTMLButtonElement,
    exportCsvBtn: document.getElementById('export-csv-btn') as HTMLButtonElement,
    dashboardContainer: document.getElementById('dashboard-container') as HTMLDivElement,
    lastUpdate: document.getElementById('last-update') as HTMLParagraphElement,
    kpiContainer: document.getElementById('kpi-container') as HTMLDivElement,
    totalEv: document.getElementById('total-ev') as HTMLParagraphElement,
    usedEv: document.getElementById('used-ev') as HTMLParagraphElement,
    pendingUseEv: document.getElementById('pending-use-ev') as HTMLParagraphElement,
    balanceEv: document.getElementById('balance-ev') as HTMLParagraphElement,
    totalPhev: document.getElementById('total-phev') as HTMLParagraphElement,
    usedPhev: document.getElementById('used-phev') as HTMLParagraphElement,
    pendingUsePhev: document.getElementById('pending-use-phev') as HTMLParagraphElement,
    balancePhev: document.getElementById('balance-phev') as HTMLParagraphElement,
    usedVehiclesEv: document.getElementById('used-vehicles-ev') as HTMLParagraphElement,
    pendingUseVehiclesEv: document.getElementById('pending-use-vehicles-ev') as HTMLParagraphElement,
    usedVehiclesPhev: document.getElementById('used-vehicles-phev') as HTMLParagraphElement,
    pendingUseVehiclesPhev: document.getElementById('pending-use-vehicles-phev') as HTMLParagraphElement,
    totalRequests: document.getElementById('total-requests') as HTMLParagraphElement,
    usedRequests: document.getElementById('used-requests') as HTMLParagraphElement,
    pendingRequests: document.getElementById('pending-requests') as HTMLParagraphElement,
    integralRequests: document.getElementById('integral-requests') as HTMLParagraphElement,
    integralValue: document.getElementById('integral-value') as HTMLParagraphElement,
    integralVehicles: document.getElementById('integral-vehicles') as HTMLParagraphElement,
    dashboardContent: document.getElementById('dashboard-content') as HTMLElement,
    liSearchInput: document.getElementById('li-search-input') as HTMLInputElement,
    pendingList: document.getElementById('pending-list') as HTMLDivElement,
    usedList: document.getElementById('used-list') as HTMLDivElement,
    integralList: document.getElementById('integral-list') as HTMLDivElement,
    placeholder: document.getElementById('placeholder') as HTMLDivElement,
    toastContainer: document.getElementById('toast-container') as HTMLDivElement,
    chartsContainer: document.getElementById('charts-container') as HTMLDivElement,
    evChartCanvas: document.getElementById('ev-chart') as HTMLCanvasElement,
    phevChartCanvas: document.getElementById('phev-chart') as HTMLCanvasElement,
    langPtBtn: document.getElementById('lang-pt-btn') as HTMLButtonElement,
    langZhBtn: document.getElementById('lang-zh-btn') as HTMLButtonElement,
};

// --- CONSTANTS ---
const QUOTAS = {
    EV: 54965275.00,
    PHEV: 114838640.00
};

// --- APP STATE ---
let originalData: QuotaData[] = [];
let evChart: any = null;
let phevChart: any = null;
let currentLanguage: 'pt-BR' | 'zh-CN' = 'pt-BR';
let currentSheetInfo: { name: string, date: string } | null = null;

Chart.register(ChartDataLabels);

// --- FUNÇÕES DO FIREBASE ---
const salvarDadosNoFirebase = async (dataToSave: { data: QuotaData[], sheetInfo: { name: string, date: string } | null }) => {
    try {
        await db.collection("quoteControl").doc("latestSheet").set(dataToSave);
        console.log("Dados salvos no Firebase em segundo plano.");
    } catch (e) {
        console.error("Erro ao salvar dados no Firebase: ", e);
    }
};

const escutarMudancasEmTempoReal = () => {
    db.collection("quoteControl").doc("latestSheet").onSnapshot((doc: any) => {
        if (doc.exists) {
            const firestoreData = doc.data();
            originalData = firestoreData.data || [];
            currentSheetInfo = firestoreData.sheetInfo || null;

            processAndRenderAll(originalData);
            
            UIElements.kpiContainer.classList.remove('hidden');
            UIElements.dashboardContent.classList.remove('hidden');
            UIElements.chartsContainer.classList.remove('hidden');
            UIElements.placeholder.classList.add('hidden');
            
            if (currentSheetInfo) {
                const date = new Date(currentSheetInfo.date);
                UIElements.lastUpdate.textContent = translations[currentLanguage].lastUpdate(currentSheetInfo.name, date.toLocaleString(currentLanguage));
            }
        } else {
            resetUI();
        }
    });
};


// --- LANGUAGE & FORMATTING FUNCTIONS ---
function setLanguage(lang: 'pt-BR' | 'zh-CN') {
    currentLanguage = lang;
    
    UIElements.langPtBtn.classList.toggle('bg-blue-600', lang === 'pt-BR');
    UIElements.langPtBtn.classList.toggle('text-white', lang === 'pt-BR');
    UIElements.langZhBtn.classList.toggle('bg-blue-600', lang === 'zh-CN');
    UIElements.langZhBtn.classList.toggle('text-white', lang === 'zh-CN');

    const t = translations[currentLanguage];
    document.querySelectorAll('[data-lang-key]').forEach(el => {
        const key = el.getAttribute('data-lang-key') as keyof typeof t;
        const element = el as HTMLElement;
        if (key && t[key] && typeof t[key] === 'string') {
            if (element.tagName === 'INPUT') {
                (element as HTMLInputElement).placeholder = t[key] as string;
            } else {
                element.textContent = t[key] as string;
            }
        }
    });

    document.title = t.pageTitle;

    if (originalData.length > 0) {
        processAndRenderAll(originalData);
        if (currentSheetInfo) {
            const date = new Date(currentSheetInfo.date);
            UIElements.lastUpdate.textContent = t.lastUpdate(currentSheetInfo.name, date.toLocaleString(currentLanguage));
        }
    }
}

function showToast(messageKey: keyof typeof translations['pt-BR'], type: 'success' | 'error' = 'success') {
    const message = translations[currentLanguage][messageKey] as string;
    const toast = document.createElement('div');
    toast.className = `toast p-4 rounded-lg shadow-lg text-white ${type === 'success' ? 'bg-green-500' : 'bg-red-500'}`;
    toast.textContent = message;
    UIElements.toastContainer.appendChild(toast);
    setTimeout(() => { toast.remove(); }, 5000);
}

function parseCurrency(value: string | number | null): number {
    if (typeof value === 'number') return value;
    if (typeof value !== 'string' || !value) return 0;
    const cleanedValue = String(value).replace(/[^0-9,.]/g, '');
    const lastComma = cleanedValue.lastIndexOf(',');
    const lastDot = cleanedValue.lastIndexOf('.');
    if (lastComma > lastDot) return parseFloat(cleanedValue.replace(/\./g, '').replace(',', '.')) || 0;
    return parseFloat(cleanedValue.replace(/,/g, '')) || 0;
}

function formatCurrency(value: number): string {
    const locale = currentLanguage === 'zh-CN' ? 'en-US' : currentLanguage;
    return new Intl.NumberFormat(locale, { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
}

function formatNumber(value: number): string {
    return new Intl.NumberFormat(currentLanguage).format(value);
}


// --- UI RENDERING FUNCTIONS ---
function resetUI() {
    UIElements.kpiContainer.classList.add('hidden');
    UIElements.dashboardContent.classList.add('hidden');
    UIElements.chartsContainer.classList.add('hidden');
    UIElements.placeholder.classList.remove('hidden');
    originalData = [];
    currentSheetInfo = null;
    UIElements.liSearchInput.value = '';
    UIElements.lastUpdate.textContent = translations[currentLanguage].promptToUpload;
}

function processAndRenderAll(data: QuotaData[]) {
    // ... (Esta função permanece a mesma da versão anterior)
}

function renderList(container: HTMLElement, items: QuotaData[], isUsed: boolean) {
    // ... (Esta função permanece a mesma da versão anterior)
}

function updateCharts(usedEv: number, balanceEv: number, usedPhev: number, balancePhev: number) {
    // ... (Esta função permanece a mesma da versão anterior)
}

function filterAndRenderLists() {
    // ... (Esta função permanece a mesma da versão anterior)
}


// --- ACTION HANDLERS & EVENT LISTENERS ---
function handleRegister(id: number, type: 'QUOTA' | 'INTEGRAL') {
    const itemIndex = originalData.findIndex(item => item.__id === id);
    if (itemIndex > -1) {
        originalData[itemIndex]['DATA REGISTRO DI'] = new Date().toLocaleDateString(currentLanguage);
        originalData[itemIndex]['REGISTRATION_TYPE'] = type;
        
        // ATUALIZA A TELA IMEDIATAMENTE
        processAndRenderAll(originalData); 
        // SALVA NO FIREBASE EM SEGUNDO PLANO
        salvarDadosNoFirebase({ data: originalData, sheetInfo: currentSheetInfo });
        
        showToast('toastRegisterSuccess', 'success');
    }
}
function handleCancelDI(id: number) {
    const itemIndex = originalData.findIndex(item => item.__id === id);
    if (itemIndex > -1) {
        originalData[itemIndex]['DATA REGISTRO DI'] = null;
        originalData[itemIndex]['REGISTRATION_TYPE'] = null;

        // ATUALIZA A TELA IMEDIATAMENTE
        processAndRenderAll(originalData);
        // SALVA NO FIREBASE EM SEGUNDO PLANO
        salvarDadosNoFirebase({ data: originalData, sheetInfo: currentSheetInfo });

        showToast('toastCancelSuccess', 'success');
    }
}
UIElements.fileUpload.addEventListener('change', (event) => {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;
    const reader = new FileReader();
    const uploadLabelElement = document.querySelector('label[for="file-upload"]')!;
    const originalHTML = uploadLabelElement.innerHTML;
    uploadLabelElement.innerHTML = `<i class="fas fa-spinner fa-spin mr-2"></i> ${translations[currentLanguage].loadingProcess}`;
    (uploadLabelElement as HTMLLabelElement).style.pointerEvents = 'none';

    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames.find(name => name.toUpperCase().includes('SHEET1') || name.toUpperCase().includes('PLANILHA1')) || workbook.SheetNames[0];
            if (!sheetName) { throw new Error("No sheets found"); }
            const jsonData: any[] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false, defval: null });
            if (jsonData.length === 0) { const err = new Error("Sheet is empty"); err.name = 'toastEmptySheet'; throw err; }
            
            originalData = jsonData.map((row, index) => ({ ...row, __id: index, REGISTRATION_TYPE: null }));
            currentSheetInfo = { name: sheetName, date: new Date().toISOString() };
            
            // ATUALIZA A TELA IMEDIATAMENTE
            processAndRenderAll(originalData);
            UIElements.kpiContainer.classList.remove('hidden');
            UIElements.dashboardContent.classList.remove('hidden');
            UIElements.chartsContainer.classList.remove('hidden');
            UIElements.placeholder.classList.add('hidden');
            showToast('toastLoaded', 'success');
            
            // SALVA NO FIREBASE EM SEGUNDO PLANO
            salvarDadosNoFirebase({ data: originalData, sheetInfo: currentSheetInfo });

        } catch (err: any) {
            showToast(err.name === 'toastEmptySheet' ? 'toastEmptySheet' : 'toastProcessError', 'error');
            resetUI();
        } finally {
            uploadLabelElement.innerHTML = originalHTML;
            (uploadLabelElement as HTMLLabelElement).style.pointerEvents = 'auto';
            (event.target as HTMLInputElement).value = '';
        }
    };
    reader.readAsArrayBuffer(file);
});

// ... (O restante dos listeners permanece igual)

// --- INICIALIZAÇÃO ---
document.addEventListener('DOMContentLoaded', () => {
    setLanguage('pt-BR');
    escutarMudancasEmTempoReal();
});
