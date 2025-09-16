// Make libraries available globally from CDN
declare var XLSX: any;
declare var jspdf: any;
declare var html2canvas: any;

// Resolve ReferenceError for CDN libraries in a module script by assigning them to constants from the window object.
const Chart = (window as any).Chart;
const ChartDataLabels = (window as any).ChartDataLabels;


// --- TYPE DEFINITIONS ---
interface QuotaData {
    __id: number; // Internal unique ID for tracking clicks
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
        totalVehicles: 'Total (Veículos)',
        usedUSD: 'Utilizado (USD)',
        usedVehicles: 'Utilizado (Veículos)',
        pendingToUseUSD: 'Pendente de Uso (USD)',
        pendingToUseVehicles: 'Pendente de Uso (Veículos)',
        balanceUSD: 'SALDO (USD)',
        balanceVehicles: 'SALDO (Veículos)',
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
        lastUpdate: (sheetName: string, date: string) => `Dados de "${sheetName}" | Carregado em: ${date}`,
        noItemsFound: 'Nenhum item encontrado.',
        vehicleAlertTooltip: (count: number) => `Atenção: ${count} pedido(s) na planilha está(ão) com 0 veículos. Os totais de veículos podem estar incorretos.`
    },
    'zh-CN': {
        pageTitle: '进口配额控制面板',
        headerTitle: '配额控制面板',
        promptToUpload: '请上传优先注册电子表格以开始',
        exportPDF: '导出 PDF',
        exportCSV: '导出 CSV',
        uploadSheet: '上传文件',
        quotaEV: '电动汽车配额',
        quotaPHEV: '插电混动车配额',
        totalUSD: '总计 (美元)',
        totalVehicles: '总计 (车辆)',
        usedUSD: '已用 (美元)',
        usedVehicles: '已用 (车辆)',
        pendingToUseUSD: '待使用 (美元)',
        pendingToUseVehicles: '待使用 (车辆)',
        balanceUSD: '余额 (美元)',
        balanceVehicles: '余额 (车辆)',
        summary: '概览',
        totalOrders: '表格订单总数',
        usedOrders: '已用订单 (DI 已注册)',
        pendingOrders: '待处理订单',
        integralSummary: '整体注册摘要',
        integralOrders: '整体订单总数',
        integralTotalValue: '总价值 (美元)',
        integralTotalVehicles: '整体车辆总数',
        quotaVisualization: '配额可视化 (美元)',
        filterByLIPO: '按 LI 或 PO 编号筛选...',
        pendingOrdersList: '待处理订单',
        usedOrdersList: '已用订单 (DI 已注册)',
        integralOrdersList: '整体注册订单',
        waitingForFile: '等待文件...',
        selectFileToStart: '请选择优先注册电子表格以开始。',
        toastLoaded: '配额面板已加载！',
        toastNoSheet: '未找到有效的表格（例如 "Sheet1" 或 "Planilha1"）。',
        toastEmptySheet: '电子表格为空。',
        toastProcessError: '处理文件时出错。',
        toastRegisterSuccess: '订单注册成功！',
        toastRegisterError: '错误：未找到订单。',
        toastCancelSuccess: '注册已成功取消。',
        toastNoDataExport: '无数据可导出。',
        toastExportCsvError: '导出 CSV 时出错。',
        toastExportCsvSuccess: '数据已导出为 CSV！',
        toastExportPdfError: '生成 PDF 时出错。',
        statusRegistered: (date: string) => `DI 注册于：${date}`,
        statusIntegral: (date: string) => `整体注册于: ${date}`,
        statusPending: '待处理',
        noPO: '无采购订单',
        noProject: '无项目',
        notAvailable: '不适用',
        noQuota: '无配额',
        vehicles: '辆车',
        registerDI: '注册 DI',
        registerIntegral: '整体注册',
        cancelRegister: '取消注册',
        chartUsed: '已用',
        chartBalance: '余额',
        chartEVQuota: '电动汽车配额',
        chartPHEVQuota: '插电混动车配额',
        loadingProcess: '处理中...',
        loadingGenerate: '生成中...',
        loadingExport: '导出中...',
        lastUpdate: (sheetName: string, date: string) => `数据来源："${sheetName}" | 加载于：${date}`,
        noItemsFound: '未找到任何项目。',
        vehicleAlertTooltip: (count: number) => `注意：电子表格中有 ${count} 个订单的车辆数量为 0。车辆总数可能不正确。`
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
    totalVehiclesEv: document.getElementById('total-vehicles-ev') as HTMLParagraphElement,
    usedVehiclesEv: document.getElementById('used-vehicles-ev') as HTMLParagraphElement,
    pendingUseVehiclesEv: document.getElementById('pending-use-vehicles-ev') as HTMLParagraphElement,
    totalVehiclesPhev: document.getElementById('total-vehicles-phev') as HTMLParagraphElement,
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
    evVehicleAlert: document.getElementById('ev-vehicle-alert') as HTMLDivElement,
    phevVehicleAlert: document.getElementById('phev-vehicle-alert') as HTMLDivElement,
    evVehicleTooltip: document.getElementById('ev-vehicle-tooltip') as HTMLDivElement,
    phevVehicleTooltip: document.getElementById('phev-vehicle-tooltip') as HTMLDivElement,
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
let currentSheetInfo: { name: string, date: Date } | null = null;

// Register ChartJS plugins
Chart.register(ChartDataLabels);

// --- LANGUAGE & FORMATTING FUNCTIONS ---

function setLanguage(lang: 'pt-BR' | 'zh-CN') {
    currentLanguage = lang;
    
    // Update button styles
    UIElements.langPtBtn.classList.toggle('bg-blue-600', lang === 'pt-BR');
    UIElements.langPtBtn.classList.toggle('text-white', lang === 'pt-BR');
    UIElements.langPtBtn.classList.toggle('ring-2', lang === 'pt-BR');
    UIElements.langPtBtn.classList.toggle('ring-blue-700', lang === 'pt-BR');
    UIElements.langPtBtn.classList.toggle('bg-gray-200', lang !== 'pt-BR');
    UIElements.langPtBtn.classList.toggle('text-gray-700', lang !== 'pt-BR');

    UIElements.langZhBtn.classList.toggle('bg-blue-600', lang === 'zh-CN');
    UIElements.langZhBtn.classList.toggle('text-white', lang === 'zh-CN');
    UIElements.langZhBtn.classList.toggle('ring-2', lang === 'zh-CN');
    UIElements.langZhBtn.classList.toggle('ring-blue-700', lang === 'zh-CN');
    UIElements.langZhBtn.classList.toggle('bg-gray-200', lang !== 'zh-CN');
    UIElements.langZhBtn.classList.toggle('text-gray-700', lang !== 'zh-CN');

    // Update static text from data-lang-key attributes
    const t = translations[currentLanguage];
    document.querySelectorAll('[data-lang-key]').forEach(el => {
        const key = el.getAttribute('data-lang-key') as keyof typeof t;
        if (key && t[key]) {
            const translation = t[key];
            if (typeof translation === 'string') {
                 if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
                    el.placeholder = translation;
                } else {
                    el.textContent = translation;
                }
            }
        }
    });

    document.title = t.pageTitle;

    // Re-render dynamic content if data exists
    if (originalData.length > 0) {
        processAndRenderAll(originalData);
        if (currentSheetInfo) {
            UIElements.lastUpdate.textContent = t.lastUpdate(currentSheetInfo.name, currentSheetInfo.date.toLocaleString(currentLanguage));
        }
    }
}

function showToast(messageKey: keyof typeof translations['pt-BR'], type: 'success' | 'error' = 'success') {
    const message = translations[currentLanguage][messageKey] as string;
    const toast = document.createElement('div');
    toast.className = `toast p-4 rounded-lg shadow-lg text-white ${type === 'success' ? 'bg-green-500' : 'bg-red-500'}`;
    toast.textContent = message;
    UIElements.toastContainer.appendChild(toast);
    setTimeout(() => {
        toast.remove();
    }, 5000);
}

function parseCurrency(value: string | number | null): number {
    if (typeof value === 'number') return value;
    if (typeof value !== 'string' || !value) return 0;

    const cleanedValue = String(value).replace(/[^0-9,.]/g, '');
    const lastComma = cleanedValue.lastIndexOf(',');
    const lastDot = cleanedValue.lastIndexOf('.');

    if (lastComma > lastDot) {
        return parseFloat(cleanedValue.replace(/\./g, '').replace(',', '.')) || 0;
    }
    
    return parseFloat(cleanedValue.replace(/,/g, '')) || 0;
}

function formatCurrency(value: number): string {
    const locale = currentLanguage === 'zh-CN' ? 'en-US' : currentLanguage;
    return new Intl.NumberFormat(locale, { 
        style: 'currency', 
        currency: 'USD', 
        minimumFractionDigits: 2, 
        maximumFractionDigits: 2 
    }).format(value);
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

function renderList(container: HTMLElement, items: QuotaData[], isUsed: boolean) {
    container.innerHTML = '';
    const t = translations[currentLanguage];

    if (items.length === 0) {
        container.innerHTML = `<p class="text-center text-gray-500 p-4">${t.noItemsFound}</p>`;
        return;
    }
    items.forEach(item => {
        const quoteType = (item['QUOTE'] || '').toUpperCase();
        const borderColor = quoteType === 'EV' ? 'border-green-500' : 'border-blue-500';
        const card = document.createElement('div');
        card.className = `item-card p-3 ${borderColor}`;

        const poNumber = item['PO '] || item['PO'] || t.noPO;
        const project = item['PROJECT'] || t.noProject;
        const liNumber = item['LI NUMBER'] || t.notAvailable;
        const registroDI = item['DATA REGISTRO DI'];

        let status;
        if (isUsed && registroDI) {
            status = item.REGISTRATION_TYPE === 'INTEGRAL'
                ? t.statusIntegral(registroDI)
                : t.statusRegistered(registroDI);
        } else {
            status = t.statusPending;
        }

        const value = parseCurrency(item['VALOR USD']);
        const vehicleCount = parseInt(String(item['QTD VEÍCULOS'] || 0));

        let actionButton = '';
        if (!isUsed && (quoteType === 'EV' || quoteType === 'PHEV')) {
            actionButton = `
                <div class="mt-2 text-right flex items-center justify-end space-x-2">
                     <button class="integral-di-btn text-xs bg-gray-500 hover:bg-gray-600 text-white font-bold py-1 px-3 rounded-full transition-transform transform hover:scale-105" data-id="${item.__id}">
                        ${t.registerIntegral}
                    </button>
                    <button class="register-di-btn text-xs bg-green-500 hover:bg-green-600 text-white font-bold py-1 px-3 rounded-full transition-transform transform hover:scale-105" data-id="${item.__id}">
                        <i class="fas fa-check mr-1"></i> ${t.registerDI}
                    </button>
                </div>
            `;
        } else if (isUsed) {
            actionButton = `
                <div class="mt-2 text-right">
                    <button class="cancel-di-btn text-xs bg-red-500 hover:bg-red-600 text-white font-bold py-1 px-3 rounded-full transition-transform transform hover:scale-105" data-id="${item.__id}">
                        <i class="fas fa-times mr-1"></i> ${t.cancelRegister}
                    </button>
                </div>
            `;
        }


        card.innerHTML = `
            <div class="flex justify-between items-start">
                <div>
                    <p class="text-xs font-bold text-gray-500">${poNumber}</p>
                    <p class="font-semibold text-gray-800">${project}</p>
                    <p class="text-xs text-gray-600 mt-1"><b>LI:</b> ${liNumber}</p>
                    <p class="text-xs text-gray-600"><b>Status:</b> ${status}</p>
                </div>
                <div class="text-right flex-shrink-0 ml-2">
                    <p class="text-lg font-bold ${isUsed ? 'text-red-600' : 'text-gray-700'}">${formatCurrency(value)}</p>
                    <div class="mt-1">
                        <span class="text-xs font-semibold px-2 py-1 rounded-full ${borderColor.replace('border', 'bg').replace('-500', '-100')} ${borderColor.replace('border', 'text')}">${quoteType || t.noQuota}</span>
                        <span class="text-xs font-semibold px-2 py-1 rounded-full bg-gray-100 text-gray-700 ml-1">${vehicleCount} ${t.vehicles}</span>
                    </div>
                </div>
            </div>
            ${actionButton}
        `;
        container.appendChild(card);
    });
}

function updateCharts(usedEv: number, balanceEv: number, usedPhev: number, balancePhev: number) {
    const t = translations[currentLanguage];
    const chartOptions = (total: number) => ({
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                position: 'bottom' as const,
            },
            tooltip: {
                callbacks: {
                    label: function(context: any) {
                        return `${context.label}: ${formatCurrency(context.parsed)}`;
                    }
                }
            },
            datalabels: {
                formatter: (value: number) => {
                    if (total === 0) return '0%';
                    const percentage = (value / total) * 100;
                    return percentage > 5 ? `${percentage.toFixed(1)}%` : '';
                },
                color: '#fff',
                font: {
                    weight: 'bold' as const,
                }
            }
        }
    });

    if (evChart) evChart.destroy();
    evChart = new Chart(UIElements.evChartCanvas, {
        type: 'doughnut',
        data: {
            labels: [t.chartUsed, t.chartBalance],
            datasets: [{
                label: t.chartEVQuota,
                data: [usedEv, balanceEv],
                backgroundColor: ['#EF4444', '#22C55E'],
            }]
        },
        options: chartOptions(QUOTAS.EV)
    });

    if (phevChart) phevChart.destroy();
    phevChart = new Chart(UIElements.phevChartCanvas, {
        type: 'doughnut',
        data: {
            labels: [t.chartUsed, t.chartBalance],
            datasets: [{
                label: t.chartPHEVQuota,
                data: [usedPhev, balancePhev],
                backgroundColor: ['#EF4444', '#3B82F6'],
            }]
        },
        options: chartOptions(QUOTAS.PHEV)
    });
}

function filterAndRenderLists() {
    const searchTerm = UIElements.liSearchInput.value.toLowerCase().trim();

    let pendingList: QuotaData[] = [];
    let usedList: QuotaData[] = [];
    let integralList: QuotaData[] = [];

    originalData.forEach(item => {
        const isUsed = item['DATA REGISTRO DI'] && String(item['DATA REGISTRO DI']).trim() !== '';
        if (isUsed) {
            if (item.REGISTRATION_TYPE === 'INTEGRAL') {
                integralList.push(item);
            } else {
                usedList.push(item);
            }
        } else {
            pendingList.push(item);
        }
    });

    if (searchTerm) {
        const filterFn = (item: QuotaData) => {
            const poNumber = String(item['PO '] || item['PO'] || '').toLowerCase();
            const liNumber = String(item['LI NUMBER'] || '').toLowerCase();
            return liNumber.includes(searchTerm) || poNumber.includes(searchTerm);
        };
        pendingList = pendingList.filter(filterFn);
        usedList = usedList.filter(filterFn);
        integralList = integralList.filter(filterFn);
    }
    
    renderList(UIElements.pendingList, pendingList, false);
    renderList(UIElements.usedList, usedList, true);
    renderList(UIElements.integralList, integralList, true);
}


function processAndRenderAll(data: QuotaData[]) {
    let usedEv = 0, usedPhev = 0;
    let pendingEv = 0, pendingPhev = 0;
    let totalVehiclesEv = 0, usedVehiclesEv = 0, pendingVehiclesEv = 0;
    let totalVehiclesPhev = 0, usedVehiclesPhev = 0, pendingVehiclesPhev = 0;
    let zeroVehicleCount = 0;
    
    let usedCount = 0;
    let integralCount = 0;
    let integralValue = 0;
    let integralVehicles = 0;

    data.forEach(row => {
        const registroDI = row['DATA REGISTRO DI'];
        const isUsed = registroDI !== null && registroDI !== undefined && String(registroDI).trim() !== '';
        
        const quoteType = (row['QUOTE'] || '').toUpperCase();
        const value = parseCurrency(row['VALOR USD']);
        const vehicleCount = parseInt(String(row['QTD VEÍCULOS'] || '0'));

        if(vehicleCount === 0) {
            zeroVehicleCount++;
        }

        if (quoteType === 'EV') {
            totalVehiclesEv += vehicleCount;
        } else if (quoteType === 'PHEV') {
            totalVehiclesPhev += vehicleCount;
        }
        
        if (isUsed) {
            if (row.REGISTRATION_TYPE === 'QUOTA') {
                usedCount++;
                if (quoteType === 'EV') {
                    usedEv += value;
                    usedVehiclesEv += vehicleCount;
                } else if (quoteType === 'PHEV') {
                    usedPhev += value;
                    usedVehiclesPhev += vehicleCount;
                }
            } else if (row.REGISTRATION_TYPE === 'INTEGRAL') {
                integralCount++;
                integralValue += value;
                integralVehicles += vehicleCount;
            }
        } else { // Is Pending
             if (quoteType === 'EV') {
                pendingEv += value;
                pendingVehiclesEv += vehicleCount;
            } else if (quoteType === 'PHEV') {
                pendingPhev += value;
                pendingVehiclesPhev += vehicleCount;
            }
        }
    });

    // Update KPIs
    UIElements.totalEv.textContent = formatCurrency(QUOTAS.EV);
    UIElements.usedEv.textContent = formatCurrency(usedEv);
    UIElements.pendingUseEv.textContent = formatCurrency(pendingEv);
    const balanceEv = QUOTAS.EV - usedEv;
    UIElements.balanceEv.textContent = formatCurrency(balanceEv);
    
    UIElements.totalPhev.textContent = formatCurrency(QUOTAS.PHEV);
    UIElements.usedPhev.textContent = formatCurrency(usedPhev);
    UIElements.pendingUsePhev.textContent = formatCurrency(pendingPhev);
    const balancePhev = QUOTAS.PHEV - usedPhev;
    UIElements.balancePhev.textContent = formatCurrency(balancePhev);

    UIElements.totalVehiclesEv.textContent = formatNumber(totalVehiclesEv);
    UIElements.usedVehiclesEv.textContent = formatNumber(usedVehiclesEv);
    UIElements.pendingUseVehiclesEv.textContent = formatNumber(pendingVehiclesEv);
    
    UIElements.totalVehiclesPhev.textContent = formatNumber(totalVehiclesPhev);
    UIElements.usedVehiclesPhev.textContent = formatNumber(usedVehiclesPhev);
    UIElements.pendingUseVehiclesPhev.textContent = formatNumber(pendingVehiclesPhev);

    const pendingCount = data.length - usedCount - integralCount;
    UIElements.totalRequests.textContent = data.length.toString();
    UIElements.usedRequests.textContent = usedCount.toString();
    UIElements.pendingRequests.textContent = pendingCount.toString();
    
    UIElements.integralRequests.textContent = integralCount.toString();
    UIElements.integralValue.textContent = formatCurrency(integralValue);
    UIElements.integralVehicles.textContent = formatNumber(integralVehicles);

    // Update Vehicle Alert
    const t = translations[currentLanguage];
    const alertElements = [UIElements.evVehicleAlert, UIElements.phevVehicleAlert];
    const tooltipElements = [UIElements.evVehicleTooltip, UIElements.phevVehicleTooltip];
    if (zeroVehicleCount > 0) {
        const tooltipText = t.vehicleAlertTooltip(zeroVehicleCount);
        alertElements.forEach(el => el.classList.remove('hidden'));
        tooltipElements.forEach(el => el.textContent = tooltipText);
    } else {
        alertElements.forEach(el => el.classList.add('hidden'));
    }

    filterAndRenderLists();
    updateCharts(usedEv, balanceEv, usedPhev, balancePhev);
}

// --- ACTION HANDLERS ---
function handleRegister(id: number, type: 'QUOTA' | 'INTEGRAL') {
     const itemIndex = originalData.findIndex(item => item.__id === id);
    if (itemIndex > -1) {
        originalData[itemIndex]['DATA REGISTRO DI'] = new Date().toLocaleDateString(currentLanguage);
        originalData[itemIndex]['REGISTRATION_TYPE'] = type;
        processAndRenderAll(originalData);
        showToast('toastRegisterSuccess', 'success');
    } else {
        showToast('toastRegisterError', 'error');
    }
}

function handleCancelDI(id: number) {
    const itemIndex = originalData.findIndex(item => item.__id === id);
    if (itemIndex > -1) {
        originalData[itemIndex]['DATA REGISTRO DI'] = null;
        originalData[itemIndex]['REGISTRATION_TYPE'] = null;
        processAndRenderAll(originalData);
        showToast('toastCancelSuccess', 'success');
    } else {
        showToast('toastRegisterError', 'error');
    }
}

function handleExportCSV() {
    if (originalData.length === 0) {
        showToast('toastNoDataExport', 'error');
        return;
    }

    const btn = UIElements.exportCsvBtn;
    const originalText = btn.querySelector('span')!.textContent;
    btn.querySelector('span')!.textContent = translations[currentLanguage].loadingExport;
    btn.disabled = true;

    try {
        const headerSet = new Set<string>();
        originalData.forEach(row => {
            Object.keys(row).forEach(key => headerSet.add(key));
        });
        
        headerSet.delete('__id');
        const headers = Array.from(headerSet);
        let csvContent = headers.join(',') + '\n';

        originalData.forEach(row => {
            const values = headers.map(header => {
                let value = row[header];
                if (value === null || value === undefined) return '';
                let stringValue = String(value).replace(/"/g, '""');
                if (stringValue.includes(',')) stringValue = `"${stringValue}"`;
                return stringValue;
            });
            csvContent += values.join(',') + '\n';
        });

        const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' }); // Add BOM for Excel
        const link = document.createElement("a");
        const url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.setAttribute("download", "dashboard-quotas-export.csv");
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showToast('toastExportCsvSuccess', 'success');

    } catch (err: any) {
        showToast('toastExportCsvError', 'error');
        console.error("CSV Export Error:", err);
    } finally {
        btn.querySelector('span')!.textContent = originalText;
        btn.disabled = false;
    }
}

// --- EVENT LISTENERS ---

UIElements.fileUpload.addEventListener('change', (event) => {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;

    const uploadLabelElement = document.querySelector('label[for="file-upload"]');
    if (!uploadLabelElement) return;

    const originalHTML = uploadLabelElement.innerHTML;
    
    // Set loading state
    uploadLabelElement.innerHTML = `<i class="fas fa-spinner fa-spin mr-2"></i> ${translations[currentLanguage].loadingProcess}`;
    (uploadLabelElement as HTMLLabelElement).style.pointerEvents = 'none';

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames.find((name: string) => 
                name.toUpperCase().includes('SHEET1') || name.toUpperCase().includes('PLANILHA1')
            );

            if (!sheetName) {
                const err = new Error("Sheet not found");
                err.name = 'toastNoSheet';
                throw err;
            }
            
            const jsonData: Omit<QuotaData, '__id' | 'REGISTRATION_TYPE'>[] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false, defval: null });
            
            if (jsonData.length === 0) {
                const err = new Error("Sheet is empty");
                err.name = 'toastEmptySheet';
                throw err;
            }

            originalData = jsonData.map((row, index) => ({ 
                ...row, 
                __id: index, 
                REGISTRATION_TYPE: null // Initialize new property
            } as QuotaData));
            currentSheetInfo = { name: sheetName, date: new Date() };

            processAndRenderAll(originalData);
            
            UIElements.kpiContainer.classList.remove('hidden');
            UIElements.dashboardContent.classList.remove('hidden');
            UIElements.chartsContainer.classList.remove('hidden');
            UIElements.placeholder.classList.add('hidden');
            UIElements.lastUpdate.textContent = translations[currentLanguage].lastUpdate(currentSheetInfo.name, currentSheetInfo.date.toLocaleString(currentLanguage));
            showToast('toastLoaded', 'success');

        } catch (err: any) {
            if (err.name === 'toastNoSheet' || err.name === 'toastEmptySheet') {
                showToast(err.name as keyof typeof translations['pt-BR'], 'error');
            } else {
                showToast('toastProcessError', 'error');
            }
            console.error("File processing error:", err);
            resetUI();
        } finally {
            // Restore original state
            uploadLabelElement.innerHTML = originalHTML;
            (uploadLabelElement as HTMLLabelElement).style.pointerEvents = 'auto';
            (event.target as HTMLInputElement).value = '';
        }
    };
    reader.readAsArrayBuffer(file);
});


UIElements.liSearchInput.addEventListener('input', filterAndRenderLists);

const listClickListener = (event: MouseEvent) => {
    const target = event.target as HTMLElement;
    const registerBtn = target.closest('.register-di-btn');
    const integralBtn = target.closest('.integral-di-btn');
    const cancelBtn = target.closest('.cancel-di-btn');

    if (registerBtn instanceof HTMLElement && registerBtn.dataset.id) {
        handleRegister(parseInt(registerBtn.dataset.id, 10), 'QUOTA');
    } else if (integralBtn instanceof HTMLElement && integralBtn.dataset.id) {
        handleRegister(parseInt(integralBtn.dataset.id, 10), 'INTEGRAL');
    } else if (cancelBtn instanceof HTMLElement && cancelBtn.dataset.id) {
        handleCancelDI(parseInt(cancelBtn.dataset.id, 10));
    }
};

UIElements.pendingList.addEventListener('click', listClickListener);
UIElements.usedList.addEventListener('click', listClickListener);
UIElements.integralList.addEventListener('click', listClickListener);


UIElements.exportPdfBtn.addEventListener('click', () => {
    const btn = UIElements.exportPdfBtn;
    const originalText = btn.querySelector('span')!.textContent;
    btn.querySelector('span')!.textContent = translations[currentLanguage].loadingGenerate;
    btn.disabled = true;

    html2canvas(UIElements.dashboardContainer, { scale: 2 }) // Increased scale for better quality
        .then((canvas: HTMLCanvasElement) => {
            const imgData = canvas.toDataURL('image/png');
            const { jsPDF } = jspdf;
            const pdf = new jsPDF('p', 'mm', 'a4');
            const pdfWidth = pdf.internal.pageSize.getWidth();
            const imgProps = pdf.getImageProperties(imgData);
            const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;
            let position = 0;
            let heightLeft = imgHeight;

            pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, imgHeight);
            heightLeft -= pdf.internal.pageSize.getHeight();

            while (heightLeft > 0) {
                position -= pdf.internal.pageSize.getHeight();
                pdf.addPage();
                pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, imgHeight);
                heightLeft -= pdf.internal.pageSize.getHeight();
            }
            pdf.save('dashboard-quotas.pdf');
        }).catch((err: any) => {
            showToast('toastExportPdfError', 'error');
            console.error("PDF Export Error: ", err);
        }).finally(() => {
            btn.querySelector('span')!.textContent = originalText;
            btn.disabled = false;
        });
});

UIElements.exportCsvBtn.addEventListener('click', handleExportCSV);
UIElements.langPtBtn.addEventListener('click', () => setLanguage('pt-BR'));
UIElements.langZhBtn.addEventListener('click', () => setLanguage('zh-CN'));

// Initial Load
document.addEventListener('DOMContentLoaded', () => {
    setLanguage('pt-BR');
});