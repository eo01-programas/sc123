const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyvLG9IxHRddHMWN12opa0mRq6Zbv7pTyJiWTPEHvtG4j-dCG2hbKbA4TXIDAGPXTqEOQ/exec';
const SHEET_ID = '18cQuwqerdMggAeJ8TCUKA7-gujXsA91-CRMUTNpr8aQ';
const MONTH_ABBR_ES = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];
const ESTADO_COSTURA_OPTIONS = ['', 'Proceso', 'Liquidado', 'Anaquel', 'En Habilitado'];
let allData = [];
const PLANTA_FILTER_OTHERS = '__OTHERS__';
const PLANTA_ORDER = ['COFACO', 'COFACO 2', 'CITI1', 'CITI2', 'CITI3', 'CITI4'];
const PLANTA_TOTAL_FILTERS = ['COFACO', 'COFACO 2', 'CITI1', 'CITI2', 'CITI3', 'CITI4', PLANTA_FILTER_OTHERS];
const LINEA_BAND_CLICK_DELAY_MS = 220;
const SHIFT_PILL_OPTIONS_BY_PLANT = {
    COFACO: ['S1', 'S2', 'TN'],
    CITI1: ['S1', 'S2'],
    CITI2: ['S1', 'S2']
};
const collapsedLineasByFilter = new Map();
const LINEA_MODAL_PLANTA_OPTIONS = [
    { value: 'COFACO', label: 'COFACO' },
    { value: 'COFACO 2', label: 'COFACO 2' },
    { value: 'CITI1', label: 'CITI1' },
    { value: 'CITI2', label: 'CITI2' },
    { value: 'CITI3', label: 'CITI3' },
    { value: 'CITI4', label: 'CITI4' },
    { value: 'S/DESTINO', label: 'S/DESTINO' },
    { value: '', label: 'VACIO' }
];
const OC_SEARCH_PLANT_ORDER = ['COFACO', 'COFACO 2', 'CITI1', 'CITI2', 'CITI3', 'CITI4', PLANTA_FILTER_OTHERS];
let selectedPlantaFilter = 'COFACO';
let activeShifts = [];
let activeOcSearchQuery = '';
let ocSearchState = null;
let ocSearchDataStamp = 0;
let selectedColorFilter = '';
let splitOcContextMenuState = {
    menu: null,
    rowIndex: null
};
let splitOcModalState = null;
const splitOcModalRefs = {
    overlay: null,
    title: null,
    subtitle: null,
    note: null,
    total: null,
    tbody: null,
    btnCancel: null,
    btnSave: null
};
let devolucionOcModalState = null;
const devolucionOcModalRefs = {
    overlay: null,
    title: null,
    subtitle: null,
    tbody: null,
    plantaSelect: null,
    btnCancel: null,
    btnSave: null
};
let lineaModalState = null;
const lineaModalRefs = {
    overlay: null,
    planta: null,
    linea: null,
    btnCancel: null,
    btnSave: null
};
let liquidacionModalState = null;
const liquidacionModalRefs = {
    overlay: null,
    btnNo: null,
    btnYes: null
};
const REORDER_MODAL_DELAY_MS = 280;
let reorderModalState = {
    overlay: null
};
let colorFilterPopoverState = {
    popover: null,
    select: null,
    btnClear: null,
    currentOptions: [],
    anchorRect: null
};
const READ_ONLY_ACCESS_PROFILE = Object.freeze({
    key: 'READ_ONLY',
    label: 'SOLO LECTURA',
    editableFilters: [],
    canEditAll: false
});
let activeAccessProfile = null;

function normalizeFilterForAccess(filterValue) {
    const normalized = normalizePlantFilterValue(filterValue);
    if (normalized === 'OTROS' || normalized === 'OTHER' || normalized === 'OTHERS') {
        return PLANTA_FILTER_OTHERS;
    }
    return normalized;
}

function normalizeAccessProfile(rawProfile) {
    const profile = rawProfile && typeof rawProfile === 'object' ? rawProfile : {};
    const editableFilters = Array.isArray(profile.editableFilters)
        ? profile.editableFilters.map(value => normalizeFilterForAccess(value)).filter(Boolean)
        : [];
    return {
        key: String(profile.key || READ_ONLY_ACCESS_PROFILE.key).trim().toUpperCase(),
        label: String(profile.label || READ_ONLY_ACCESS_PROFILE.label).trim().toUpperCase(),
        canEditAll: profile.canEditAll === true,
        editableFilters
    };
}

function resolveRowFilterFromPlant(plantaValue) {
    const plantNorm = normalizePlantFilterValue(plantaValue);
    if (plantNorm === '' || plantNorm === 'S/DESTINO') return PLANTA_FILTER_OTHERS;
    return plantNorm;
}

function updateAccessBadge() {
    const headerRow = document.querySelector('header .header-row');
    if (!headerRow) return;

    let badge = document.getElementById('access-role-badge');
    if (!badge) {
        badge = document.createElement('span');
        badge.id = 'access-role-badge';
        badge.className = 'meta-pill access-role-badge';
        headerRow.appendChild(badge);
    }

    const profile = getActiveAccessProfile();
    const modeLabel = profile.canEditAll
        ? 'Acceso total'
        : (profile.editableFilters.length ? 'Edicion parcial' : 'Solo lectura');
    badge.textContent = `${profile.label} | ${modeLabel}`;
}

function setActiveAccessProfile(rawProfile) {
    activeAccessProfile = normalizeAccessProfile(rawProfile || READ_ONLY_ACCESS_PROFILE);
    window.STOCK_COSTURA_ACTIVE_PROFILE = activeAccessProfile;
    updateAccessBadge();
}

function getActiveAccessProfile() {
    if (!activeAccessProfile) {
        const fromWindow = window.STOCK_COSTURA_ACTIVE_PROFILE || window.STOCK_COSTURA_ROLE_PRESET;
        setActiveAccessProfile(fromWindow || READ_ONLY_ACCESS_PROFILE);
    }
    return activeAccessProfile;
}

function canEditFilter(filterValue) {
    const profile = getActiveAccessProfile();
    if (profile.canEditAll) return true;
    const normalized = normalizeFilterForAccess(filterValue);
    return profile.editableFilters.includes(normalized);
}

function canEditCurrentFilter() {
    return canEditFilter(selectedPlantaFilter);
}

function normalizeColorFilterValue(value) {
    return String(value === null || value === undefined ? '' : value).trim().toUpperCase();
}

function resolveInitialPlantaFilterForProfile(profile) {
    const rawKey = String(profile && profile.key ? profile.key : '').trim().toUpperCase();
    const editableFilters = profile && Array.isArray(profile.editableFilters) ? profile.editableFilters : [];
    const supportedFilters = new Set(['COFACO', 'COFACO 2', 'CITI1', 'CITI2', 'CITI3', 'CITI4', PLANTA_FILTER_OTHERS]);

    if (supportedFilters.has(rawKey)) {
        return rawKey;
    }

    const firstEditable = editableFilters
        .map(value => normalizeFilterForAccess(value))
        .find(value => supportedFilters.has(value));
    if (firstEditable) return firstEditable;

    return selectedPlantaFilter || 'COFACO';
}

function canEditRowMeta(meta) {
    if (!meta || !Number.isFinite(Number(meta.rowIndex))) return false;
    const row = allData.find(item => Number(item.rowIndex) === Number(meta.rowIndex));
    if (!row) return false;
    return canEditFilter(resolveRowFilterFromPlant(row.planta));
}

async function ensureAccessFromLogin() {
    const loginApi = window.StockCosturaLogin;
    if (loginApi && typeof loginApi.requireAccess === 'function') {
        const profile = await loginApi.requireAccess();
        setActiveAccessProfile(profile);
        selectedPlantaFilter = resolveInitialPlantaFilterForProfile(profile);
        return;
    }
    const fallback = window.STOCK_COSTURA_ACTIVE_PROFILE || window.STOCK_COSTURA_ROLE_PRESET || READ_ONLY_ACCESS_PROFILE;
    setActiveAccessProfile(fallback);
    selectedPlantaFilter = resolveInitialPlantaFilterForProfile(fallback);
}

function showLoader(show) {
    document.getElementById('loader').style.display = show ? 'flex' : 'none';
}

function parseGvizToMatrix(jsonResponse) {
    if (!jsonResponse || !jsonResponse.table) {
        throw new Error('Respuesta gviz invalida');
    }

    const colCount = Array.isArray(jsonResponse.table.cols) ? jsonResponse.table.cols.length : 0;
    const rowsRaw = jsonResponse.table.rows.map(r => {
        const mapped = r.c.map(cell => {
            if (!cell) return '';
            if (cell.v !== null && cell.v !== undefined) return cell.v;
            if (cell.f !== null && cell.f !== undefined) return cell.f;
            return '';
        });
        if (colCount > 0 && mapped.length < colCount) {
            while (mapped.length < colCount) mapped.push('');
        }
        return mapped;
    });
    const gvizHeaders = jsonResponse.table.cols.map(col => col.label || col.id || '');

    let matrix;
    if (gvizHeaders.includes('OP TELA')) {
        matrix = [gvizHeaders, ...rowsRaw];
    } else {
        let headerRowIndex = -1;
        for (let i = 0; i < rowsRaw.length; i++) {
            if (rowsRaw[i].includes('OP TELA')) {
                headerRowIndex = i;
                break;
            }
        }

        if (headerRowIndex !== -1) {
            const actualHeaders = rowsRaw[headerRowIndex];
            matrix = [actualHeaders, ...rowsRaw];
        } else {
            matrix = [gvizHeaders, ...rowsRaw];
        }
    }

    if (matrix.length > 0 && Array.isArray(matrix[0])) {
        matrix[0] = matrix[0].map(h => (h === null || h === undefined) ? '' : String(h).trim());
    }

    return matrix;
}

function fetchDataViaGviz() {
    return new Promise((resolve, reject) => {
        const callbackName = `costuraLoadDataCallback_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
        const script = document.createElement('script');
        let done = false;

        const cleanup = () => {
            if (script.parentNode) script.parentNode.removeChild(script);
            try { delete window[callbackName]; } catch (e) { window[callbackName] = undefined; }
        };

        window[callbackName] = function (jsonResponse) {
            if (done) return;
            done = true;
            cleanup();
            try {
                resolve(parseGvizToMatrix(jsonResponse));
            } catch (error) {
                reject(error);
            }
        };

        script.onerror = () => {
            if (done) return;
            done = true;
            cleanup();
            reject(new Error('Error de conexion con Google Sheets (gviz)'));
        };

        script.src = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=responseHandler:${callbackName}`;
        document.body.appendChild(script);
    });
}

async function fetchDataViaWebApp() {
    const response = await fetch(WEB_APP_URL);
    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    if (!data || !Array.isArray(data)) {
        throw new Error('Datos invalidos recibidos');
    }
    return data;
}

async function reloadData(options = {}) {
    showLoader(true);
    try {
        let data;
        if (options && options.preferWebApp) {
            try {
                data = await fetchDataViaWebApp();
            } catch (webAppError) {
                console.warn('Fallo WEB_APP_URL, usando fallback gviz:', webAppError);
                data = await fetchDataViaGviz();
            }
        } else {
            try {
                data = await fetchDataViaGviz();
            } catch (gvizError) {
                console.warn('Fallo gviz, usando fallback WEB_APP_URL:', gvizError);
                data = await fetchDataViaWebApp();
            }
        }

        processData(data);
        renderTable();
    } catch (error) {
        console.error('Error cargando datos:', error);
        document.getElementById('tbody-costura').innerHTML = '<tr><td colspan="19" class="no-data">Error al cargar los datos: ' + error.message + '</td></tr>';
    } finally {
        showLoader(false);
    }
}

function processData(data) {
    if (!Array.isArray(data) || data.length === 0) {
        console.log('No data received or data is not an array');
        allData = [];
        ocSearchDataStamp += 1;
        ocSearchState = null;
        renderPlantPills();
        updateStats();
        return;
    }

    const headers = data[0];
    const rows = data.slice(1);

    console.log('Headers:', headers);
    console.log('Total rows:', rows.length);

    // Encontrar ÃƒÂ­ndices de columnas
    const colMap = {};
    headers.forEach((header, idx) => {
        colMap[header] = idx;
    });
    const normalizedColMap = {};
    headers.forEach((header, idx) => {
        normalizedColMap[normalizeHeader(header)] = idx;
    });
    const getColIdx = (...aliases) => {
        for (const alias of aliases) {
            if (colMap[alias] !== undefined) return colMap[alias];
            const normalizedAlias = normalizeHeader(alias);
            if (normalizedColMap[normalizedAlias] !== undefined) return normalizedColMap[normalizedAlias];
        }
        return -1;
    };

    const estadoHabilitadoIdx = getColIdx('estado_habilitado', 'ESTADO_HABILITADO', 'ESTADO HABILITADO');
    const liquidacionCostIdx = getColIdx('liquidacion_cost', 'LIQUIDACION_COST', 'liquidacion cost', 'LIQUIDACION COST', 'liquidacion', 'LIQUIDACION');
    const opTelaIdx = getColIdx('OP TELA', 'OP_TELA', 'OP-TELA', 'OPTELA');
    const partidaIdx = getColIdx('PARTIDA');
    const pcostIdx = getColIdx('Pcost', 'PCOST', 'pcost', 'P COST', 'P.COST', 'N', 'n');
    const fDespachoIdx = getColIdx('F. DESPACHO', 'F.DESPACHO', 'F DESPACHO', 'F DESPACH', 'FDESPACHO', 'HOD');
    const fIngRealIdx = getColIdx('F.ING.REAL', 'F ING REAL', 'F. ING. REAL', 'F.ING REAL', 'FINGREAL', 'F.ING.COST', 'F. ING. COST', 'FINGCOST');
    const hiloCosturaIdx = getColIdx('HILO COSTURA', 'HILO_COSTURA', 'hilo costura', 'hilo_costura');
    const estadoAviosIdx = getColIdx('estado_avios', 'ESTADO_AVIOS', 'estado avios', 'ESTADO AVIOS');
    const stockInicioIdx = getColIdx('stock inicio', 'STOCK INICIO', 'stock_inicio', 'STOCK_INICIO');
    const salidaLineaIdx = getColIdx('salida de linea', 'SALIDA DE LINEA', 'salida_de_linea', 'SALIDA_DE_LINEA');
    const estadoCosturaIdx = getColIdx('estado_costura', 'ESTADO_COSTURA', 'estado costura', 'ESTADO COSTURA');

    // Validar que exista la columna estado_habilitado
    if (estadoHabilitadoIdx === -1) {
        console.warn('Columna estado_habilitado no encontrada. Columnas disponibles:', headers);
        allData = [];
        ocSearchDataStamp += 1;
        ocSearchState = null;
        renderPlantPills();
        updateStats();
        return;
    }

    // Filtrar registros donde estado_habilitado = OK | PROG 1T | PROG 2T | PROG 3T
    // DEPURADO no debe mostrarse en Stock Costura.
    const allowedEstadoHabilitado = new Set(['OK', 'PROG 1T', 'PROG 2T', 'PROG 3T']);
    allData = rows
        .map((row, idx) => ({ row, sourceRowIndex: idx + 2 }))
        .filter(item => {
            const estadoHabil = item.row[estadoHabilitadoIdx] || '';
            const estadoNorm = (estadoHabil || '').toString().toUpperCase().trim();
            const liquidacionRaw = liquidacionCostIdx > -1 ? item.row[liquidacionCostIdx] : '';
            if (estadoNorm === 'DEPURADO') return false;
            return allowedEstadoHabilitado.has(estadoNorm) && !isLiquidacionCostClosed(liquidacionRaw);
        })
        .map(item => {
            const row = item.row;
            const fDespachoRaw = fDespachoIdx > -1 ? row[fDespachoIdx] || '' : '';
            const fIngRealRaw = fIngRealIdx > -1 ? row[fIngRealIdx] || '' : '';
            const opVal = row[colMap['OP']] || '';
            const corteVal = row[colMap['CORTE']] || '';
            const colorVal = row[colMap['COLOR']] || '';
            const opTelaVal = opTelaIdx > -1 ? row[opTelaIdx] || '' : opVal;
            const partidaVal = partidaIdx > -1 ? row[partidaIdx] || '' : corteVal;
            const stockInicioRaw = stockInicioIdx > -1 ? row[stockInicioIdx] || '' : '';
            const salidaDeLineaRaw = salidaLineaIdx > -1 ? row[salidaLineaIdx] || '' : '';

            return {
                rowIndex: item.sourceRowIndex,
                fDespacho: fDespachoRaw,
                fDespachoDisplay: formatDate(fDespachoRaw),
                fDespachoTooltip: formatDateTooltip(fDespachoRaw),
                fIngReal: fIngRealRaw,
                fIngRealDisplay: formatDate(fIngRealRaw),
                fIngRealTooltip: formatDateTooltip(fIngRealRaw),
                planta: String(row[colMap['PLANTA']] || '').trim(),
                linea: row[colMap['LINEA']] || '',
                cliente: row[colMap['CLIENTE']] || '',
                pcost: pcostIdx > -1 ? row[pcostIdx] || '' : '',
                op: opVal,
                corte: corteVal,
                color: colorVal,
                colorTooltip: formatColorTooltipValues(colorVal, opTelaVal, partidaVal),
                opTela: opTelaVal,
                partida: partidaVal,
                pds: parseFloat(row[colMap['PDS GIRADAS']]) || 0,
                prenda: row[colMap['PRENDA']] || '',
                tipoCert: row[colMap['TIPO CERTIFICADO']] || '',
                estadoHabilitado: row[estadoHabilitadoIdx] || '',
                rib: row[colMap['estado_rib']] || row[colMap['RIB']] || '',
                collTap: row[colMap['estado_coll_tap']] || row[colMap['COLL o TAP?']] || '',
                trsfTipo: row[colMap['tipo-transfer']] || row[colMap['TIPO-TRANSFER']] || row[colMap['tipo_transfer']] || '',
                tipoBordado: row[colMap['tipo-bordado']] || row[colMap['TIPO-BORDADO']] || row[colMap['tipo_bordado']] || row[colMap['TIPO_BORDADO']] || '',
                nEstmp: row[colMap['n.ESTAMPxpda']] || row[colMap['N.ESTAMPXPDA']] || row[colMap['n.ESTAMP xpda']] || '',
                hiloCostura: hiloCosturaIdx > -1 ? row[hiloCosturaIdx] || '' : '',
                estadoAvios: estadoAviosIdx > -1 ? row[estadoAviosIdx] || '' : '',
                stockInicio: stockInicioRaw,
                salidaDeLinea: salidaDeLineaRaw,
                stockFinalDisplay: formatStockFinal(stockInicioRaw, salidaDeLineaRaw),
                estadoCostura: estadoCosturaIdx > -1 ? row[estadoCosturaIdx] || '' : '',
                liquidacionCost: liquidacionCostIdx > -1 ? row[liquidacionCostIdx] || '' : ''
            };
        });

    console.log('Filtered data (OK/PROG 1T/PROG 2T/PROG 3T):', allData.length, 'records');
    ocSearchDataStamp += 1;
    ocSearchState = null;
    renderPlantPills();
    updateStats();
}

function getFilteredDataByFilter(filterValue, options = {}) {
    const includeSearch = options.includeSearch !== false;
    const explicitQuery = (typeof options.query === 'string') ? options.query : null;
    const searchQuery = explicitQuery !== null ? explicitQuery : activeOcSearchQuery;
    let rows;

    if (filterValue === PLANTA_FILTER_OTHERS) {
        rows = allData.filter(d => {
            const plantNorm = normalizePlantFilterValue(d.planta);
            return plantNorm === '' || plantNorm === 'S/DESTINO';
        });
    } else {
        rows = allData.filter(d => normalizePlantFilterValue(d.planta) === filterValue);
    }

    if (!includeSearch || !searchQuery) return rows;
    return rows.filter(row => rowMatchesOcSearchQuery(row, searchQuery));
}

function getFilteredData() {
    return getFilteredDataByFilter(selectedPlantaFilter).filter(matchesSelectedColorFilter);
}

function isLineaEditableForCurrentFilter() {
    return canEditCurrentFilter();
}

function matchesSelectedColorFilter(row) {
    const filter = normalizeColorFilterValue(selectedColorFilter);
    if (!filter) return true;

    const rowColor = String((row && row.color) || '').trim();
    if (filter === '__EMPTY__') {
        return rowColor === '';
    }

    return normalizeColorFilterValue(rowColor) === filter;
}

function getBaseRowsForColorFilter() {
    return getFilteredDataByFilter(selectedPlantaFilter, { includeSearch: true });
}

function getUniqueColorOptionsForCurrentView() {
    const rows = getBaseRowsForColorFilter();
    const map = new Map();
    let hasEmpty = false;

    rows.forEach(row => {
        const raw = String((row && row.color) || '').trim();
        if (!raw) {
            hasEmpty = true;
            return;
        }
        const key = normalizeColorFilterValue(raw);
        if (!map.has(key)) {
            map.set(key, raw);
        }
    });

    const options = Array.from(map.entries())
        .sort((a, b) => a[1].localeCompare(b[1], undefined, { numeric: true, sensitivity: 'base' }))
        .map(([value, label]) => ({ value, label }));

    if (hasEmpty) {
        options.unshift({ value: '__EMPTY__', label: '(VACIO)' });
    }

    if (selectedColorFilter) {
        const currentNorm = normalizeColorFilterValue(selectedColorFilter);
        const alreadyExists = options.some(option => option.value === currentNorm || (option.value === '__EMPTY__' && currentNorm === '__EMPTY__'));
        if (!alreadyExists) {
            options.unshift({
                value: currentNorm,
                label: currentNorm === '__EMPTY__' ? '(VACIO)' : String(selectedColorFilter).trim()
            });
        }
    }

    return options;
}

function ensureColorFilterModalInitialized() {
    if (colorFilterPopoverState.popover) return;

    colorFilterPopoverState.popover = document.getElementById('color-filter-popover');
    colorFilterPopoverState.select = document.getElementById('color-filter-select');
    colorFilterPopoverState.btnClear = document.getElementById('color-filter-clear');
    if (!colorFilterPopoverState.popover || !colorFilterPopoverState.select || !colorFilterPopoverState.btnClear) {
        return;
    }

    colorFilterPopoverState.btnClear.addEventListener('click', () => {
        applyColorFilterValue('');
        closeColorFilterPopover();
    });
    colorFilterPopoverState.select.addEventListener('change', () => {
        applyColorFilterValue(colorFilterPopoverState.select.value);
        closeColorFilterPopover();
    });
    colorFilterPopoverState.select.addEventListener('keydown', event => {
        if (event.key === 'Enter') {
            event.preventDefault();
            applyColorFilterValue(colorFilterPopoverState.select.value);
            closeColorFilterPopover();
        } else if (event.key === 'Escape') {
            event.preventDefault();
            closeColorFilterPopover();
        }
    });
}

function updateColorFilterHeaderState() {
    const th = document.getElementById('th-color-filter');
    if (!th) return;

    const active = !!normalizeColorFilterValue(selectedColorFilter);
    th.classList.toggle('is-filter-active', active);
    if (active) {
        const label = selectedColorFilter === '__EMPTY__' ? '(VACIO)' : String(selectedColorFilter).trim();
        th.title = `Filtro activo: ${label}. Click derecho para cambiar.`;
    } else {
        th.title = 'Click derecho para filtrar por COLOR';
    }
}

function positionColorFilterPopover(anchorRect) {
    if (!colorFilterPopoverState.popover) return;

    const popover = colorFilterPopoverState.popover;
    const margin = 8;
    const width = Math.min(315, Math.max(260, popover.offsetWidth || 315));
    const viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
    const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;

    let left = anchorRect.right + margin;
    let top = anchorRect.bottom + 6;

    if (left + width > viewportWidth - margin) {
        left = Math.max(margin, anchorRect.left - width - margin);
    }

    const popoverHeight = popover.offsetHeight || 220;
    if (top + popoverHeight > viewportHeight - margin) {
        top = Math.max(margin, viewportHeight - popoverHeight - margin);
    }

    popover.style.left = `${Math.round(left)}px`;
    popover.style.top = `${Math.round(top)}px`;
    popover.style.minWidth = `${width}px`;
}

function openColorFilterPopover(anchorRect) {
    ensureColorFilterModalInitialized();
    if (!colorFilterPopoverState.popover || !colorFilterPopoverState.select) return;

    const options = getUniqueColorOptionsForCurrentView();
    colorFilterPopoverState.currentOptions = options;
    colorFilterPopoverState.select.innerHTML = [
        '<option value="">-- Seleccione un valor --</option>',
        ...options.map(option => {
            const valueAttr = escapeHtmlAttr(option.value);
            return `<option value="${valueAttr}">${escapeHtml(option.label)}</option>`;
        })
    ].join('');

    colorFilterPopoverState.select.value = normalizeColorFilterValue(selectedColorFilter);
    if (selectedColorFilter === '') {
        colorFilterPopoverState.select.value = '';
    } else if (selectedColorFilter === '__EMPTY__') {
        colorFilterPopoverState.select.value = '__EMPTY__';
    }

    colorFilterPopoverState.anchorRect = anchorRect || document.getElementById('th-color-filter')?.getBoundingClientRect?.();
    if (colorFilterPopoverState.anchorRect) {
        colorFilterPopoverState.popover.style.visibility = 'hidden';
        colorFilterPopoverState.popover.classList.add('active');
        colorFilterPopoverState.popover.setAttribute('aria-hidden', 'false');
        requestAnimationFrame(() => {
            positionColorFilterPopover(colorFilterPopoverState.anchorRect);
            colorFilterPopoverState.popover.style.visibility = 'visible';
        });
    } else {
        colorFilterPopoverState.popover.classList.add('active');
        colorFilterPopoverState.popover.setAttribute('aria-hidden', 'false');
    }

    attachColorFilterOutsideHandlers();
    setTimeout(() => {
        try { colorFilterPopoverState.select.focus(); } catch (e) { }
    }, 0);
}

function closeColorFilterPopover() {
    if (!colorFilterPopoverState.popover) return;
    colorFilterPopoverState.popover.classList.remove('active');
    colorFilterPopoverState.popover.setAttribute('aria-hidden', 'true');
    colorFilterPopoverState.popover.style.visibility = '';
    detachColorFilterOutsideHandlers();
}

function applyColorFilterValue(value) {
    selectedColorFilter = String(value || '').trim();
    ocSearchState = null;
    updateColorFilterHeaderState();
    updateStats();
    renderTable();
}

function openColorFilterFromHeader(event) {
    if (event) {
        event.preventDefault();
        event.stopPropagation();
    }
    const targetRect = event && event.currentTarget && typeof event.currentTarget.getBoundingClientRect === 'function'
        ? event.currentTarget.getBoundingClientRect()
        : null;
    openColorFilterPopover(targetRect);
}

function onColorFilterDocumentPointerDown(event) {
    if (!colorFilterPopoverState.popover || !colorFilterPopoverState.popover.classList.contains('active')) return;
    const popover = colorFilterPopoverState.popover;
    const header = document.getElementById('th-color-filter');
    if (popover.contains(event.target) || (header && header.contains(event.target))) return;
    closeColorFilterPopover();
}

function onColorFilterViewportChanged() {
    closeColorFilterPopover();
}

function attachColorFilterOutsideHandlers() {
    document.addEventListener('mousedown', onColorFilterDocumentPointerDown, true);
    document.addEventListener('touchstart', onColorFilterDocumentPointerDown, true);
    window.addEventListener('resize', onColorFilterViewportChanged, true);
    window.addEventListener('scroll', onColorFilterViewportChanged, { capture: true, passive: true });
}

function detachColorFilterOutsideHandlers() {
    document.removeEventListener('mousedown', onColorFilterDocumentPointerDown, true);
    document.removeEventListener('touchstart', onColorFilterDocumentPointerDown, true);
    window.removeEventListener('resize', onColorFilterViewportChanged, true);
    window.removeEventListener('scroll', onColorFilterViewportChanged, true);
}

function normalizePlantaModalValue(value) {
    return String(value || '').trim().toUpperCase();
}

function ensureLineaModalInitialized() {
    if (lineaModalRefs.overlay) return;

    lineaModalRefs.overlay = document.getElementById('linea-modal-overlay');
    lineaModalRefs.planta = document.getElementById('linea-modal-planta');
    lineaModalRefs.linea = document.getElementById('linea-modal-linea');
    lineaModalRefs.btnCancel = document.getElementById('linea-modal-cancel');
    lineaModalRefs.btnSave = document.getElementById('linea-modal-save');
    if (!lineaModalRefs.overlay || !lineaModalRefs.planta || !lineaModalRefs.linea || !lineaModalRefs.btnCancel || !lineaModalRefs.btnSave) return;

    lineaModalRefs.btnCancel.addEventListener('click', () => closeLineaModal());
    lineaModalRefs.btnSave.addEventListener('click', () => saveLineaModalChanges());
    lineaModalRefs.overlay.addEventListener('click', (event) => {
        if (event.target === lineaModalRefs.overlay) closeLineaModal();
    });
    lineaModalRefs.linea.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
            event.preventDefault();
            saveLineaModalChanges();
        } else if (event.key === 'Escape') {
            event.preventDefault();
            closeLineaModal();
        }
    });
}

function renderLineaModalPlantaOptions(currentValue) {
    ensureLineaModalInitialized();
    if (!lineaModalRefs.planta) return;

    const normalizedCurrent = normalizePlantaModalValue(currentValue);
    const options = LINEA_MODAL_PLANTA_OPTIONS.slice();
    if (normalizedCurrent && !options.some(option => option.value === normalizedCurrent)) {
        options.unshift({ value: normalizedCurrent, label: normalizedCurrent });
    }

    lineaModalRefs.planta.innerHTML = options.map(option => {
        const selected = option.value === normalizedCurrent ? ' selected' : '';
        return `<option value="${escapeHtmlAttr(option.value)}"${selected}>${escapeHtml(option.label)}</option>`;
    }).join('');
}

function openLineaModal(meta, currentPlanta, currentLinea) {
    ensureLineaModalInitialized();
    if (!lineaModalRefs.overlay || !lineaModalRefs.planta || !lineaModalRefs.linea) return;
    if (!canEditRowMeta(meta)) {
        alert('Solo tienes permisos de lectura en esta vista.');
        return;
    }

    const plantaNorm = normalizePlantaModalValue(currentPlanta);
    const lineaNorm = String(currentLinea || '').trim();
    lineaModalState = {
        meta,
        previousPlanta: plantaNorm,
        previousLinea: lineaNorm
    };

    renderLineaModalPlantaOptions(plantaNorm);
    lineaModalRefs.planta.value = plantaNorm;
    lineaModalRefs.linea.value = lineaNorm;
    lineaModalRefs.overlay.classList.add('active');
    lineaModalRefs.overlay.setAttribute('aria-hidden', 'false');
    setTimeout(() => {
        try { lineaModalRefs.linea.focus(); lineaModalRefs.linea.select(); } catch (e) { }
    }, 0);
}

function closeLineaModal() {
    ensureLineaModalInitialized();
    if (!lineaModalRefs.overlay) return;
    if (lineaModalRefs.btnSave) lineaModalRefs.btnSave.disabled = false;
    if (lineaModalRefs.btnCancel) lineaModalRefs.btnCancel.disabled = false;
    if (lineaModalRefs.planta) lineaModalRefs.planta.disabled = false;
    if (lineaModalRefs.linea) lineaModalRefs.linea.disabled = false;
    lineaModalRefs.overlay.classList.remove('active');
    lineaModalRefs.overlay.setAttribute('aria-hidden', 'true');
    lineaModalState = null;
}

function ensureLiquidacionModalInitialized() {
    if (liquidacionModalRefs.overlay) return;

    liquidacionModalRefs.overlay = document.getElementById('liquidacion-modal-overlay');
    liquidacionModalRefs.btnNo = document.getElementById('liquidacion-modal-no');
    liquidacionModalRefs.btnYes = document.getElementById('liquidacion-modal-yes');
    if (!liquidacionModalRefs.overlay || !liquidacionModalRefs.btnNo || !liquidacionModalRefs.btnYes) return;

    liquidacionModalRefs.btnNo.addEventListener('click', () => closeLiquidacionModal());
    liquidacionModalRefs.btnYes.addEventListener('click', () => confirmLiquidacionModal());
    liquidacionModalRefs.overlay.addEventListener('click', event => {
        if (event.target === liquidacionModalRefs.overlay) {
            closeLiquidacionModal();
        }
    });
}

function openLiquidacionModal(meta, checkbox) {
    ensureLiquidacionModalInitialized();
    if (!liquidacionModalRefs.overlay || !liquidacionModalRefs.btnNo || !liquidacionModalRefs.btnYes) return;

    liquidacionModalState = {
        meta,
        checkbox
    };
    liquidacionModalRefs.btnNo.disabled = false;
    liquidacionModalRefs.btnYes.disabled = false;
    liquidacionModalRefs.overlay.classList.add('active');
    liquidacionModalRefs.overlay.setAttribute('aria-hidden', 'false');
    setTimeout(() => {
        try { liquidacionModalRefs.btnYes.focus(); } catch (e) { }
    }, 0);
}

function closeLiquidacionModal(options = {}) {
    ensureLiquidacionModalInitialized();
    if (!liquidacionModalRefs.overlay) return;

    const { keepCheckboxChecked = false } = options;
    if (!keepCheckboxChecked && liquidacionModalState && liquidacionModalState.checkbox) {
        liquidacionModalState.checkbox.checked = false;
    }
    liquidacionModalRefs.btnNo.disabled = false;
    liquidacionModalRefs.btnYes.disabled = false;
    liquidacionModalRefs.overlay.classList.remove('active');
    liquidacionModalRefs.overlay.setAttribute('aria-hidden', 'true');
    liquidacionModalState = null;
}

function ensureReorderModalInitialized() {
    if (reorderModalState.overlay) return;

    reorderModalState.overlay = document.getElementById('reorder-modal-overlay');
}

function openReorderModal() {
    ensureReorderModalInitialized();
    if (!reorderModalState.overlay) return;
    reorderModalState.overlay.classList.add('active');
    reorderModalState.overlay.setAttribute('aria-hidden', 'false');
}

function closeReorderModal() {
    ensureReorderModalInitialized();
    if (!reorderModalState.overlay) return;
    reorderModalState.overlay.classList.remove('active');
    reorderModalState.overlay.setAttribute('aria-hidden', 'true');
}

async function runWithOptionalReorderModal(task, delayMs = REORDER_MODAL_DELAY_MS) {
    let modalTimer = null;
    let modalShown = false;
    const taskPromise = Promise.resolve().then(task);

    modalTimer = setTimeout(() => {
        modalShown = true;
        openReorderModal();
    }, delayMs);

    try {
        return await taskPromise;
    } finally {
        if (modalTimer) clearTimeout(modalTimer);
        if (modalShown) closeReorderModal();
    }
}

async function confirmLiquidacionModal() {
    ensureLiquidacionModalInitialized();
    if (!liquidacionModalState || !liquidacionModalRefs.btnNo || !liquidacionModalRefs.btnYes) return;

    const { meta } = liquidacionModalState;
    if (!canEditRowMeta(meta)) {
        closeLiquidacionModal();
        alert('Solo tienes permisos de lectura en esta vista.');
        return;
    }
    liquidacionModalRefs.btnNo.disabled = true;
    liquidacionModalRefs.btnYes.disabled = true;

    try {
        const nextValue = buildLiquidacionCostStamp();
        await runWithOptionalReorderModal(async () => {
            await saveCellToSheet(meta, 'liquidacion_cost', nextValue);
            const deletedRow = removeLocalRow(meta.rowIndex);
            if (deletedRow) {
                resequenceLocalPcostByLinea(deletedRow.planta, deletedRow.linea);
            }
            updateStats();
            renderTable();
        });
        closeLiquidacionModal({ keepCheckboxChecked: true });
    } catch (err) {
        console.error('Error guardando liquidacion_cost:', err);
        alert('No se pudo guardar liquidacion_cost.');
        closeLiquidacionModal();
    }
}

function showCosturaToast(message, type = 'info') {
    let container = document.getElementById('toast-container');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toast-container';
        container.className = 'toast-container';
        document.body.appendChild(container);
    }

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;

    container.appendChild(toast);

    toast.offsetHeight; // trigger reflow
    toast.classList.add('show');

    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => {
            if (toast.parentNode === container) {
                container.removeChild(toast);
            }
        }, 300);
    }, 3000);
}

async function saveLineaModalChanges() {
    ensureLineaModalInitialized();
    if (!lineaModalState || !lineaModalRefs.planta || !lineaModalRefs.linea || !lineaModalRefs.btnSave || !lineaModalRefs.btnCancel) return;

    const nextPlanta = normalizePlantaModalValue(lineaModalRefs.planta.value);
    const nextLinea = String(lineaModalRefs.linea.value || '').trim();
    const { meta, previousPlanta, previousLinea } = lineaModalState;
    if (!canEditRowMeta(meta)) {
        closeLineaModal();
        alert('Solo tienes permisos de lectura en esta vista.');
        return;
    }
    const hasPlantaChange = nextPlanta !== previousPlanta;
    const hasLineaChange = nextLinea !== previousLinea;
    if (hasPlantaChange && !canEditFilter(resolveRowFilterFromPlant(nextPlanta))) {
        alert('No puedes mover registros a una vista sin permiso de edicion.');
        return;
    }

    if (!hasPlantaChange && !hasLineaChange) {
        closeLineaModal();
        return;
    }

    const originalPlanta = previousPlanta;
    const originalLinea = previousLinea;

    let localRowUpdated = false;
    if (hasPlantaChange) {
        updateLocalRowValue(meta.rowIndex, 'PLANTA', nextPlanta);
        localRowUpdated = true;
    }
    if (hasLineaChange) {
        updateLocalRowValue(meta.rowIndex, 'LINEA', nextLinea);
        localRowUpdated = true;
    }

    closeLineaModal();
    if (localRowUpdated) {
        renderPlantPills();
        updateStats();
        renderTable();
    }

    showCosturaToast('Guardando cambios...', 'info');

    try {
        const promises = [];
        if (hasPlantaChange) {
            promises.push(saveCellToSheet(meta, 'PLANTA', nextPlanta));
        }
        if (hasLineaChange) {
            promises.push(saveCellToSheet(meta, 'LINEA', nextLinea));
        }
        await Promise.all(promises);
        showCosturaToast('Cambios guardados', 'success');
    } catch (err) {
        console.error('Error guardando PLANTA/LINEA:', err);
        showCosturaToast('Error al guardar. Revirtiendo...', 'error');
        if (hasPlantaChange) updateLocalRowValue(meta.rowIndex, 'PLANTA', originalPlanta);
        if (hasLineaChange) updateLocalRowValue(meta.rowIndex, 'LINEA', originalLinea);
        renderPlantPills();
        updateStats();
        renderTable();
    }
}

function normalizePlantFilterValue(value) {
    return String(value || '')
        .trim()
        .toUpperCase();
}

function getShiftPillOptionsForPlant(plant) {
    const normalizedPlant = normalizePlantFilterValue(plant);
    return SHIFT_PILL_OPTIONS_BY_PLANT[normalizedPlant] || [];
}

function getShiftUserKeyFromAccessProfile() {
    const profile = getActiveAccessProfile();
    const normalized = normalizePlantFilterValue(profile && profile.key);
    return SHIFT_PILL_OPTIONS_BY_PLANT[normalized] ? normalized : '';
}

function normalizeShiftTokenValue(token) {
    return String(token || '')
        .trim()
        .toUpperCase();
}

function parseShiftStateTokens(rawState) {
    return String(rawState || '')
        .split(/[,;\n]+/)
        .map(token => normalizeShiftTokenValue(token))
        .filter(Boolean);
}

function isShiftTokenForPlant(token, plant, shift) {
    const normalizedToken = normalizeShiftTokenValue(token);
    const normalizedPlant = normalizePlantFilterValue(plant);
    const normalizedShift = normalizeShiftTokenValue(shift);
    if (!normalizedToken || !normalizedPlant || !normalizedShift) return false;

    if (normalizedPlant === 'COFACO' && normalizedToken === normalizedShift) {
        return true;
    }

    return normalizedToken === `${normalizedPlant}-${normalizedShift}`;
}

function hasShiftTokenForCurrentPlant(shift) {
    const userKey = getShiftUserKeyFromAccessProfile();
    if (!userKey) return false;
    return activeShifts.some(token => isShiftTokenForPlant(token, userKey, shift));
}

function isShiftOptionEnabledForCurrentPlant(shift) {
    const userKey = getShiftUserKeyFromAccessProfile();
    if (!userKey) return false;
    return getShiftPillOptionsForPlant(userKey).includes(normalizeShiftTokenValue(shift));
}

function removeShiftTokensForPlantShift(plant, shift) {
    const normalizedPlant = normalizePlantFilterValue(plant);
    const normalizedShift = normalizeShiftTokenValue(shift);
    activeShifts = activeShifts.filter(token => {
        const normalizedToken = normalizeShiftTokenValue(token);
        if (normalizedPlant === 'COFACO' && normalizedToken === normalizedShift) {
            return false;
        }
        return normalizedToken !== `${normalizedPlant}-${normalizedShift}`;
    });
}

function addShiftTokenForPlant(plant, shift) {
    const normalizedPlant = normalizePlantFilterValue(plant);
    const normalizedShift = normalizeShiftTokenValue(shift);
    const nextToken = `${normalizedPlant}-${normalizedShift}`;
    if (!activeShifts.includes(nextToken)) {
        activeShifts.push(nextToken);
    }
}

function serializeShiftStateTokens(tokens) {
    const normalized = [];
    const seen = new Set();

    (tokens || []).forEach(token => {
        const normalizedToken = normalizeShiftTokenValue(token);
        if (!normalizedToken) return;

        const migratedToken = normalizedToken.includes('-')
            ? normalizedToken
            : `COFACO-${normalizedToken}`;

        if (seen.has(migratedToken)) return;
        seen.add(migratedToken);
        normalized.push(migratedToken);
    });

    return normalized.join(',');
}

function sanitizeShiftStateTokens(tokens) {
    const cleaned = [];
    const seen = new Set();

    (tokens || []).forEach(token => {
        const normalizedToken = normalizeShiftTokenValue(token);
        if (!normalizedToken) return;

        let plant = '';
        let shift = '';

        if (normalizedToken.includes('-')) {
            const parts = normalizedToken.split('-');
            plant = normalizePlantFilterValue(parts.shift());
            shift = normalizeShiftTokenValue(parts.join('-'));
        } else {
            plant = 'COFACO';
            shift = normalizedToken;
        }

        const allowedShifts = getShiftPillOptionsForPlant(plant);
        if (!allowedShifts.includes(shift)) return;

        const canonicalToken = `${plant}-${shift}`;
        if (seen.has(canonicalToken)) return;
        seen.add(canonicalToken);
        cleaned.push(canonicalToken);
    });

    return cleaned;
}

function getShiftDisplayTokensForPlant(plant) {
    const plantKey = normalizePlantFilterValue(plant);
    const allowed = getShiftPillOptionsForPlant(plantKey);
    if (!allowed.length) return [];

    return allowed.filter(shift => {
        return activeShifts.some(token => isShiftTokenForPlant(token, plantKey, shift));
    });
}

function renderPlantPills() {
    renderShiftPills();
    const container = document.getElementById('plant-pills');
    if (!container) return;

    const plantSet = new Set(allData.map(d => normalizePlantFilterValue(d.planta)));
    const hasOthers = allData.some(d => {
        const plantNorm = normalizePlantFilterValue(d.planta);
        return plantNorm === '' || plantNorm === 'S/DESTINO';
    });

    const allPills = [
        { value: 'COFACO', label: 'COFACO', enabled: plantSet.has('COFACO') },
        { value: 'COFACO 2', label: 'COFACO 2', enabled: plantSet.has('COFACO 2') },
        { value: 'CITI1', label: 'CITI1', enabled: plantSet.has('CITI1') },
        { value: 'CITI2', label: 'CITI2', enabled: plantSet.has('CITI2') },
        { value: 'CITI3', label: 'CITI3', enabled: plantSet.has('CITI3') },
        { value: 'CITI4', label: 'CITI4', enabled: plantSet.has('CITI4') },
        { value: PLANTA_FILTER_OTHERS, label: 'Otros', enabled: hasOthers, title: 'Incluye S/DESTINO y VACIO' }
    ];

    const currentTarget = allPills.find(p => p.value === selectedPlantaFilter);
    const firstEnabled = allPills.find(p => p.enabled !== false);
    if (!currentTarget || currentTarget.enabled === false) {
        selectedPlantaFilter = firstEnabled ? firstEnabled.value : 'COFACO';
    }

    container.innerHTML = allPills.map(p => {
        const active = selectedPlantaFilter === p.value ? ' active' : '';
        const disabled = p.enabled === false ? ' style="opacity:.45;"' : '';
        const title = p.title ? ` title="${escapeHtmlAttr(p.title)}"` : '';
        return `<button type="button" class="plant-pill${active}" data-value="${escapeHtmlAttr(p.value)}"${disabled}${title}>${escapeHtml(p.label)}</button>`;
    }).join('');

    container.querySelectorAll('.plant-pill').forEach(btn => {
        const value = btn.getAttribute('data-value') || 'COFACO';
        btn.addEventListener('click', () => {
            const target = allPills.find(p => p.value === value);
            if (target && target.enabled === false) return;
            selectedPlantaFilter = value;
            renderShiftPills();
            renderPlantPills();
            updateStats();
            renderTable();
        });
    });
}

function escapeHtml(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function escapeHtmlAttr(text) {
    return escapeHtml(text);
}

function calculateMetricsForRows(rows) {
    let proceso = 0;
    let xHabilitar = 0;

    (rows || []).forEach(row => {
        const estadoNorm = String(row.estadoHabilitado || '').toUpperCase().trim();
        const pdsVal = Number(row.pds) || 0;

        const stockInicio = parseEditableInteger(row.stockInicio);
        const salidaLinea = parseEditableInteger(row.salidaDeLinea);
        const sFinalVal = (stockInicio === null ? 0 : stockInicio) - (salidaLinea === null ? 0 : salidaLinea);
        proceso += sFinalVal;

        if (estadoNorm === 'PROG 1T' || estadoNorm === 'PROG 2T' || estadoNorm === 'PROG 3T') {
            xHabilitar += pdsVal;
        }
    });

    return { proceso, xHabilitar };
}

function updateStats() {
    const filtered = getFilteredData();
    const currentMetrics = calculateMetricsForRows(filtered);
    let totalProceso = 0;
    let totalXHabilitar = 0;

    PLANTA_TOTAL_FILTERS.forEach(filterValue => {
        const rowsByFilter = getFilteredDataByFilter(filterValue, { includeSearch: false }).filter(matchesSelectedColorFilter);
        const m = calculateMetricsForRows(rowsByFilter);
        totalProceso += m.proceso;
        totalXHabilitar += m.xHabilitar;
    });

    const pillProceso = document.getElementById('metric-pds-proceso');
    if (pillProceso) {
        const procesoLabel = currentMetrics.proceso.toLocaleString('es-ES', { maximumFractionDigits: 0 });
        pillProceso.textContent = `${procesoLabel} prendas en proceso`;
    }

    const pillXHab = document.getElementById('metric-pds-xhabilitar');
    if (pillXHab) pillXHab.textContent = `${formatNumber(currentMetrics.xHabilitar)} prendas x habilitar`;

    const totalProcesoPill = document.getElementById('metric-total-proceso');
    if (totalProcesoPill) {
        const totalProcesoLabel = totalProceso.toLocaleString('es-ES', { maximumFractionDigits: 0 });
        totalProcesoPill.textContent = `TOTAL: ${totalProcesoLabel} PDS proceso`;
    }

    const totalXHabPill = document.getElementById('metric-total-xhabilitar');
    if (totalXHabPill) {
        const totalXHabLabel = totalXHabilitar.toLocaleString('es-ES', { maximumFractionDigits: 0 });
        totalXHabPill.textContent = `${totalXHabLabel} PDS x habilitar`;
    }
}

function renderTable() {
    const filtered = getFilteredData();
    const canEditInCurrentView = canEditCurrentFilter();
    const allowLineaEdit = canEditInCurrentView && isLineaEditableForCurrentFilter();
    updateColorFilterHeaderState();

    const tbody = document.getElementById('tbody-costura');

    if (filtered.length === 0) {
        if (allData.length === 0) {
            tbody.innerHTML = '<tr><td colspan="19" class="no-data">No hay registros de costura registrados aÃƒÂºn</td></tr>';
        } else if (activeOcSearchQuery) {
            tbody.innerHTML = '<tr><td colspan="19" class="no-data">No se encontraron coincidencias para la OC buscada en la planta seleccionada</td></tr>';
        } else if (normalizeColorFilterValue(selectedColorFilter)) {
            const colorLabel = selectedColorFilter === '__EMPTY__' ? '(VACIO)' : String(selectedColorFilter).trim();
            tbody.innerHTML = `<tr><td colspan="19" class="no-data">No hay registros para el color ${escapeHtml(colorLabel)}</td></tr>`;
        } else {
            tbody.innerHTML = '<tr><td colspan="19" class="no-data">No hay registros para la planta seleccionada</td></tr>';
        }
        return;
    }

    const columnsCount = 19;
    const collapsedSet = getCollapsedLineSetForCurrentFilter();
    const sectorConfig = getSectorConfigForPlant(selectedPlantaFilter);
    let lineGroupNames = [];

    if (sectorConfig) {
        const sectorGroups = groupRowsBySector(filtered, selectedPlantaFilter);
        const validGroupNames = new Set();
        sectorGroups.forEach(sectorGroup => {
            (sectorGroup.lineGroups || []).forEach(group => validGroupNames.add(group.name));
        });
        Array.from(collapsedSet).forEach(name => {
            if (!validGroupNames.has(name)) collapsedSet.delete(name);
        });

        tbody.innerHTML = sectorGroups.map(sectorGroup => {
            const sectorBand = `
                    <tr class="sector-band" data-sector="${escapeHtmlAttr(sectorGroup.name)}">
                        <th colspan="${columnsCount}">
                            <span class="sector-badge">${escapeHtml(sectorGroup.name)}</span>
                            <span class="band-meta">${formatNumber(sectorGroup.stockFinalTotal)} pds</span>
                        </th>
                    </tr>
                `;

            const lineGroupsHtml = (sectorGroup.lineGroups || []).map(group => {
                lineGroupNames.push(group.name);
                return buildLineaGroupHtml(group, columnsCount, collapsedSet, canEditInCurrentView, allowLineaEdit);
            }).join('');

            return sectorBand + lineGroupsHtml;
        }).join('');
    } else {
        const groups = groupRowsByLinea(filtered);
        const validGroupNames = new Set(groups.map(group => group.name));
        Array.from(collapsedSet).forEach(name => {
            if (!validGroupNames.has(name)) collapsedSet.delete(name);
        });

        tbody.innerHTML = groups.map(group => {
            lineGroupNames.push(group.name);
            return buildLineaGroupHtml(group, columnsCount, collapsedSet, canEditInCurrentView, allowLineaEdit);
        }).join('');
    }

    bindLineaBandInteractions(lineGroupNames);
    updateColorFilterHeaderState();
}

function getLineaGroupName(value) {
    const raw = String(value || '').trim();
    return raw ? raw : 'SIN LINEA';
}

function compareLineaGroupNames(a, b) {
    const isANum = /^\d+$/.test(a);
    const isBNum = /^\d+$/.test(b);
    if (isANum && isBNum) return parseInt(a, 10) - parseInt(b, 10);
    if (a === 'SIN LINEA') return 1;
    if (b === 'SIN LINEA') return -1;
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
}

function getPcostSortValue(row) {
    const parsed = parseEditableInteger(row && row.pcost);
    return parsed === null ? null : parsed;
}

function compareOcGroupValues(a, b) {
    const opA = String((a && a.op) || '').trim();
    const opB = String((b && b.op) || '').trim();
    const opDiff = opA.localeCompare(opB, undefined, { numeric: true, sensitivity: 'base' });
    if (opDiff !== 0) return opDiff;

    const corteA = String((a && a.corte) || '').trim();
    const corteB = String((b && b.corte) || '').trim();
    const corteDiff = corteA.localeCompare(corteB, undefined, { numeric: true, sensitivity: 'base' });
    if (corteDiff !== 0) return corteDiff;

    return Number((a && a.rowIndex) || 0) - Number((b && b.rowIndex) || 0);
}

function compareRowsWithinLinea(a, b) {
    const pcostA = getPcostSortValue(a);
    const pcostB = getPcostSortValue(b);
    const hasPcostA = pcostA !== null;
    const hasPcostB = pcostB !== null;

    if (hasPcostA && hasPcostB && pcostA !== pcostB) {
        return pcostA - pcostB;
    }

    if (hasPcostA !== hasPcostB) {
        return hasPcostA ? -1 : 1;
    }

    return compareOcGroupValues(a, b);
}

function groupRowsByLinea(rows) {
    const map = new Map();
    (rows || []).forEach(row => {
        const name = getLineaGroupName(row.linea);
        if (!map.has(name)) {
            map.set(name, { name, stockFinalTotal: 0, rows: [] });
        }
        const group = map.get(name);
        group.rows.push(row);
        group.stockFinalTotal += getLineaGroupStockFinalNumericValue(row);
    });

    const groups = Array.from(map.values());
    groups.sort((a, b) => compareLineaGroupNames(a.name, b.name));
    groups.forEach(group => {
        group.rows.sort((a, b) => compareRowsWithinLinea(a, b));
    });
    return groups;
}

function getSectorConfigForPlant(plantValue) {
    const plant = String(plantValue || '').trim().toUpperCase();
    if (plant === 'COFACO') {
        return [
            { name: 'SECTOR 1', min: 1, max: 10 },
            { name: 'SECTOR 2', min: 11, max: 19 },
            { name: 'TURNO NOCHE', min: 20, max: 29 }
        ];
    }
    if (plant === 'CITI1') {
        return [
            { name: 'SECTOR 1', min: 1, max: 14 },
            { name: 'SECTOR 2', min: 15, max: 30 }
        ];
    }
    if (plant === 'CITI2') {
        return [
            { name: 'SECTOR 1', min: 1, max: 6 },
            { name: 'SECTOR 2', min: 7, max: 9 }
        ];
    }
    return null;
}

function getSectorNameForPlant(lineaValue, plantValue) {
    const config = getSectorConfigForPlant(plantValue);
    if (!config) return '';

    const raw = String(lineaValue || '').trim();
    if (!/^\d+$/.test(raw)) return 'SIN SECTOR';

    const lineNum = parseInt(raw, 10);
    const matched = config.find(range => lineNum >= range.min && lineNum <= range.max);
    return matched ? matched.name : 'SIN SECTOR';
}

function compareSectorNames(a, b) {
    const order = {
        'SECTOR 1': 1,
        'SECTOR 2': 2,
        'TURNO NOCHE': 3,
        'SIN SECTOR': 4
    };
    return (order[a] || 99) - (order[b] || 99);
}

function groupRowsBySector(rows, plantValue) {
    const map = new Map();

    (rows || []).forEach(row => {
        const sectorName = getSectorNameForPlant(row.linea, plantValue);
        if (!map.has(sectorName)) {
            map.set(sectorName, { name: sectorName, stockFinalTotal: 0, rows: [] });
        }
        const sectorGroup = map.get(sectorName);
        sectorGroup.rows.push(row);
        sectorGroup.stockFinalTotal += getLineaGroupStockFinalNumericValue(row);
    });

    const sectors = Array.from(map.values());
    sectors.sort((a, b) => compareSectorNames(a.name, b.name));
    sectors.forEach(sectorGroup => {
        sectorGroup.lineGroups = groupRowsByLinea(sectorGroup.rows);
    });
    return sectors;
}

function updateRenderedLineaBandTotal(lineaValue) {
    const lineaName = getLineaGroupName(lineaValue);
    const tbody = document.getElementById('tbody-costura');
    if (!tbody) return;

    const bandRow = Array.from(tbody.querySelectorAll('tr.linea-band')).find(row => {
        return String(row.getAttribute('data-linea') || '') === lineaName;
    });
    if (!bandRow) return;

    const bandMeta = bandRow.querySelector('.band-meta');
    if (!bandMeta) return;

    const total = getFilteredData()
        .filter(row => getLineaGroupName(row.linea) === lineaName)
        .reduce((sum, row) => sum + getLineaGroupStockFinalNumericValue(row), 0);

    bandMeta.textContent = `${formatNumber(total)} pds`;
}

function updateRenderedSectorBandTotal(sectorValue) {
    const plantValue = String(selectedPlantaFilter || '').toUpperCase();
    if (!getSectorConfigForPlant(plantValue)) return;
    const sectorName = String(sectorValue || '');
    if (!sectorName) return;

    const tbody = document.getElementById('tbody-costura');
    if (!tbody) return;

    const bandRow = Array.from(tbody.querySelectorAll('tr.sector-band')).find(row => {
        return String(row.getAttribute('data-sector') || '') === sectorName;
    });
    if (!bandRow) return;

    const bandMeta = bandRow.querySelector('.band-meta');
    if (!bandMeta) return;

    const total = getFilteredData()
        .filter(row => getSectorNameForPlant(row.linea, plantValue) === sectorName)
        .reduce((sum, row) => sum + getLineaGroupStockFinalNumericValue(row), 0);

    bandMeta.textContent = `${formatNumber(total)} pds`;
}

function buildLineaGroupHtml(group, columnsCount, collapsedSet, canEditInCurrentView, allowLineaEdit) {
    const isCollapsed = collapsedSet.has(group.name);
    const band = `
                    <tr class="linea-band${isCollapsed ? ' is-collapsed' : ''}" data-linea="${escapeHtmlAttr(group.name)}" title="1 click: contraer/expandir linea | doble click: contraer/expandir todas">
                        <th colspan="${columnsCount}">
                            <span class="band-toggle">${isCollapsed ? '+' : '-'}</span>
                            ${escapeHtml(group.name)}
                            <span class="band-meta">${formatNumber(group.stockFinalTotal)} pds</span>
                        </th>
                    </tr>
                `;

    const rowsHtml = isCollapsed ? '' : group.rows.map((row, idx) => `
                    <tr class="${idx % 2 === 0 ? 'group-a' : 'group-b'}">
                        <td title="${escapeHtmlAttr(row.fDespachoTooltip)}"><strong>${row.fDespachoDisplay}</strong></td>
                        <td title="${escapeHtmlAttr(row.fIngRealTooltip)}"><strong>${row.fIngRealDisplay}</strong></td>
                        <td class="${allowLineaEdit ? 'linea-editable-cell' : ''}" ${allowLineaEdit ? `title="Doble click para editar" ${buildRowUpdateMetaAttrs(row)} data-col-name="LINEA" data-current-linea="${escapeHtmlAttr(String(row.linea || '').trim())}" data-current-planta="${escapeHtmlAttr(String(row.planta || '').trim())}"` : ''}>${escapeHtml(String(row.linea || ''))}</td>
                        <td class="client-cell">${normalizeClientName(row.cliente)}</td>
                        <td class="pds-cell${canEditInCurrentView ? ' editable-int-cell pcost-editable-cell' : ''}" ${canEditInCurrentView ? `title="1 click para editar" ${buildRowUpdateMetaAttrs(row)} data-col-name="Pcost" data-raw="${escapeHtmlAttr(formatEditableIntegerRaw(row.pcost))}" data-allow-empty="1"` : ''}>${formatEditableIntegerDisplay(row.pcost)}</td>
                        <td class="oc-cell">${row.op}-${row.corte}</td>
                        <td class="${canEditInCurrentView ? 'color-split-cell' : ''}" ${canEditInCurrentView ? `title="${escapeHtmlAttr(row.colorTooltip)} | Click derecho: partir la OC o devolucion" ${buildRowUpdateMetaAttrs(row)}` : `title="${escapeHtmlAttr(row.colorTooltip)}"`}>${abbreviate(row.color)}</td>
                        <td class="pds-cell">${formatNumber(row.pds)}</td>
                        <td>${normalizePrenda(row.prenda)}</td>
                        <td>${normalizeCert(row.tipoCert)}</td>
                        <td>${formatStatusHabilitado(row.estadoHabilitado)}</td>
                        <td>${formatCompOtros(row)}</td>
                        <td class="hilo-cell">${formatHiloCostura(row.hiloCostura)}</td>
                        <td class="avios-cell">${formatEstadoAvios(row.estadoAvios)}</td>
                        <td class="pds-cell${canEditInCurrentView ? ' editable-int-cell' : ''}" ${canEditInCurrentView ? `title="Doble click para editar" ${buildRowUpdateMetaAttrs(row)} data-col-name="stock inicio" data-raw="${escapeHtmlAttr(formatEditableIntegerRaw(row.stockInicio))}"` : ''}>${formatEditableIntegerDisplay(row.stockInicio)}</td>
                        <td class="pds-cell${canEditInCurrentView ? ' editable-int-cell' : ''}" ${canEditInCurrentView ? `title="Doble click para editar" ${buildRowUpdateMetaAttrs(row)} data-col-name="salida de linea" data-raw="${escapeHtmlAttr(formatEditableIntegerRaw(row.salidaDeLinea))}"` : ''}>${formatEditableIntegerDisplay(row.salidaDeLinea)}</td>
                        <td class="pds-cell s-final-cell"><strong>${row.stockFinalDisplay}</strong></td>
                        <td>${renderEstadoCosturaSelect(row, canEditInCurrentView)}</td>
                        <td class="liquidacion-cell">${renderLiquidacionCheckbox(row, canEditInCurrentView)}</td>
                    </tr>
                `).join('');

    return band + rowsHtml;
}

function getCollapsedLineSetForCurrentFilter() {
    const key = String(selectedPlantaFilter || '');
    if (!collapsedLineasByFilter.has(key)) {
        collapsedLineasByFilter.set(key, new Set());
    }
    return collapsedLineasByFilter.get(key);
}

function toggleCollapsedStateForLinea(lineaName) {
    const name = String(lineaName || '');
    if (!name) return;
    const set = getCollapsedLineSetForCurrentFilter();
    if (set.has(name)) set.delete(name);
    else set.add(name);
}

function toggleCollapseForAllLineas(lineaNames) {
    const names = (lineaNames || []).map(name => String(name || '')).filter(Boolean);
    const set = getCollapsedLineSetForCurrentFilter();
    const allCollapsed = names.length > 0 && names.every(name => set.has(name));
    if (allCollapsed) {
        names.forEach(name => set.delete(name));
    } else {
        names.forEach(name => set.add(name));
    }
}

function bindLineaBandInteractions(lineaNames) {
    const tbody = document.getElementById('tbody-costura');
    if (!tbody) return;
    const names = (lineaNames || []).map(name => String(name || '')).filter(Boolean);

    tbody.querySelectorAll('tr.linea-band').forEach(band => {
        const lineaName = String(band.getAttribute('data-linea') || '');
        let clickTimer = null;

        band.addEventListener('click', (event) => {
            event.preventDefault();
            if (clickTimer) clearTimeout(clickTimer);
            clickTimer = setTimeout(() => {
                toggleCollapsedStateForLinea(lineaName);
                renderTable();
                clickTimer = null;
            }, LINEA_BAND_CLICK_DELAY_MS);
        });

        band.addEventListener('dblclick', (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (clickTimer) {
                clearTimeout(clickTimer);
                clickTimer = null;
            }
            toggleCollapseForAllLineas(names);
            renderTable();
        });
    });
}

function buildRowUpdateMetaAttrs(row) {
    return `data-row-index="${escapeHtmlAttr(row.rowIndex)}" data-source-op="${escapeHtmlAttr(row.op)}" data-source-corte="${escapeHtmlAttr(row.corte)}" data-source-op-tela="${escapeHtmlAttr(row.opTela)}" data-source-partida="${escapeHtmlAttr(row.partida)}" data-source-color="${escapeHtmlAttr(row.color)}" data-source-pds="${escapeHtmlAttr(formatEditableIntegerRaw(row.pds))}"`;
}

function parseEditableInteger(value) {
    const raw = String(value === null || value === undefined ? '' : value).trim();
    if (!raw) return null;
    const cleaned = raw.replace(/[\s,\.]/g, '');
    if (!/^\d+$/.test(cleaned)) return null;
    return parseInt(cleaned, 10);
}

function formatEditableIntegerRaw(value) {
    const parsed = parseEditableInteger(value);
    return parsed === null ? '' : String(parsed);
}

function formatEditableIntegerDisplay(value) {
    const raw = formatEditableIntegerRaw(value);
    if (raw === '') return '';
    const num = parseInt(raw, 10);
    if (!Number.isFinite(num)) return '';
    return num.toLocaleString('es-PE', { maximumFractionDigits: 0 });
}

function formatStockFinal(stockInicio, salidaDeLinea) {
    const stock = parseEditableInteger(stockInicio);
    const salida = parseEditableInteger(salidaDeLinea);
    if (stock === null && salida === null) return '';
    const diff = (stock === null ? 0 : stock) - (salida === null ? 0 : salida);
    return diff.toLocaleString('es-PE', { maximumFractionDigits: 0 });
}

function getStockFinalNumericValue(row) {
    if (!row) return 0;
    const stock = parseEditableInteger(row.stockInicio);
    const salida = parseEditableInteger(row.salidaDeLinea);
    if (stock === null && salida === null) return 0;
    return (stock === null ? 0 : stock) - (salida === null ? 0 : salida);
}

function getLineaGroupStockFinalNumericValue(row) {
    const baseValue = getStockFinalNumericValue(row);
    const currentFilter = String(selectedPlantaFilter || '').toUpperCase();
    if (currentFilter !== 'CITI1' && currentFilter !== 'CITI2' && currentFilter !== 'CITI3' && currentFilter !== 'CITI4') {
        return baseValue;
    }

    const estado = normalizeEstadoCostura(row && row.estadoCostura);
    if (estado === 'OK COSTURA') {
        return 0;
    }

    return baseValue;
}

function normalizeEstadoCostura(value) {
    const raw = String(value || '').trim();
    if (!raw) return '';

    const normalized = raw
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase();

    if (normalized === 'PROCESO') return 'Proceso';
    if (normalized === 'LIQUIDADO') return 'Liquidado';
    if (normalized === 'ANAQUEL') return 'Anaquel';
    if (normalized === 'EN HABILITADO') return 'En Habilitado';
    if (normalized === 'EN HAB CITI1') return 'En hab CITI1';
    if (normalized === 'EN HAB CITI2') return 'En hab CITI2';
    if (normalized === 'EN HAB CITI3') return 'En hab CITI3';
    if (normalized === 'EN HAB CITI4') return 'En hab CITI4';
    if (normalized === 'OK COSTURA') return 'OK COSTURA';
    return '';
}

function getEstadoCosturaOptionsForCurrentFilter() {
    const options = ESTADO_COSTURA_OPTIONS.slice();
    const currentFilter = String(selectedPlantaFilter || '').toUpperCase();
    if (currentFilter === 'CITI1') {
        options.push('En hab CITI1', 'OK COSTURA');
    } else if (currentFilter === 'CITI2') {
        options.push('En hab CITI2', 'OK COSTURA');
    } else if (currentFilter === 'CITI3') {
        options.push('En hab CITI3', 'OK COSTURA');
    } else if (currentFilter === 'CITI4') {
        options.push('En hab CITI4', 'OK COSTURA');
    }
    return options;
}

function getDefaultEstadoCosturaFromHab(row) {
    const habNorm = String((row && row.estadoHabilitado) || '').trim().toUpperCase();
    if (habNorm === 'PROG 1T' || habNorm === 'PROG 2T' || habNorm === 'PROG 3T') {
        return 'En Habilitado';
    }
    return '';
}

function renderEstadoCosturaSelect(row, editable = true) {
    const current = normalizeEstadoCostura(row.estadoCostura) || getDefaultEstadoCosturaFromHab(row);
    if (!editable) {
        return `<span class="estado-costura-readonly">${escapeHtml(current || '-')}</span>`;
    }
    const optionsHtml = getEstadoCosturaOptionsForCurrentFilter().map(option => {
        const label = option || 'Seleccionar';
        const selected = option === current ? ' selected' : '';
        return `<option value="${escapeHtmlAttr(option)}"${selected}>${escapeHtml(label)}</option>`;
    }).join('');

    return `<select class="estado-costura-select" ${buildRowUpdateMetaAttrs(row)} data-col-name="estado_costura" data-current="${escapeHtmlAttr(current)}">${optionsHtml}</select>`;
}

function renderLiquidacionCheckbox(row, editable = true) {
    if (!editable) {
        return '<input type="checkbox" class="liquidacion-checkbox" aria-label="Liquidar corte" disabled title="Solo lectura">';
    }
    return `<input type="checkbox" class="liquidacion-checkbox" aria-label="Liquidar corte" ${buildRowUpdateMetaAttrs(row)}>`;
}

function getLocalFieldByColName(colName) {
    const key = normalizeHeader(colName);
    if (key === 'PCOST') return 'pcost';
    if (key === 'PLANTA') return 'planta';
    if (key === 'LINEA') return 'linea';
    if (key === 'STOCKINICIO') return 'stockInicio';
    if (key === 'SALIDADELINEA') return 'salidaDeLinea';
    if (key === 'ESTADOCOSTURA') return 'estadoCostura';
    return null;
}

function recomputeRowDerivedFields(row) {
    if (!row) return;
    row.stockFinalDisplay = formatStockFinal(row.stockInicio, row.salidaDeLinea);
}

function updateLocalRowValue(rowIndex, colName, value) {
    const field = getLocalFieldByColName(colName);
    if (!field) return null;
    const target = allData.find(r => Number(r.rowIndex) === Number(rowIndex));
    if (!target) return null;
    target[field] = value;
    recomputeRowDerivedFields(target);
    ocSearchDataStamp += 1;
    ocSearchState = null;
    return target;
}

function removeLocalRow(rowIndex) {
    const idx = allData.findIndex(r => Number(r.rowIndex) === Number(rowIndex));
    if (idx === -1) return null;
    const removed = allData[idx];
    allData.splice(idx, 1);
    ocSearchDataStamp += 1;
    ocSearchState = null;
    return removed;
}

function resequenceLocalPcostByLinea(plantaValue, lineaValue) {
    const plantNorm = normalizePlantFilterValue(plantaValue);
    const lineaNorm = String(lineaValue || '').trim();
    if (!lineaNorm) return;

    const rows = allData.filter(row => {
        return normalizePlantFilterValue(row.planta) === plantNorm
            && String(row.linea || '').trim() === lineaNorm;
    });

    if (!rows.length) return;

    rows.sort((a, b) => compareRowsWithinLinea(a, b));
    rows.forEach((row, idx) => {
        row.pcost = idx + 1;
    });

    ocSearchDataStamp += 1;
    ocSearchState = null;
}

function readRowMeta(dataset) {
    const rowIndex = parseInt(dataset.rowIndex, 10);
    if (!Number.isFinite(rowIndex) || rowIndex < 2) return null;
    return {
        rowIndex,
        sourceOp: dataset.sourceOp || '',
        sourceCorte: dataset.sourceCorte || '',
        sourceOpTela: dataset.sourceOpTela || '',
        sourcePartida: dataset.sourcePartida || '',
        sourceColor: dataset.sourceColor || '',
        sourcePds: dataset.sourcePds || ''
    };
}

async function saveCellToSheet(meta, colName, value) {
    if (!canEditRowMeta(meta)) {
        throw new Error('No autorizado para editar esta vista');
    }
    const payload = {
        action: 'update',
        row: meta.rowIndex - 1,
        colName: colName,
        value: value,
        sourceOp: meta.sourceOp,
        sourceCorte: meta.sourceCorte,
        sourceOpTela: meta.sourceOpTela,
        sourcePartida: meta.sourcePartida,
        sourceColor: meta.sourceColor
    };

    const response = await fetch(WEB_APP_URL, {
        method: 'POST',
        body: JSON.stringify(payload)
    });
    if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
    }

    const result = await response.json();
    if (!result || result.result !== 'success') {
        throw new Error((result && result.message) || 'No se pudo guardar');
    }

    return result;
}

function normalizeSplitCorteBase(corteValue) {
    const raw = String(corteValue || '').trim();
    if (!raw) return '';
    const dotMatch = raw.match(/^(.*)\.(\d+)$/);
    if (dotMatch) {
        return dotMatch[1];
    }
    const parts = raw.split('-');
    if (parts.length >= 2 && /^\d+$/.test(parts[parts.length - 1])) {
        return parts.slice(0, -1).join('-');
    }
    return raw;
}

function normalizeSplitLineaValue(lineaValue) {
    return String(lineaValue || '')
        .trim()
        .toUpperCase()
        .replace(/[^0-9A-Z]/g, '');
}

function getRowByRowIndex(rowIndex) {
    return allData.find(row => Number(row.rowIndex) === Number(rowIndex)) || null;
}

function getSplitOcBaseFromRow(row) {
    const op = String((row && row.op) || '').trim();
    const corteBase = normalizeSplitCorteBase(row && row.corte);
    const baseOc = op && corteBase ? `${op}-${corteBase}` : (op || corteBase);
    return {
        op,
        corteBase,
        baseOc
    };
}

function formatSplitOcLabel(op, corteBase, suffixIndex) {
    const suffix = Number(suffixIndex);
    const safeSuffix = Number.isFinite(suffix) && suffix > 0 ? suffix : 1;
    const safeOp = String(op || '').trim();
    const safeBase = String(corteBase || '').trim();
    return safeOp && safeBase ? `${safeOp}-${safeBase}.${safeSuffix}` : '';
}

function ensureSplitOcModalInitialized() {
    if (splitOcModalRefs.overlay) return;

    splitOcModalRefs.overlay = document.getElementById('split-oc-modal-overlay');
    splitOcModalRefs.title = document.getElementById('split-oc-modal-title');
    splitOcModalRefs.subtitle = document.getElementById('split-oc-modal-subtitle');
    splitOcModalRefs.note = document.getElementById('split-oc-modal-note');
    splitOcModalRefs.total = document.getElementById('split-oc-modal-total');
    splitOcModalRefs.tbody = document.getElementById('split-oc-modal-tbody');
    splitOcModalRefs.btnCancel = document.getElementById('split-oc-modal-cancel');
    splitOcModalRefs.btnSave = document.getElementById('split-oc-modal-save');

    if (!splitOcModalRefs.overlay || !splitOcModalRefs.tbody) return;

    splitOcModalRefs.tbody.addEventListener('click', onSplitOcModalClick);
    splitOcModalRefs.tbody.addEventListener('input', onSplitOcModalInput);
}

function ensureSplitOcContextMenuInitialized() {
    if (splitOcContextMenuState.menu) return;
    splitOcContextMenuState.menu = document.getElementById('split-oc-context-menu');
}

function updateSplitOcModalSummary() {
    if (!splitOcModalState || !splitOcModalRefs.total) return;
    const parts = Array.isArray(splitOcModalState.parts) ? splitOcModalState.parts : [];
    let total = 0;
    let hasPending = false;

    parts.forEach(part => {
        const raw = String(part && part.pds !== undefined && part.pds !== null ? part.pds : '').trim();
        if (!raw) {
            hasPending = true;
            return;
        }
        if (!/^\d+$/.test(raw)) {
            hasPending = true;
            return;
        }
        total += parseInt(raw, 10);
    });

    const originalRaw = splitOcModalState.originalPds;
    const original = (originalRaw === null || originalRaw === undefined || originalRaw === '')
        ? null
        : Number(originalRaw);
    const totalLabel = formatNumber(total);
    const originalLabel = original === null ? '-' : formatNumber(original);
    splitOcModalRefs.total.textContent = `Total PDS: ${totalLabel} / ${originalLabel}`;
    splitOcModalRefs.total.classList.toggle('ok', original !== null && total === original && !hasPending);
    splitOcModalRefs.total.classList.toggle('warn', original !== null && (total !== original || hasPending));

    if (splitOcModalRefs.note) {
        if (original === null) {
            splitOcModalRefs.note.textContent = 'Complete los PDS y use + para agregar mas particiones.';
        } else if (total === original && !hasPending) {
            splitOcModalRefs.note.textContent = 'La suma coincide con el PDS original.';
        } else {
            splitOcModalRefs.note.textContent = 'La suma de PDS debe coincidir con el total original.';
        }
    }
}

function renderSplitOcModalRows() {
    if (!splitOcModalRefs.tbody || !splitOcModalState) return;
    const parts = Array.isArray(splitOcModalState.parts) && splitOcModalState.parts.length
        ? splitOcModalState.parts
        : [{
            pds: '',
            linea: String(splitOcModalState.sourceLinea || '').trim()
        }];
    splitOcModalState.parts = parts;

    const op = String(splitOcModalState.op || '').trim();
    const corteBase = String(splitOcModalState.corteBase || '').trim();

    splitOcModalRefs.tbody.innerHTML = parts.map((part, idx) => {
        const ocLabel = formatSplitOcLabel(op, corteBase, idx + 1);
        const rawLinea = String(part && part.linea !== undefined && part.linea !== null ? part.linea : '').trim();
        const rawPds = String(part && part.pds !== undefined && part.pds !== null ? part.pds : '').trim();
        const actionButton = idx === 0
            ? '<button type="button" class="split-oc-row-btn add" data-split-action="add" title="Agregar otra particion">+</button>'
            : `<button type="button" class="split-oc-row-btn remove" data-split-action="remove" data-split-index="${idx}" title="Eliminar particion">&times;</button>`;

        return `
            <tr data-split-index="${idx}">
                <td class="split-oc-label" title="${escapeHtmlAttr(ocLabel)}">${escapeHtml(ocLabel)}</td>
                <td>
                    <input type="text" class="split-oc-linea-input" inputmode="text" maxlength="12" autocomplete="off"
                        data-split-index="${idx}" value="${escapeHtmlAttr(rawLinea)}" placeholder="Ingrese LINEA">
                </td>
                <td>
                    <input type="text" class="split-oc-pds-input" inputmode="numeric" pattern="\\d*"
                        data-split-index="${idx}" value="${escapeHtmlAttr(rawPds)}" placeholder="Ingrese PDS">
                </td>
                <td class="split-oc-action-cell">${actionButton}</td>
            </tr>
        `;
    }).join('');

    updateSplitOcModalSummary();
}

function closeSplitOcContextMenu() {
    ensureSplitOcContextMenuInitialized();
    const menu = splitOcContextMenuState.menu;
    if (menu) {
        menu.classList.remove('active');
        menu.setAttribute('aria-hidden', 'true');
    }
    splitOcContextMenuState.rowIndex = null;
}

function positionSplitOcContextMenu(event) {
    const menu = splitOcContextMenuState.menu;
    if (!menu) return;

    const menuWidth = menu.offsetWidth || 220;
    const menuHeight = menu.offsetHeight || 48;
    const margin = 8;
    const maxLeft = Math.max(margin, window.innerWidth - menuWidth - margin);
    const maxTop = Math.max(margin, window.innerHeight - menuHeight - margin);
    const left = Math.min(Math.max(margin, event.clientX), maxLeft);
    const top = Math.min(Math.max(margin, event.clientY), maxTop);

    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;
}

function openSplitOcContextMenu(event, rowIndex) {
    ensureSplitOcContextMenuInitialized();
    const menu = splitOcContextMenuState.menu;
    if (!menu) return false;

    closeSplitOcContextMenu();
    splitOcContextMenuState.rowIndex = rowIndex;
    menu.classList.add('active');
    menu.setAttribute('aria-hidden', 'false');

    if (event) {
        positionSplitOcContextMenu(event);
    }

    return true;
}

function closeSplitOcModal() {
    ensureSplitOcModalInitialized();
    const overlay = splitOcModalRefs.overlay;
    if (overlay) {
        overlay.classList.remove('active');
        overlay.setAttribute('aria-hidden', 'true');
    }
    splitOcModalState = null;
}

function openSplitOcModalFromRow(rowIndex) {
    const row = getRowByRowIndex(rowIndex);
    if (!row) return;
    if (!canEditRowMeta({ rowIndex })) return;

    const baseInfo = getSplitOcBaseFromRow(row);
    if (!baseInfo.baseOc) {
        alert('No se pudo resolver la OC de la fila seleccionada.');
        return;
    }
    const originalPds = parseEditableInteger(row.pds);
    splitOcModalState = {
        rowIndex: Number(row.rowIndex),
        sourceOp: String(row.op || '').trim(),
        sourceCorte: String(row.corte || '').trim(),
        sourceOpTela: String(row.opTela || '').trim(),
        sourcePartida: String(row.partida || '').trim(),
        sourceColor: String(row.color || '').trim(),
        sourceLinea: String(row.linea || '').trim(),
        sourcePds: String(row.pds === null || row.pds === undefined ? '' : row.pds).trim(),
        op: baseInfo.op,
        corteBase: baseInfo.corteBase,
        baseOc: baseInfo.baseOc,
        color: String(row.color || '').trim(),
        originalPds: Number.isFinite(originalPds) ? originalPds : null,
        parts: [{
            pds: Number.isFinite(originalPds) ? String(originalPds) : '',
            linea: String(row.linea || '').trim()
        }]
    };

    ensureSplitOcModalInitialized();
    if (!splitOcModalRefs.overlay) return;

    if (splitOcModalRefs.title) {
        splitOcModalRefs.title.textContent = `OC: ${splitOcModalState.baseOc || '-'}`;
    }
    if (splitOcModalRefs.subtitle) {
        splitOcModalRefs.subtitle.textContent = `COLOR: ${splitOcModalState.color || '-'} | PDS: ${splitOcModalState.originalPds === null ? '-' : formatNumber(splitOcModalState.originalPds)}`;
    }
    if (splitOcModalRefs.note) {
        splitOcModalRefs.note.textContent = 'La primera fila se guardara en la fila original como sufijo .1. Puede editar LINEA por fila.';
    }

    renderSplitOcModalRows();
    splitOcModalRefs.overlay.classList.add('active');
    splitOcModalRefs.overlay.setAttribute('aria-hidden', 'false');
}

function onSplitOcModalClick(event) {
    const button = event.target.closest('button[data-split-action]');
    if (!button || !splitOcModalState) return;

    const action = String(button.dataset.splitAction || '').trim().toLowerCase();
    if (action === 'add') {
        splitOcModalState.parts.push({
            pds: '',
            linea: String(splitOcModalState.sourceLinea || '').trim()
        });
        renderSplitOcModalRows();
        const lastInput = splitOcModalRefs.tbody ? splitOcModalRefs.tbody.querySelector('tr:last-child input.split-oc-linea-input') : null;
        if (lastInput) {
            setTimeout(() => {
                try {
                    lastInput.focus();
                    lastInput.select();
                } catch (e) { }
            }, 0);
        }
        return;
    }

    if (action === 'remove') {
        const idx = parseInt(button.dataset.splitIndex, 10);
        if (!Number.isInteger(idx) || idx < 1) return;
        splitOcModalState.parts.splice(idx, 1);
        renderSplitOcModalRows();
    }
}

function onSplitOcModalInput(event) {
    const lineaInput = event.target.closest('input.split-oc-linea-input');
    if (lineaInput && splitOcModalState) {
        const idx = parseInt(lineaInput.dataset.splitIndex, 10);
        if (!Number.isInteger(idx) || idx < 0 || !splitOcModalState.parts[idx]) return;

        const sanitized = normalizeSplitLineaValue(lineaInput.value);
        if (lineaInput.value !== sanitized) {
            lineaInput.value = sanitized;
        }
        splitOcModalState.parts[idx].linea = sanitized;
        return;
    }

    const input = event.target.closest('input.split-oc-pds-input');
    if (!input || !splitOcModalState) return;

    const idx = parseInt(input.dataset.splitIndex, 10);
    if (!Number.isInteger(idx) || idx < 0 || !splitOcModalState.parts[idx]) return;

    const sanitized = String(input.value || '').replace(/[^\d]/g, '');
    if (input.value !== sanitized) {
        input.value = sanitized;
    }
    splitOcModalState.parts[idx].pds = sanitized;
    updateSplitOcModalSummary();
}

window.abrirSplitOcModalDesdeMenu = function () {
    const rowIndex = splitOcContextMenuState.rowIndex;
    closeSplitOcContextMenu();
    if (!Number.isInteger(rowIndex) || rowIndex < 1) return;
    openSplitOcModalFromRow(rowIndex);
};

window.addSplitOcRow = function () {
    if (!splitOcModalState) return;
    splitOcModalState.parts.push({
        pds: '',
        linea: String(splitOcModalState.sourceLinea || '').trim()
    });
    renderSplitOcModalRows();
};

window.removeSplitOcRow = function (idx) {
    if (!splitOcModalState || !Array.isArray(splitOcModalState.parts)) return;
    if (!Number.isInteger(idx) || idx < 1) return;
    splitOcModalState.parts.splice(idx, 1);
    renderSplitOcModalRows();
};

window.cerrarSplitOcModal = function () {
    closeSplitOcModal();
};

window.guardarSplitOcModal = async function () {
    try {
        ensureSplitOcModalInitialized();
        if (!splitOcModalState || !Array.isArray(splitOcModalState.parts) || splitOcModalState.parts.length === 0) {
            alert('No hay particiones para guardar.');
            return;
        }

        const payloadParts = [];
        let total = 0;
        for (let i = 0; i < splitOcModalState.parts.length; i++) {
            const rawLinea = String(splitOcModalState.parts[i].linea || '').trim();
            const raw = String(splitOcModalState.parts[i].pds || '').trim();
            const linea = normalizeSplitLineaValue(rawLinea || splitOcModalState.sourceLinea || '');
            if (!/^\d+$/.test(raw)) {
                alert(`Ingrese un PDS entero en la fila ${i + 1}.`);
                return;
            }
            const pds = parseInt(raw, 10);
            if (!Number.isFinite(pds) || pds < 0) {
                alert(`Ingrese un PDS valido en la fila ${i + 1}.`);
                return;
            }
            payloadParts.push({
                pds: pds,
                linea: linea
            });
            total += pds;
        }

        if (splitOcModalState.originalPds !== null && total !== splitOcModalState.originalPds) {
            alert(`La suma de PDS (${formatNumber(total)}) debe coincidir con el total original (${formatNumber(splitOcModalState.originalPds)}).`);
            return;
        }

        if (splitOcModalRefs.btnSave) splitOcModalRefs.btnSave.disabled = true;

        const response = await fetch(WEB_APP_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'splitOC',
                row: splitOcModalState.rowIndex - 1,
                baseCorte: splitOcModalState.corteBase,
                parts: payloadParts,
                sourceOp: splitOcModalState.sourceOp,
                sourceCorte: splitOcModalState.sourceCorte,
                sourceOpTela: splitOcModalState.sourceOpTela,
                sourcePartida: splitOcModalState.sourcePartida,
                sourceColor: splitOcModalState.sourceColor,
                sourcePds: splitOcModalState.sourcePds
            })
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const result = await response.json();
        if (!result || result.result !== 'success') {
            throw new Error((result && result.message) || 'No se pudo partir la OC');
        }

        closeSplitOcModal();
        showCosturaToast('OC partida correctamente', 'success');
        await reloadData({ preferWebApp: true });
    } catch (err) {
        console.error('Error guardando split OC:', err);
        alert((err && err.message) ? err.message : 'No se pudo partir la OC.');
    } finally {
        if (splitOcModalRefs.btnSave) splitOcModalRefs.btnSave.disabled = false;
    }
};

function onSplitOcGlobalPointerDown(event) {
    ensureSplitOcContextMenuInitialized();
    const menu = splitOcContextMenuState.menu;
    if (!menu || !menu.classList.contains('active')) return;
    if (menu.contains(event.target)) return;
    closeSplitOcContextMenu();
}

function onSplitOcGlobalKeyDown(event) {
    if (event.key !== 'Escape') return;
    closeSplitOcContextMenu();
    closeSplitOcModal();
    closeDevolucionOcModal();
}

function ensureDevolucionOcModalInitialized() {
    if (devolucionOcModalRefs.overlay) return;

    devolucionOcModalRefs.overlay = document.getElementById('devolucion-oc-modal-overlay');
    devolucionOcModalRefs.title = document.getElementById('devolucion-oc-modal-title');
    devolucionOcModalRefs.subtitle = document.getElementById('devolucion-oc-modal-subtitle');
    devolucionOcModalRefs.tbody = document.getElementById('devolucion-oc-modal-tbody');
    devolucionOcModalRefs.plantaSelect = document.getElementById('devolucion-oc-planta');
    devolucionOcModalRefs.btnCancel = document.getElementById('devolucion-oc-modal-cancel');
    devolucionOcModalRefs.btnSave = document.getElementById('devolucion-oc-modal-save');

    if (!devolucionOcModalRefs.overlay || !devolucionOcModalRefs.tbody || !devolucionOcModalRefs.plantaSelect) return;
}

function getDevolucionViewLabel() {
    const view = String(selectedPlantaFilter || '').trim();
    if (!view) return 'COFACO';
    if (typeof getPlantFilterLabel === 'function') {
        return getPlantFilterLabel(view);
    }
    return view;
}

function buildDevolucionObservationLabel() {
    return `DEVOLUCION DE CORTE DESDE ${getDevolucionViewLabel()}`;
}

function closeDevolucionOcModal() {
    ensureDevolucionOcModalInitialized();
    const overlay = devolucionOcModalRefs.overlay;
    if (overlay) {
        overlay.classList.remove('active');
        overlay.setAttribute('aria-hidden', 'true');
    }
    devolucionOcModalState = null;
}

function renderDevolucionOcModal() {
    if (!devolucionOcModalState || !devolucionOcModalRefs.tbody) return;

    const row = devolucionOcModalState.row;
    const cliente = String(row.cliente || '').trim();
    const oc = String(row.op || '').trim() && String(row.corte || '').trim()
        ? `${String(row.op || '').trim()}-${String(row.corte || '').trim()}`
        : String(row.corte || '').trim();
    const pds = Number.isFinite(Number(devolucionOcModalState.pds)) ? Number(devolucionOcModalState.pds) : 0;
    const viewLabel = String(devolucionOcModalState.viewLabel || getDevolucionViewLabel()).trim() || 'COFACO';
    const obsText = `DEVOLUCION DE CORTE DESDE ${viewLabel}`;

    devolucionOcModalRefs.tbody.innerHTML = `
        <tr>
            <td class="devolucion-label">CLIENTE</td>
            <td class="devolucion-value">${escapeHtml(cliente || '-')}</td>
        </tr>
        <tr>
            <td class="devolucion-label">OC</td>
            <td class="devolucion-value">${escapeHtml(oc || '-')}</td>
        </tr>
        <tr>
            <td class="devolucion-label">PDS</td>
            <td class="devolucion-value">${escapeHtml(formatNumber(pds))}</td>
        </tr>
        <tr>
            <td class="devolucion-label">PLANTA</td>
            <td class="devolucion-value">
                <select id="devolucion-oc-planta" class="devolucion-select">
                    <option value="COFACO">COFACO</option>
                </select>
            </td>
        </tr>
        <tr>
            <td class="devolucion-label">LINEA</td>
            <td class="devolucion-value devolucion-preview">X</td>
        </tr>
        <tr>
            <td class="devolucion-label">HAB</td>
            <td class="devolucion-value devolucion-preview">X PROG</td>
        </tr>
        <tr>
            <td class="devolucion-label">OBSERVACIONES</td>
            <td class="devolucion-value devolucion-preview">${escapeHtml(obsText)}</td>
        </tr>
    `;

    devolucionOcModalRefs.plantaSelect = document.getElementById('devolucion-oc-planta');
    if (devolucionOcModalRefs.plantaSelect) {
        devolucionOcModalRefs.plantaSelect.value = String(devolucionOcModalState.planta || 'COFACO');
    }

    if (devolucionOcModalRefs.title) {
        devolucionOcModalRefs.title.textContent = `OC: ${oc || '-'}`;
    }
    if (devolucionOcModalRefs.subtitle) {
        devolucionOcModalRefs.subtitle.textContent = `CLIENTE: ${cliente || '-'} | PDS: ${formatNumber(pds)}`;
    }
}

function openDevolucionOcModalFromRow(rowIndex) {
    const row = getRowByRowIndex(rowIndex);
    if (!row) return;
    if (!canEditRowMeta({ rowIndex })) return;

    const oc = String(row.op || '').trim() && String(row.corte || '').trim()
        ? `${String(row.op || '').trim()}-${String(row.corte || '').trim()}`
        : String(row.corte || '').trim();

    devolucionOcModalState = {
        rowIndex: Number(row.rowIndex),
        row: row,
        planta: 'COFACO',
        oc: oc,
        pds: Number(row.pds) || 0,
        viewLabel: getDevolucionViewLabel()
    };

    ensureDevolucionOcModalInitialized();
    if (!devolucionOcModalRefs.overlay) return;

    renderDevolucionOcModal();
    devolucionOcModalRefs.overlay.classList.add('active');
    devolucionOcModalRefs.overlay.setAttribute('aria-hidden', 'false');
}

window.abrirDevolucionOcModalDesdeMenu = function () {
    const rowIndex = splitOcContextMenuState.rowIndex;
    closeSplitOcContextMenu();
    if (!Number.isInteger(rowIndex) || rowIndex < 1) return;
    openDevolucionOcModalFromRow(rowIndex);
};

window.cerrarDevolucionOcModal = function () {
    closeDevolucionOcModal();
};

window.generarDevolucionOcModal = async function () {
    try {
        ensureDevolucionOcModalInitialized();
        if (!devolucionOcModalState) return;

        const plantaSel = devolucionOcModalRefs.plantaSelect;
        const targetPlanta = String(plantaSel && plantaSel.value ? plantaSel.value : 'COFACO').trim() || 'COFACO';
        const viewLabel = String(devolucionOcModalState.viewLabel || getDevolucionViewLabel()).trim() || 'COFACO';
        const obsText = `DEVOLUCION DE CORTE DESDE ${viewLabel}`;

        if (devolucionOcModalRefs.btnSave) devolucionOcModalRefs.btnSave.disabled = true;

        const response = await fetch(WEB_APP_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'devolucionCorte',
                row: devolucionOcModalState.rowIndex - 1,
                sourceOp: String(devolucionOcModalState.row.op || '').trim(),
                sourceCorte: String(devolucionOcModalState.row.corte || '').trim(),
                sourceOpTela: String(devolucionOcModalState.row.opTela || '').trim(),
                sourcePartida: String(devolucionOcModalState.row.partida || '').trim(),
                sourceColor: String(devolucionOcModalState.row.color || '').trim(),
                sourcePds: String(devolucionOcModalState.row.pds === null || devolucionOcModalState.row.pds === undefined ? '' : devolucionOcModalState.row.pds).trim(),
                targetPlanta: targetPlanta,
                obsText: obsText,
                viewLabel: viewLabel
            })
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const result = await response.json();
        if (!result || result.result !== 'success') {
            throw new Error((result && result.message) || 'No se pudo generar la devolucion');
        }

        closeDevolucionOcModal();
        showCosturaToast('Devolucion generada correctamente', 'success');
        await reloadData({ preferWebApp: true });
    } catch (err) {
        console.error('Error generando devolucion:', err);
        alert((err && err.message) ? err.message : 'No se pudo generar la devolucion.');
    } finally {
        if (devolucionOcModalRefs.btnSave) devolucionOcModalRefs.btnSave.disabled = false;
    }
};

function onDevolucionOcGlobalPointerDown(event) {
    ensureDevolucionOcModalInitialized();
    const overlay = devolucionOcModalRefs.overlay;
    if (!overlay || !overlay.classList.contains('active')) return;
    if (overlay.contains(event.target)) return;
    closeDevolucionOcModal();
}

function findNextEditableCellInColumn(currentCell, colName) {
    const tbody = currentCell ? currentCell.closest('tbody') : null;
    if (!tbody) return null;
    const sameColumnCells = Array.from(tbody.querySelectorAll('td.editable-int-cell')).filter(cell => {
        return String(cell.dataset.colName || '') === String(colName || '');
    });
    const currentIdx = sameColumnCells.indexOf(currentCell);
    if (currentIdx === -1) return null;
    return sameColumnCells[currentIdx + 1] || null;
}

function openEditableIntCellEditor(intCell, options = {}) {
    if (!intCell || intCell.querySelector('input.int-editor')) return;

    const meta = readRowMeta(intCell.dataset);
    const colName = intCell.dataset.colName || '';
    if (!meta || !colName) return;
    if (!canEditRowMeta(meta)) {
        return;
    }

    const allowEmpty = options.allowEmpty === true;
    const previousRaw = formatEditableIntegerRaw(intCell.dataset.raw || '');
    const startValue = previousRaw;
    const isPcostColumn = normalizeHeader(colName) === 'PCOST';
    const normalizedColName = normalizeHeader(colName);
    const shouldMoveNextOnEnter = normalizedColName === 'STOCKINICIO' || normalizedColName === 'SALIDADELINEA';

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'int-editor';
    input.inputMode = 'numeric';
    input.pattern = '\\d*';
    input.value = startValue;

    intCell.textContent = '';
    intCell.appendChild(input);
    input.focus();
    input.select();

    let finished = false;
    const restoreCell = rawValue => {
        const nextRaw = formatEditableIntegerRaw(rawValue);
        intCell.dataset.raw = nextRaw;
        intCell.textContent = formatEditableIntegerDisplay(nextRaw);
    };

    const commit = async (commitOptions = {}) => {
        if (finished) return;
        const moveNext = commitOptions.moveNext === true && shouldMoveNextOnEnter;
        const entered = String(input.value || '').trim();

        if (entered === '') {
            if (!allowEmpty) {
                cancel();
                return;
            }

            finished = true;
            input.disabled = true;

            try {
                let updatedRow = null;
                if (previousRaw !== '') {
                    await saveCellToSheet(meta, colName, '');
                    updatedRow = updateLocalRowValue(meta.rowIndex, colName, '');
                }
                restoreCell('');
                if (updatedRow && isPcostColumn) {
                    renderTable();
                }
            } catch (err) {
                console.error('Error guardando valor vacio:', err);
                alert('No se pudo guardar el valor.');
                restoreCell(previousRaw);
            }
            return;
        }

        if (!/^\d+$/.test(entered)) {
            alert('Solo se permite numero entero.');
            input.focus();
            return;
        }

        finished = true;
        const normalized = String(parseInt(entered, 10));
        input.disabled = true;

        try {
            let updatedRow = null;
            if (normalized !== previousRaw) {
                await saveCellToSheet(meta, colName, normalized);
                updatedRow = updateLocalRowValue(meta.rowIndex, colName, normalized);
            }
            restoreCell(normalized);
            if (updatedRow && isPcostColumn) {
                renderTable();
                return;
            }
            if (updatedRow) {
                const tr = intCell.closest('tr');
                const finalCell = tr ? tr.querySelector('.s-final-cell') : null;
                if (finalCell) {
                    finalCell.innerHTML = `<strong>${updatedRow.stockFinalDisplay}</strong>`;
                }
                updateRenderedLineaBandTotal(updatedRow.linea);
                updateRenderedSectorBandTotal(getSectorNameForPlant(updatedRow.linea, selectedPlantaFilter));
                updateStats();
            }
            if (moveNext) {
                const nextCell = findNextEditableCellInColumn(intCell, colName);
                if (nextCell) {
                    setTimeout(() => {
                        openEditableIntCellEditor(nextCell, {
                            allowEmpty: nextCell.dataset.allowEmpty === '1'
                        });
                    }, 0);
                }
            }
        } catch (err) {
            console.error('Error guardando valor entero:', err);
            alert('No se pudo guardar el valor.');
            restoreCell(previousRaw);
        }
    };

    const cancel = () => {
        if (finished) return;
        finished = true;
        restoreCell(previousRaw);
    };

    input.addEventListener('blur', commit);
    input.addEventListener('keydown', eventKey => {
        if (eventKey.key === 'Enter') {
            eventKey.preventDefault();
            commit({ moveNext: true });
        } else if (eventKey.key === 'Escape') {
            eventKey.preventDefault();
            cancel();
        }
    });
}

async function fetchShiftState() {
    try {
        const response = await fetch(`${WEB_APP_URL}?action=getShiftState`);
        const data = await response.json();
        const stateStr = data.state || "";
        activeShifts = sanitizeShiftStateTokens(parseShiftStateTokens(stateStr));
        renderShiftPills();
    } catch (err) {
        console.error('Error fetching shift state:', err);
    }
}

function renderShiftPills() {
    const profile = getActiveAccessProfile();
    const shiftOwner = getShiftUserKeyFromAccessProfile();
    const headerTotals = document.querySelector('.header-totals');
    if (!headerTotals) return;

    let container = document.getElementById('shift-pills-container');

    if (!shiftOwner) {
        if (container) container.remove();
        return;
    }

    if (!container) {
        container = document.createElement('div');
        container.id = 'shift-pills-container';
        container.className = 'shift-pills-container';
        // Insertar al inicio de header-totals
        headerTotals.insertBefore(container, headerTotals.firstChild);
    }

    container.innerHTML = getShiftPillOptionsForPlant(shiftOwner).map(s => {
        const active = hasShiftTokenForCurrentPlant(s) ? ' active' : '';
        return `
            <div class="shift-pill${active}" onclick="toggleShift('${s}')">
                <span class="check-icon">✔</span>
                ${s}
            </div>
        `;
    }).join('');
}

window.toggleShift = async function (shift) {
    const plant = getShiftUserKeyFromAccessProfile();
    if (!plant) return;
    if (!isShiftOptionEnabledForCurrentPlant(shift)) return;

    const wasActive = hasShiftTokenForCurrentPlant(shift);
    if (wasActive) {
        removeShiftTokensForPlantShift(plant, shift);
    } else {
        addShiftTokenForPlant(plant, shift);
    }
    renderShiftPills(); // Actualización optimista

    activeShifts = sanitizeShiftStateTokens(activeShifts);
    const stateStr = serializeShiftStateTokens(activeShifts);
    showCosturaToast(`Actualizando ${plant}-${shift}...`, 'info');

    try {
        const response = await fetch(WEB_APP_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'updateShiftState',
                state: stateStr
            })
        });
        const result = await response.json();
        if (result.result === 'success') {
            showCosturaToast(`Turno ${plant}-${shift} actualizado`, 'success');
        } else {
            throw new Error(result.message || 'Error en servidor');
        }
    } catch (err) {
        console.error('Error updating shift state:', err);
        showCosturaToast('Error al sincronizar turno', 'error');
        // Revertir localmente en caso de error
        if (wasActive) {
            addShiftTokenForPlant(plant, shift);
        } else {
            removeShiftTokensForPlantShift(plant, shift);
        }
        renderShiftPills();
    }
};

function initTableInteractions() {
    const tbody = document.getElementById('tbody-costura');
    if (!tbody || tbody.dataset.interactionsBound === '1') return;
    tbody.dataset.interactionsBound = '1';

    tbody.addEventListener('click', event => {
        const intCell = event.target.closest('td.pcost-editable-cell');
        if (!intCell) return;
        openEditableIntCellEditor(intCell, {
            allowEmpty: intCell.dataset.allowEmpty === '1'
        });
    });

    tbody.addEventListener('dblclick', event => {
        const intCell = event.target.closest('td.editable-int-cell');
        if (intCell) {
            if (intCell.classList.contains('pcost-editable-cell')) return;
            openEditableIntCellEditor(intCell, {
                allowEmpty: intCell.dataset.allowEmpty === '1'
            });
            return;
        }

        const lineaCell = event.target.closest('td.linea-editable-cell');
        if (!lineaCell) return;

        const meta = readRowMeta(lineaCell.dataset);
        if (!meta) return;
        const currentLinea = String(lineaCell.dataset.currentLinea || '').trim();
        const currentPlanta = String(lineaCell.dataset.currentPlanta || '').trim();
        openLineaModal(meta, currentPlanta, currentLinea);
    });

    tbody.addEventListener('contextmenu', event => {
        const colorCell = event.target.closest('td.color-split-cell');
        if (!colorCell) return;

        const meta = readRowMeta(colorCell.dataset);
        if (!meta || !canEditRowMeta(meta)) return;

        event.preventDefault();
        event.stopPropagation();
        openSplitOcContextMenu(event, meta.rowIndex);
    });

    tbody.addEventListener('change', async event => {
        const liquidacionCheckbox = event.target.closest('input.liquidacion-checkbox');
        if (liquidacionCheckbox) {
            if (!liquidacionCheckbox.checked) return;
            const meta = readRowMeta(liquidacionCheckbox.dataset);
            if (!meta) {
                liquidacionCheckbox.checked = false;
                return;
            }
            openLiquidacionModal(meta, liquidacionCheckbox);
            return;
        }

        const select = event.target.closest('select.estado-costura-select');
        if (!select) return;

        const meta = readRowMeta(select.dataset);
        const colName = select.dataset.colName || 'estado_costura';
        if (!meta) return;

        const previous = normalizeEstadoCostura(select.dataset.current || '');
        const next = normalizeEstadoCostura(select.value);
        if (previous === next) return;

        select.disabled = true;
        try {
            await saveCellToSheet(meta, colName, next);
            select.dataset.current = next;
            select.value = next;
            const updatedRow = updateLocalRowValue(meta.rowIndex, colName, next);
            if (updatedRow && updatedRow.linea) {
                updateRenderedLineaBandTotal(updatedRow.linea);
                updateRenderedSectorBandTotal(getSectorNameForPlant(updatedRow.linea, selectedPlantaFilter));
            }
        } catch (err) {
            console.error('Error guardando estado_costura:', err);
            alert('No se pudo guardar estado_costura.');
            select.value = previous;
        } finally {
            select.disabled = false;
        }
    });
}

function isNoLleva(value) {
    return String(value || '').toUpperCase().trim().includes('NO LLEVA');
}

function cleanCollTapValue(value) {
    let out = String(value || '').trim();
    out = out.replace(/en\s*hab/ig, '').trim();
    out = out.replace(/^[\s\-:|]+|[\s\-:|]+$/g, '').trim();
    return out;
}

function abbreviateCompOtrosText(value) {
    let out = String(value || '').trim();
    if (!out) return '';
    out = out.replace(/\bprenda(?:s)?\b/ig, 'PDA');
    out = out.replace(/\bpieza(?:s)?\b/ig, 'PZA');
    out = out.replace(/\ben\s+PDA\b/ig, 'PDA');
    out = out.replace(/\ben\s+PZA\b/ig, 'PZA');
    out = out.replace(/\s+/g, ' ').trim();
    return out;
}

function normalizeRibCompOtros(value) {
    const raw = String(value || '').trim();
    if (!raw) return '';
    const cleaned = raw.replace(/^RIB\s*:\s*/i, '').trim();
    if (!cleaned) return '';
    const normalized = cleaned
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase()
        .replace(/\s+/g, ' ')
        .trim();
    if (normalized === 'EN HAB' || normalized === 'EN HABILITADO' || normalized === 'SI LLEVA') {
        return '✔';
    }
    return abbreviateCompOtrosText(cleaned);
}

function formatCompOtros(row) {
    const parts = [];

    const ribVal = String(row.rib || '').trim();
    if (ribVal && !isNoLleva(ribVal)) {
        const ribShort = normalizeRibCompOtros(ribVal);
        if (ribShort) parts.push(`RIB: ${ribShort}`);
    }

    const collTapRaw = String(row.collTap || '').trim();
    const collTapClean = abbreviateCompOtrosText(cleanCollTapValue(collTapRaw));
    if (collTapClean && !isNoLleva(collTapRaw)) {
        parts.push(collTapClean);
    }

    const trfVal = abbreviateCompOtrosText(String(row.trsfTipo || '').trim());
    if (trfVal && !isNoLleva(trfVal)) {
        parts.push(`TRF: ${trfVal}`);
    }

    const bdVal = abbreviateCompOtrosText(String(row.tipoBordado || '').trim());
    if (bdVal && !isNoLleva(bdVal)) {
        parts.push(`BD: ${bdVal}`);
    }

    const estmpVal = abbreviateCompOtrosText(String(row.nEstmp || '').trim());
    if (estmpVal && !isNoLleva(estmpVal)) {
        parts.push(`ESTMPADO: ${estmpVal}`);
    }

    return parts.join(' | ');
}

function normalizeOcSearchToken(value) {
    return String(value || '')
        .trim()
        .toUpperCase()
        .replace(/\s+/g, '');
}

function getRowOcCandidates(row) {
    const out = [];
    const push = (value) => {
        const normalized = normalizeOcSearchToken(value);
        if (!normalized) return;
        if (!out.includes(normalized)) out.push(normalized);
    };

    const op = String((row && row.op) || '').trim();
    const corte = String((row && row.corte) || '').trim();
    const opTela = String((row && row.opTela) || '').trim();
    const partida = String((row && row.partida) || '').trim();

    if (op && corte) push(`${op}-${corte}`);
    if (opTela && partida) push(`${opTela}-${partida}`);
    push(op);
    push(corte);

    return out;
}

function rowMatchesOcSearchQuery(row, query) {
    const q = normalizeOcSearchToken(query);
    if (!q) return true;
    const qFlat = q.replace(/-/g, '');
    return getRowOcCandidates(row).some(candidate => {
        const cFlat = candidate.replace(/-/g, '');
        return candidate.includes(q) || cFlat.includes(qFlat);
    });
}

function getPlantFilterLabel(filterValue) {
    if (filterValue === PLANTA_FILTER_OTHERS) return 'OTROS';
    return String(filterValue || '').toUpperCase();
}

function showOcSearchMessage(text, isSuccess) {
    const msgEl = document.getElementById('search-oc-msg');
    if (!msgEl) return;
    msgEl.textContent = text;
    msgEl.className = 'search-oc-msg' + (isSuccess ? ' success' : '');
    msgEl.style.display = 'block';
    clearTimeout(window._searchOcMsgTimer);
    window._searchOcMsgTimer = setTimeout(() => {
        msgEl.style.display = 'none';
    }, 2600);
}

window.clearCosturaOcSearch = function (keepInput = false) {
    activeOcSearchQuery = '';
    ocSearchState = null;

    const input = document.getElementById('search-oc-input');
    if (input && !keepInput) input.value = '';

    const msgEl = document.getElementById('search-oc-msg');
    if (msgEl) {
        msgEl.textContent = '';
        msgEl.style.display = 'none';
    }
    try { clearTimeout(window._searchOcMsgTimer); } catch (e) { }

    renderPlantPills();
    updateStats();
    renderTable();
};

function collectOcSearchResults(query) {
    const q = normalizeOcSearchToken(query);
    if (!q) return [];

    const results = [];
    OC_SEARCH_PLANT_ORDER.forEach(filterValue => {
        const rows = getFilteredDataByFilter(filterValue, {
            includeSearch: false,
            query: q
        }).filter(row => rowMatchesOcSearchQuery(row, q))
            .filter(matchesSelectedColorFilter);
        if (rows.length > 0) {
            results.push({
                plantFilter: filterValue,
                plantLabel: getPlantFilterLabel(filterValue),
                matches: rows.length
            });
        }
    });
    return results;
}

window.buscarOCCostura = function (ocQuery) {
    if (!allData || allData.length === 0) {
        showOcSearchMessage('Sin datos cargados', false);
        return;
    }

    const q = normalizeOcSearchToken(ocQuery);
    if (!q) return;

    const canReuseState = !!(
        ocSearchState
        && ocSearchState.query === q
        && Number(ocSearchState.dataStamp || 0) === Number(ocSearchDataStamp || 0)
        && Array.isArray(ocSearchState.results)
        && ocSearchState.results.length > 0
    );

    if (!canReuseState) {
        const results = collectOcSearchResults(q);
        if (!results.length) {
            ocSearchState = {
                query: q,
                dataStamp: ocSearchDataStamp,
                results: [],
                index: -1
            };
            activeOcSearchQuery = q;
            renderPlantPills();
            updateStats();
            renderTable();
            showOcSearchMessage(`OC "${String(ocQuery || '').trim()}" no encontrada`, false);
            return;
        }

        ocSearchState = {
            query: q,
            dataStamp: ocSearchDataStamp,
            results,
            index: -1
        };
    }

    ocSearchState.index = (ocSearchState.index + 1) % ocSearchState.results.length;
    const currentResult = ocSearchState.results[ocSearchState.index];
    selectedPlantaFilter = currentResult.plantFilter;
    activeOcSearchQuery = q;
    renderPlantPills();
    updateStats();
    renderTable();
    showOcSearchMessage(
        `${currentResult.plantLabel}: ${formatNumber(currentResult.matches)} coincidencias (${ocSearchState.index + 1}/${ocSearchState.results.length})`,
        true
    );
};

function formatStatusHabilitado(text) {
    if (!text) return '-';
    const s = String(text).toUpperCase().trim();
    if (s === 'OK') {
        return '<span class="pill pill-ok">OK</span>';
    }
    if (s === 'PROG 1T' || s === 'PROG 2T' || s === 'PROG 3T') {
        return `<span class="pill pill-prog">${s}</span>`;
    }
    return escapeHtml(s);
}

function isLiquidacionCostClosed(value) {
    const normalized = String(value || '').trim().toUpperCase();
    return normalized.startsWith('LIQUIDADO');
}

function normalizeHeader(value) {
    return String(value || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, '');
}

function parseDateValue(rawValue) {
    if (rawValue === null || rawValue === undefined || rawValue === '') return null;
    if (rawValue instanceof Date && !Number.isNaN(rawValue.getTime())) return rawValue;
    if (typeof rawValue === 'number' && Number.isFinite(rawValue)) {
        // Serial de fecha de Excel/Sheets.
        const parsedSerial = new Date(Math.round((rawValue - 25569) * 86400000));
        if (!Number.isNaN(parsedSerial.getTime())) return parsedSerial;
    }

    const dateStr = String(rawValue).trim();
    if (!dateStr) return null;
    const dateStrNorm = dateStr
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase();

    // Mantener comportamiento correcto para ISO con zona horaria.
    if (dateStr.includes('T') || /[zZ]$/.test(dateStr) || /[+-]\d{2}:?\d{2}$/.test(dateStr)) {
        const parsedIso = new Date(dateStr);
        if (!Number.isNaN(parsedIso.getTime())) return parsedIso;
    }

    // Formato gviz: Date(2026,2,20,15,26,0)
    let partsGviz = dateStr.match(/^Date\((\d{4}),\s*(\d{1,2}),\s*(\d{1,2})(?:,\s*(\d{1,2}),\s*(\d{1,2}),\s*(\d{1,2}))?\)$/i);
    if (partsGviz) {
        const year = parseInt(partsGviz[1], 10);
        const monthZeroBased = parseInt(partsGviz[2], 10);
        const day = parseInt(partsGviz[3], 10);
        const hour = parseInt(partsGviz[4] || '0', 10);
        const minute = parseInt(partsGviz[5] || '0', 10);
        const second = parseInt(partsGviz[6] || '0', 10);
        const parsedGviz = new Date(year, monthZeroBased, day, hour, minute, second);
        if (!Number.isNaN(parsedGviz.getTime())) return parsedGviz;
    }

    // dd/mm/yyyy [hh:mm[:ss]]
    let parts = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:[T\s,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (parts) {
        const day = parseInt(parts[1], 10);
        const month = parseInt(parts[2], 10);
        let year = parseInt(parts[3], 10);
        if (year < 100) year += 2000;
        const hour = parseInt(parts[4] || '0', 10);
        const minute = parseInt(parts[5] || '0', 10);
        const second = parseInt(parts[6] || '0', 10);
        const parsedDMY = new Date(year, month - 1, day, hour, minute, second);
        if (!Number.isNaN(parsedDMY.getTime())) return parsedDMY;
    }

    // yyyy-mm-dd [hh:mm[:ss]]
    parts = dateStr.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (parts) {
        const year = parseInt(parts[1], 10);
        const month = parseInt(parts[2], 10);
        const day = parseInt(parts[3], 10);
        const hour = parseInt(parts[4] || '0', 10);
        const minute = parseInt(parts[5] || '0', 10);
        const second = parseInt(parts[6] || '0', 10);
        const parsedYMD = new Date(year, month - 1, day, hour, minute, second);
        if (!Number.isNaN(parsedYMD.getTime())) return parsedYMD;
    }

    // dd/mmm/yy [hh:mm[:ss]][am|pm] con meses ENE..DIC (tambien set)
    const monthMap = {
        ene: 0, feb: 1, mar: 2, abr: 3, may: 4, jun: 5,
        jul: 6, ago: 7, sep: 8, set: 8, oct: 9, nov: 10, dic: 11
    };
    parts = dateStrNorm.match(/^(\d{1,2})[\/\-\s](ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[\/\-\s](\d{2,4})(?:[T\s,]+(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?\s*(am|pm)?)?/);
    if (parts) {
        const day = parseInt(parts[1], 10);
        const month = monthMap[parts[2]];
        let year = parseInt(parts[3], 10);
        if (year < 100) year += 2000;
        let hour = parseInt(parts[4] || '0', 10);
        const minute = parseInt(parts[5] || '0', 10);
        const second = parseInt(parts[6] || '0', 10);
        const ampm = String(parts[7] || '').toLowerCase();
        if (ampm === 'pm' && hour < 12) hour += 12;
        if (ampm === 'am' && hour === 12) hour = 0;
        const parsedDMYTextMonth = new Date(year, month, day, hour, minute, second);
        if (!Number.isNaN(parsedDMYTextMonth.getTime())) return parsedDMYTextMonth;
    }

    const parsedNative = new Date(dateStr);
    if (!Number.isNaN(parsedNative.getTime())) return parsedNative;

    return null;
}

function formatDate(dateStr) {
    const parsed = parseDateValue(dateStr);
    if (!parsed) return '-';
    const day = String(parsed.getDate()).padStart(2, '0');
    const month = MONTH_ABBR_ES[parsed.getMonth()] || '---';
    return `${day}/${month}`;
}

function formatDateTooltip(dateStr) {
    const parsed = parseDateValue(dateStr);
    if (!parsed) return '-';
    const hours24 = parsed.getHours();
    const hours12 = hours24 % 12 || 12;
    const minutes = String(parsed.getMinutes()).padStart(2, '0');
    const ampm = hours24 >= 12 ? 'pm' : 'am';
    return `${String(hours12).padStart(2, '0')}:${minutes}${ampm}`;
}

function getLimaDateInfo(dateObj) {
    const formatter = new Intl.DateTimeFormat('es-PE', {
        timeZone: 'America/Lima',
        weekday: 'short',
        day: '2-digit',
        month: '2-digit',
        hour: '2-digit',
        hour12: false
    });
    const parts = formatter.formatToParts(dateObj);
    const map = {};
    parts.forEach(part => {
        if (part.type !== 'literal') map[part.type] = part.value;
    });

    const weekday = String(map.weekday || '')
        .replace(/\./g, '')
        .trim()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .slice(0, 3)
        .toUpperCase();
    const day = String(map.day || '').padStart(2, '0');
    const monthNum = parseInt(map.month || '1', 10);
    const monthIndex = Number.isFinite(monthNum) ? Math.max(0, Math.min(11, monthNum - 1)) : 0;
    const hour = parseInt(map.hour || '0', 10);

    return {
        weekday: weekday || '---',
        day: day || '--',
        monthIndex,
        hour: Number.isFinite(hour) ? hour : 0
    };
}

function buildSalidaHeaderHtml() {
    const now = new Date();
    const nowInfo = getLimaDateInfo(now);
    const targetDate = nowInfo.hour < 7 ? new Date(now.getTime() - 86400000) : now;
    const targetInfo = getLimaDateInfo(targetDate);
    const month = MONTH_ABBR_ES[targetInfo.monthIndex] || '---';
    return `SALIDA<br>${targetInfo.weekday} ${targetInfo.day}/${month}`;
}

function updateSalidaHeaderLabel() {
    const salidaHeader = document.getElementById('th-salida-linea');
    if (!salidaHeader) return;
    salidaHeader.innerHTML = buildSalidaHeaderHtml();
}

function buildLiquidacionCostStamp() {
    const formatter = new Intl.DateTimeFormat('es-PE', {
        timeZone: 'America/Lima',
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
    });
    const parts = formatter.formatToParts(new Date());
    const map = {};
    parts.forEach(part => {
        if (part.type !== 'literal') map[part.type] = part.value;
    });

    const day = String(map.day || '').padStart(2, '0');
    const monthNum = parseInt(map.month || '1', 10);
    const month = MONTH_ABBR_ES[Math.max(0, Math.min(11, monthNum - 1))] || '---';
    const year = String(map.year || '');
    const hour = String(map.hour || '').padStart(2, '0');
    const minute = String(map.minute || '').padStart(2, '0');
    return `LIQUIDADO ${day}/${month}/${year} ${hour}:${minute}`;
}

function formatColorTooltipValues(colorValue, opTelaValue, partidaValue) {
    const parts = [];
    const color = String(colorValue || '').trim();
    const opTela = String(opTelaValue || '').trim();
    const partida = String(partidaValue || '').trim();
    if (color) parts.push(color);
    if (opTela) parts.push(opTela);
    if (partida) parts.push(partida);
    return parts.join(', ') || '-';
}

function formatColorTooltip(row) {
    return formatColorTooltipValues(row.color, row.opTela, row.partida);
}

function formatNumber(num) {
    return num > 0 ? num.toLocaleString('es-ES', { maximumFractionDigits: 0 }) : '0';
}

function abbreviate(text) {
    if (!text) return '';
    return String(text).substring(0, 15);
}

function normalizePrenda(text) {
    if (!text) return '';
    const s = String(text).toUpperCase().trim();
    return s.substring(0, 20);
}

function normalizeCert(text) {
    if (!text) return '';
    const s = String(text).toUpperCase().trim();
    const map = {
        'SIN CERTIFICADO': 'S/C',
        'OE': 'OE',
        'OCS': 'OCS'
    };
    return map[s] || s.substring(0, 10);
}

function normalizeClientName(clientName) {
    if (!clientName) return '';
    const name = clientName.toString().toUpperCase();
    if (name.includes('LACOSTE')) return 'LAC';
    if (name.includes('ATHLETA, INC.')) return 'ATH';
    if (name.includes('BANANA REPUBLIC, LLC')) return 'BNN';
    if (name.includes('THEORY LLC,')) return 'THE';
    if (name.includes('DISH & DUER')) return 'DDU';
    if (name.includes('SKECHERS PERFORMANCE')) return 'SKE';
    if (name.includes('LULULEMON ATHLETICA CANADA INC')) return 'LLL';
    if (name.includes('AM RETAIL S.A.C.')) return 'AMR';
    if (name.includes('ALLBIRDS')) return 'ALLB';
    return clientName;
}

function formatEstadoAvios(value) {
    const raw = String(value || '').trim();
    if (!raw) return '?';
    const normalized = raw
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase()
        .replace(/\s+/g, ' ')
        .trim();
    if (normalized === 'DESPACHADO') {
        return '<span class="avios-check">✔</span>';
    }
    if (normalized === 'HC PREP' || normalized === 'OK PREP') {
        return 'F';
    }
    return escapeHtml(raw);
}

function formatHiloCostura(value) {
    const raw = String(value || '').trim();
    if (!raw) return '?';
    const normalized = raw
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase()
        .replace(/\s+/g, ' ')
        .trim();

    if (normalized === 'DESPACHADO') {
        return '<span class="avios-check">✔</span>';
    }

    if (
        normalized === 'EN LABORATORIO'
        || normalized === 'POR TENIR'
        || normalized === 'EN CALIDAD'
        || normalized === 'X REENCONAR'
    ) {
        return 'F';
    }

    return escapeHtml(raw);
}

function formatEstado(text) {
    if (!text) return '-';
    const s = String(text).toUpperCase().trim();
    if (s === 'OK ENM' || s === 'OK PAQUETEO') {
        return '<span class="pill pill-ok">OK</span>';
    } else if (s.includes('PROG')) {
        return '<span class="pill pill-prog">PROG</span>';
    }
    return String(text).substring(0, 15);
}

function formatRib(text) {
    if (!text) return '-';
    const s = String(text).toUpperCase().trim();
    if (s === 'NO LLEVA') {
        return '<span class="pill pill-x">X</span>';
    } else if (s === 'EN HAB') {
        return '<span class="pill pill-prog">EN HAB</span>';
    } else if (s === 'OK') {
        return '<span class="pill pill-ok">OK</span>';
    }
    return String(text).substring(0, 15);
}

function formatCollTap(text) {
    if (!text) return '-';
    const s = String(text).toUpperCase().trim();
    if (s === 'NO LLEVA') {
        return '<span class="pill pill-x">X</span>';
    }
    return String(text).substring(0, 15);
}

function formatTrsf(text) {
    if (!text) return '-';
    return String(text).substring(0, 20);
}

function formatBord(text) {
    if (!text) return '-';
    const s = String(text).toUpperCase().trim();
    if (s === 'NO LLEVA') {
        return '<span class="pill pill-x">X</span>';
    } else if (s === 'PROG') {
        return '<span class="pill pill-prog">PROG</span>';
    } else if (s === 'OK') {
        return '<span class="pill pill-ok">OK</span>';
    }
    return String(text).substring(0, 15);
}

function formatEstmp(text) {
    if (!text) return '-';
    const s = String(text).toUpperCase().trim();
    if (s === 'NO LLEVA') {
        return '<span class="pill pill-x">X</span>';
    } else if (s === 'PROG') {
        return '<span class="pill pill-prog">PROG</span>';
    } else if (s === 'OK') {
        return '<span class="pill pill-ok">OK</span>';
    }
    return String(text).substring(0, 15);
}

// Cargar datos al abrir la pagina
window.addEventListener('DOMContentLoaded', async () => {
    try {
        await ensureAccessFromLogin();
    } catch (error) {
        console.error('No se pudo validar acceso:', error);
        alert('No se pudo validar el acceso del usuario.');
        return;
    }

    initTableInteractions();
    ensureLineaModalInitialized();
    ensureLiquidacionModalInitialized();
    ensureColorFilterModalInitialized();
    ensureSplitOcModalInitialized();
    ensureDevolucionOcModalInitialized();
    ensureSplitOcContextMenuInitialized();
    const colorHeader = document.getElementById('th-color-filter');
    if (colorHeader) {
        colorHeader.addEventListener('contextmenu', openColorFilterFromHeader);
    }
    document.addEventListener('mousedown', onSplitOcGlobalPointerDown, true);
    document.addEventListener('touchstart', onSplitOcGlobalPointerDown, true);
    document.addEventListener('mousedown', onDevolucionOcGlobalPointerDown, true);
    document.addEventListener('touchstart', onDevolucionOcGlobalPointerDown, true);
    document.addEventListener('keydown', onSplitOcGlobalKeyDown, true);
    window.addEventListener('resize', closeSplitOcContextMenu, true);
    window.addEventListener('scroll', closeSplitOcContextMenu, { capture: true, passive: true });
    const searchInput = document.getElementById('search-oc-input');
    if (searchInput) {
        searchInput.addEventListener('input', () => {
            if (String(searchInput.value || '').trim() === '' && activeOcSearchQuery) {
                clearCosturaOcSearch(true);
            }
        });
    }
    updateSalidaHeaderLabel();
    setInterval(updateSalidaHeaderLabel, 60000);
    reloadData();
    fetchShiftState();
});

function renderShiftPills() {
    const shiftOwner = getShiftUserKeyFromAccessProfile();
    const headerTotals = document.querySelector('.header-totals');
    if (!headerTotals) return;

    let container = document.getElementById('shift-pills-container');

    if (!shiftOwner) {
        if (container) container.remove();
        return;
    }

    if (!container) {
        container = document.createElement('div');
        container.id = 'shift-pills-container';
        container.className = 'shift-pills-container';
        headerTotals.insertBefore(container, headerTotals.firstChild);
    }

    const shifts = getShiftPillOptionsForPlant(shiftOwner);
    container.innerHTML = shifts.map(s => {
        const active = hasShiftTokenForCurrentPlant(s) ? ' active' : '';
        return `
            <div class="shift-pill${active}" onclick="toggleShift('${s}')">
                <span class="check-icon">&#10003;</span>
                ${s}
            </div>
        `;
    }).join('');
}

window.toggleShift = async function (shift) {
    const plant = getShiftUserKeyFromAccessProfile();
    if (!plant) return;
    if (!isShiftOptionEnabledForCurrentPlant(shift)) return;

    const wasActive = hasShiftTokenForCurrentPlant(shift);
    if (wasActive) {
        removeShiftTokensForPlantShift(plant, shift);
    } else {
        addShiftTokenForPlant(plant, shift);
    }
    renderShiftPills();

    activeShifts = sanitizeShiftStateTokens(activeShifts);
    const stateStr = serializeShiftStateTokens(activeShifts);
    showCosturaToast(`Actualizando ${plant}-${shift}...`, 'info');

    try {
        const response = await fetch(WEB_APP_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'updateShiftState',
                state: stateStr
            })
        });
        const result = await response.json();
        if (result.result === 'success') {
            showCosturaToast(`Turno ${plant}-${shift} actualizado`, 'success');
        } else {
            throw new Error(result.message || 'Error en servidor');
        }
    } catch (err) {
        console.error('Error updating shift state:', err);
        showCosturaToast('Error al sincronizar turno', 'error');
        if (wasActive) {
            addShiftTokenForPlant(plant, shift);
        } else {
            removeShiftTokensForPlantShift(plant, shift);
        }
        renderShiftPills();
    }
};
