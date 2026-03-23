(function registerCiti3Role(global) {
    if (!global) return;
    global.STOCK_COSTURA_ROLE_PRESETS = global.STOCK_COSTURA_ROLE_PRESETS || {};
    global.STOCK_COSTURA_ROLE_PRESETS.CITI3 = {
        key: 'CITI3',
        label: 'CITI3',
        editableFilters: ['CITI3'],
        canEditAll: false
    };
})(window);
