(function registerCiti2Role(global) {
    if (!global) return;
    global.STOCK_COSTURA_ROLE_PRESETS = global.STOCK_COSTURA_ROLE_PRESETS || {};
    global.STOCK_COSTURA_ROLE_PRESETS.CITI2 = {
        key: 'CITI2',
        label: 'CITI2',
        editableFilters: ['CITI2'],
        canEditAll: false
    };
})(window);
