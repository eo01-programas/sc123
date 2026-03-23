(function registerCiti4Role(global) {
    if (!global) return;
    global.STOCK_COSTURA_ROLE_PRESETS = global.STOCK_COSTURA_ROLE_PRESETS || {};
    global.STOCK_COSTURA_ROLE_PRESETS.CITI4 = {
        key: 'CITI4',
        label: 'CITI4',
        editableFilters: ['CITI4'],
        canEditAll: false
    };
})(window);
