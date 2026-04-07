(function registerCofacoRole(global) {
    if (!global) return;
    global.STOCK_COSTURA_ROLE_PRESETS = global.STOCK_COSTURA_ROLE_PRESETS || {};
    global.STOCK_COSTURA_ROLE_PRESETS.COFACO = {
        key: 'COFACO',
        label: 'COFACO',
        editableFilters: ['COFACO'],
        canEditAll: false
    };
})(window);
