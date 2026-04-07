(function registerOtrosRole(global) {
    if (!global) return;
    global.STOCK_COSTURA_ROLE_PRESETS = global.STOCK_COSTURA_ROLE_PRESETS || {};
    global.STOCK_COSTURA_ROLE_PRESETS.OTROS = {
        key: 'OTROS',
        label: 'OTROS',
        editableFilters: ['__OTHERS__'],
        canEditAll: false
    };
})(window);
