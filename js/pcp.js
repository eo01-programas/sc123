(function registerPcpRole(global) {
    if (!global) return;
    global.STOCK_COSTURA_ROLE_PRESETS = global.STOCK_COSTURA_ROLE_PRESETS || {};
    global.STOCK_COSTURA_ROLE_PRESETS.PCP = {
        key: 'PCP',
        label: 'PCP',
        editableFilters: [],
        canEditAll: true
    };
})(window);
