(function initStockCosturaLogin(global) {
    if (!global) return;

    const USER_CREDENTIALS = Object.freeze({
        COFACO: '097mp',
        CITI1: '106vd',
        CITI2: '925rrs',
        CITI3: '977lg',
        CITI4: '104jv',
        PCP: '236wt'
    });
    const LOGIN_USER_ORDER = ['COFACO', 'CITI1', 'CITI2', 'CITI3', 'CITI4', 'PCP'];
    let pendingLoginPromise = null;
    let resolvedProfile = null;

    function normalizeUserKey(value) {
        return String(value || '').trim().toUpperCase();
    }

    function getRoleProfile(userKey) {
        const presets = global.STOCK_COSTURA_ROLE_PRESETS || {};
        if (presets[userKey] && typeof presets[userKey] === 'object') {
            return presets[userKey];
        }
        return {
            key: userKey,
            label: userKey,
            editableFilters: [],
            canEditAll: false
        };
    }

    function buildUserOptionsHtml() {
        return LOGIN_USER_ORDER.map(user => `<option value="${user}">${user}</option>`).join('');
    }

    function ensureLoginDialog() {
        let overlay = document.getElementById('stock-costura-login-overlay');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.id = 'stock-costura-login-overlay';
            overlay.className = 'stock-login-overlay active';
            overlay.setAttribute('aria-hidden', 'false');
            overlay.innerHTML = `
                <div class="stock-login-card" role="dialog" aria-modal="true" aria-labelledby="stock-login-title">
                    <h2 id="stock-login-title" class="stock-login-title">Acceso Stock Costura</h2>
                    <p class="stock-login-subtitle">Selecciona usuario e ingresa contrasena.</p>
                    <form id="stock-login-form" class="stock-login-form" autocomplete="off">
                        <label for="stock-login-user">Usuario</label>
                        <select id="stock-login-user" required>
                            ${buildUserOptionsHtml()}
                        </select>
                        <label for="stock-login-password">Contrasena</label>
                        <input id="stock-login-password" type="password" required />
                        <p id="stock-login-error" class="stock-login-error" aria-live="polite"></p>
                        <button id="stock-login-submit" type="submit">Ingresar</button>
                    </form>
                </div>
            `;
            document.body.appendChild(overlay);
        }

        return {
            overlay,
            form: overlay.querySelector('#stock-login-form'),
            userSelect: overlay.querySelector('#stock-login-user'),
            passwordInput: overlay.querySelector('#stock-login-password'),
            errorEl: overlay.querySelector('#stock-login-error'),
            submitBtn: overlay.querySelector('#stock-login-submit')
        };
    }

    function openLoginDialog(resolve) {
        const refs = ensureLoginDialog();
        const { overlay, form, userSelect, passwordInput, errorEl, submitBtn } = refs;
        if (!form || !userSelect || !passwordInput || !errorEl || !submitBtn) {
            resolve({
                key: 'READ_ONLY',
                label: 'SOLO LECTURA',
                editableFilters: [],
                canEditAll: false
            });
            return;
        }

        overlay.classList.add('active');
        overlay.setAttribute('aria-hidden', 'false');
        errorEl.textContent = '';
        errorEl.style.display = 'none';
        submitBtn.disabled = false;
        userSelect.value = LOGIN_USER_ORDER[0];
        passwordInput.value = '';
        setTimeout(() => {
            try { passwordInput.focus(); } catch (e) {}
        }, 0);

        const showError = (message) => {
            errorEl.textContent = message;
            errorEl.style.display = 'block';
        };

        const onSubmit = (event) => {
            event.preventDefault();
            const userKey = normalizeUserKey(userSelect.value);
            const password = String(passwordInput.value || '');
            const expectedPassword = USER_CREDENTIALS[userKey];
            if (!expectedPassword || password !== expectedPassword) {
                showError('Usuario o contrasena incorrecta.');
                try {
                    passwordInput.focus();
                    passwordInput.select();
                } catch (e) {}
                return;
            }

            const profile = getRoleProfile(userKey);
            global.STOCK_COSTURA_ACTIVE_USER = userKey;
            global.STOCK_COSTURA_ACTIVE_PROFILE = profile;
            resolvedProfile = profile;
            overlay.classList.remove('active');
            overlay.setAttribute('aria-hidden', 'true');
            form.removeEventListener('submit', onSubmit);
            resolve(profile);
        };

        form.addEventListener('submit', onSubmit);
    }

    function requireAccess() {
        if (resolvedProfile) {
            return Promise.resolve(resolvedProfile);
        }
        if (pendingLoginPromise) {
            return pendingLoginPromise;
        }
        pendingLoginPromise = new Promise(resolve => {
            openLoginDialog(resolve);
        }).finally(() => {
            pendingLoginPromise = null;
        });
        return pendingLoginPromise;
    }

    global.StockCosturaLogin = {
        requireAccess
    };
})(window);
