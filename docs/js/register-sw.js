// ==========================================
// 註冊 Service Worker (離線快取)
// 獨立成檔案而非 inline script，CSP 才能維持不使用 'unsafe-inline'
// ==========================================
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./service-worker.js')
            .then(reg => console.log('[SW] 註冊成功:', reg.scope))
            .catch(err => console.warn('[SW] 註冊失敗:', err));
    });
}
