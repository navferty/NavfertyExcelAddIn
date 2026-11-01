document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a[data-lang]');
    if (a) {
        try { localStorage.setItem('lng', a.getAttribute('data-lang')); } catch (err) { }
    }
});

var yearEl = document.getElementById('current-year');
if (yearEl) yearEl.textContent = new Date().getFullYear();
