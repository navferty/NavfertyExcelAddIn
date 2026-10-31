function getUserLanguage() {
    return localStorage.getItem('lng') === 'ru' ? 'ru' : 'en';
}

function setUserLanguage(name) {
    localStorage.setItem('lng', name);
    applyLanguage();
}

function applyLanguage() {
    var lang = getUserLanguage();
    document.documentElement.lang = lang;
    document.querySelectorAll('.lang-en').forEach(function (e) { e.hidden = lang !== 'en'; });
    document.querySelectorAll('.lang-ru').forEach(function (e) { e.hidden = lang !== 'ru'; });
    document.title = document.body.getAttribute('data-title-' + lang);
    var desc = document.querySelector('meta[name="description"]');
    if (lang === 'ru' && desc) {
        if (!desc.hasAttribute('data-en')) desc.setAttribute('data-en', desc.content);
        desc.content = document.body.getAttribute('data-desc-ru');
    } else if (desc && desc.hasAttribute('data-en')) {
        desc.content = desc.getAttribute('data-en');
    }
    document.getElementById('head_name').firstElementChild.textContent =
        lang === 'ru' ? 'Надстройка Navferty Excel Add-In' : "Navferty's Excel Add-In";
}

window.addEventListener('load', function () {
    applyLanguage();
    document.getElementById('current-year').textContent = new Date().getFullYear();
});
