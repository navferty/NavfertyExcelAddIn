document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a[data-lang]');
    if (a) {
        try { localStorage.setItem('lng', a.getAttribute('data-lang')); } catch (err) { }
    }
});

// test

var yearEl = document.getElementById('current-year');
if (yearEl) yearEl.textContent = new Date().getFullYear();

var article = document.querySelector('#content.article');
if (article) {
    var usedIds = new Set(Array.from(document.querySelectorAll('[id]'), function (element) {
        return element.id;
    }));

    article.querySelectorAll('h2, h3, h4').forEach(function (heading) {
        if (!heading.id) {
            var slug = heading.textContent
                .toLowerCase()
                .normalize('NFKD')
                .replace(/\p{M}/gu, '')
                .replace(/[^\p{L}\p{N}]+/gu, '-')
                .replace(/^-|-$/g, '') || 'section';
            var id = slug;
            var suffix = 2;
            while (usedIds.has(id)) id = slug + '-' + suffix++;
            heading.id = id;
            usedIds.add(id);
        }

        var copyButton = document.createElement('button');
        copyButton.type = 'button';
        copyButton.className = 'heading-anchor';
        var copyLabel = document.documentElement.lang === 'ru'
            ? 'Скопировать ссылку на этот раздел'
            : 'Copy link to this section';
        copyButton.setAttribute('aria-label', copyLabel);
        copyButton.title = copyLabel;
        copyButton.textContent = '#';

        var status = document.createElement('span');
        status.className = 'heading-anchor-status';
        status.setAttribute('aria-live', 'polite');
        status.setAttribute('aria-atomic', 'true');
        copyButton.appendChild(status);
        heading.insertBefore(copyButton, heading.firstChild);
    });

    article.addEventListener('click', function (event) {
        var button = event.target.closest('.heading-anchor');
        if (!button) return;

        var heading = button.parentElement;
        var link = new URL(window.location.href);
        link.hash = heading.id;
        var status = button.querySelector('.heading-anchor-status');

        Promise.resolve().then(function () {
            if (!navigator.clipboard) throw new Error('Clipboard access is unavailable.');
            return navigator.clipboard.writeText(link.href);
        }).then(function () {
            status.textContent = document.documentElement.lang === 'ru'
                ? 'Ссылка скопирована'
                : 'Link copied';
        }).catch(function (error) {
            console.error('Could not copy the section link.', error);
            status.textContent = document.documentElement.lang === 'ru'
                ? 'Не удалось скопировать ссылку. Скопируйте её из адресной строки.'
                : 'Could not copy link. Please copy it from the address bar.';
        });
    });

    var fragment = window.location.hash.slice(1);
    var target = Array.from(article.querySelectorAll('[id]')).find(function (element) {
        return element.id === fragment || encodeURIComponent(element.id) === fragment;
    });
    if (target) target.scrollIntoView();
}
