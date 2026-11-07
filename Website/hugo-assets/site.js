(function () {
    'use strict';
    var russian = document.documentElement.lang === 'ru';
    var searchToggle = document.querySelector('.header-search-toggle');
    var headerForm = document.getElementById('header-search-form');
    var headerInput = document.getElementById('header-search-query');
    if (searchToggle && headerForm && headerInput) {
        function setSearchOpen(open) {
            headerForm.hidden = !open;
            searchToggle.setAttribute('aria-expanded', String(open));
            if (open) headerInput.focus();
        }
        searchToggle.addEventListener('click', function (event) {
            event.preventDefault();
            setSearchOpen(headerForm.hidden);
        });
        document.addEventListener('keydown', function (event) {
            var target = event.target;
            var editing = target instanceof Element &&
                (target.matches('input, textarea, select') || target.isContentEditable);
            if (event.key === '/' && !event.ctrlKey && !event.metaKey && !event.altKey && !editing) {
                event.preventDefault();
                setSearchOpen(true);
            } else if (event.key === 'Escape' && !headerForm.hidden) {
                setSearchOpen(false);
                searchToggle.focus();
            }
        });
    }

    var form = document.querySelector('.site-search');
    if (form) {
        var status = document.getElementById('search-status');
        var results = document.getElementById('search-results');
        var indexPromise;
        var requestNumber = 0;
        form.addEventListener('submit', function (event) {
            event.preventDefault();
            var terms = document.getElementById('search-query').value.trim().toLocaleLowerCase().split(/\s+/);
            if (!terms[0]) return;
            var request = ++requestNumber;
            status.textContent = russian ? 'Поиск…' : 'Searching…';
            results.replaceChildren();
            if (!indexPromise) {
                indexPromise = fetch(form.dataset.index).then(function (response) {
                    if (!response.ok) throw new Error('Search index: HTTP ' + response.status);
                    return response.json();
                }).catch(function (error) {
                    indexPromise = null;
                    throw error;
                });
            }
            indexPromise.then(function (pages) {
                if (request !== requestNumber) return;
                var matches = pages.filter(function (page) {
                    var text = (page.title + ' ' + page.description + ' ' + page.content).toLocaleLowerCase();
                    return terms.every(function (term) { return text.includes(term); });
                });
                status.textContent = russian ? 'Результатов: ' + matches.length : 'Results: ' + matches.length;
                matches.forEach(function (page) {
                    var item = document.createElement('li');
                    var link = document.createElement('a');
                    var description = document.createElement('p');
                    link.href = page.url;
                    link.textContent = page.title;
                    description.textContent = page.description;
                    item.append(link, description);
                    results.appendChild(item);
                });
            }).catch(function (error) {
                if (request !== requestNumber) return;
                console.error('Search failed:', error);
                status.textContent = russian ? 'Не удалось выполнить поиск. Попробуйте ещё раз.' : 'Search failed. Please try again.';
            });
        });
        var initialQuery = new URLSearchParams(window.location.search).get('q');
        if (initialQuery && initialQuery.trim()) {
            document.getElementById('search-query').value = initialQuery;
            form.requestSubmit();
        }
    }

    if (typeof HTMLDialogElement === 'undefined' || !HTMLDialogElement.prototype.showModal) return;
    var dialog = document.createElement('dialog');
    dialog.className = 'image-dialog';
    dialog.setAttribute('aria-label', russian ? 'Просмотр изображения' : 'Image viewer');
    var close = document.createElement('button');
    close.type = 'button';
    close.textContent = russian ? 'Закрыть' : 'Close';
    var enlarged = document.createElement('img');
    var trigger;
    dialog.append(close, enlarged);
    document.body.appendChild(dialog);
    close.addEventListener('click', function () { dialog.close(); });
    dialog.addEventListener('close', function () {
        if (trigger) trigger.focus();
    });
    dialog.addEventListener('click', function (event) {
        if (event.target === dialog) dialog.close();
    });
    document.querySelectorAll('#content img:not(.icons)').forEach(function (image) {
        if (!image.alt || image.closest('a, button')) return;
        var button = document.createElement('button');
        button.type = 'button';
        button.className = 'image-open';
        button.setAttribute('aria-label', (russian ? 'Увеличить изображение: ' : 'Enlarge image: ') + image.alt);
        image.parentNode.insertBefore(button, image);
        button.appendChild(image);
        button.addEventListener('click', function () {
            trigger = button;
            enlarged.src = image.currentSrc || image.src;
            enlarged.alt = image.alt;
            dialog.showModal();
            close.focus();
        });
    });
}());
