(function () {
    'use strict';
    try {
        window.localStorage.setItem('websiteLanguage', document.documentElement.lang);
    } catch (error) {
        console.warn('Could not save language preference:', error);
    }
    var messagesElement = document.getElementById('site-messages');
    var messages = JSON.parse(messagesElement.textContent);
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
            var locale = document.documentElement.lang;
            var terms = document.getElementById('search-query').value.trim().toLocaleLowerCase(locale).split(/\s+/);
            if (!terms[0]) return;
            var request = ++requestNumber;
            status.textContent = messages.searching;
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
                var matches = pages.map(function (page, index) {
                    var title = page.title.toLocaleLowerCase(locale);
                    var description = page.description.toLocaleLowerCase(locale);
                    var text = title + ' ' + description + ' ' + page.content.toLocaleLowerCase(locale);
                    return {
                        page: page,
                        index: index,
                        matches: terms.every(function (term) { return text.includes(term); }),
                        titleMatches: terms.filter(function (term) { return title.includes(term); }).length,
                        descriptionMatches: terms.filter(function (term) { return description.includes(term); }).length
                    };
                }).filter(function (result) {
                    return result.matches;
                }).sort(function (a, b) {
                    return b.titleMatches - a.titleMatches ||
                        b.descriptionMatches - a.descriptionMatches || a.index - b.index;
                });
                status.textContent = messages.searchResults.replace('{count}', matches.length);
                matches.forEach(function (result) {
                    var page = result.page;
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
                status.textContent = messages.searchFailed;
            });
        });
        var initialQuery = new URLSearchParams(window.location.search).get('q');
        if (initialQuery && initialQuery.trim()) {
            document.getElementById('search-query').value = initialQuery;
            form.requestSubmit();
        }
    }

    var year = document.getElementById('current-year');
    if (year) year.textContent = new Date().getFullYear();

    var article = document.querySelector('#content.article');
    if (article) {
        var usedIds = new Set(Array.from(document.querySelectorAll('[id]'), function (element) {
            return element.id;
        }));

        article.querySelectorAll('h2, h3, h4').forEach(function (heading) {
            if (heading.closest('.installation-steps')) return;
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
            copyButton.setAttribute('aria-label', messages.copySectionLink);
            copyButton.title = messages.copySectionLink;
            copyButton.textContent = '#';

            var copyStatus = document.createElement('span');
            copyStatus.className = 'heading-anchor-status';
            copyStatus.setAttribute('aria-live', 'polite');
            copyStatus.setAttribute('aria-atomic', 'true');
            copyButton.appendChild(copyStatus);
            heading.insertBefore(copyButton, heading.firstChild);
        });

        article.addEventListener('click', function (event) {
            var button = event.target.closest('.heading-anchor');
            if (!button) return;

            var heading = button.parentElement;
            var link = new URL(window.location.href);
            link.hash = heading.id;
            var copyStatus = button.querySelector('.heading-anchor-status');

            Promise.resolve().then(function () {
                if (!navigator.clipboard) throw new Error('Clipboard access is unavailable.');
                return navigator.clipboard.writeText(link.href);
            }).then(function () {
                copyStatus.textContent = messages.linkCopied;
            }).catch(function (error) {
                console.error('Could not copy the section link.', error);
                copyStatus.textContent = messages.copyLinkFailed;
            });
        });

        var fragment = window.location.hash.slice(1);
        var target = Array.from(article.querySelectorAll('[id]')).find(function (element) {
            return element.id === fragment || encodeURIComponent(element.id) === fragment;
        });
        if (target) target.scrollIntoView();
    }

    if (typeof HTMLDialogElement === 'undefined' || !HTMLDialogElement.prototype.showModal) return;
    var dialog = document.createElement('dialog');
    dialog.className = 'image-dialog';
    dialog.setAttribute('aria-label', messages.imageViewer);
    var close = document.createElement('button');
    close.type = 'button';
    close.className = 'image-close';
    close.setAttribute('aria-label', messages.close);
    close.title = messages.close;
    close.textContent = '×';
    var enlarged = document.createElement('img');
    var imageFrame = document.createElement('div');
    imageFrame.className = 'image-frame';
    var trigger;
    imageFrame.append(enlarged, close);
    dialog.appendChild(imageFrame);
    document.body.appendChild(dialog);
    close.addEventListener('click', function () { dialog.close(); });
    dialog.addEventListener('close', function () {
        imageFrame.classList.remove('image-frame-hover-close');
        if (trigger) trigger.focus();
    });
    dialog.addEventListener('click', function (event) {
        if (event.target === dialog) dialog.close();
    });
    imageFrame.addEventListener('pointermove', function (event) {
        if (event.pointerType === 'touch') return;
        var bounds = enlarged.getBoundingClientRect();
        var inCorner = event.clientX >= bounds.left && event.clientX <= bounds.left + bounds.width / 3 &&
            event.clientY >= bounds.top && event.clientY <= bounds.top + bounds.height / 3;
        imageFrame.classList.toggle('image-frame-hover-close', inCorner);
    });
    imageFrame.addEventListener('pointerleave', function () {
        imageFrame.classList.remove('image-frame-hover-close');
    });
    document.querySelectorAll('#content img:not(.icons)').forEach(function (image) {
        if (!image.alt || image.closest('a, button')) return;
        var button = document.createElement('button');
        button.type = 'button';
        button.className = 'image-open';
        button.setAttribute('aria-label', messages.enlargeImage + image.alt);
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
