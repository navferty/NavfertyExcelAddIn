getUserLanguage = () => localStorage.getItem('lng');

setUserLanguage = (name) => {
    localStorage.setItem('lng', name);
    changeLanguage();
}

if (!getUserLanguage())
    setUserLanguage("en");

async function changeLanguage() {
    let lang = getUserLanguage();
    document.getElementById('content').innerHTML = await loadContents(lang + '.html');
    let title;
    if (lang == "ru") {
        title = "Надстройка Navferty Excel Add-In";
    } else if (lang == "en") {
        title = "Navferty's Excel Add-In";
    } else {
        console.warn(`Invalid lang: ${lang}!`);
        title = "Navferty's Excel Add-In";
    }

    document.getElementById('head_name').innerHTML = title;
    document.title = title;

    var allHeaders = document.querySelectorAll('h2, h3');
    var navigationMenu = document.querySelector("#left_nav");
    navigationMenu.innerHTML = '';
    for (let header of allHeaders) {
        var href = header.id;
        if (!href)
            continue;
        var link = document.createElement("a");
        link.className = 'nav_tab';
        link.href = "#" + href;
        var textContent = header.textContent || header.innerText;
        link.innerHTML = '<p>' + textContent.trim() + '</p>';
        navigationMenu.appendChild(link);
    }
}

function loadContents(url) {
    return new Promise(function (resolve, reject) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', url, true);
        xhr.responseType = 'text';
        xhr.onload = function () {
            var status = xhr.status;
            if (status == 200) {
                resolve(xhr.responseText);
            } else {
                reject(status);
            }
        };
        xhr.send();
    });
}

function initializeCurrentYear() {
    document.getElementById('current-year').textContent = new Date().getFullYear();
}

window.addEventListener('load', function() {
    changeLanguage();
    initializeCurrentYear();
});
