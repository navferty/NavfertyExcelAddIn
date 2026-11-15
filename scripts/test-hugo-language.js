const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const vm = require('node:vm');

const root = fs.readFileSync(path.join(process.argv[2], 'index.html'), 'utf8');
const redirect = root.match(/<script>([\s\S]*?)<\/script>/)[1];
const site = fs.readFileSync('Website/hugo-assets/site.js', 'utf8');
const quietConsole = { warn() {}, error: console.error };

function chooseLanguage(saved, languages, language = 'en', blocked = false, script = redirect) {
    let destination;
    vm.runInNewContext(script, {
        localStorage: {
            getItem(key) {
                assert.equal(key, 'websiteLanguage');
                if (blocked) throw new Error('Storage blocked');
                return saved;
            }
        },
        navigator: { languages, language },
        location: { replace(url) { destination = url; } },
        console: quietConsole
    });
    return destination;
}

assert.equal(chooseLanguage(null, ['ru-RU', 'en-US']), '/ru/');
assert.equal(chooseLanguage(null, ['fr-FR', 'en-US']), '/en/');
assert.equal(chooseLanguage(null, ['fr-FR']), '/en/');
assert.equal(chooseLanguage('en', ['ru-RU']), '/en/');
assert.equal(chooseLanguage('RU', ['en-US']), '/ru/');
assert.equal(chooseLanguage('unsupported', ['ru-RU']), '/ru/');
assert.equal(chooseLanguage(null, [], 'ru-RU'), '/ru/');
assert.equal(chooseLanguage('ru', ['en-US'], 'en', true), '/en/');

function visitLanguage(language, blocked = false) {
    let saved;
    const document = {
        documentElement: { lang: language },
        getElementById(id) { return id === 'site-messages' ? { textContent: '{}' } : null; },
        querySelector() { return null; }
    };
    const window = {
        localStorage: {
            setItem(key, value) {
                assert.equal(key, 'websiteLanguage');
                if (blocked) throw new Error('Storage blocked');
                saved = value;
            }
        }
    };
    vm.runInNewContext(site, { document, window, console: quietConsole });
    return saved;
}

assert.equal(visitLanguage('en'), 'en');
assert.equal(visitLanguage('ru'), 'ru');
assert.equal(chooseLanguage(visitLanguage('ru'), ['en-US']), '/ru/');
assert.equal(visitLanguage('ru', true), undefined);

const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'hugo-language-test-'));
try {
    for (const name of ['layouts', 'i18n', 'hugo-assets', 'style.css', 'favicon.ico', 'digital-certificate.jpg', 'cert-install.png']) {
        fs.cpSync(path.join('Website', name), path.join(temporary, name), { recursive: true });
    }
    fs.mkdirSync(path.join(temporary, 'images'));
    for (const language of ['en', 'ru', 'pt-br']) {
        const content = path.join(temporary, 'content', language);
        fs.mkdirSync(content, { recursive: true });
        fs.writeFileSync(path.join(content, '_index.md'), '---\ntitle: Test\n---\nTest\n');
    }
    fs.copyFileSync(path.join(temporary, 'i18n', 'en.toml'), path.join(temporary, 'i18n', 'pt-br.toml'));
    fs.writeFileSync(path.join(temporary, 'hugo.toml'),
        fs.readFileSync('Website/hugo.toml', 'utf8') + `
[languages.pt-br]
  languageName = "Português"
  languageCode = "pt-BR"
  contentDir = "content/pt-br"
  weight = 3
[[module.mounts]]
  source = "content/pt-br"
  target = "content"
  lang = "pt-br"
`);
    const output = path.join(temporary, 'output');
    execFileSync('hugo', ['--source', temporary, '--destination', output, '--minify'], { stdio: 'pipe' });
    const html = fs.readFileSync(path.join(output, 'index.html'), 'utf8');
    const regionalRedirect = html.match(/<script>([\s\S]*?)<\/script>/)[1];
    assert.equal(chooseLanguage(null, ['pt-BR'], 'en', false, regionalRedirect), '/pt-br/');
    assert.equal(chooseLanguage(null, ['pt-BR-x-test'], 'en', false, regionalRedirect), '/pt-br/');
    assert.equal(chooseLanguage('PT-BR', ['ru'], 'en', false, regionalRedirect), '/pt-br/');
    assert.equal(chooseLanguage(null, ['pt-PT'], 'en', false, regionalRedirect), '/en/');
} finally {
    fs.rmSync(temporary, { recursive: true });
}
