const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

async function testCopy() {
    const messages = {
        copySectionLink: 'Copy link',
        linkCopied: 'Copied',
        copyLinkFailed: 'Copy failed'
    };
    const handlers = {};
    const timers = new Map();
    const writes = [];
    let timerId = 0;

    function element(tag) {
        return {
            tag, children: [], attributes: {}, textContent: '',
            setAttribute(name, value) { this.attributes[name] = value; },
            appendChild(child) { child.parentElement = this; this.children.push(child); },
            querySelector(selector) {
                return this.children.find(child => `.${child.className}` === selector);
            }
        };
    }

    function heading(id) {
        const node = element('h2');
        node.id = id;
        node.closest = () => null;
        node.insertBefore = child => node.appendChild(child);
        return node;
    }

    const headings = [heading('first'), heading('second')];
    const article = {
        querySelectorAll(selector) {
            return selector === 'h2, h3, h4' ? headings : headings;
        },
        addEventListener(type, handler) { handlers[type] = handler; }
    };
    const document = {
        documentElement: { lang: 'en' },
        getElementById(id) {
            return id === 'site-messages' ? { textContent: JSON.stringify(messages) } : null;
        },
        querySelector(selector) { return selector === '#content.article' ? article : null; },
        querySelectorAll() { return headings; },
        createElement: element
    };
    const navigator = {
        clipboard: {
            writeText(url) {
                return new Promise((resolve, reject) => writes.push({ url, resolve, reject }));
            }
        }
    };
    const context = {
        document, navigator, URL, WeakMap, Set,
        window: {
            localStorage: { setItem() {} },
            location: { href: 'https://example.test/en/article/', hash: '' }
        },
        console: { error() {}, warn() {} },
        setTimeout(callback, delay) {
            assert.equal(delay, 2000);
            const id = ++timerId;
            timers.set(id, callback);
            return id;
        },
        clearTimeout(id) { timers.delete(id); }
    };
    vm.runInNewContext(fs.readFileSync('Website/hugo-assets/site.js', 'utf8'), context);
    const [first, second] = headings.map(node => node.children[0]);
    const icon = button => button.querySelector('.heading-anchor-icon');
    const status = button => button.querySelector('.heading-anchor-status');
    const click = button => handlers.click({ target: { closest: () => button } });
    const flush = () => new Promise(resolve => setImmediate(resolve));

    assert.equal(first.attributes['aria-label'], messages.copySectionLink);
    assert.equal(icon(first).attributes['aria-hidden'], 'true');
    assert.equal(status(first).attributes['aria-live'], 'polite');
    click(first);
    await flush();
    assert.equal(writes[0].url, 'https://example.test/en/article/#first');
    assert.equal(icon(first).textContent, '#');
    writes[0].resolve();
    await flush();
    assert.equal(icon(first).textContent, '✓');
    assert.equal(status(first).textContent, messages.linkCopied);
    assert.equal(timers.size, 1);
    const [expiredId, expiredCallback] = timers.entries().next().value;
    timers.delete(expiredId);
    expiredCallback();
    assert.equal(icon(first).textContent, '#');
    assert.equal(timers.size, 0);

    click(first);
    await flush();
    writes[1].reject(new Error('Clipboard denied'));
    await flush();
    assert.equal(icon(first).textContent, '#');
    assert.equal(status(first).textContent, messages.copyLinkFailed);

    click(first);
    click(first);
    click(second);
    await flush();
    writes[4].resolve();
    writes[3].resolve();
    await flush();
    assert.equal(icon(first).textContent, '✓');
    assert.equal(icon(second).textContent, '✓');
    writes[2].reject(new Error('Older attempt failed'));
    await flush();
    assert.equal(icon(first).textContent, '✓');
    assert.equal(status(first).textContent, messages.linkCopied);
    assert.equal(timers.size, 2);

    click(first);
    await flush();
    assert.equal(icon(first).textContent, '#');
    assert.equal(status(first).textContent, '');
    assert.equal(icon(second).textContent, '✓');
    assert.equal(timers.size, 1);
    writes[5].resolve();
    await flush();
    assert.equal(icon(first).textContent, '✓');
    assert.equal(timers.size, 2);
}

testCopy().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
