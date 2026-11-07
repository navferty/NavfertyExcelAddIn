const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

async function testSearch() {
    let submit;
    const query = { value: 'error' };
    const status = { textContent: '' };
    const results = {
        children: [],
        replaceChildren() { this.children = []; },
        appendChild(child) { this.children.push(child); }
    };
    const form = {
        dataset: { index: '/en/search-index.json' },
        addEventListener(type, handler) { if (type === 'submit') submit = handler; }
    };
    const pages = [
        { title: 'Body match', description: 'Guide', content: 'error', url: '/body/' },
        { title: 'Description match', description: 'an error', content: 'guide', url: '/description/' },
        { title: 'Find Errors', description: 'Find problems', content: 'guide', url: '/title/' },
        { title: 'Other Errors', description: 'Find problems', content: 'guide', url: '/second-title/' }
    ];
    const document = {
        documentElement: { lang: 'en' },
        querySelector(selector) { return selector === '.site-search' ? form : null; },
        getElementById(id) {
            return { 'search-query': query, 'search-status': status, 'search-results': results }[id] || null;
        },
        createElement(tag) {
            return {
                tag, children: [],
                append(...children) { this.children.push(...children); }
            };
        }
    };
    const context = {
        document, window: { location: { search: '' } }, URLSearchParams, console,
        fetch: async () => ({ ok: true, json: async () => pages })
    };
    vm.runInNewContext(fs.readFileSync('Website/hugo-assets/site.js', 'utf8'), context);
    submit({ preventDefault() {} });
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(status.textContent, 'Results: 4');
    assert.deepEqual(results.children.map(item => item.children[0].href),
        ['/title/', '/second-title/', '/description/', '/body/']);

    query.value = 'find error';
    submit({ preventDefault() {} });
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(status.textContent, 'Results: 2');
    assert.deepEqual(results.children.map(item => item.children[0].href), ['/title/', '/second-title/']);
}

testSearch().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
