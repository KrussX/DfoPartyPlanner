/**
 * DFO Gang Playwright Proxy Server (Optimized)
 *
 * Replaces the Cloudflare Workers proxy by using a real browser (Playwright)
 * to navigate to dfogang.com and intercept the API responses.
 *
 * Optimizations:
 *   - Blocks images, CSS, fonts, ads, and analytics (we only need the API call)
 *   - Resolves IMMEDIATELY when the API response is captured (no waiting for full page load)
 *   - Reuses a persistent browser context with cookies/sessions across requests
 *   - Request queue prevents Cloudflare rate-limiting from concurrent tabs
 *
 * Usage:
 *   npm install
 *   npm run install-browser
 *   npm start
 *
 * Then point your front-end API_PROXY_URL to http://localhost:3001
 */

const express = require('express');
const { chromium } = require('playwright');

const app = express();
const PORT = process.env.PORT || 3001;

// --- CORS ---
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') return res.sendStatus(204);
    next();
});

app.use(express.json());

// --- Browser + Persistent Context ---
// We reuse a single browser AND context (keeps cookies/session alive,
// so Cloudflare challenges only need to be solved once).
let browser = null;
let persistentContext = null;

async function getBrowser() {
    if (!browser || !browser.isConnected()) {
        console.log('[Browser] Launching Chromium...');
        browser = await chromium.launch({
            headless: true, // Set to false for debugging
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',  // Use /tmp instead of /dev/shm (often too small in containers)
                '--disable-gpu',
            ],
        });
        persistentContext = await browser.newContext({
            userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
            viewport: { width: 1280, height: 720 },
        });
        console.log('[Browser] Ready.');
    }
    return { browser, context: persistentContext };
}

// --- Request Queue ---
// Process one search at a time to avoid rate-limiting and resource contention.
let requestQueue = Promise.resolve();

function enqueue(fn) {
    const p = requestQueue.then(fn, fn);
    requestQueue = p.catch(() => {}); // prevent unhandled rejection chain
    return p;
}

// --- Resource Blocking ---
// Block everything except documents and XHR/fetch — massive speed boost.
const BLOCKED_TYPES = new Set([
    'image', 'stylesheet', 'font', 'media', 'texttrack', 'eventsource',
    'manifest', 'other',
]);

const BLOCKED_DOMAINS = [
    'googlesyndication.com',
    'googletagmanager.com',
    'google-analytics.com',
    'doubleclick.net',
    'adservice.google',
    'facebook.net',
    'fbcdn.net',
];

// --- Health Check ---
app.get('/', (req, res) => {
    res.json({ status: 'ok', engine: 'playwright-optimized' });
});

// --- Main Endpoint: POST /api/search_explorer ---
app.post('/api/search_explorer', async (req, res) => {
    const { name, server = 'explorer', exact_match = true } = req.body;

    if (!name) {
        return res.status(400).json({ error: 'Missing "name" in request body' });
    }

    const startTime = Date.now();
    console.log(`[Search] Looking up "${name}" (server=${server}, exact=${exact_match})`);

    // Queue the request so we don't overwhelm the browser
    try {
        const result = await enqueue(() => performSearch(name, server, exact_match));
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
        console.log(`[Search] ✅ "${name}" → ${(result.results || []).length} results in ${elapsed}s`);
        return res.json(result);
    } catch (err) {
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
        console.error(`[Search] ❌ "${name}" failed in ${elapsed}s:`, err.message);
        return res.status(502).json({ error: `Proxy request failed: ${err.message}` });
    }
});

async function performSearch(name, server, exact_match) {
    const { context } = await getBrowser();
    const page = await context.newPage();

    try {
        // Block unnecessary resources for speed
        await page.route('**/*', (route) => {
            const req = route.request();
            const type = req.resourceType();
            const url = req.url();

            // Block heavy resource types
            if (BLOCKED_TYPES.has(type)) {
                return route.abort();
            }

            // Block ad/analytics domains
            if (BLOCKED_DOMAINS.some(domain => url.includes(domain))) {
                return route.abort();
            }

            return route.continue();
        });

        // Create a promise that resolves as soon as we capture the API response.
        // This way we don't wait for the full page — we bail out early.
        let resolveCapture, rejectCapture;
        const capturePromise = new Promise((resolve, reject) => {
            resolveCapture = resolve;
            rejectCapture = reject;
        });

        // Listen for the API response
        page.on('response', async (response) => {
            const url = response.url();

            // Primary: direct API call to api.dfogang.com
            if (url.includes('api.dfogang.com') && url.includes('search_explorer')) {
                try {
                    const contentType = response.headers()['content-type'] || '';
                    if (contentType.includes('application/json')) {
                        const json = await response.json();
                        if (json.results || Array.isArray(json)) {
                            resolveCapture(json);
                        }
                    }
                } catch (e) { /* ignore parse errors */ }
            }
        });

        // Build URL
        const params = new URLSearchParams({ server, name, view: 'card' });
        if (exact_match) params.set('exact', 'true');
        const targetUrl = `https://dfogang.com/?${params.toString()}`;

        // Navigate — use 'commit' which resolves as soon as the server responds
        // (doesn't wait for all sub-resources to load)
        const gotoPromise = page.goto(targetUrl, {
            waitUntil: 'commit',
            timeout: 20000,
        });

        // Race: resolve as soon as either the API response is captured
        // or we hit a timeout
        const timeoutPromise = new Promise((_, reject) =>
            setTimeout(() => reject(new Error('Timed out waiting for API response (15s)')), 15000)
        );

        const result = await Promise.race([capturePromise, timeoutPromise]);
        return result;

    } finally {
        // Always close the page to free memory
        await page.close().catch(() => {});
    }
}

// --- Graceful Shutdown ---
process.on('SIGINT', async () => {
    console.log('\n[Server] Shutting down...');
    if (browser) await browser.close();
    process.exit(0);
});

process.on('SIGTERM', async () => {
    if (browser) await browser.close();
    process.exit(0);
});

// --- Start ---
app.listen(PORT, async () => {
    console.log(`\n🚀 DFO Gang Playwright Proxy (Optimized) running on http://localhost:${PORT}`);
    console.log(`   Point your front-end API_PROXY_URL to this address.\n`);
    // Pre-launch browser so first request is fast
    await getBrowser();
});
