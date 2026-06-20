/**
 * DFO Gang Playwright Proxy Server
 *
 * Replaces the Cloudflare Workers proxy by using a real browser (Playwright)
 * to navigate to dfogang.com and intercept the API responses.
 *
 * The website makes internal API calls that are blocked for external access,
 * but since we load the real page in a real browser, the calls go through
 * normally and we capture the responses via network interception.
 *
 * Usage:
 *   npm install
 *   npm run install-browser
 *   npm start
 *
 * Then point your front-end API_PROXY_URL to http://localhost:3000
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

// --- Browser Pool ---
// We keep a single persistent browser instance to avoid the overhead
// of launching a new browser for every request (~2-3s saved per call).
let browser = null;

async function getBrowser() {
    if (!browser || !browser.isConnected()) {
        console.log('[Browser] Launching Chromium...');
        browser = await chromium.launch({
            headless: true, // Set to false to see the browser for debugging
        });
        console.log('[Browser] Ready.');
    }
    return browser;
}

// --- Health Check ---
app.get('/', (req, res) => {
    res.json({ status: 'ok', engine: 'playwright' });
});

// --- Main Endpoint: POST /api/search_explorer ---
// Matches the same contract as the old Cloudflare proxy.
// Body: { name: string, server: string, average_set_dmg?: bool, exact_match?: bool }
// Returns: { results: [...] }
app.post('/api/search_explorer', async (req, res) => {
    const { name, server = 'explorer', exact_match = true } = req.body;

    if (!name) {
        return res.status(400).json({ error: 'Missing "name" in request body' });
    }

    console.log(`[Search] Looking up "${name}" (server=${server}, exact=${exact_match})`);

    let context = null;
    try {
        const b = await getBrowser();
        context = await b.newContext({
            // Mimic a normal browser to avoid bot detection
            userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
            viewport: { width: 1280, height: 720 },
        });
        const page = await context.newPage();

        // Set up network interception to capture the API response.
        // dfogang.com is a Next.js app that fetches data via internal routes.
        // We look for the response that contains character search results.
        let capturedData = null;

        page.on('response', async (response) => {
            const url = response.url();

            // The site makes requests to api.dfogang.com or its own Next.js
            // data routes. We capture any response containing search results.
            if (url.includes('api.dfogang.com') || url.includes('search_explorer')) {
                try {
                    const contentType = response.headers()['content-type'] || '';
                    if (contentType.includes('application/json')) {
                        const json = await response.json();
                        if (json.results || Array.isArray(json)) {
                            capturedData = json;
                            console.log(`[Search] Captured API response from: ${url} (${(json.results || json).length} results)`);
                        }
                    }
                } catch (e) {
                    // Not JSON or parsing failed, ignore
                }
            }
        });

        // Also capture from Next.js RSC data (the site uses server components)
        // which may embed the data in the page payload
        page.on('response', async (response) => {
            const url = response.url();
            // Next.js data fetches often use __next_data or RSC format
            if (capturedData) return; // Already captured

            try {
                if (url.includes('dfogang.com') && url.includes('?server=')) {
                    const contentType = response.headers()['content-type'] || '';
                    if (contentType.includes('text/x-component') || contentType.includes('application/json')) {
                        const text = await response.text();
                        // Try to extract JSON results from RSC payload
                        const match = text.match(/"results"\s*:\s*(\[[\s\S]*?\])/);
                        if (match) {
                            try {
                                const results = JSON.parse(match[1]);
                                capturedData = { results };
                                console.log(`[Search] Captured RSC data (${results.length} results)`);
                            } catch (e) { /* not valid JSON array */ }
                        }
                    }
                }
            } catch (e) { /* ignore */ }
        });

        // Build the URL - the website uses query params:
        // dfogang.com/?server=explorer&name=<name>&exact=true
        const params = new URLSearchParams({
            server,
            name,
            view: 'card',
        });
        if (exact_match) params.set('exact', 'true');

        const targetUrl = `https://dfogang.com/?${params.toString()}`;
        console.log(`[Search] Navigating to: ${targetUrl}`);

        // Navigate and wait for network to settle
        await page.goto(targetUrl, {
            waitUntil: 'networkidle',
            timeout: 30000,
        });

        // If we haven't captured the API response yet, wait a bit more
        // (some requests may fire after initial load)
        if (!capturedData) {
            console.log('[Search] Waiting for additional network activity...');
            await page.waitForTimeout(3000);
        }

        // If still no API response captured, try to extract data from the DOM
        if (!capturedData) {
            console.log('[Search] No API response captured, attempting DOM extraction...');
            capturedData = await page.evaluate(() => {
                // The Next.js app stores data in __NEXT_DATA__ or similar
                const nextData = document.getElementById('__NEXT_DATA__');
                if (nextData) {
                    try {
                        const parsed = JSON.parse(nextData.textContent);
                        if (parsed?.props?.pageProps?.results) {
                            return { results: parsed.props.pageProps.results };
                        }
                    } catch (e) { /* ignore */ }
                }

                // Try window.__next_f (RSC flight data)
                if (window.__next_f) {
                    const allText = window.__next_f
                        .filter(item => typeof item[1] === 'string')
                        .map(item => item[1])
                        .join('');

                    // Look for results array in the flight data
                    const match = allText.match(/"results"\s*:\s*(\[[\s\S]*?\](?=\s*[,}]))/);
                    if (match) {
                        try {
                            return { results: JSON.parse(match[1]) };
                        } catch (e) { /* ignore */ }
                    }
                }

                return null;
            });
        }

        await context.close();
        context = null;

        if (capturedData) {
            console.log(`[Search] ✅ Returning ${(capturedData.results || []).length} results`);
            return res.json(capturedData);
        } else {
            console.log('[Search] ❌ No data captured');
            return res.status(502).json({
                error: 'Could not capture API response from dfogang.com',
                hint: 'The site may have changed its structure or is blocking automated access.',
            });
        }

    } catch (err) {
        console.error('[Search] Error:', err.message);
        if (context) {
            try { await context.close(); } catch (e) { /* ignore */ }
        }
        return res.status(502).json({ error: `Proxy request failed: ${err.message}` });
    }
});

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
    console.log(`\n🚀 DFO Gang Playwright Proxy running on http://localhost:${PORT}`);
    console.log(`   Point your front-end API_PROXY_URL to this address.\n`);
    // Pre-launch browser so first request is fast
    await getBrowser();
});
