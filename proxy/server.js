const http = require('http');

const PORT = process.env.PORT || 3001;
const API_BASE = 'https://api.dfogang.com';

const server = http.createServer(async (req, res) => {
    // CORS headers for all responses
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    // Handle preflight
    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    // Health check
    if (req.method === 'GET' && req.url === '/') {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ status: 'ok' }));
        return;
    }

    // Proxy POST /api/*  →  https://api.dfogang.com/*
    if (req.method === 'POST' && req.url.startsWith('/api/')) {
        const endpoint = req.url.replace('/api/', '');
        const targetUrl = `${API_BASE}/${endpoint}`;

        try {
            // Read request body
            const body = await new Promise((resolve, reject) => {
                const chunks = [];
                req.on('data', c => chunks.push(c));
                req.on('end', () => resolve(Buffer.concat(chunks).toString()));
                req.on('error', reject);
            });

            const response = await fetch(targetUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body
            });

            const data = await response.text();
            res.writeHead(response.status, { 'Content-Type': 'application/json' });
            res.end(data);
        } catch (err) {
            console.error(`Proxy error: ${err.message}`);
            res.writeHead(502, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'Proxy request failed' }));
        }
        return;
    }

    res.writeHead(404);
    res.end('Not found');
});

server.listen(PORT, () => {
    console.log(`CORS proxy running on http://localhost:${PORT}`);
});
