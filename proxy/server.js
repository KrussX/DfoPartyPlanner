const API_BASE = 'https://api.dfogang.com';

const CORS_HEADERS = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
};

export default {
    async fetch(request) {
        // Handle preflight
        if (request.method === 'OPTIONS') {
            return new Response(null, { status: 204, headers: CORS_HEADERS });
        }

        const url = new URL(request.url);

        // Health check
        if (request.method === 'GET' && url.pathname === '/') {
            return Response.json({ status: 'ok' }, { headers: CORS_HEADERS });
        }

        // Proxy POST /api/*  →  https://api.dfogang.com/*
        if (request.method === 'POST' && url.pathname.startsWith('/api/')) {
            const endpoint = url.pathname.replace('/api/', '');
            const targetUrl = `${API_BASE}/${endpoint}`;

            try {
                const response = await fetch(targetUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: request.body,
                });

                const data = await response.text();
                return new Response(data, {
                    status: response.status,
                    headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
                });
            } catch (err) {
                return Response.json(
                    { error: 'Proxy request failed' },
                    { status: 502, headers: CORS_HEADERS }
                );
            }
        }

        return new Response('Not found', { status: 404, headers: CORS_HEADERS });
    },
};
