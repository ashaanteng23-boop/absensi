/**
 * Vercel Function proxy untuk Google Apps Script.
 * Set APPS_SCRIPT_URL di Vercel Project Settings > Environment Variables.
 */
export default async function handler(request, response) {
    const appsScriptUrl = process.env.APPS_SCRIPT_URL;

    if (!appsScriptUrl) {
        return response.status(500).json({
            ok: false,
            error: 'APPS_SCRIPT_URL belum dikonfigurasi di Vercel.'
        });
    }

    try {
        const targetUrl = new URL(appsScriptUrl);
        const requestUrl = new URL(request.url, 'http://localhost');
        requestUrl.searchParams.forEach((value, key) => targetUrl.searchParams.set(key, value));

        const options = {
            method: request.method,
            headers: { 'Content-Type': 'application/json' }
        };

        if (request.method !== 'GET' && request.method !== 'HEAD') {
            options.body = typeof request.body === 'string'
                ? request.body
                : JSON.stringify(request.body || {});
        }

        const upstream = await fetch(targetUrl, options);
        const body = await upstream.text();
        let result;
        try {
            result = JSON.parse(body);
        } catch {
            result = { ok: false, error: 'Respons Apps Script bukan JSON yang valid.', detail: body.slice(0, 300) };
        }

        return response.status(upstream.ok ? 200 : 502).json(result);
    } catch (error) {
        console.error(error);
        return response.status(502).json({ ok: false, error: 'Tidak dapat terhubung ke Google Apps Script.' });
    }
}
