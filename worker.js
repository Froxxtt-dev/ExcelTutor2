/**
 * Excel Tutor — Groq proxy (Cloudflare Worker)
 * ------------------------------------------------------------
 * Purpose: keep your Groq API key out of every learner's browser.
 * The frontend calls this worker instead of api.groq.com directly;
 * this worker attaches the real key (stored as a secret) and
 * streams the response straight back through.
 *
 * Setup:
 *   1. npm install -g wrangler   (one-time)
 *   2. wrangler login
 *   3. wrangler secret put GROQ_API_KEY     (paste your free Groq key)
 *   4. wrangler deploy
 *   5. Paste the resulting *.workers.dev URL into config.js as BACKEND_URL
 *
 * Free tier: Cloudflare Workers gives 100,000 requests/day free,
 * no credit card required — plenty for personal or small-class use.
 */

const ALLOWED_ORIGIN = '*'; // tighten to your GitHub Pages origin in production, e.g. 'https://yourname.github.io'

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders() });
    }

    if (request.method !== 'POST') {
      return new Response('Method not allowed', { status: 405, headers: corsHeaders() });
    }

    const url = new URL(request.url);
    if (url.pathname !== '/chat') {
      return new Response('Not found', { status: 404, headers: corsHeaders() });
    }

    if (!env.GROQ_API_KEY) {
      return new Response(
        JSON.stringify({ error: 'Server is missing GROQ_API_KEY. Run: wrangler secret put GROQ_API_KEY' }),
        { status: 500, headers: { ...corsHeaders(), 'Content-Type': 'application/json' } }
      );
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return new Response(JSON.stringify({ error: 'Invalid JSON body' }), {
        status: 400,
        headers: { ...corsHeaders(), 'Content-Type': 'application/json' },
      });
    }

    // Basic safety: cap message count/size so one client can't send huge payloads.
    if (!Array.isArray(body.messages) || body.messages.length > 40) {
      return new Response(JSON.stringify({ error: 'Invalid or oversized messages array' }), {
        status: 400,
        headers: { ...corsHeaders(), 'Content-Type': 'application/json' },
      });
    }

    const groqRes = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${env.GROQ_API_KEY}`,
      },
      body: JSON.stringify({
        model: body.model || 'llama-3.3-70b-versatile',
        messages: body.messages,
        stream: true,
        temperature: body.temperature ?? 0.6,
      }),
    });

    // Stream the upstream response straight through to the client.
    return new Response(groqRes.body, {
      status: groqRes.status,
      headers: {
        ...corsHeaders(),
        'Content-Type': 'text/event-stream',
      },
    });
  },
};

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
}
