// ─────────────────────────────────────────────────────────────
// Cloudflare Pages Function
// File path in your repo MUST be:  functions/api/submit.js
//   → this automatically serves the route  POST /api/submit
//   (so your frontend fetch('/api/submit') needs NO changes)
//
// Updated for the SHORTENED assessment: 5 regimes (was 6 dimensions).
// Payload keys are now reg01..reg05 + overall + overallUnweighted.
//
// Set the secret in: Cloudflare Dashboard → your Pages project →
//   Settings → Environment variables → add TEAMS_WEBHOOK_URL
// ─────────────────────────────────────────────────────────────
const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST,OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

// Handle CORS preflight (OPTIONS)
export async function onRequestOptions() {
  return new Response(null, { status: 204, headers: CORS });
}

// Handle the actual submission (POST)
export async function onRequestPost(context) {
  const { request, env } = context;
  const jsonHeaders = { ...CORS, 'Content-Type': 'application/json' };
  const WEBHOOK_URL = env.TEAMS_WEBHOOK_URL;

  if (!WEBHOOK_URL) {
    console.error('TEAMS_WEBHOOK_URL not set');
    return new Response(
      JSON.stringify({ error: 'Webhook not configured' }),
      { status: 500, headers: jsonHeaders }
    );
  }

  try {
    const d = await request.json();

    // Build RAG emoji
    const ragEmoji = d.rag === 'Red' ? '🔴' : d.rag === 'Amber' ? '🟡' : '🟢';

    // Build answer details if available
    let answerDetails = '';
    if (d.answers && Array.isArray(d.answers)) {
      answerDetails = d.answers
        .map((a) => `[${a.regime}] ${a.scenario}: ${a.selectedOption} (${a.score}/4)`)
        .join('\n');
    }

    // Format as Adaptive Card for clean Teams channel rendering
    const card = {
      type: 'message',
      text: JSON.stringify(d),
      attachments: [
        {
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: {
            $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
            type: 'AdaptiveCard',
            version: '1.4',
            body: [
              {
                type: 'TextBlock',
                text: `📊 New Assessment: ${d.company || 'Unknown'}`,
                weight: 'Bolder',
                size: 'Medium',
                wrap: true,
              },
              {
                type: 'TextBlock',
                text: `${d.name || 'Anonymous'} · ${d.contact || 'no contact'} · ${d.role || 'N/A'} · ${d.model || 'N/A'}${d.entities ? ' · ' + d.entities + ' entities' : ''}`,
                isSubtle: true,
                wrap: true,
              },
              {
                type: 'TextBlock',
                text: `${ragEmoji} Overall Score: ${d.overall}% (${d.rag})${d.overallUnweighted != null ? ` · Unweighted ${d.overallUnweighted}%` : ''}`,
                weight: 'Bolder',
                color: d.rag === 'Red' ? 'Attention' : d.rag === 'Amber' ? 'Warning' : 'Good',
              },
              {
                type: 'FactSet',
                facts: [
                  { title: 'Regime 01 - Daily Operations', value: `${d.reg01}%` },
                  { title: 'Regime 02 - Market Volatility', value: `${d.reg02}%` },
                  { title: 'Regime 03 - Crisis Resilience', value: `${d.reg03}%` },
                  { title: 'Regime 04 - Strategic Capital', value: `${d.reg04}%` },
                  { title: 'Regime 05 - Group Scale', value: `${d.reg05}%` },
                ],
              },
              {
                type: 'TextBlock',
                text: `Submitted: ${d.timestamp || new Date().toISOString()}`,
                isSubtle: true,
                size: 'Small',
              },
              {
                type: 'TextBlock',
                text: JSON.stringify(d),
                isVisible: false,
                id: 'rawPayload',
              },
            ],
          },
        },
      ],
      // Full payload as JSON string for Power Automate to parse
      summary: JSON.stringify(d),
    };

    const webhookRes = await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(card),
    });

    if (!webhookRes.ok) {
      const errText = await webhookRes.text();
      console.error('Webhook error:', webhookRes.status, errText);
      return new Response(
        JSON.stringify({ error: 'Webhook delivery failed', status: webhookRes.status }),
        { status: 502, headers: jsonHeaders }
      );
    }

    return new Response(
      JSON.stringify({ ok: true, message: 'Assessment submitted successfully' }),
      { status: 200, headers: jsonHeaders }
    );
  } catch (err) {
    console.error('Submit error:', err);
    return new Response(
      JSON.stringify({ error: 'Internal server error' }),
      { status: 500, headers: jsonHeaders }
    );
  }
}
