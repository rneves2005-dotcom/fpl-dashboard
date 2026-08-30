// FPL API CORS proxy — deploy as a Cloudflare Worker (free tier).
// It forwards <worker-url>/<path> to https://fantasy.premierleague.com/api/<path>
// and adds the CORS headers the browser needs. Caches 30s at the edge.
//
// Deploy: dash.cloudflare.com -> Workers & Pages -> Create -> Worker ->
//   name it "fpl-proxy" -> Deploy -> Edit code -> paste this -> Deploy.
// Then the URL is https://fpl-proxy.<your-subdomain>.workers.dev

export default {
  async fetch(request) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: cors() });
    }
    const url = new URL(request.url);
    const target = 'https://fantasy.premierleague.com/api' + url.pathname + url.search;
    const upstream = await fetch(target, {
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' },
      cf: { cacheTtl: 30, cacheEverything: true }
    });
    const body = await upstream.text();
    return new Response(body, {
      status: upstream.status,
      headers: { ...cors(), 'Content-Type': 'application/json' }
    });
  }
};

function cors() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': '*',
    'Cache-Control': 'public, max-age=30'
  };
}
