# Optional: server-side key proxy

By default, Excel Tutor runs entirely client-side — each learner pastes their
own free Groq key, stored only in their browser. That's the simplest setup
and needs zero deployment.

If you're distributing this to other people (e.g. your import-business
training program) and don't want each person to need their own Groq account,
deploy this Cloudflare Worker instead. It holds *your* key server-side and
the app talks to it instead of Groq directly — no key prompt for end users.

## Deploy (free, no credit card)

```bash
npm install -g wrangler
wrangler login
cd server
wrangler secret put GROQ_API_KEY   # paste your key when prompted
wrangler deploy
```

You'll get a URL like `https://exceltutor-proxy.<your-subdomain>.workers.dev`.

## Wire it up

Open `config.js` in the app root and set:

```js
const BACKEND_URL = 'https://exceltutor-proxy.<your-subdomain>.workers.dev';
```

Reload the app — the API-key onboarding screen disappears entirely, and all
requests go through your worker instead.

## Notes

- Free tier: 100,000 requests/day on Cloudflare Workers, no card required —
  comfortably enough for personal or small-group use.
- Tighten `ALLOWED_ORIGIN` in `worker.js` from `'*'` to your actual GitHub
  Pages URL before sharing this widely, so only your deployed app can call it.
- This worker doesn't rate-limit per learner yet — for a public class rollout
  you'd want to add a Cloudflare KV-based counter per IP/session; ask me if
  you want that added later.
- You can switch back to bring-your-own-key mode any time by clearing
  `BACKEND_URL` back to `''`.
