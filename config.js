/* ============================================================
   Excel Tutor config
   ------------------------------------------------------------
   Leave BACKEND_URL empty to run in "bring your own key" mode
   (the learner pastes a free Groq key, stored only in their
   browser — this is the default and needs zero setup).

   To hide the API key step entirely and proxy requests through
   your own serverless function instead (recommended if you're
   deploying this for other people, e.g. a class or training
   program), deploy the Cloudflare Worker in /server and paste
   its URL here, e.g.:
     const BACKEND_URL = 'https://exceltutor-proxy.yourname.workers.dev';
   ============================================================ */
const BACKEND_URL = '';
