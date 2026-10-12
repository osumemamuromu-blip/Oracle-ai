# Oracle — Cloudflare Workers starter

This package is prepared for Cloudflare Workers with static assets. The API key is read only from the Cloudflare Worker secret named `GEMINI_API_KEY`; do not place it in `public/`, GitHub, or frontend code.

## Deploy from a GitHub repository
1. Upload the contents of this folder to the root of a new GitHub repository (not the ZIP file itself). The repository root should contain `worker.js`, `wrangler.jsonc`, `README.md`, and `public/index.html`.
2. In Cloudflare, open Workers & Pages and create a Worker from the GitHub repository, or connect the repository from the Workers & Pages Git integration.
3. Configure the project to use Wrangler configuration (`wrangler.jsonc`). If asked for build settings, use no build command and no output directory; the Wrangler config defines the Worker and `public` assets.
4. After the first deploy, open the Worker Settings > Variables and Secrets area. Add a secret named `GEMINI_API_KEY` with your Google AI Studio key. Do not share the key in chat or commit it to GitHub. Save and redeploy if prompted.
5. Open the deployed `workers.dev` URL and test the writing tools. `https://YOUR-WORKER-URL/api/health` should return `aiConfigured: true` once the secret is set.

## Notes
- This is a starter, not a finished commercial product. Authentication, persistent per-user free-use tracking, subscriptions/payments, and abuse prevention are not implemented yet. Do not advertise the two-free-checks or Pro billing as active until those systems are built and tested.
- Cloudflare and Google AI Studio have separate quotas and terms. Check current usage limits before launch.
