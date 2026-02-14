# Money Report (Approval Draft)

Upload a quotation file (Excel/PDF) and generate a Korean official-style approval draft (품의서).

## Deploy (Cloudflare Pages)

1. Cloudflare Dashboard -> `Workers & Pages` -> `Pages` -> `Create application`
2. `Connect to Git` -> choose `ninanonanssilo/money-report`
3. Build settings:
   - Framework preset: `None`
   - Build command: (empty)
   - Build output directory: `/`
4. Deploy

### OpenAI (Optional)

Recommended: set an environment variable in Cloudflare Pages:

- Variable: `OPENAI_API_KEY`

If you do not set it, the UI falls back to a deterministic template.

## Local Dev

This is a static app plus Pages Functions.

- Static preview:
  - `python3 -m http.server 8080`
- Functions require Cloudflare dev tooling (optional):
  - `npx wrangler pages dev .`
