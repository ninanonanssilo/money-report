# Blueprint: Automated Approval Draft (Money Report)

## Goal

Upload a quotation file (Excel or PDF) and automatically draft a Korean official-style approval request document (품의서 / 품의 공문).

## MVP Scope

- Client (static):
  - Upload: `.xlsx`, `.pdf`
  - Extract: basic vendor/items/amount/date fields when possible
  - Input: requester/department/subject/purpose/budget info
  - Generate: 품의서 HTML preview + "Print to PDF"
  - Optional: OpenAI API key can be provided by the user (text input or .txt file)
- Server (Cloudflare Pages Functions):
  - `/api/generate`: builds a clean, formal 품의서 draft
    - Uses Cloudflare env `OPENAI_API_KEY` if configured
    - Otherwise accepts an `apiKey` passed from the browser (not stored)
    - Falls back to a deterministic template if AI is not configured

## Non-Goals (MVP)

- Perfect parsing of all PDF quotation formats (scans/images)
- OCR (requires extra service)
- Storage/authentication

## Next

- Add structured extraction for common Excel templates
- Add OCR pipeline for scanned PDFs
- Add "company letterhead" export and approval workflow steps
