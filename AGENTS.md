# AGENTS.md

<!-- handoff:learnings:begin -->
## Session Learnings (auto-updated by handoff)

### Implicit Rules
- App is a single self-contained HTML file — all state lives in embedded JSON seed blocks, no external dependencies except Google Fonts and XLSX.js CDN
- API base: https://api.tesis.io/v2/ — CORS origin: https://anagan.app.tesis.io — Auth: Bearer JWT in Authorization header
- JWT token expires in ~12 hours (exp - iat = 43200s); iat/exp are Unix timestamps in the payload
- API pagination: GET /listadoClientes?grupos=...&page=N&limit=50 returns {items:[], total:733, totalPages:15, ...}
- GET /countClientes returns {total:733, totalPages:N} — use for preflight count
- Client ID in seed data: a[0] (field 'id') matches API response field 'id' exactly — this is the join key
- Seed rawRows format: [id, cia, ef, vt, ramo, tom, rie, colid, col, tipo, sub, prima, est, pol, dni]
- Hub conventions: catalan language, DM Sans font, #DC0028 red, #202020 dark, no external deps
- Windows environment — Python unavailable, use PowerShell COM or Invoke-RestMethod for data tasks
- gstack browse skill (headless Chromium) does not work on this Windows machine — use computer-use screenshot of Edge instead

### Key Decisions
- Decision: API connector (Camino B) over name-based Excel matching — Reason: Excel ID column empty for all 733 rows, name matching ~85-90% accuracy insufficient, direct API access confirmed available
- Decision: PowerShell script updates HTML file directly instead of in-page fetch — Reason: HTML served from file:// origin, API CORS only permits https://anagan.app.tesis.io, browser fetch would be blocked
- Decision: Inject client data as separate <script id='clients-seed'> block alongside existing <script id='seed'> — Reason: preserves existing save/export flow unmodified, easy to update independently
- Decision: Limit=50 with 15 pages for API pagination — Reason: API returns wrapper {items, total, totalPages} not bare array; 50 items/page balances speed and server load
- Decision: Fields extracted: id, nif, tomador, telefonoMovil/telefonoPrincipal (first non-empty), email, direccionCompletaFiscal, localidad, provincia, cp, tipoDocumento — Reason: confirmed present in API response via field inspection

<!-- handoff:learnings:end -->
