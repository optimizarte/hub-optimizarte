# BRIEF — Pressa de venciments al modal Alta Clients → creació d'OPOs al sidebar OPTICRM

**Data:** 2026-09-13
**Preparat per:** Sessió de continuació amb Dany present
**Font:** Memòria persistent (`/areas/opticrm.md`, `/areas/alta-clients.md`, projects/CRM-OPTIMIZARTE) + `conversation_search` a 4 sessions rellevants d'abril-maig 2026

---

## 1. Estat del que tenim documentat

### 1.1 Formulari Alta Clients (`alta-clientes.html`, HTML monolític, ~5.911 línies)

- Desplegat a OneDrive `hub-optimizarte/alta-clientes.html`, carregat dins un iframe a `catalanaaplicaciones.gco.global`
- 11 cards de ramos (hogar, auto, moto, vida, salud, decesos, ahorro-g/i, embarcaciones, comunidades, otro)
- Cada pill té:
  - **Estat** — un dels 7 valors: `vto-inmediato`, `necesita`, `facilita`, `no-cambia-vto`, `no-cambia`, `no-necesita`, `en-vigor`
  - **Modal de dades adicionals** amb camps segons ramo (mesVcto, prima, matrícula, capital, personas, notes...)
  - **Estructura de dades global** `ramoData = {'hogar-1': {estado, mesVcto, ...}}`

### 1.2 Bridge postMessage (app.js OPTICRM, secció `// ── FORM BRIDGE`, ~línia 15906)

Handlers documentats:
- `save_client`, `load_client`, `list_clients`, `search_client_by_nif` — CRUD de fitxers `cli_*.json` a OneDrive
- `save_alta` — desar draft del formulari
- `generate_report` — passar per generar HTML + mailto
- `send_mail`, `_opticrm_form_alta_payload` — endpoint on el formulari envia payload complet

### 1.3 Motor d'anàlisi de payload — `_altaAnalyzePayload()` (app.js ~línia 4311)

Rep el payload amb `ramoData` i genera OPOs + ACTs segons taula:

| Estat pill | Genera | Columna kanban | Títol |
|---|---|---|---|
| `vto-inmediato` | **OPO** | `tarifa` | `VTO <ramo> · <client>` |
| `facilita` | **OPO** | `espera` | `VTO <ramo> · <client>` |
| `no-cambia-vto` | **OPO** | `espera` | `VTO <ramo> · <client>` |
| `necesita` | **ACT** | — | `Registrar Hot Candis <RAMO> · <client>` |
| `no-cambia` | **ACT** | — | `Registrar Cold Candis <RAMO> · <client>` |
| Camp `primera_accion` del formulari | **ACT** | — | Mapejat segons tipus (Llamada/Cita/Visita/Team) |

Notes autogenerades via `_altaBuildOpoNotes()`: companyia, VTO, prima, capital, matrícula, observacions, dades del client (F.nac, sexo) + llista d'altres ramos.

Textos **CANDI** amb pattern `CANDI<YYYY><RAMO_MAYUS_SIN_ESPAIS>` (ex: `CANDI2026HOGAR`, `CANDI2026RCPYME`) — es mostren a l'informe HTML en caixa groga/vermella per copiar manualment al camp d'etiquetes de Gestiona.

### 1.4 Modal previ — `_showAltaPreviewModal()` (app.js ~línia 4415)

Modal que mostra el llistat d'OPOs+ACTs a crear ABANS d'aplicar-les. Documentat: **no valida** (pot mostrar-se sense OPO/ACT). Té un botó "Aplicar dades" que és el punt crític del flux.

### 1.5 AGENDA VTOS (històric del bookmarklet, abril 2026)

Un mòdul dissenyat inicialment al **bookmarklet** predecessor (mai confirmat si es va migrar a l'extensió OPTICRM actual):
- Pseudo-usuari local "AGENDA VTOS" al sidebar per sobre de "Colaboradores"
- 4a columna a Vista Día amb color ambar `#F59E0B`
- Items guardats a `localStorage.optimizarte_vtos` amb estructura `{id, asunto, cliente, fecha, tipo, notas, creado}`
- Al drag&drop VTO → columna d'usuari real: crear activitat al CRM via
  ```js
  crmPost('ISC.Gaan.Actividad.FE.DetalleActividad', 'alta', {
    UID_RACF, DESASUN, TIM_INICI, INDACTI_DESC,
    AE_COD_MEDIADOR_RACF: '00017167', CIAGRUPO: 'SCO'
  })
  ```
  → **aquest endpoint mai va ser validat en producció**, era una hipòtesi a verificar amb DevTools.

---

## 2. Buits d'informació crítics

Aquestes són les preguntes que **cal resoldre abans de tocar codi**, ordenades per importància:

### Nivell 1 — bloqueants

1. **On queden les OPOs generades per `_altaAnalyzePayload()`?**
   Tres opcions plausibles: (a) es creen directament al Gestiona via fetch (aleshores el CRM les retorna via `_syncFromCRM()`), (b) es guarden a localStorage/OneDrive com a items pendents, (c) el "Aplicar dades" del `_showAltaPreviewModal` no fa res avui (o fa alguna cosa parcial). No hi ha confirmació documentada.

2. **Existeix una vista al sidebar OPTICRM per veure aquestes opos pendents?**
   Sé que hi ha AGENDA VTOS al bookmarklet vell, però no tinc confirmació que s'hagi migrat a l'extensió. Cal `grep -n "AGENDA VTOS\|vtosNew\|vtosToggle\|optimizarte_vtos" app.js` al fitxer actual.

3. **L'endpoint `crmPost('ISC.Gaan.Actividad.FE.DetalleActividad', 'alta', ...)` funciona?**
   És una hipòtesi de la conversa d'abril 2026. Cal validar amb DevTools capturant una alta manual d'activitat al Gestiona actual.

### Nivell 2 — importants

4. **Bugs pendents del 19 maig — estat actual:**
   - Bug 1: validació al HTML del formulari que requereix VTO/ACT abans de "Donar alta" — Dany va dir explícitament "no ha de ser necessari"
   - Bug 2: toast vermell "Dades del client llestes" no apareix + botó "Aplicar dades" manual desaparegut. Causa identificada: `chrome.storage.local` no accessible des d'app.js (injectat com a `<script>`, no com a content script). Solució proposada: postMessage entre finestres via `window.opener`.

   Cap dels dos té marker de patch aplicat a la memòria. Cal buscar-los al codi actual.

5. **Codis mediador per crear OPO** — segons `platform-notes.md`:
   > OPO mediador uses the field CodigosAsociados1 with numeric codes, handler changeCodigosAsociados(); ComboBEMediador is legacy/ghost.

   Els codis numèrics han de ser mapejats per usuari. A validar quins codis usar (Dany, Fadoua, Silvia).

6. **Preselec de client** — `platform-notes.md`:
   > New OPO/ACT forms open via HTTP POST (not GET); IdReferencia in the URL preselects the client automatically.

   Això és útil si la propagació al Gestiona s'ha de fer amb el client ja seleccionat.

### Nivell 3 — a decidir

7. **Fluxe desitjat** — la petició de Dany diu: "crear noves opo's al sidebar perquè l'usuari les pugui propagar al gestiona i del gestiona al CRM". Hi ha dues interpretacions:
   - **A** (semiautomàtic): OPOs es creen al sidebar OPTICRM com a pendents locals. L'usuari les revisa, decideix quines propagar, i clica un botó per crear-les al Gestiona. La sync automàtica del CRM (5 min) les torna com a OPOs reals al OPTICRM.
   - **B** (assistit): OPOs es creen al sidebar amb un enllaç/botó que obre el formulari Nova OPO del Gestiona amb tots els camps preomplerts (via POST amb IdReferencia + `_altaBuildOpoNotes()` al camp de notes). L'usuari revisa i clica "Guardar" al Gestiona.

   L'opció B és més coherent amb "propagar-al gestiona i del gestiona al CRM" perquè manté el Gestiona com a font única de veritat i evita duplicats. Però cal confirmar amb Dany.

---

## 3. Ordre de treball proposat per la propera sessió (amb Dany present)

**Pas 0 (5 min)** — Confirmació de fluxe. Dany decideix A vs B del punt 7 anterior.

**Pas 1 (15 min)** — Lectura estàtica. Necessito app.js sencer per fer:
```bash
grep -n "_altaAnalyzePayload\|_showAltaPreviewModal\|_opticrm_form_alta_payload\|AGENDA VTOS\|optimizarte_vtos\|_altaBuildOpoNotes" app.js
```
Amb això sé exactament què fa avui el flux.

**Pas 2 (15 min)** — Test amb Claude in Chrome (amb Dany). Obrir un client real al CRM, filar 2-3 ramos al formulari, clicar "Donar alta", capturar tot el network que surti per veure:
- Quins missatges postMessage surten
- Si hi ha fetch cap al CRM
- Què acaba passant al final del flux

**Pas 3 (30 min)** — Validar endpoints CRM per crear OPO/ACT nova. Capturar una alta manual d'ACT i d'OPO al Gestiona amb DevTools Network per tenir el payload exacte.

**Pas 4 (variable)** — Disseny del patch. Amb tota la info anterior, proposar canvis amb taula de risc i esperar confirmació.

---

## 4. El que NO faré ara sense Dany present

- Tocar `app.js` a producció (3 PCs de l'agència)
- Fer test al Gestiona real (dades de clients reals)
- Assumir que l'endpoint `ISC.Gaan.Actividad.FE.DetalleActividad` funciona sense verificar
- Reactivar codi orfe/experimental (com les funcions eliminades del header d'alta-clientes.html)

---

## 5. Nota sobre el que sí es va fer avui

Aquesta sessió (13 setembre 2026) va aplicar només aquests canvis a `alta-clientes.html`, **no a `app.js`**:
- Netejar header (eliminar `hdr-actions` sencer)
- Afegir pill "Nou client (no està al CRM ni a OD)"
- CSS `body[data-flow="found"]` per amagar `#sec-datos` i `#sec-contacto`
- Banner informatiu `#existingClientBanner`
- ZIP de skill actualitzat generat a outputs

Cap d'aquests canvis afecta el flux de vencimientos → opos.
