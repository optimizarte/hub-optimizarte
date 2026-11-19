# Estil unificat «OPTIMIZARTE by Occident» — apps del HUB

Guia curta perquè totes les apps de les targetes del hub semblin **el mateix producte**.
Pilot de referència: `no-renovacion-poliza.html` (arrel i `!!!APPS/`, còpies idèntiques).

Fonts d'on surt: `index.html` (capçalera fosca, DM Sans), simuladors `autos/hogar/copago`
(neutres càlids, radis, ombres OPTICRM) i l'Occident Design System d'OPTICRM
(`!!OPTICRM/extension/PROMPT_CLAUDE_DESIGN.md`).

## Regles d'or

1. **Només estètica.** No es toquen `id`, `name`, `data-*`, classes que usa el JS, l'estructura del DOM
   que llegeix el JS, ni cap text legal o de client. L'idioma de la UI es manté el de l'app
   (les cartes per a clients poden ser en castellà).
2. **Autocontingut.** Un sol HTML amb el CSS a dins. Únic extern d'estil: Google Fonts.
   Llibreries funcionals que ja hi siguin (pdf.js, XLSX.js) es mantenen tal qual.
3. **Tokens, no hex.** Tot color/radi/ombra surt de `:root`. Cap hex solt fora del bloc de tokens
   (excepte dins SVG inline d'icones).
4. **Sempre `.bak`** abans de tocar (`fitxer.html.bak-pre-estil-AAAAMMDD`).

## 1. Tokens (`:root`) — copiar tal qual

```css
:root{
  /* marca */
  --brand:#DC0028; --brand-ink:#B20021; --brand-dim:rgba(220,0,40,.07); --brand-dim2:rgba(220,0,40,.15);
  --dark:#202020;                         /* capçalera, títols de bloc foscos */
  /* fons */
  --bg:#FAFAF7; --bg2:#FFFFFF; --bg3:#F4F3EF; --bg4:#ECEBE6;
  /* vores */
  --br:#E5E3DE; --brl:#CFCDC7;
  /* text */
  --t1:#15161A; --t2:#4F5059; --t3:#8A8B93; --t4:#B4B5BC;
  /* estat (semàntics, no canviar) */
  --ok:#1f9d55;  --ok-bg:#EDFAF3;  --ok-br:#BFE6CF;
  --warn:#B45309; --warn-bg:#FFFBEB; --warn-br:#F6DE9C;
  --err:#B20021; --err-bg:#FFF0F2; --err-br:#F8C9D1;
  --info:#3B82F6;
  /* forma */
  --radius-sm:6px; --radius:10px; --radius-lg:14px; --pill:999px;
  --shadow-sm:0 1px 2px rgba(15,15,30,.05);
  --shadow-card:0 1px 0 rgba(15,15,30,.03),0 1px 3px rgba(15,15,30,.05);
  --shadow-md:0 6px 20px -8px rgba(15,15,30,.10),0 2px 4px rgba(15,15,30,.05);
  /* tipografia */
  --font:'DM Sans',system-ui,-apple-system,'Segoe UI',sans-serif;
  --font-title:Arial,Helvetica,sans-serif;   /* títols: la tipografia «de sempre» de les apps */
}
```

## 2. Tipografia

Text en **DM Sans**; títols (H1 i subtítol de la capçalera) en **Arial/Helvetica negreta**, com les apps
de sempre (decisió de Dany, 12/11/2026).

```html
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&display=swap" rel="stylesheet">
```

| Rol | Font | Mida / pes |
|---|---|---|
| Base | `--font` | 14px / 400, `line-height:1.5`, color `--t1` |
| H1 de pàgina | `--font-title` | 30px (26px mòbil) / 700, paraula clau en `--brand` |
| Títol de bloc (h2) | `--font` | 17px / 600, `letter-spacing:-.01em` |
| Subtítol capçalera | `--font-title` | 15px / 700, blanc 88 % |
| Label de camp («kick») | `--font` | 11px / 600, MAJÚSCULES, `letter-spacing:.7px`, `--t3` |
| Nota / ajuda | `--font` | 11.5–12px, `--t3` |
| Xifra destacada | `--font` | 20–34px / 700, `letter-spacing:-.02em` |

Inter, Roboto o «Catalana» al cos del text queden substituïdes per DM Sans.

## 3. Capçalera i logo

Barra fosca enganxosa (com el hub). Logo de text, sempre igual:

```html
<header class="hdr">
  <a class="hdr-logo" href="https://optimizarte.github.io/hub-optimizarte/" title="Tornar al Hub">
    Optimizarte <span class="by">by</span> <span class="occ">Occident</span>
  </a>
  <span class="hdr-div"></span>
  <span class="hdr-sub">Nom de l'app</span>
  <span class="hdr-sp"></span>
  <span class="hdr-meta">Context curt (p. ex. Art. 22 — Ley 50/1980)</span>
</header>
```

```css
.hdr{position:sticky;top:0;z-index:50;height:60px;background:var(--dark);display:flex;align-items:center;gap:14px;padding:0 20px;box-shadow:0 2px 20px rgba(0,0,0,.25)}
.hdr-logo{font-size:15px;font-weight:700;letter-spacing:-.01em;color:#fff;text-decoration:none;white-space:nowrap}
.hdr-logo .by{font-size:9.5px;font-weight:500;color:#999;margin:0 3px}
.hdr-logo .occ{color:var(--brand)}
.hdr-div{width:1px;height:26px;background:rgba(255,255,255,.14)}
.hdr-sub{font-family:var(--font-title);font-size:15px;font-weight:700;color:rgba(255,255,255,.88);white-space:nowrap}
.hdr-sp{flex:1}
.hdr-meta{font-size:12px;color:rgba(255,255,255,.55)}
```

El logo enllaça al hub amb URL absoluta (`https://optimizarte.github.io/hub-optimizarte/`), perquè
la mateixa còpia funcioni a l'arrel i a `!!!APPS/`. En mòbil s'amaga `.hdr-meta`; sota 420px també el subtítol.

## 4. Botons

Forma **píndola** (`--pill`), 600, 13.5px, padding `10px 18px`, icona SVG 16px amb `gap:8px`.

| Classe | Ús | Estil |
|---|---|---|
| `.btn--primary` | acció principal (1 per vista) | fons `--brand`, text blanc; hover `--brand-ink` |
| `.btn--secondary` | accions secundàries | fons `--bg2`, vora `--br`; hover vora+text `--brand` |
| `.btn--ghost` | reset / tornar | transparent, text `--t2`; hover fons `--bg3` |
| `.btn--green` | confirmació temporal («✓ Copiat») | fons `--ok`, text blanc |

Focus visible: `outline:2px solid var(--brand);outline-offset:2px`.

## 5. Formularis

```css
.field-label{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;color:var(--t3);margin-bottom:6px}
input,select,textarea{width:100%;font:inherit;color:var(--t1);background:var(--bg2);border:1px solid var(--br);border-radius:var(--radius);padding:9px 12px;transition:border-color .15s,box-shadow .15s}
input:focus,select:focus,textarea:focus{outline:none;border-color:var(--brand);box-shadow:0 0 0 3px var(--brand-dim)}
input::placeholder{color:var(--t4)}
```

Zona d'arrossegar fitxer: fons `--bg2`, vora `2px dashed var(--brl)`, radi `--radius-lg`;
estat actiu vora `--brand` i fons `--brand-dim`.

## 6. Targetes i avisos

```css
.card{background:var(--bg2);border:1px solid var(--br);border-radius:var(--radius-lg);box-shadow:var(--shadow-card);padding:16px}
.card--accent{position:relative;overflow:hidden}             /* franja vermella esquerra com el hero dels simuladors */
.card--accent::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--brand)}
.card--calc{background:var(--bg3)}                            /* valors calculats, no editables */
```

Avisos (`.banner--ok/--warn/--error`): fons `--*-bg`, vora `--*-br`, `border-left:4px solid` del color d'estat,
radi `--radius`. Pastilles d'estat (`.badge`): radi `--pill`, 12px/600, mateixos parells de color.

Capçalera de bloc fosca (vista prèvia de document, etc.): fons `--dark`, text blanc, 13.5px/600, radi superior de la targeta.

## 7. Taules

```css
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:var(--bg3);color:var(--t2);font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;text-align:left;padding:9px 12px;border-bottom:1px solid var(--br)}
td{padding:9px 12px;border-bottom:1px solid var(--br)}
tbody tr:hover{background:var(--bg3)}
td.num{text-align:right;font-variant-numeric:tabular-nums}
```

Taula dins `.card` amb `overflow-x:auto` perquè no trenqui el mòbil.

## 8. Maquetació i mòbil

- Contenidor: `max-width:980px` (formularis/documents) o `1180px` (taulers), `margin:0 auto`, padding `28px 20px 56px`.
- Graelles amb `grid-template-columns:repeat(N,minmax(0,1fr))` i `gap:12–14px`.
- **≤ 720px:** graelles a 1–2 columnes, H1 24px, padding lateral 16px, capçalera sense `.hdr-meta`,
  botons d'acció a amplada completa apilats. Cap scroll horitzontal de pàgina.
- Objectius tàctils ≥ 40px d'alt.

## 9. Impressió A4

```css
@page{size:A4;margin:8mm}           /* 6–8 mm */
@media print{
  .hdr,.no-print,.actions,footer{display:none!important}
  body{background:#fff;font-size:11pt}
  .card{box-shadow:none;border:none}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
}
```

Només s'imprimeix el document útil (carta, informe…). Si és una carta, el marge de la pàgina és 8 mm
i la carta afegeix el seu propi «marge de document» interior (~12 mm) perquè respiri.

## 10. Peu

Franja `--bg3`, vora superior `--br`, 11.5px `--t3`, centrat: **OPTIMIZARTE by Occident** · context legal o de l'app.

## Checklist per app

- [ ] `.bak` fet · [ ] tokens enganxats i hex substituïts · [ ] DM Sans carregada i fonts velles fora
- [ ] capçalera/logo estàndard · [ ] botons/formularis/targetes amb les classes de la guia
- [ ] mòbil 375px sense scroll horitzontal · [ ] previsualització d'impressió A4 correcta
- [ ] funcionalitat idèntica (mateixos `id`, mateixos textos legals) · [ ] còpies arrel / `!!!APPS` iguals
