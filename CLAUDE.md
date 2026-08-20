# Hub Optimizarte — Instruccions per a Claude

## Carpeta del projecte
Aquesta carpeta és la raó de treball principal per a totes les eines d'OPTIMIZARTE by Occident.

**Ruta completa:**
```
C:\Users\primo\OneDrive - OPTIMIZARTE 3.0 - SCO\!!IA\!!GitHub-Repository-HUB\HUB-OPTIMIZARTE
```

**Al inici de cada sessió**, si l'usuari menciona el hub, eines d'oficina o qualsevol dels fitxers
de la llista de sota, sol·licita immediatament accés a la carpeta amb:
```
mcp__cowork__request_cowork_directory(path="C:\\Users\\primo\\OneDrive - OPTIMIZARTE 3.0 - SCO\\!!IA\\!!GitHub-Repository-HUB\\HUB-OPTIMIZARTE")
```
Aixó evita que l'usuari hagi de seleccionar-la manualment cada vegada.

## Fitxers actuals del hub

| Fitxer | Descripció |
|--------|------------|
| `index.html` | Hub principal (català, DM Sans, sidebar + grid de targetes) |
| `CalendarioPagosPolizas_Occident.html` | Calendari de pagaments de pòlisses |
| `historics-seguiment.html` | Tancaments mensuals (DAFO, Tancament, Fitxa Mediador) |
| `alta-clientes.html` + `.css` + `.js` | Formulari d'alta de client |
| `generador-cambio-mediador.html` | Generador de cartes de canvi de mediador |
| `no-renovacion-poliza.html` | Avís de no renovació (Art. 22, Llei 50/1980) |

## Convencions del hub (`index.html`)

- Idioma: **català** (`lang="ca"`)
- Fonts: DM Sans + DM Serif Display (Google Fonts)
- Colors: `#202020` fosc, `#DC0028` vermell Occident
- Noves targetes: afegir a `#grid-live` amb `data-category` i `data-status="live"`
- Badge de categoria al sidebar: actualitzar manualment el comptador
- Sempre escriure canvis directament a aquesta carpeta (no només a outputs)

## Convencions tècniques

- Fitxers grans: usar Python per escriure (bash heredoc trunca)
- No afegir dependències externes (excepte Google Fonts)
- Tots els HTML inclouen `@media print` (A4, marges 6–8mm)
