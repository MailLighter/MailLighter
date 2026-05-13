# Debug Taskpane — HTML Viewer (dev)

Spec de restauration de la taskpane de diagnostic **"HTML Viewer (dev)"**.

Quand cette doc est suivie à la lettre, la taskpane peut être remise en place à l'identique sans rien avoir à réexpliquer. Commit d'origine : `cab831da` — `feat: add HTML Viewer dev taskpane + fix nbsp in wrote-attributions`.

## Objectif

Outil interne de diagnostic, **jamais publié au Store**. Surface le ruban Outlook en mode composition pour permettre à un développeur de :

- voir le HTML brut du corps de l'email en cours d'édition
- copier ce HTML dans le presse-papier
- recalculer et copier les positions détectées par `findReplySeparators(html)` (`src/shared/reply-detection.js`)

La taskpane n'est branchée qu'au `MessageComposeCommandSurface`. Le mode lecture n'est pas couvert (pas besoin du debug dans ce mode).

## UI

Titre : `HTML Viewer (dev)`.

Toolbar, dans l'ordre :

| id | Style | Libellé |
|----|-------|---------|
| `btnRefresh` | primary (bleu `#0078d4`) | `Refresh` |
| `btnCopy` | default | `Copy HTML` |
| `btnCopySep` | default | `Copy Separators` |

Zones d'affichage :

- `<div id="status">` : messages courts au-dessus des zones. Départ : `"Click Refresh to load the email HTML."`
- `<pre id="htmlOutput">` : HTML brut affiché via `textContent`. `white-space: pre-wrap`, `word-break: break-all`, `max-height: 45vh`, police `Consolas, "Courier New", monospace` 11 px
- `<div id="separators">` : conteneur avec un `<h3>Detected separators</h3>` et un `<div id="sepList">`. État initial : `<span class="sep-none">No data yet.</span>`. Chaque séparateur rendu en `<div class="sep-item">` contenant :
  - `<div class="label">Separator N — position P</div>` (orange `#d83b01`)
  - un excerpt `html.substring(pos-20, pos+120)` avec un `<b style="color:#d83b01">|CUT|</b>` inséré exactement à la position `pos`, le reste échappé via une fonction locale `escapeHtml`

Styles : **tous inline dans `<head><style>…</style>`** du fichier HTML. Pas de CSS externe. Pas de i18n.

## Logique JS

Fichier : `src/taskpane/dev-viewer.js`.

Imports :

```js
/* global Office */
import { findReplySeparators } from "../shared/reply-detection";
```

État : `let lastHtml = ""` au niveau module.

Init :

```js
Office.onReady(() => {
  document.getElementById("btnRefresh").addEventListener("click", refresh);
  document.getElementById("btnCopy").addEventListener("click", copyToClipboard);
  document.getElementById("btnCopySep").addEventListener("click", copySeparators);
});
```

`refresh()` :

1. Récupérer `const body = Office.context.mailbox.item && Office.context.mailbox.item.body;`
2. Garde-fou : `if (!body || typeof body.getAsync !== "function") { setStatus("body.getAsync not available in this context."); return; }`
3. `body.getAsync(Office.CoercionType.Html, (result) => { … })`
4. Sur `result.status !== Office.AsyncResultStatus.Succeeded` : `setStatus("Error: " + (result.error ? result.error.message : "unknown"))`
5. Sur succès :
   - `lastHtml = result.value || "";`
   - `document.getElementById("htmlOutput").textContent = lastHtml;`
   - `setStatus("Loaded " + lastHtml.length + " chars — " + new Date().toLocaleTimeString());`
   - `showSeparators(lastHtml);`

`showSeparators(html)` :

- `positions = findReplySeparators(html)` (tableau de `number`)
- Si `positions.length === 0` : `sepList.innerHTML = '<span class="sep-none">No separators detected.</span>'`
- Sinon : pour chaque position, `before = html.substring(Math.max(0, pos-20), pos)`, `after = html.substring(pos, Math.min(html.length, pos+120))`. Excerpt = `escapeHtml(before) + '<b style="color:#d83b01">|CUT|</b>' + escapeHtml(after)`. Les items sont joints et injectés dans `sepList.innerHTML`

`escapeHtml(str)` local (ne pas importer depuis `office-helpers.js`) :

```js
return str
  .replace(/&/g, "&amp;")
  .replace(/</g, "&lt;")
  .replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;");
```

`copyToClipboard()` :

- Si `!lastHtml` : `setStatus("Nothing to copy — click Refresh first.")`
- Sinon préférer `navigator.clipboard.writeText(lastHtml)` → `"Copied to clipboard!"`
- Fallback `fallbackCopy()` → crée un `<textarea>` hidden (`position: fixed; opacity: 0`), `select()`, `document.execCommand("copy")` → `"Copied to clipboard! (fallback)"` ou `"Copy failed — select the HTML manually."`

`copySeparators()` : même logique, source = `document.getElementById("sepList").innerText`. Si texte est `""`, `"No data yet."` ou `"No separators detected."` → `setStatus("No separator data to copy.")`. Messages de succès : `"Separators copied!"` / `"Separators copied! (fallback)"` / `"Copy failed."`

`setStatus(msg)` : `document.getElementById("status").textContent = msg;`

## Webpack

Fichier : `webpack.config.js`.

Dans `entry` ajouter :

```js
"dev-viewer": "./src/taskpane/dev-viewer.js",
```

Dans `plugins`, après le `HtmlWebpackPlugin` de `commands.html`, ajouter :

```js
new HtmlWebpackPlugin({
  filename: "dev-viewer.html",
  template: "./src/taskpane/dev-viewer.html",
  chunks: ["polyfill", "dev-viewer"],
}),
```

## Manifest

Fichier : `manifest.xml`. **Chaque modification doit être dupliquée dans les deux blocs `VersionOverrides`** (V1_0 puis V1_1). Le deuxième `<Control>` Button prend l'id `msgComposeDevViewer2` pour éviter le conflit.

### Bouton dans chaque `<ExtensionPoint xsi:type="MessageComposeCommandSurface">`

À ajouter **après** le dernier `<Control>` du groupe Quick Actions (après le bouton Full cleanup) :

```xml
<Control xsi:type="Button" id="msgComposeDevViewer">
  <Label resid="DevViewer.Label"/>
  <Supertip>
    <Title resid="DevViewer.Label"/>
    <Description resid="DevViewer.Tooltip"/>
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="Icon.16x16"/>
    <bt:Image size="32" resid="Icon.32x32"/>
    <bt:Image size="80" resid="Icon.80x80"/>
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="DevViewer.Url"/>
  </Action>
</Control>
```

Dans le 2ème bloc `VersionOverrides`, remplacer `id="msgComposeDevViewer"` par `id="msgComposeDevViewer2"`.

Icônes réutilisées : `Icon.16x16` / `Icon.32x32` / `Icon.80x80` (logo MailLighter déjà présent dans `<bt:Images>`). **Aucune nouvelle icône à créer.**

### Ressources à ajouter dans chaque `<Resources>`

```xml
<bt:Urls>
  <!-- existant -->
  <bt:Url id="DevViewer.Url" DefaultValue="https://localhost:3000/dev-viewer.html"/>
</bt:Urls>
```

```xml
<bt:ShortStrings>
  <!-- existant -->
  <bt:String id="DevViewer.Label" DefaultValue="HTML Viewer"/>
</bt:ShortStrings>
```

```xml
<bt:LongStrings>
  <!-- existant -->
  <bt:String id="DevViewer.Tooltip" DefaultValue="Show the raw HTML of the email body (dev tool)."/>
</bt:LongStrings>
```

Libellés volontairement en anglais uniquement, pas de `<bt:Override>` : outil interne.

## i18n

Aucune clé à ajouter dans `src/i18n/*.json`. Tous les libellés (titre, boutons, messages de statut) sont en anglais en dur dans `dev-viewer.html` / `dev-viewer.js`.

## Vérification end-to-end

1. `npm run build:dev` → build réussit, `dist/dev-viewer.html` et `dist/dev-viewer.js` présents
2. `npm run validate` → manifeste valide
3. `npm run dev-server` → https://localhost:3000/dev-viewer.html retourne le HTML
4. Sideload Outlook Desktop → ouvrir un nouveau mail (compose) → le bouton **HTML Viewer** apparaît dans le groupe "Quick Actions" du ruban MailLighter
5. Clic sur le bouton → la taskpane s'ouvre à droite, titre "HTML Viewer (dev)"
6. Clic **Refresh** → `<pre>` affiche le HTML complet du body, statut indique `Loaded N chars — HH:MM:SS`, la zone "Detected separators" liste les séparateurs trouvés (ou indique `No separators detected.`)
7. Clic **Copy HTML** → statut `Copied to clipboard!`, le presse-papier contient le HTML
8. Clic **Copy Separators** → statut `Separators copied!`, le presse-papier contient la liste formatée des séparateurs

## Suppression propre

Dans l'ordre :

1. Supprimer les fichiers `src/taskpane/dev-viewer.html` et `src/taskpane/dev-viewer.js`. Si le répertoire `src/taskpane/` devient vide, le supprimer aussi
2. Retirer de `webpack.config.js` :
   - la ligne d'entry `"dev-viewer": "./src/taskpane/dev-viewer.js"`
   - le `HtmlWebpackPlugin` pour `dev-viewer.html`
3. Retirer de `manifest.xml`, **dans chacun des 2 blocs `VersionOverrides`** :
   - les `<Control>` Button `msgComposeDevViewer` et `msgComposeDevViewer2`
   - les 2 occurrences de `<bt:Url id="DevViewer.Url" …/>`
   - les 2 occurrences de `<bt:String id="DevViewer.Label" …/>`
   - les 2 occurrences de `<bt:String id="DevViewer.Tooltip" …/>`
4. Lancer `npm run build:dev` puis `npm run validate` — tout doit passer
5. Commit : `chore: remove HTML Viewer dev taskpane`

## Restauration

Phrase à donner à Claude : **« Remets en place la taskpane de debug selon `docs/debug-taskpane.md`. »**

Variante plus rapide si le commit d'origine est encore atteignable dans l'historique : **« Cherry-pick les parties dev-viewer du commit `cab831da` en suivant `docs/debug-taskpane.md`. »**
