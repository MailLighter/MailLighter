# MailLighter - Contexte Projet

## Vue d'ensemble

MailLighter est un add-in Outlook construit sur la base du template Office Add-in JavaScript.
Le projet cible principalement Outlook Desktop et fournit des actions rapides pour nettoyer ou réutiliser le contenu d'un email.

Le projet expose deux surfaces UI principales :

- le menu déroulant dans le ruban Outlook (actions directes)
- un dialogue de paramètres (message écologique + compteur d'économies)

Une taskpane `dev-viewer` est également livrée comme outil de debug (HTML viewer).

## Fonctionnalités métier actuelles

### En composition

- Ne garder que la sélection (remplace le corps par le texte sélectionné)
- Supprimer les images
- Supprimer les pièces jointes
- Conserver uniquement les 2 dernières réponses
- Nettoyage complet (images + pièces jointes + réponses)
- Ouvrir les paramètres (message écologique, économies cumulées)

### En lecture de message

- Ouvrir les paramètres (mêmes options qu'en composition)

## Structure du dépôt

### Fichiers racine importants

- `manifest.xml` : manifeste Office de l'add-in, avec les commandes, labels, tooltips et ressources localisées
- `package.json` : scripts npm et dépendances
- `webpack.config.js` : build webpack, copie des assets, génération de `commands.html`, `dev-viewer.html` et `settings.html`
- `claude.md` : ce fichier de contexte

### Dossiers importants

- `src/commands/` : logique des commandes Outlook exécutées depuis le ruban
- `src/settings/` : dialogue des paramètres (message écologique, économies)
- `src/taskpane/` : taskpane `dev-viewer` (outil de debug du HTML du mail)
- `src/i18n/` : traductions JSON (`en-US`, `fr-FR`, `es-ES`)
- `src/shared/` : utilitaires partagés (i18n, office-helpers, reply-detection, savings-storage)
- `assets/` : logos et icônes PNG utilisés par le manifeste et les pages HTML
- `dist/` : sortie générée par webpack, à ne pas modifier manuellement

## Fichiers source clés

### `src/commands/commands.js`

Contient la logique métier principale côté commandes Outlook.

Points importants :

- `notify()` affiche une notification Outlook via `notificationMessages.replaceAsync`
- `officeAsync()` centralise les wrappers Promise autour des API Office async
- `keepSelectionOnlyCore()` capture la sélection HTML puis remplace le corps en conservant la zone utilisateur en amont de `_MailOriginal` / `divRplyFwdMsg`
- `keepTwoRepliesCore()` travaille sur le HTML et détecte les séparateurs de réponses via `findReplySeparators` (voir `src/shared/reply-detection.js`)
- `openSettingsCore()` ouvre le dialogue de paramètres en passant les économies cumulées et l'état du message écologique via URL search params
- les commandes sont enregistrées via `Office.actions.associate(...)`

### `src/settings/settings.js` + `src/settings/settings.html`

Dialogue ouvert via `Office.context.ui.displayDialogAsync`.

- permet d'activer/désactiver le « message écologique » ajouté en fin d'email après un nettoyage des réponses
- affiche les économies cumulées (images / réponses / pièces jointes / total) via `getSavings()`
- communique vers le parent via `Office.context.ui.messageParent` (JSON)

### `src/taskpane/dev-viewer.js` + `src/taskpane/dev-viewer.html`

Outil de debug exposé en compose. Permet d'afficher le HTML brut du mail, de le copier et de visualiser le résultat de `findReplySeparators`.

### `src/shared/i18n.js`

Gestion simple des traductions :

- détection de langue via `Office.context.displayLanguage` puis `navigator.language`
- mapping vers `en-US`, `fr-FR`, `es-ES`
- fallback systématique sur l'anglais

## Localisation

Langues actuellement prises en charge :

- anglais
- français
- espagnol

Les labels existent à deux niveaux :

- dans `src/i18n/*.json` pour l'UI JavaScript et la taskpane
- dans `manifest.xml` pour les labels et tooltips Office du ruban

Attention : les ressources du manifeste sont dupliquées entre plusieurs blocs `VersionOverrides`.
Quand un label du ruban change, il faut vérifier toutes les occurrences pertinentes dans `manifest.xml`, pas seulement la première.

## Libellés actuels du ruban/taskpane dans le manifeste

Le manifeste a récemment été ajusté pour afficher :

- menu : `Quick Actions` / `Actions directes` / `Acciones directas`
- panneau : `Action Panel` / `Panneau d'actions` / `Panel de acciones`

Ne pas réintroduire `MailLighter` dans ces labels de surface sans vérifier l'intention produit.
Le nom global de l'add-in reste toutefois `MailLighter` dans le `DisplayName` du manifeste.

## Assets et branding

Assets notables présents dans `assets/` :

- `MailLighter_Logo_transp.png` : logo principal (dialogue settings)
- `icon-remove-images-*`, `icon-remove-attachments-*`, `icon-keep-replies-*`, `icon-keep-selection-*`, `icon-clean-all-*`, `icon-settings-*`
- `icon-16.png`, `icon-32.png`, `icon-64.png`, `icon-80.png`, `icon-128.png` pour l'add-in lui-même

## Build et exécution

### Commandes utiles

- `npm run build:dev` : build de développement
- `npm run build` : build de production
- `npm run dev-server` : serveur webpack en local HTTPS sur le port 3000
- `npm run start -- desktop --app outlook` ou la tâche VS Code équivalente : debug Outlook Desktop
- `npm run validate` : validation du manifeste
- `npm run lint` : lint du projet

### Ce qui marche de façon fiable

- `npm run build:dev` compile correctement actuellement
- webpack copie les assets vers `dist/assets/`
- webpack génère `commands.html`, `dev-viewer.html` et `settings.html`

### Points de debug connus

- Le sideload Outlook Desktop peut échouer si la loopback exemption WebView n'est pas présente
- Sur cette machine, l'ajout automatique de loopback a échoué avec `Accès refusé`, ce qui indique qu'une exécution admin peut être nécessaire
- Le sideload a aussi échoué via `@microsoft/teamsapp-cli` avec un `invalid_scope` / HTTP 403
- Le dev server local écoute sur le port `3000`

Conséquence pratique :

- un build réussi ne garantit pas que le sideload Outlook fonctionne immédiatement
- avant de diagnostiquer le code, vérifier d'abord les problèmes d'environnement M365 / loopback

## Dépendances et compatibilité

### Stack principale

- JavaScript ES via Babel
- webpack 5
- Office.js
- HTML simple sans framework UI

### Dépendances métier minimales

- `core-js`
- `regenerator-runtime`

### Particularité actuelle

Le dépôt contient une modification locale de `office-addin-debugging` vers `^5.1.6` dans `package.json` et `package-lock.json`.
Ne pas écraser ce changement sans raison, car il peut être lié aux essais de sideload en cours.

## Conventions de modification

- Ne pas modifier `dist/` manuellement
- Préférer les changements ciblés et minimaux
- Si un changement touche le ruban Outlook, vérifier `manifest.xml` dans tous les blocs concernés
- Si un changement touche les libellés UI, vérifier aussi `src/i18n/en-US.json`, `src/i18n/fr-FR.json`, `src/i18n/es-ES.json`
- Si un changement touche les icônes de fonctionnalités, vérifier à la fois `manifest.xml`, `assets/` et `src/settings/settings.html`
- Pour les actions Outlook, conserver les fallbacks quand une API Office n'est pas disponible dans un contexte donné

## Pièges techniques importants

- `manifest.xml` contient des ressources dupliquées pour différentes versions Office : il est facile d'oublier une occurrence
- Les chemins d'images dans `src/settings/settings.html` sont résolus depuis le fichier source HTML pendant le build webpack
- `prependAsync` n'est pas toujours disponible dans tous les contextes Outlook (fallback silencieux dans `keepSelectionOnlyCore`)
- La récupération de sélection doit tolérer l'absence de HTML et basculer vers du texte simple

## État fonctionnel récent

Changements déjà intégrés dans le dépôt :

- ajout des icônes fonctionnelles dédiées dans `assets/` (remove-images, remove-attachments, keep-replies, keep-selection, clean-all, settings)
- dialogue de paramètres (`src/settings/`) avec message écologique et compteur d'économies (`src/shared/savings-storage.js`)
- harmonisation des labels du menu en anglais, français et espagnol
- refactorisation de `commands.js` autour de `officeAsync()`
- extraction de la détection de séparateurs dans `src/shared/reply-detection.js` (unit-testée)
- factorisation de `escapeHtml`, `formatFileSize` et `MAILLIGHTER_SITE_URL` dans `src/shared/office-helpers.js`

## Taskpane de debug

La taskpane "HTML Viewer (dev)" (`src/taskpane/dev-viewer.*`) est livrée dans le manifest en compose.
Elle affiche le HTML brut du body, le copie et liste les séparateurs détectés par `findReplySeparators`.

## Si tu dois reprendre le projet rapidement

Lis dans cet ordre :

1. `package.json`
2. `manifest.xml`
3. `src/commands/commands.js`
4. `src/settings/settings.js`
5. `src/shared/i18n.js`
6. `src/shared/office-helpers.js`
7. `src/i18n/*.json`

## Consignes de modification

Rechercher dans le code source avant de modifier. Ne jamais changer du code que vous n'avez pas lu.

## Résumé opérationnel

Projet Outlook add-in JavaScript orienté actions rapides sur le contenu d'email.
Le coeur métier est dans `src/commands/commands.js`, le dialogue de paramètres dans `src/settings/`, et `manifest.xml` reste la source critique pour le ruban, les labels et les ressources Office.