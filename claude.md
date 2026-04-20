# MailLighter - Contexte Projet

## Vue d'ensemble

MailLighter est un add-in Outlook construit sur la base du template Office Add-in JavaScript.
Le projet cible principalement Outlook Desktop et fournit des actions rapides pour nettoyer ou réutiliser le contenu d'un email.

Le projet expose deux surfaces UI principales :

- le menu déroulant dans le ruban Outlook
- une taskpane dédiée aux actions partielles de réponse/transfert

## Fonctionnalités métier actuelles

### En lecture de message

- Réponse partielle : ouvre un brouillon de réponse avec le texte sélectionné
- Répondre à tous partiel : ouvre un brouillon de réponse à tous avec le texte sélectionné
- Transfert partiel : ouvre un brouillon de transfert avec le texte sélectionné

### En composition

- Supprimer les images
- Supprimer les pièces jointes
- Conserver uniquement les 2 dernières réponses
- Nettoyage complet

## Structure du dépôt

### Fichiers racine importants

- `manifest.xml` : manifeste Office de l'add-in, avec les commandes, labels, tooltips et ressources localisées
- `package.json` : scripts npm et dépendances
- `webpack.config.js` : build webpack, copie des assets, génération de `commands.html` et `taskpane.html`
- `generate-icons.ps1` : script PowerShell de génération d'icônes PNG
- `claude.md` : ce fichier de contexte

### Dossiers importants

- `src/commands/` : logique des commandes Outlook exécutées depuis le ruban
- `src/taskpane/` : UI et logique de la taskpane
- `src/i18n/` : traductions JSON (`en-US`, `fr-FR`, `es-ES`)
- `src/shared/` : utilitaires partagés, notamment l'i18n
- `assets/` : logos et icônes PNG utilisés par le manifeste et la taskpane
- `dist/` : sortie générée par webpack, à ne pas modifier manuellement

## Fichiers source clés

### `src/commands/commands.js`

Contient la logique métier principale côté commandes Outlook.

Points importants :

- `notify()` affiche une notification Outlook via `notificationMessages.replaceAsync`
- le code supprime aussi l'ancienne notification `ActionPerformanceNotification` pour éviter les restes de builds précédents
- `officeAsync()` centralise les wrappers Promise autour des API Office async
- les actions partielles utilisent la sélection HTML si disponible, sinon la sélection texte
- `keepTwoRepliesCore()` travaille maintenant sur le HTML et détecte les séparateurs de réponses avec plusieurs stratégies
- les commandes sont enregistrées via `Office.actions.associate(...)`

### `src/taskpane/taskpane.html`

Taskpane minimaliste qui expose actuellement 3 boutons :

- réponse partielle
- répondre à tous partiel
- transfert partiel

État actuel :

- le header affiche `assets/MailLighter_Logo_transp.png`
- les boutons affichent les mêmes icônes que le menu ruban
- le style est en CSS inline dans le fichier HTML

### `src/taskpane/taskpane.js`

Logique de la taskpane.

Points importants :

- même principe de récupération de sélection que dans `commands.js`
- ouvre les formulaires Outlook avec `displayReplyForm`, `displayReplyAllForm`, `displayForwardForm`
- fallback sur `displayNewMessageForm` pour le forward si nécessaire
- localise les labels avec `t()` depuis `src/shared/i18n.js`

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

- `MailLighter_Logo_transp.png` : logo principal utilisé dans la taskpane
- `icon-reply-*`, `icon-reply-all-*`, `icon-forward-*`
- `icon-remove-images-*`, `icon-remove-attachments-*`, `icon-keep-replies-*`, `icon-clean-all-*`
- `icon-16.png`, `icon-32.png`, `icon-64.png`, `icon-80.png`, `icon-128.png` pour l'add-in lui-même

Le script `generate-icons.ps1` permet de régénérer les icônes fonctionnelles en tailles 16, 32 et 80.

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
- webpack génère `commands.html` et `taskpane.html`

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
- Si un changement touche les libellés taskpane, vérifier aussi `src/i18n/en-US.json`, `src/i18n/fr-FR.json`, `src/i18n/es-ES.json`
- Si un changement touche les icônes de fonctionnalités, vérifier à la fois `manifest.xml`, `assets/` et `src/taskpane/taskpane.html`
- Pour les actions Outlook, conserver les fallbacks quand une API Office n'est pas disponible dans un contexte donné

## Pièges techniques importants

- `manifest.xml` contient des ressources dupliquées pour différentes versions Office : il est facile d'oublier une occurrence
- Les chemins d'images dans `src/taskpane/taskpane.html` sont résolus depuis le fichier source HTML pendant le build webpack
- Les APIs `displayReplyForm`, `displayReplyAllForm` et `displayForwardForm` ne sont pas toujours disponibles dans tous les contextes Outlook
- La récupération de sélection doit tolérer l'absence de HTML et basculer vers du texte simple

## État fonctionnel récent

Changements déjà intégrés dans le dépôt :

- ajout d'une vraie taskpane pour les actions partielles
- ajout des icônes fonctionnelles dédiées dans `assets/`
- affichage des icônes correspondantes dans la taskpane
- remplacement du header texte de la taskpane par le logo `MailLighter_Logo_transp.png`
- harmonisation des labels du menu et du panneau en anglais, français et espagnol
- refactorisation de `commands.js` autour de `officeAsync()`
- suppression de vestiges de sample files vides (`parameters.html`, `parameters.js`)

## Si tu dois reprendre le projet rapidement

Lis dans cet ordre :

1. `package.json`
2. `manifest.xml`
3. `src/commands/commands.js`
4. `src/taskpane/taskpane.html`
5. `src/taskpane/taskpane.js`
6. `src/shared/i18n.js`
7. `src/i18n/*.json`

## Consignes de modification

Rechercher dans le code source avant de modifier. Ne jamais changer du code que vous n'avez pas lu.

## Résumé opérationnel

Projet Outlook add-in JavaScript orienté actions rapides sur le contenu d'email.
Le coeur métier est dans `src/commands/commands.js`, la taskpane actuelle couvre les actions partielles, et `manifest.xml` reste la source critique pour le ruban, les labels et les ressources Office.