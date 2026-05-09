# Architecture — MailLighter

> Cible : Phase 1a (refactor en couches + cycle send-time du compteur Community).
> Phase 1b (Enterprise) et Phase 2 (backend) sont décrites dans `REFACTOR_PLAN.md`.

## Vue d'ensemble en couches

```
┌─────────────────────────────────────────────────────────────────┐
│  src/commands/commands.js                                        │
│    Routeur Office.actions + exposition globale du sendHandler   │
└──────────────┬──────────────────────────────────┬───────────────┘
               │                                  │
               ▼                                  ▼
┌──────────────────────────┐        ┌─────────────────────────────┐
│  src/core/                │        │  src/ui/                    │
│   - cleaners/*            │        │   - i18n.js                 │
│   - savings/              │        │   - (settings/, taskpane/   │
│       savingsCalculator   │        │      gardés à la racine     │
│       pendingSavings      │        │      pour minimiser le diff)│
│   - lifecycle/sendHandler │        └─────────────────────────────┘
│   - types/CleanupEvent    │
│   - replyDetection        │
│   - htmlSanitizer         │
│   - ecoMessage            │
│  Logique pure — aucune    │
│  dépendance Office.*      │
└─────┬────────────────┬────┘
      │                │
      ▼                ▼
┌──────────────┐  ┌─────────────────────────────────────────────┐
│  src/utils/  │  │  src/platforms/                              │
│   - format   │  │   - PlatformAdapter (interface)              │
│   - logger   │  │   - outlook/outlookAdapter (Office.* ici)    │
└──────────────┘  └─────────────────────────────────────────────┘
                  ┌─────────────────────────────────────────────┐
                  │  src/config/                                 │
                  │   - constants                                │
                  │   - featureFlags (vide en 1a, peuplé en 1b) │
                  └─────────────────────────────────────────────┘
```

## Règles d'architecture (enforcées par ESLint)

1. **`Office.*` ne peut être appelé qu'à l'intérieur de `src/platforms/outlook/`.**
   Bloqué par `no-restricted-globals` sur `src/core/`, `src/utils/`, `src/config/`.

2. **Aucun cleaner ne peut importer `savingsCalculator`.**
   Bloqué par `no-restricted-imports` dans `.eslintrc.json`. Les cleaners n'ont
   que l'API `addPendingEvent` ; le compteur Settings n'est mis à jour que par
   `sendHandler` au moment du send effectif.

3. **`core/` ne dépend jamais de `platforms/`.** Les cleaners reçoivent un
   `PlatformAdapter` en argument, jamais un import direct.

4. **Aucun `getUserIdHash` n'existera jamais dans le `PlatformAdapter`.**
   Invariant RGPD non négociable. Seul `getTenantIdHash` sera ajouté en 1b.

## Flow d'une commande utilisateur

Exemple : l'utilisateur clique « Supprimer les images » dans le ruban.

```
Office.actions.associate("removeImagesCommand", removeImagesCommand)
              │
              ▼
commands.js : removeImagesCommand(event)
   → executeWithNotification(event, async () => {
        result = await removeImages(platform);    ← cleaner
        return formatNotification(result);
     })
              │
              ▼
core/cleaners/imageCleaner.js : removeImages(platform)
   1. html = await platform.getBodyHtml()
   2. { cleaned, bytesRemoved, imagesRemoved } = stripInlineImages(html)
   3. await platform.setBodyHtml(cleaned)
   4. composeId = await platform.getComposeId()
   5. recipients = await platform.getRecipients()
   6. addPendingEvent(composeId, createCleanupEvent({...}))
   7. return CleanupResult
              │
              ▼
        platform.notify("3 images supprimées — 215 KB")
   ⚠️ La notification mentionne 215 KB, mais le COMPTEUR Settings n'a pas
      encore bougé. Il ne bougera qu'au send effectif.
```

## Flow send-time (cœur du refactor 1a)

```
                 Compose ouvert
                       │
                       ▼
       ┌──────────────────────────────┐
       │ user clicks "Remove images"  │
       │ user clicks "Keep replies"   │
       │ user adds 2 recipients       │
       └─────────────┬────────────────┘
                     │
                     ▼
              pendingSavings (Map)
              { composeId →
                  [imageEvent(1024B, 3), replyEvent(2048B, 3)] }
              ─────────────────────────────────
                     │
                     ▼
           user clicks "Send"
                     │
                     ▼
       Office.js fires OnMessageSend
                     │
                     ▼
   commands.js : globalThis.onMessageSendGlobalHandler(eventArgs)
                     │
                     ▼
   core/lifecycle/sendHandler : onMessageSend(eventArgs, platform, storage)
       1. composeId = await platform.getComposeId()
       2. events = consumePendingEvents(composeId)     ← clears the queue
       3. recipients = await platform.getRecipients()
          // user added 2 recipients since cleanup → now 5 (was 3)
       4. for each event:
              recordConfirmedSavings(storage, {...event, recipientCount: 5})
                  ↓
              localStorage.setItem(USER_SAVINGS_IMAGES,        +1024)
              localStorage.setItem(USER_SAVINGS_TRANSMISSION,  +1024×5)
       5. eventArgs.completed({ allowEvent: true })   ← in finally{} :
                                                       NEVER blocks the send
                     │
                     ▼
           Email is sent
                     │
                     ▼
        Settings reads getSavings(storage)
        → displays raw + transmission totals
```

### Invariants critiques

- **Si l'utilisateur abandonne** (close window without send) : aucun appel à
  `recordConfirmedSavings` ne se produit. La file `pendingSavings` expire via
  `purgeStale(24h)`.
- **Si l'utilisateur ajoute/retire des destinataires entre cleanup et send** :
  le `recipientCount` est recalculé à 100 % au send. La valeur capturée au
  cleanup est ignorée.
- **Si le `sendHandler` échoue** : `eventArgs.completed({ allowEvent: true })`
  est appelé dans le `finally`, garantissant que l'envoi de l'email ne soit
  jamais bloqué par MailLighter.

## Manifest

Le `manifest.xml` enregistre le handler `OnMessageSend` dans le bloc
`VersionOverridesV1_1` :

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnMessageSend"
                 FunctionName="onMessageSendGlobalHandler"
                 SendMode="SoftBlock"/>
  </LaunchEvents>
  <SourceLocation resid="Commands.Url"/>
</ExtensionPoint>
```

- `MinVersion="1.10"` (requis par `OnMessageSend`).
- Les clients < 1.10 utilisent le fallback V1.0 sans LaunchEvent → comportement
  dégradé : aucun comptage. C'est cohérent avec la philosophie "comptage juste
  ou rien".
- Aucun `<AppDomain>` ajouté (pas de permissions réseau en 1a).

## Compose ID

`platform.getComposeId()` génère un UUID stable pendant toute la durée d'une
fenêtre de composition. L'implémentation Outlook persiste l'UUID via
`Office.context.mailbox.item.sessionData` (clé `ml.composeId`). Si `sessionData`
n'est pas disponible (clients legacy), un fallback in-memory est utilisé. Cet
edge case est acceptable en 1a (multi-compose simultané sans sessionData
pourrait collisionner — non bloquant pour le MVP).

## Préparation Phase 1b et au-delà

Tout est en place pour que la Phase 1b se contente d'**ajouter** des modules
et **un seul appel** dans `sendHandler.js` :

- `src/integrations/license/licenseManager.js` (stub)
- `src/integrations/telemetry/eventCollector.js` etc.
- `getTenantIdHash()` ajouté à `PlatformAdapter` et `OutlookAdapter`
- Dans `sendHandler.onMessageSend`, **après** `recordConfirmedSavings`,
  ajouter `if (FEATURE_FLAGS.ENTERPRISE_MODE) await collectEvent(updated)`

Aucune réécriture de cleaner ni de routeur ne sera nécessaire en 1b.

Pour Gmail (Phase ultérieure) : créer `src/platforms/gmail/gmailAdapter.js` qui
implémente la même interface `PlatformAdapter`. Note : Gmail tourne dans un
runtime différent (Apps Script ou Gmail Add-on) — il s'agira probablement d'un
**bundle distinct** partageant uniquement les fichiers `core/` purs.
