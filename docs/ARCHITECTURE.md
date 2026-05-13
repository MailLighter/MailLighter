# Architecture — MailLighter

Add-in Outlook (Office.js) qui propose des actions de nettoyage rapides
sur l'email en composition et tient un compteur d'économies de
transmission dans Settings.

## Vue d'ensemble en couches

```
┌─────────────────────────────────────────────────────────────────┐
│  src/commands/commands.js                                        │
│    Routeur Office.actions (6 commandes)                          │
└──────────────┬──────────────────────────────────┬───────────────┘
               │                                  │
               ▼                                  ▼
┌──────────────────────────┐        ┌─────────────────────────────┐
│  src/core/                │        │  src/ui/                    │
│   - cleaners/             │        │   - i18n.js                 │
│       imageCleaner        │        │   - (settings/ dialog       │
│       attachmentCleaner   │        │      à la racine src/)      │
│       replyCleaner        │        └─────────────────────────────┘
│       selectionCleaner    │
│       fullCleaner         │
│   - savings/              │
│       savingsCalculator   │
│   - types/                │
│       CleanupEvent        │
│       CleanupResult       │
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
                  └─────────────────────────────────────────────┘
```

## Règles d'architecture

Conventions à respecter manuellement (pas d'enforcement ESLint
automatique) :

1. **`Office.*` ne s'appelle qu'à l'intérieur de `src/platforms/outlook/`.**
   Les cleaners passent par le `PlatformAdapter` qui leur est injecté.
2. **`core/` ne dépend jamais de `platforms/`.** Les cleaners reçoivent
   un `PlatformAdapter` en argument, jamais un import direct.
3. **Aucun `getUserIdHash`** n'existera dans le `PlatformAdapter`.
   Invariant RGPD non négociable.

## Flow d'une commande utilisateur

Exemple : l'utilisateur clique « Supprimer les images » dans le ruban.

```
Office.actions.associate("removeImagesCommand", removeImagesCommand)
              │
              ▼
commands.js : removeImagesCommand(event)
   → executeWithNotification(event, async () => {
        result = await removeImages(platform, storage);
        return formatNotification(result);
     })
              │
              ▼
core/cleaners/imageCleaner.js : removeImages(platform, storage)
   1. html = await platform.getBodyHtml()
   2. { cleaned, bytesRemoved, imagesRemoved } = stripInlineImages(html)
   3. await platform.setBodyHtml(cleaned)
   4. recipients = await platform.getRecipients()
   5. recordConfirmedSavings(storage, createCleanupEvent({
        elementType: "image",
        bytesRemoved,
        recipientCount: recipients.length,
      }))
      // localStorage.setItem(USER_SAVINGS_TRANSMISSION_IMAGES,
      //   +bytesRemoved × recipientCount)
   6. return CleanupResult
              │
              ▼
   platform.notify("3 images supprimées — 215 KB")
```

L'économie est enregistrée **immédiatement** au moment du cleanup. Le
`recipientCount` utilisé est celui présent dans le brouillon au moment
du clic ; ajouter/retirer des destinataires après le cleanup ne modifie
pas le compteur. Si l'utilisateur ferme la fenêtre sans envoyer, le
compteur reste tout de même crédité — limitation acceptée.

## Settings — compteur de transmission

`getSavings(storage)` retourne :

```js
{
  transmission: {
    images: <number>,        // bytes × recipients cumulés
    replies: <number>,
    attachments: <number>,
    total: <number>,
  }
}
```

`selection` est rangé sous `replies` (cohérent avec l'affichage UI
"Text & replies"). Le dialogue Settings reçoit ces valeurs en URL params
depuis `commands.js : openSettingsCommand`.

## Manifest

Deux blocs `VersionOverrides` :

- **V1_0** : fallback pour clients Mailbox < 1.10. Aucun runtime LaunchEvent.
- **V1_1** : runtime moderne avec WebViewRuntime, Mailbox ≥ 1.10.

Les deux blocs déclarent le même menu (compose + read) avec 6 items :
Remove images, Keep two replies, Remove attachments, Full cleanup,
Keep selection, Settings. Toute modification de libellé du ruban doit
être appliquée dans les deux blocs.

Aucun `<AppDomain>` réseau ajouté — pas de communication externe.

## Plateformes futures (Gmail)

Pour ajouter Gmail : créer `src/platforms/gmail/gmailAdapter.js` qui
implémente l'interface `PlatformAdapter`. Gmail tourne dans un runtime
différent (Apps Script ou Gmail Add-on) ; il s'agira probablement d'un
bundle distinct partageant uniquement les fichiers `core/` purs.
