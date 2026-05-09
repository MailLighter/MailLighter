# Cycle de vie d'une économie — MailLighter

> Phase 1a. Le compteur Settings n'est mis à jour qu'au moment où l'email
> quitte effectivement le poste de l'utilisateur.

## Pourquoi ce cycle ?

Avant la Phase 1a, le compteur Settings montait dès qu'un cleanup était
effectué. Problème : si l'utilisateur nettoyait son email puis l'abandonnait
(fermait la fenêtre, restaurait une version antérieure, changeait d'avis),
le compteur affichait des économies **fantômes** — aucun octet n'avait
réellement été économisé en transmission ni en stockage.

La Phase 1a corrige ce bug méthodologique en deux temps :

```
ÉTAPE 1 — Cleanup
    Le cleaner enregistre l'économie LOCALEMENT comme "candidate" (pending)
    Aucune mise à jour du compteur

ÉTAPE 2 — Envoi effectif
    Le sendHandler confirme la candidate → le compteur monte
    Recalcul du nombre actuel de destinataires
    Multiplicateur appliqué : bytesRemoved × recipientCount

ÉTAPE 2 (alternative) — Abandon
    L'entrée pending expire silencieusement (TTL 24 h)
    Le compteur ne bouge jamais
```

## Diagramme détaillé

```
┌───────────────────────────────────────────────────────────────────┐
│                      USER OPENS A COMPOSE                          │
│                                                                    │
│  Office.context.mailbox.item.sessionData                           │
│    ml.composeId = UUID                                             │
└────────────────────┬──────────────────────────────────────────────┘
                     │
                     ▼
┌───────────────────────────────────────────────────────────────────┐
│             USER CLICKS "Remove images" (or any cleaner)           │
│                                                                    │
│  imageCleaner.removeImages(platform)                               │
│    ├─ stripInlineImages(html) → bytesRemoved=1024, count=3         │
│    ├─ platform.setBodyHtml(cleaned)                                │
│    ├─ composeId = await platform.getComposeId()                    │
│    ├─ recipients = await platform.getRecipients()  // 3            │
│    └─ addPendingEvent(composeId, {                                 │
│          elementType: "image",                                     │
│          bytesRemoved: 1024,                                       │
│          recipientCount: 3,    // INDICATIVE only                  │
│          timestamp, platform                                       │
│        })                                                          │
│                                                                    │
│  notify("3 images supprimées — 1 KB")  ← informational only        │
│  ⚠ Settings counter has NOT moved                                  │
└────────────────────┬──────────────────────────────────────────────┘
                     │
                     ▼
   pendingSavings Map:
     { "<UUID>" → { events: [imageEvent], createdAt: 1714... } }


─── Optionally, more cleanups happen, all queued under the same composeId ──


─── Optionally, the user adds 2 more recipients (now 5) ────────────────────


                     │
                     ▼ user clicks SEND
┌───────────────────────────────────────────────────────────────────┐
│                  Office.js fires OnMessageSend                     │
│                                                                    │
│  globalThis.onMessageSendGlobalHandler(eventArgs)                  │
│    → onMessageSend(eventArgs, platform, storage)                   │
│         try {                                                      │
│           composeId = await platform.getComposeId()                │
│           events    = consumePendingEvents(composeId)              │
│           //         → returns [imageEvent], clears the queue      │
│                                                                    │
│           current = await platform.getRecipients() // NOW 5        │
│                                                                    │
│           for each event:                                          │
│             recordConfirmedSavings(storage, {                      │
│               ...event,                                            │
│               recipientCount: 5                                    │
│             })                                                     │
│             // localStorage.setItem(USER_SAVINGS_IMAGES,    +1024) │
│             // localStorage.setItem(TRANSMISSION_IMAGES,    +5120) │
│         } catch (e) { logger.warn(...) }                           │
│         finally {                                                  │
│           // CRITICAL: never block the send                        │
│           eventArgs.completed({ allowEvent: true })                │
│         }                                                          │
└────────────────────┬──────────────────────────────────────────────┘
                     │
                     ▼
                 Email sent
                     │
                     ▼
        Settings reads getSavings(storage)
        Returns:
          {
            raw:          { images: 1024,  ..., total: 1024  },
            transmission: { images: 5120,  ..., total: 5120  }
          }
        Displays Option B layout:
          "Cleaned data            : 1 KB"
          "Transmission savings    : 5 KB"
```

## Cas d'abandon

```
USER OPENS COMPOSE → CLEANUP → CLOSES WINDOW WITHOUT SENDING

  pendingSavings Map:
    { "<UUID>" → { events: [...], createdAt: 1714... } }
                  │
                  │  No send → no consumePendingEvents call
                  │  Entry stays in memory
                  ▼
  After 24 h   purgeStale(24h)  removes the stale entry
                  ▼
              Map cleared
              Counter never moved.   ✓
```

## Recalcul des destinataires

C'est la subtilité méthodologique principale.

| Moment      | recipientCount capturé | Utilisation                         |
|-------------|------------------------|-------------------------------------|
| Cleanup     | 3                      | Stocké dans le `CleanupEvent` pending. **Indicatif uniquement** ; jamais utilisé tel quel pour calculer les économies. |
| Send        | 5                      | Lu par `sendHandler` via `platform.getRecipients()`. Réécrit dans le `CleanupEvent` avant `recordConfirmedSavings`. **C'est cette valeur qui multiplie `bytesRemoved` dans le total transmission.** |

Si l'utilisateur ajoute 2 destinataires entre le cleanup et le send,
l'économie de transmission reflète bien `bytesRemoved × 5`, pas
`bytesRemoved × 3`.

Si l'utilisateur retire des destinataires, idem : on prend la valeur au send.

Si l'envoi échoue côté serveur Outlook (pas notre cas, on ne le sait pas) :
les économies sont déjà enregistrées localement. Petit défaut acceptable.

## Tests qui verrouillent ces invariants

- `tests/sendHandler.test.js`
  - `confirms pending events into the Settings counter when send completes`
  - `uses the LATEST recipient count at send time, not the one captured at cleanup`
  - `aggregates multiple pending events into one send confirmation`
  - `does nothing when no events are pending for the composeId`
  - `a different composeId does not consume another's pending events`
  - `calls eventArgs.completed({ allowEvent: true }) on success`
  - `calls eventArgs.completed({ allowEvent: true }) even when getRecipients throws`

- `tests/savingsCalculator.test.js`
  - vérifie `bytesRemoved × recipientCount` pour chaque catégorie

- `tests/pendingSavings.test.js`
  - vérifie le cycle add → consume → discard → purgeStale

## Validation manuelle (à effectuer dans Outlook après sideload)

Trois scénarios à valider explicitement :

### Scénario 1 — Cleanup + Send

1. Ouvrir un nouveau brouillon avec 3 destinataires en `To`.
2. Cliquer « Remove images » (ou tout autre cleaner).
3. Cliquer « Send ».
4. Ouvrir Settings. Vérifier que `Cleaned data` a augmenté de la taille
   nettoyée, et que `Transmission savings` a augmenté de `taille × 3`.

### Scénario 2 — Cleanup + Close sans send

1. Ouvrir un nouveau brouillon.
2. Cliquer « Remove images ».
3. Fermer la fenêtre **sans** envoyer (clic sur la croix → Discard).
4. Ouvrir Settings. Vérifier que **rien n'a changé** dans les compteurs.

### Scénario 3 — Cleanup + ajout destinataires + Send

1. Ouvrir un nouveau brouillon avec 3 destinataires en `To`.
2. Cliquer « Remove images » (capture indicative `recipientCount=3`).
3. Ajouter 2 adresses en `Cc`.
4. Cliquer « Send ».
5. Ouvrir Settings. Vérifier que `Transmission savings` a augmenté de
   `taille × 5` (et non `× 3`).
