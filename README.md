<p align="center">
  <img src="assets/MailLighter_Logo_transp.png" alt="MailLighter" width="200">
</p>

<h1 align="center">MailLighter</h1>

<p align="center">
  An Outlook add-in that provides quick actions to lighten your emails before sending them.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/platform-Outlook_Desktop-0078D4?logo=microsoftoutlook" alt="Outlook Desktop">
  <img src="https://img.shields.io/badge/license-MIT-green" alt="MIT License">
  <img src="https://img.shields.io/badge/i18n-EN_|_FR_|_ES-blue" alt="Languages">
</p>

---

## Why MailLighter?

Long email threads accumulate images, attachments, and dozens of quoted replies that bloat your messages. MailLighter sits in the Outlook ribbon and gives you **one-click actions** to clean all of that up — right where you compose your emails.

## Features

### While composing (compose mode)

| Action | Description |
|--------|-------------|
| **Remove images** | Strips all inline images from the email body |
| **Remove attachments** | Removes all attached files |
| **Keep 2 replies** | Keeps only the 2 most recent replies in the thread and removes everything below |
| **Full cleanup** | Runs all of the above in a single click |

### While reading (read mode)

| Action | Description |
|--------|-------------|
| **Partial reply** | Opens a reply draft containing only the text you selected |
| **Partial reply all** | Opens a reply-all draft with your selection |
| **Partial forward** | Opens a forward draft with your selection |

These read-mode actions are available both from the **ribbon menu** and from a dedicated **side panel** (taskpane).

## How it works

MailLighter integrates directly into the Outlook ribbon as a dropdown menu:

```
Quick Actions
├── Remove images
├── Remove attachments
├── Keep 2 replies
├── Full cleanup
├── ─────────────
├── Partial reply
├── Partial reply all
└── Partial forward
```

The partial reply/forward actions use your **current text selection** in the email body. Select the relevant paragraph, click the action, and a new draft opens with only that content pre-filled.

## Languages

MailLighter adapts to your Outlook display language:

- **English** (en-US) — default
- **French** (fr-FR)
- **Spanish** (es-ES)

Both the ribbon labels and in-app notifications are localized.

## Tech stack

- **JavaScript** (ES6+ via Babel)
- **Office.js** — Outlook add-in APIs
- **webpack 5** — build & bundling
- **HTML/CSS** — no UI framework

## Installation

> *Coming soon*

## License

[MIT](LICENSE)
