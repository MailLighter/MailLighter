<p align="center">
  <img src="assets/MailLighter_Logo_transp.png" alt="MailLighter" width="200">
</p>

<p align="center">
  An Outlook add-in that provides quick actions to lighten your emails before sending them.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/platform-Outlook_Desktop-0078D4?logo=microsoftoutlook" alt="Outlook Desktop">
  <img src="https://img.shields.io/badge/platform-Outlook_Web-0078D4?logo=microsoftoutlook" alt="Outlook Web">
  <img src="https://img.shields.io/badge/license-MIT-green" alt="MIT License">
  <img src="https://img.shields.io/badge/i18n-EN_|_FR_|_ES-blue" alt="Languages">
</p>

---

## Why MailLighter?

Long email threads accumulate images, attachments, and dozens of quoted replies that bloat your messages. MailLighter sits in the Outlook ribbon and gives you **one-click actions** to clean all of that up — right where you compose your emails.

## Features

All actions are available in **compose mode** from the Outlook ribbon dropdown:

| Action | Description |
|--------|-------------|
| **Remove images** | Strips all inline images from the email body and shows the space saved |
| **Keep 2 replies** | Keeps only the 2 most recent replies in the thread and removes everything below |
| **Remove attachments** | Removes all attached files (excluding inline) and shows the space saved |
| **Full cleanup** | Runs images + attachments + replies cleanup in one click with a detailed summary |
| **Keep selection only** | Replaces the quoted content with only the text you selected |
| **Settings** | Configure the eco message and view your cumulative savings |

### Full cleanup summary

The full cleanup action provides a detailed report showing what was done:

```
✅ Full cleanup completed — Images: 3 (150 KB) | Attachments: 2 (252.4 KB) | Replies: 2 | Total saved: 402.4 KB
```

Each category shows the count of items processed, the space saved when applicable, and for replies the reduction (e.g. 5 → 2). A total space saved is shown at the end when applicable.

## How it works

MailLighter integrates directly into the Outlook ribbon as a dropdown menu.

**When composing an email:**

```
Quick Actions
├── Remove images
├── Keep 2 replies
├── Remove attachments
├── Full cleanup
├── Keep selection only
└── Settings
```

**When reading an email:**

```
Quick Actions
└── Settings
```

### Settings

The Settings dialog gives you two things:

- **Eco message** — optionally append an eco-friendly footer to your outgoing emails. You can toggle it on/off and customize the message text.
- **My savings** — a running total of everything MailLighter has saved across all your emails: bytes removed from images, replies trimmed, and attachments stripped.

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
- **Jest** — unit tests
- **GitHub Actions** — CI (tests, lint, build)

## Installation

### For end users

MailLighter will be available on Microsoft AppSource soon.  
Visit [maillighter.com](https://www.maillighter.com) for updates.

### For developers

1. Clone the repository and install dependencies: `npm install`
2. Build the project: `npm run build`
3. Start the local dev server: `npm run dev-server`
4. Sideload `manifest.xml` in Outlook:
   - **Desktop**: File → Manage Add-ins → Upload a custom add-in → `manifest.xml`
   - **Web**: Settings → Manage add-ins → Upload → `manifest.xml`

> Requires a Microsoft 365 account with add-in sideloading permissions.

## License

[MIT](LICENSE)
