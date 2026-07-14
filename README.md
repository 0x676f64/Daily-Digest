# Executive Daily Digest

An automated 7:00 AM email that gives each leader a one-glance recap of the workday they
just finished: meeting invites they never answered, mail that arrived unread, and anything
declined or cancelled. Built entirely on Microsoft 365 — no per-seat Copilot license, no
Power Platform, no AI token bill for the core digest.

It comes in two parts:

- **The engine** — a small script that runs by itself on a schedule and sends the emails. Nobody touches it after setup.
- **The console** — a web app where a non-technical person manages who gets the digest, when it sends, and how it looks, entirely by clicking. No PowerShell, no editing files.

---

## What a recipient sees

Each leader gets *their own* recap of *their own* previous day. The email opens with three
big numbers (need response / unread / cancelled) so it can be triaged in about three seconds,
then lists the details underneath in color-coded sections. The most important section —
meeting invites you never responded to — sits at the top in red.

Recipients never log into anything. They just receive an email.

---

## How it works

### The two layers

```
        ┌─────────────────────────┐        writes        ┌──────────────────────┐
        │   DIGEST CONSOLE (web)  │  ──────────────────▶  │  digest-config.json  │
        │  clicks, no terminal    │                       │  who / when / brand  │
        └─────────────────────────┘                       └──────────┬───────────┘
                 ▲  admin signs in                                    │ reads
                 │  with M365 account                                 ▼
                 │                                        ┌──────────────────────┐
                 │                                        │   ENGINE (scheduled) │
                 │                                        │  runs 7 AM weekdays  │
                 │                                        └──────────┬───────────┘
                 │                                                   │ Microsoft Graph
                 │                                                   ▼
                 │                                        ┌──────────────────────┐
                 └────────────────────────────────────── │  Leadership mailboxes │
                        recipients just receive email     │  (read + send only)   │
                                                          └──────────────────────┘
```

The console never emails anyone directly; it only edits the config. The engine never has a
UI; it only reads the config and sends. That separation is what lets non-technical people
safely manage a system that reads executive mailboxes.

### Authentication (no one logs in)

The engine signs in as *itself* — an app identity registered in Entra, not a person. Every
run it exchanges three stored values (tenant ID, client ID, client secret) for a short-lived
access token and attaches that token to each Microsoft Graph call. These are **application
permissions**, which is what allows an unattended job to read several mailboxes with nobody
present. Because that power is tenant-wide by default, an **Application Access Policy** fences
the app to the leadership group only — so it is physically unable to read anyone else's mail.

### How it targets the right mailbox

The access token identifies the *app*, not a mailbox. Each Graph request names the mailbox
explicitly in its URL (`/users/ceo@ceasusa.com/mailFolders/inbox/messages`). The engine simply
loops over the enabled recipients from the config, substituting each address in turn. It reads
exactly the mailboxes you list and nothing else.

### What "missed" means

- **Needs your response** = a calendar invite where your response is `none`/`notResponded` — the CEO's "a meeting was sent but you never answered."
- **Declined or cancelled** = invites you declined, or events the organizer cancelled.
- Actual physical no-show attendance is *not* included; Graph only exposes that via attendance reports for meetings you organized, which is a much heavier lift.

---

## Who does what

| Role | Does | Touches a terminal? |
|------|------|---------------------|
| IT (you) | One-time setup: app registration, scope the access policy, deploy the engine + console | Once, for a single scoping command |
| Admin (office manager / exec assistant) | Day to day: add/remove recipients, change send time, tweak branding, send test | Never — the console |
| Recipients (leadership) | Read the 7 AM email | Never |

---

## One-time setup (IT)

Mostly portal clicks. Budget ~15 minutes.

1. **Register the app.** Entra admin center → App registrations → New registration. Name it "Executive Daily Digest," single tenant. Copy the **Application (client) ID** and **Directory (tenant) ID**.
2. **Grant permissions.** API permissions → Add → Microsoft Graph → **Application permissions** → add `Mail.Read`, `Calendars.Read`, `Mail.Send` → **Grant admin consent**. Confirm each shows a green check.
3. **Create a secret.** Certificates & secrets → New client secret → copy the value immediately (shown once). A certificate is preferable long-term.
4. **Scope the mailboxes (the one command).** In Exchange Online PowerShell, create a mail-enabled security group with the leadership members plus the sender mailbox, then:
   ```powershell
   New-ApplicationAccessPolicy -AppId <clientId> `
     -PolicyScopeGroupId leadership-digest@ceasusa.com -AccessRight RestrictAccess
   ```
   This is the only command-line step, done once. After it, the app cannot read anyone outside that group.
5. **Deploy the engine.** Put `Send-ExecDailyDigest.ps1` in an Azure Automation account as a runbook, add the secrets as **encrypted Automation variables** (`DIGEST_TENANT_ID`, `DIGEST_CLIENT_ID`, `DIGEST_CLIENT_SECRET`), and attach a daily 7 AM schedule. (A Managed Identity can replace the secret entirely once it's running in Azure.)
6. **Host the console.** Serve `digest-console.html` behind Microsoft Entra sign-in (App Service with Easy Auth, or an Azure Static Web App with authentication) so only your admins can reach it. Point its Save action at wherever the engine reads config.

> Ideal end state: the console is a URL your admin bookmarks. They sign in with their normal
> M365 account and click. No files, no scripts, ever.

---

## Testing on your own account first

Before it goes near leadership:

1. Do steps 1–4 above, but scope the access policy to **just yourself**.
2. Set the three secret values as environment variables in your PowerShell session.
3. Open the console, set the only recipient to your own address, Save / Download config next to the script.
4. Run the engine once against a day you know had activity.
5. Check how it actually lands in your inbox and adjust branding in the console until it feels right.

Only once *you'd* be happy to receive it does anyone else get added.

---

## The console

- **Recipients** — add by name + email, remove, or toggle someone off to pause them without deleting.
- **Schedule** — send time and timezone.
- **What's included** — turn each section on/off; optional one-line AI summary at the top.
- **Subject & branding** — subject line, org name, header and accent colors, optional logo.
- **Live preview** — the real digest, updating as you click, so there's never a "hope it looks right" send.
- **Download / Copy config** — produces the `digest-config.json` the engine reads.

---

## Configuration reference (`digest-config.json`)

| Field | Meaning |
|-------|---------|
| `orgName` | Shown in the header and footer |
| `sender` | Mailbox the digest sends *from*; must be inside the access-policy group |
| `subjectPrefix` | Subject line; the date is appended automatically |
| `sendTimeLocal` / `timeZoneId` | When it sends, and in whose clock |
| `headerColor` / `accentColor` | Branding; accent drives the priority section and the "need response" number |
| `logoUrl` | Optional header logo |
| `useAiNarrative` | If true, one cheap Azure OpenAI call per recipient adds a summary line |
| `active` | Master on/off for the whole send |
| `sections` | Which of the four sections to include |
| `recipients` | List of `{name, email, enabled}` |

Secrets are **never** in this file — they live in environment or Automation variables.

---

## Security notes

- The app is scoped to the leadership group by the Application Access Policy — it cannot read other mailboxes even if the code were changed.
- Secrets stay in encrypted Automation variables or, better, are eliminated with a Managed Identity.
- The console sits behind Entra sign-in; only your named admins can open it.
- Reading executive inboxes is sensitive by nature — keep the app registration, the access policy, and the admin list documented.

## Cost

- Azure Automation / Functions at this volume: effectively free.
- Table/blob storage for config: pennies.
- Optional AI summary line: a fraction of a cent per recipient per day. The core digest uses no AI at all — which is the whole reason this exists instead of the Power Automate + AI Builder flow.

## Limitations & roadmap

- No physical meeting-attendance ("did they no-show") — only invite-response state.
- The one PowerShell scoping command has no GUI equivalent today; everything else is clickable.
- Next steps worth considering: a one-click "Deploy to Azure" button so even the initial deploy is button-driven; a "Send test to me" wired to the live engine; per-recipient section preferences.
