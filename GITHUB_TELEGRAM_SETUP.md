# GitHub Pages + Telegram Mini App Setup

This project publishes `webapp/index.html` as a static Telegram Mini App and uses `Code.gs` as the Google Apps Script backend.

## 1. Apps Script Backend

1. Open your Apps Script project.
2. Add the updated `Code.gs`.
3. Open **Project Settings > Script Properties**.
4. Add:

```text
TELEGRAM_BOT_TOKEN=your_bot_token_from_botfather
```

5. Deploy as **Web app**:

```text
Execute as: Me
Who has access: Anyone
```

6. Copy the `/exec` deployment URL and make sure it matches `GAS_API_URL` in `webapp/index.html`.

## 2. Google Sheet Access

In the `Teammates` sheet, each Telegram user must have their Telegram numeric ID in one of these columns:

```text
Telegram Chat ID
Telegram ID
Chat ID
```

Roles are resolved from the teammate row:

```text
Super Admin
Admin / Lead
Manager / APM
User / CSE
```

## 3. GitHub Pages

Create a new GitHub repository, then run these commands from this folder:

```powershell
git init
git branch -M main
git add webapp .github/workflows/deploy-pages.yml .gitignore GITHUB_TELEGRAM_SETUP.md
git commit -m "Set up Telegram Mini App GitHub Pages deployment"
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

If the repository is private and you want backend version history there too, also add `Code.gs`. For a public repo, keep `Code.gs` out unless you are comfortable publishing the backend source.

In GitHub:

1. Open **Settings > Pages**.
2. Set **Source** to **GitHub Actions**.
3. Open the **Actions** tab and wait for **Deploy GitHub Pages** to finish.

Your site will be available at:

```text
https://YOUR_USERNAME.github.io/YOUR_REPO/
```

## 4. BotFather

In Telegram BotFather:

1. Open your bot.
2. Configure the Mini App / Web App URL.
3. Use the GitHub Pages URL from the previous step.

After opening the Mini App from Telegram, the frontend sends signed Telegram `initData` to Apps Script. Apps Script verifies it with `TELEGRAM_BOT_TOKEN`, then matches the verified Telegram ID against the `Teammates` sheet.
