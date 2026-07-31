KEYZAR FENIX BOT — GITHUB AUTO-UPDATER

Repository:
https://github.com/stopthinkin99/keyzar-fenix-bot-updates

REPOSITORY STRUCTURE

keyzar-fenix-bot-updates/
└── bot/
    ├── app_runtime.py
    ├── config.py
    ├── processing/
    ├── fenix/
    ├── email_reader/
    ├── excel_reports/
    └── all other updateable Python files

DO NOT PUSH

- .env
- data/
- playwright_profile/
- fenix_storage_state.json
- keyzar_jobs.db
- Excel invoices
- logs
- screenshots
- passwords or tokens

INSTALLED APPLICATION STRUCTURE

KeyzarFenixBot/
├── KeyzarFenixBot.exe
├── updater.py
├── bot/
├── data/
├── playwright_profile/
└── .env

HOW IT WORKS

1. The manager opens KeyzarFenixBot.exe.
2. launcher.py runs first.
3. updater.py downloads everything under GitHub's bot/ folder.
4. Existing files under the installed bot/ folder are replaced.
5. If GitHub is unavailable, the installed copy still starts.
6. launcher.py runs bot/app_runtime.py.

IMPORTANT

The launcher EXE and updater.py do not update themselves.
Normal application changes belong inside the GitHub bot/ folder.

When launcher/updater behavior itself changes, build and install a new
installer. Normal GUI, Outlook, Fenix, Excel, and workflow changes do not
require reinstalling.

INITIAL GITHUB COMMANDS

Create a new public GitHub repository named:
keyzar-fenix-bot-updates

From a clean local repository folder:

git init
git branch -M main
git remote add origin https://github.com/stopthinkin99/keyzar-fenix-bot-updates.git
git add bot
git commit -m "Initial Keyzar Fenix Bot update files"
git push -u origin main

FUTURE UPDATE COMMANDS

Copy changed files into the repository's bot/ folder, then:

git add bot
git commit -m "Update Keyzar Fenix Bot"
git push origin main

The manager receives the update the next time the app opens.
