# Maintenace-Project

Collection of 6 Google Apps Script web apps, each writing to the same
submissions sheet. Each project lives in its own subfolder with its own
`.clasp.json` (own script ID) — they are deployed independently, only the
code lives together in this one repo.

Layout (subfolders added as each project is brought in):

```
project-name/
  .clasp.json     # scriptId + file extensions (clasp clone creates this)
  appsscript.json
  Code.gs
  index.html
```

## Working on a project

```bash
cd project-name
clasp pull      # get latest from Apps Script editor
clasp push      # push local edits back
clasp deploy    # ship a new version of the web app
```

## Shared changes

If a fix applies to all 6 (e.g. a change to how they write to the shared
sheet), it's logged in [MODIFICATIONS.md](MODIFICATIONS.md) so it doesn't get
missed in the other 5.
