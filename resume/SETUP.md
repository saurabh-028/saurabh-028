# Resume sync

Edit `resume/resume.json`, push to `main`. The GitHub Action (`.github/workflows/sync-resume.yml`) rebuilds:

| Output | Where |
|---|---|
| Resume (Word + PDF) | `dist/Saurabh_Shinde_AI_Engineer_Resume.docx` / `.pdf` |
| GitHub profile README | `README.md` (generated, do not edit by hand) |
| LinkedIn copy-paste sheet | `dist/linkedin.md` |
| Portfolio data | `dist/portfolio.json`, optionally pushed to your portfolio repo |

Run locally: `npm install && npm run build`.

## Portfolio sync (optional)
1. Create a fine-grained token with Contents: write on your portfolio repo.
2. In this repo: Settings > Secrets and variables > Actions.
   - Secret `PORTFOLIO_PAT` = the token
   - Variable `PORTFOLIO_REPO` = `owner/portfolio-repo`
   - Variable `PORTFOLIO_DATA_PATH` (optional) = where the JSON goes, default `data/resume.json`
3. Make your site read that JSON. Netlify redeploys on every push.

## LinkedIn
LinkedIn has no API for editing your own profile, so `dist/linkedin.md` is the paste-ready sheet.
