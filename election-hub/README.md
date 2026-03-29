# 2026 Election Research Hub

Interactive reference tool for the 2026 U.S. election cycle covering all 50 states:

- **Filing deadlines & primary dates** for every state
- **Races on ballot** — Governor, Senate, House, statewide offices, legislature
- **Where to get filed candidate data** — the actual databases and search tools
- **County office directories** — links to every county election office for local race data
- **School board data sources** — the hardest tier to find, broken out per state

## Run locally

```bash
npm install
npm run dev
```

Opens at http://localhost:5173

## Deploy to Vercel

### Option A: CLI
```bash
npm i -g vercel
vercel
```

### Option B: GitHub
1. Push this folder to a GitHub repo
2. Go to [vercel.com/new](https://vercel.com/new)
3. Import the repo
4. Vercel auto-detects Vite — click Deploy

That's it. No config needed.

## Data as of March 29, 2026

- 5 states have held primaries (AR, NC, TX, MS, IL)
- 23 states have closed filing
- 22 states still have open filing windows
- 4 states centralize ALL candidate data including local (NC, IL, MN, LA)
