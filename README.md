# Dossche Mills Internal AI Opportunity Explorer

Internal Astro website built from the AI Compass scorecard v2.

## What is inside

- All 50 curated core use cases from the scorecard v2
- All 7 reviewed add-ons from the external draft
- Navigation by group and priority
- A detailed page for every core use case
- A detailed page for every reviewed add-on
- A generated data layer based on the workbook in `../deliverables/Dossche_Mills_AI_Scorecard_v2.xlsx`

## Publish safely

- Do not publish the whole `Dossche Mills Codex` workspace as a public repository.
- Publish only this `internal-site` folder in its own repository if you want a public GitHub Pages URL.
- This folder already includes a GitHub Pages workflow in `.github/workflows/deploy-pages.yml`.

## Run locally

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

The static build output is written to `./dist`.

## GitHub Pages

1. Create a new GitHub repository for this folder only.
2. Push the contents of this folder to the repository root.
3. In GitHub, open `Settings > Pages` and set `Source` to `GitHub Actions`.
4. Push to `main` and the workflow will publish the site automatically.

The Astro config auto-detects the repository name during the GitHub Pages build, so assets and routes should work correctly on the public Pages URL.

## Data flow

1. The source workbook is `../deliverables/Dossche_Mills_AI_Scorecard_v2.xlsx`.
2. `npm run build:data` runs `./scripts/build_usecase_data.py`.
3. That script generates `./src/data/usecases.generated.json`.
4. Astro pages read the generated JSON and build the website.

## Python dependency note

- `openpyxl` is only required when you want to regenerate the JSON from the workbook.
- If `./src/data/usecases.generated.json` already exists, the site can still run without `openpyxl`.
- If you want to refresh the data and your Python environment does not have it, install:

```bash
python3 -m pip install openpyxl
```

## Main routes

- `/`
- `/groups/thera/`
- `/groups/procurement/`
- `/groups/rd-product/`
- `/priorities/quick-win/`
- `/priorities/strategic-bet/`
- `/add-ons/`
- `/use-cases/<slug>/`
- `/reviewed-add-ons/<slug>/`
