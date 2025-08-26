# VTS Tools Browser (Web)

Client-side web UI that mimics the PowerShell TUI. It scrapes any public GitHub repo folder for PowerShell `.ps1` files using the GitHub API in the browser, parses comment-based help/param blocks, and builds the menu. Enter copies the command to the clipboard.

## Run locally

```pwsh
# From repo root
cd web-app
npm install
npm run dev
```

Open the URL shown (default: http://localhost:5173/).

## Configure source via URL params

All params are optional; sensible defaults point at this repo's `functions` folder.

- `repo`  — `owner/repo` (default: `roberto-ryan/Public`)
- `branch` — branch or tag (default: `main`)
- `path` — subfolder to scan (default: `functions`)
- `depth` — how many folder levels to use for categories (1-5, default: 2)
- `ignore` — comma-separated folder names to skip from categories (default includes `functions,scripts,src,docs,test,...`)
- `token` — optional GitHub token to raise rate limits and access private repos (browser-only; avoid using long-lived tokens)

Examples:

- Default repo/folder:
  - http://localhost:5173/
- A different repo and folder on a tag:
  - http://localhost:5173/?repo=someuser/somerepo&branch=v1.2.3&path=scripts
- Custom category depth and ignore list:
  - http://localhost:5173/?repo=org/ops&path=pwsh/tools&depth=1&ignore=tools,legacy

## Notes

- Public repos work anonymously but are rate-limited (~60 requests/hour). Provide a short-lived `token` param to raise limits.
- The token is read only in the browser and never stored. For production, prefer an OAuth flow or a small proxy if needed.
- The parser is pragmatic: it reads `#< ... #>` comment-based help and a simple `param(...)` block to infer parameters.
- Categories are derived from folders under `path`, after skipping any `ignore` names.

## Build

```pwsh
npm run build
npm run preview
```
