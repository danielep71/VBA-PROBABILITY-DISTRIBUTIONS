# traffic-history

Data-only branch. It is written by the **Daily traffic export** workflow
(`.github/workflows/daily-traffic.yml` on `main`) and is never merged.

GitHub keeps repository traffic for 14 days only; this branch keeps the longer
history.

| File | Content |
|---|---|
| `data/traffic.csv` | One row per day: 14-day views, unique views, clones and unique clones, plus stars, forks, watchers and open issues |
| `data/traffic_daily.csv` | Per-day views, unique views, clones and unique clones |
| `data/referrers.csv` | Top referring sites for each snapshot |
| `data/popular_paths.csv` | Most visited repository paths for each snapshot |
| `data/badges/*.json` | shields.io endpoint files behind the README clones, visitors and visits badges |

All values are aggregates that GitHub reports; nothing here identifies an
individual visitor. The credential and data policy is in `SECURITY.md` on
`main`, under "Repository automation credentials".
