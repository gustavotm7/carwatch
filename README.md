# CarWatch

Used manual BMW M / E46 / E91 and Nissan Z cars in Canada, checked every morning, shown in a phone-friendly web page.

- **Web app:** `https://gustavotm7.github.io/carwatch/` (add it to your phone's home screen).
- **Update now:** Actions tab > **update** > **Run workflow** (or the menu in the app). Takes about 10 minutes.
- **Daily run:** automatic, early every morning. Nothing runs on your PC.

## Add a car

Edit `cars_config.json` and push (the GitHub web editor works). Copy a block under `models`, change it, set
`"enabled": true`. The S2000 block at the end is a ready example: switch it on and the next run starts tracking
AP1 and AP2 S2000s.

| Field | What it does |
|---|---|
| `make`, `match`, `none` | Words that must / must not be in the ad title or address (regular expressions) |
| `generations` | One per chassis: `years`, the `name` shown in the app (e.g. `AP1 S2000`), `color` |
| `name_by_body` | Optional: name by body, like E90 / E92 / E93 M3 |
| `smg_risk` | `true` if an automated gearbox existed, so the ad text must also say manual |
| `bodies` | Bodies the car came in (first one is the default) |
| `trims` | `[regex, label]` pairs, first match wins |
| `comps` | Optional sold prices from Bring a Trailer: `bat` = results page id(s) of the model, `match` = regex the auction title must contain, `body` = body when the title doesn't say |
| `searches` | Search pages read every day, per site |

## Files

- `collector/car_collector.py` – checks ads, finds new ones, scores them (also runs on Windows; see the Excel version).
- `data/` – daily snapshots (last 14 kept). Each listing carries its price history (`price_hist`) and the date the
  ad went up on the site (`listed`). `comps_bat.csv` holds Bring a Trailer sales (last 3 years, in CAD at the
  Bank of Canada rate of the sale day, `fx_usdcad.csv`).
- `site/` – the web app; `site/data/cars.json` is rebuilt every run.

Scores, fair prices and flags are guides, not appraisals: flags only catch what the seller wrote.
Sold prices are US auction results (hammer price, before the 5% buyer's fee and import costs); they are shown
next to each car but do not change the score.
CarPages blocks GitHub's servers, so the cloud run reads Autotrader and Kijiji only.
