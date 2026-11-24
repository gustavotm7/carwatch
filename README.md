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

## Brazil

Tap the flag next to the name to switch between Canada and Brazil (the app remembers the choice and keeps
filters per country). The Brazil store is in Brazilian Portuguese, numbers and dates included; the Canada store is in
English. The data stays in English inside (flag names, gearbox codes), and the app translates it on screen
(`L(en, pt)` and `FLAG_PT` in `site/index.html`). In Brazil the serious flag is "Leilão ou sinistro". The Brazil list comes from `cars_config_br.json`, same format as `cars_config.json`:

- Read every day: **NaPista** and **Mobiauto** (data in `data_br/`, app data in `site/data/cars_br.json`).
- **OLX and Webmotors** block automated reading, so each car has one-tap search links (`links` in the config),
  shown at the top of the Brazil list and in the menu.
- Prices in R$. The score works the same way, against the other Brazilian cars of the same model.
- Manual only, except the FB Civic and the Yaris sedan (`allow_auto` in the config): their automatics are
  shown with the Manual | Automatic toggle at the start of the car chips. The app always opens on Manual.
- The FA5 Civic Si is its own car (`split` in the config): own chip and its own fair price.
- On NaPista and Mobiauto, ads seen in the day's searches are confirmed from the search data without opening
  them; automatics are added from the search data and their pages read later for the description.
- Brazilian years are model years (a "2010/2011" car is a 2011).

## Files

- `collector/car_collector.py` – checks ads, finds new ones, scores them (also runs on Windows; see the Excel version).
- `data/` – daily snapshots (last 14 kept). Each listing carries its price history (`price_hist`) and the date the
  ad went up on the site (`listed`). `comps_bat.csv` holds Bring a Trailer sales (last 3 years, in CAD at the
  Bank of Canada rate of the sale day, `fx_usdcad.csv`).
- `site/` – the web app; `site/data/cars.json` is rebuilt every run. `site/data/guides.json` holds the
  "what to check" lists (edit it by hand).
- `data_br/fipe_cache.json` – FIPE reference prices (Brazil), kept between runs.

Scores, fair prices and flags are guides, not appraisals: flags only catch what the seller wrote.
Sold prices are US auction results (hammer price, before the 5% buyer's fee and import costs); they are shown
next to each car but do not change the score.
CarPages blocks GitHub's servers, so the cloud run reads Autotrader and Kijiji only.

### Location filter (Brazil)

Each car's town is matched to IBGE municipality coordinates (`collector/geo_br.csv`, 5,570 towns) when the
JSON is exported. In the app, the row of chips above the list filters by distance from a favourite town
(São Paulo, SP and Teresina, PI come pre-saved). Tap the pin or **Edit** to change the distance
(town only to 400 km) or to add and remove favourites. Favourites are saved on the phone.

### Local prices (Brazil)

Asking prices differ by region (Rio de Janeiro sits about 10% below the national level for the same model, km
and year). The collector first measures each state's price level, shrunk toward its macro-region when a state has
few cars, and fits the fair-price model on prices corrected for it, so the national fair price is the typical
state's. It exports a table of towns with each town's gap to the national price (manual and automatic cars).
When a town is picked in the location filter, the app works out the local level from the cars in the radius
(leaning on the state level when there are fewer than 10) and uses the local fair price for the % badge, the score
and sorting. The details show both prices. Settings: `K_AREA` and `AREA_CLIP` in the collector.

### Gearbox-aware flags

Automatics are priced only against automatics and lose no points for the gearbox. Clutch flags ("Clutch done",
"Clutch slipping") only apply to manuals; "Gearbox serviced" and "Gearbox problem" apply to both. Ads read before
these flags existed are checked again from their stored text on every run.

### Prices tab

The second tab under the title analyses the market for the model chosen in the chips (All: a table of every model
and, in Brazil, the price level by state). Per model: headline numbers, price against km with the fair-price line,
what moves the price (km, year, convertible, from the fair-price fit), price by year, time on sale (ads with a site
listing date only), price changes, the weekly market (shown after 3 weeks of data), ads that left, by state (Brazil)
and the best deals. The Manual/Automatic button switches the analysis too. The collector exports what the app can't
work out from the cars on sale: `market` (per model: fit numbers, weekly ads on sale and price level, stats on ads
that left) and `b_conv`. The convertible effect is only measured when a model is listed both ways.

### Verdict and buying guide

Every ad gets a price label against its fair price: Great, Good, Fair, High, Overpriced, and **Be careful** when it is
far below (often a scam, an auction car, a write-off or a hidden fault). The cut-offs follow how spread the
model's prices are (`deal` in the export). The fair price also counts the version (LXR, XLS, Nismo...): each
version's level is measured per model and shrunk toward the average when few cars have it (`K_VER`).

The car details open with a verdict: the label on a scale, what to offer (open with, good deal at, don't pay over,
from the price cuts seen on that model), the reasons (km against the typical km for the year, version, price cuts,
days listed, flags, FIPE) and **What to check before buying**: each model's known weak points plus the general
steps for the country, every item with its source (`site/data/guides.json`). Prices tab, per model: a buying guide
(which year and version, what to pay, how to negotiate, warning signs), a budget finder (the years and versions
that fit an amount, and the best fair-or-better ads up to it; the list link applies it as the price filter),
the years x km grid (tap a cell for its best ads), the version ladder and, in Brazil, a state tile map. The details
also show similar cars for sale and how the fair price is built.

### FIPE (Brazil)

The collector looks up the FIPE price of each version and model year on sale with the free FIPE API
(`fipe.parallelum.com.br`, 500 requests a day; a run uses at most 300 and only asks for what the cache lacks, so
after the first day a run costs one request until the month changes). Each generation's `fipe` rule in
`cars_config_br.json` says how to find it: brand, a regex on the FIPE model name, whether the ad's version must be
in the name, and how the brand marks automatics when it doesn't say "Aut." (`330iA`). Convertibles only take a
convertible's FIPE entry. The app shows the FIPE price in the verdict and as a tick on its scale, and the Prices tab
has **Preço x Tabela FIPE**: how the ads ask against FIPE, where the fair price sits, how FIPE moved over the
months in the cache, and FIPE, fair and asking prices for the most common years and versions. `--no-fipe` skips it.
