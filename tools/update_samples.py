"""
Refresh the sample datasets of Graphite (line charts), Graphite Combo and Graphite Histogram from official
Singapore open data.

    python tools/update_samples.py

1. Downloads the latest figures from the SingStat Table Builder and data.gov.sg APIs.
2. Rebuilds samples/, combo/samples/ and histogram/samples/<name>.xlsx, but only for datasets whose
   numbers actually changed.
3. Updates the sample lists in index.html, combo/index.html and histogram/index.html (date ranges + a version tag
   per file, so the CDN's 7-day cache can't serve an old copy once the new file is uploaded).

Nothing is uploaded; upload the changed files yourself afterwards.
Needs Python 3.8+ and openpyxl (update_samples.bat installs it if missing).
"""
import hashlib
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import date

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is missing. Install it with:  python -m pip install --user openpyxl")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MONTHS = {m: i for i, m in enumerate("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split(), 1)}
BROWSER_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"

# print "·" and "–" properly in the Windows console
for stream in (sys.stdout, sys.stderr):
    try:
        stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ---------------------------------------------------------------------------------------------
# Downloading
# ---------------------------------------------------------------------------------------------

def get_json(url, tries=4):
    for attempt in range(1, tries + 1):
        try:
            # SingStat's firewall rejects Python's default client identity, so present a browser's
            req = urllib.request.Request(url, headers={"User-Agent": BROWSER_UA, "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=60) as res:
                return json.loads(res.read().decode("utf-8"))
        except Exception as err:  # network blips, 429 rate limits, 5xx
            if attempt == tries:
                raise RuntimeError(f"Couldn't download {url}\n  {err}") from err
            wait = 5 * attempt
            print(f"    retrying in {wait}s ({err})")
            time.sleep(wait)


def num(v):
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None  # "na", "-" etc. -> empty cell


def singstat_rows(table):
    """SingStat Table Builder table -> [(seriesNo, row name, {period or column: value})], in table order.
    The API pages by number of data points, so keep asking until a page comes back empty."""
    rows, index, offset, step = [], {}, 0, 2000
    while True:
        data = get_json(f"https://tablebuilder.singstat.gov.sg/api/table/tabledata/{table}?offset={offset}&limit={step}")
        page = (data.get("Data") or {}).get("row") or []
        if not page:
            break
        for r in page:
            key = (r.get("seriesNo"), r["rowText"])
            if key not in index:
                index[key] = len(rows)
                rows.append((r.get("seriesNo"), r["rowText"], {}))
            rows[index[key]][2].update({c["key"]: num(c["value"]) for c in r["columns"]})
        offset += step
        if len(page) == 0 or data.get("DataCount") == 0:
            break
    if not rows:
        raise RuntimeError(f"SingStat table {table} came back empty.")
    return rows


def singstat(table):
    """SingStat table -> {series name: {period: value}} (top-level series only)."""
    series = {}
    for no, name, values in singstat_rows(table):
        if "." in str(no or ""):
            continue  # sub-rows like "Tariff" under a water price; not used
        series.setdefault(name, {}).update(values)
    return series


def pick(series, name, table):
    if name not in series:
        raise RuntimeError(f"SingStat table {table} no longer has a series called:\n  {name!r}\n"
                           f"Available: {sorted(series)}")
    return series[name]


def datagov(resource_id):
    """data.gov.sg dataset -> list of records (all pages)."""
    records, offset = [], 0
    while True:
        q = urllib.parse.urlencode({"resource_id": resource_id, "limit": 5000, "offset": offset})
        result = get_json(f"https://data.gov.sg/api/action/datastore_search?{q}")["result"]
        records += result["records"]
        offset += 5000
        if offset >= result["total"] or not result["records"]:
            return records


def month_key(k):  # "2015 Jan" -> date(2015, 1, 1)
    y, m = k.split()
    return date(int(y), MONTHS[m], 1)


# ---------------------------------------------------------------------------------------------
# Datasets. Each returns the table for one sample file.
# ---------------------------------------------------------------------------------------------

def utilities():
    elec = pick(singstat("M890991"), "Low Tension Supplies - Domestic", "M890991")
    gas = pick(singstat("M891051"), "Piped Town Gas Tariff - General Tariff", "M891051")
    months = sorted((m for m in set(elec) & set(gas) if elec[m] is not None or gas[m] is not None), key=month_key)
    return dict(
        title="Household electricity & piped gas tariffs, Singapore",
        short="Household electricity & piped gas tariffs",
        freq="Monthly", org="SingStat (SP Group, City Energy)",
        source="Source: SingStat tables M890991 & M891051 (data from SP Group / City Energy). Cents per kWh, before GST.",
        header=["Month", "Electricity (¢/kWh)", "Town gas (¢/kWh)"],
        rows=[(month_key(m), elec[m], gas[m]) for m in months],
        label_fmt="mmm yyyy", value_fmt="0.00",
    )


def water():
    s = singstat("M890981")
    low = pick(s, "Potable Water Price For Domestic Users, 0-40 Cubic Metres", "M890981")
    high = pick(s, "Potable Water Price For Domestic Users, > 40 Cubic Metres", "M890981")
    years = sorted(set(low) & set(high), key=int)
    return dict(
        title="Household water prices, Singapore", short="Household water prices",
        freq="Yearly", org="SingStat (PUB)",
        source="Source: SingStat table M890981 (data from PUB). Total price per cubic metre "
               "(tariff + water conservation tax + waterborne fee), before GST.",
        header=["Year", "First 40 m³ ($/m³)", "Above 40 m³ ($/m³)"],
        rows=[(int(y), low[y], high[y]) for y in years],
        label_fmt="0", value_fmt='"$"0.00',
    )


RETAIL = {}


def retail():
    if not RETAIL:
        RETAIL.update(singstat("M213761"))
    return RETAIL


def retail_table(cols):
    r = retail()
    series = [pick(r, name, "M213761") for name, _ in cols]
    months = sorted(set().union(*series), key=month_key)
    # only months where the first item has a price (others may start later)
    months = [m for m in months if series[0].get(m) is not None]
    return [(month_key(m),) + tuple(s.get(m) for s in series) for m in months]


def fuel():
    cols = [("Petrol, 92 Octane (Per Litre)", "Petrol 92"), ("Petrol, 95 Octane (Per Litre)", "Petrol 95"),
            ("Petrol, 98 Octane (Per Litre)", "Petrol 98"), ("Diesel (Per Litre)", "Diesel")]
    return dict(
        title="Average pump prices, Singapore", short="Average pump prices",
        freq="Monthly", org="SingStat",
        source="Source: SingStat table M213761, Average Retail Prices of Selected Consumer Items. Dollars per litre.",
        header=["Month"] + [label + " ($/L)" for _, label in cols],
        rows=retail_table(cols), label_fmt="mmm yyyy", value_fmt='"$"0.00',
    )


def hawker():
    cols = [("Chicken Rice (Per Plate)", "Chicken rice"), ("Fishball Noodles (Per Bowl)", "Fishball noodles"),
            ("Economical Rice (1 Meat & 2 Vegetables) (Per Plate)", "Economical rice (1 meat, 2 veg)"),
            ("Chicken Nasi Briyani (Per Plate)", "Chicken nasi briyani"), ("Fried Carrot Cake (Per Plate)", "Fried carrot cake"),
            ("Roti Prata (Plain) (Per Piece)", "Roti prata (plain)"), ("Coffee/Tea With Condensed Milk (Per Cup)", "Kopi / teh"),
            ("Ice Kachang (Per Bowl)", "Ice kachang")]
    return dict(
        title="Average hawker food prices, Singapore", short="Average hawker food prices",
        freq="Monthly", org="SingStat",
        source="Source: SingStat table M213761, Average Retail Prices of Selected Consumer Items. "
               "Dollars per plate, bowl, piece or cup.",
        header=["Month"] + [label for _, label in cols],
        rows=retail_table(cols), label_fmt="mmm yyyy", value_fmt='"$"0.00',
    )


def coe():
    cats = ["Category A", "Category B", "Category C", "Category D", "Category E"]
    by_ex = {}
    for r in datagov("d_69b3380ad7e51aff3a7dcc84eba52b8a"):
        by_ex.setdefault((r["month"], int(r["bidding_no"])), {})[r["vehicle_class"]] = num(r["premium"])
    ordinal = {1: "1st", 2: "2nd", 3: "3rd"}
    rows = []
    for (m, n) in sorted(by_ex):
        y, mo = m.split("-")
        label = date(int(y), int(mo), 1).strftime("%b %Y") + f" ({ordinal.get(n, str(n) + 'th')})"
        rows.append((label,) + tuple(by_ex[(m, n)].get(c) for c in cats))
    return dict(
        title="COE premiums by category, Singapore", short="COE premiums by category",
        freq="Twice monthly", org="LTA",
        source="Source: LTA, COE Bidding Results / Prices (data.gov.sg d_69b3380ad7e51aff3a7dcc84eba52b8a). "
               "One row per bidding exercise. Cat A/B: cars, C: goods vehicles & buses, D: motorcycles, E: open.",
        header=["Bidding exercise", "Cat A ($)", "Cat B ($)", "Cat C ($)", "Cat D ($)", "Cat E ($)"],
        rows=rows, label_fmt=None, value_fmt='"$"#,##0', widths={1: 20},
    )


def ridership():
    by_year = {}
    for r in datagov("d_75248cf2fbf340de6a746dc91ec9223c"):
        by_year.setdefault(int(r["year"]), {})[r["mode"]] = num(r["ridership"])
    return dict(
        title="Average daily public transport ridership, Singapore", short="Average daily public transport ridership",
        freq="Yearly", org="LTA",
        source="Source: LTA, Public Transport Utilisation - Average Daily Public Transport Ridership "
               "(data.gov.sg d_75248cf2fbf340de6a746dc91ec9223c). Trips per day.",
        header=["Year", "MRT", "LRT", "Bus"],
        rows=[(y, by_year[y].get("MRT"), by_year[y].get("LRT"), by_year[y].get("Bus")) for y in sorted(by_year)],
        label_fmt="0", value_fmt="#,##0",
    )


def hdb():
    recs = sorted(datagov("d_14f63e595975691e7c24a27ae4c07c79"), key=lambda r: r["quarter"])
    return dict(
        title="HDB resale price index, Singapore", short="HDB resale price index",
        freq="Quarterly", org="HDB",
        source="Source: HDB, HDB Resale Price Index (1Q2009 = 100), quarterly "
               "(data.gov.sg d_14f63e595975691e7c24a27ae4c07c79).",
        header=["Quarter", "Resale price index"],
        rows=[(r["quarter"].replace("-", " "), num(r["index"])) for r in recs],
        label_fmt=None, value_fmt="0.0",
    )


# file name -> builder, in the order they appear in the app
DATASETS = [("utilities", utilities), ("water", water), ("fuel", fuel), ("coe", coe),
            ("ridership", ridership), ("hawker", hawker), ("hdb", hdb)]


# ---------------------------------------------------------------------------------------------
# Combo chart samples (combo/): bars + one line. The LAST column is drawn as the line on the
# right axis, so each table ends with the rate/price. "axes" are the chart's axis titles.
# ---------------------------------------------------------------------------------------------

def year_rows(*series, start=None):
    years = sorted(set.intersection(*(set(k for k, v in s.items() if v is not None) for s in series)), key=int)
    return [(int(y),) + tuple(s[y] for s in series) for y in years if start is None or int(y) >= start]


def month_rows(*series, start=None):
    months = sorted(set.intersection(*(set(k for k, v in s.items() if v is not None) for s in series)), key=month_key)
    return [(month_key(m),) + tuple(s[m] for s in series) for m in months if start is None or month_key(m) >= start]


def combo_hdb():
    # The source sometimes repeats a quarter with revised figures and publishes a new quarter
    # flat type by flat type. Keep the latest record per (quarter, flat type), and only use
    # quarters where every flat type is in.
    latest = {}
    for r in sorted(datagov("d_02aa4bb51bc674f3a2d0b9bb6911d934"), key=lambda r: int(r["_id"])):
        latest[(r["quarter"], r["flat_type"])] = num(r["no_of_resale_applications"])
    flat_types = {ft for (_, ft) in latest}
    by_quarter = {}
    for (q, ft), v in latest.items():
        by_quarter.setdefault(q, {})[ft] = v
    apps, partial = {}, []
    for q, counts in by_quarter.items():
        if set(counts) == flat_types and all(v is not None for v in counts.values()):
            apps[q] = sum(counts.values())
        else:
            partial.append(q)
    if partial:
        print(f"    note: skipping incomplete quarter(s) {', '.join(sorted(partial))} (not all flat types published yet)")
    rpi = {r["quarter"]: num(r["index"]) for r in datagov("d_14f63e595975691e7c24a27ae4c07c79")}
    quarters = sorted(q for q in set(apps) & set(rpi) if rpi[q] is not None)
    return dict(
        title="HDB resale applications & resale price index, Singapore", short="HDB resale volume & prices",
        freq="Quarterly", org="HDB",
        source="Source: HDB via data.gov.sg - Number of Resale Applications Registered by Flat Type "
               "(d_02aa4bb51bc674f3a2d0b9bb6911d934, all flat types added up) and HDB Resale Price Index, "
               "1Q2009 = 100 (d_14f63e595975691e7c24a27ae4c07c79).",
        header=["Quarter", "Resale applications", "Resale price index"],
        rows=[(q.replace("-", " "), apps[q], rpi[q]) for q in quarters],
        label_fmt=None, value_fmts=["#,##0", "0.0"],
        axes=dict(x="Quarter", y="Resale applications registered", y2="Resale price index (1Q2009 = 100)"),
    )


def combo_births():
    s = singstat("M810091")
    births = pick(s, "Total Live-Births", "M810091")
    tfr = pick(s, "Total Fertility Rate (TFR)", "M810091")
    return dict(
        title="Births & total fertility rate, Singapore", short="Births & fertility rate",
        freq="Yearly", org="SingStat",
        source="Source: SingStat table M810091, Births And Fertility Rates. Total fertility rate = average "
               "births per woman over her reproductive years.",
        header=["Year", "Live-births", "Total fertility rate"],
        rows=year_rows(births, tfr), label_fmt="0", value_fmts=["#,##0", "0.00"],
        axes=dict(x="Year", y="Live-births", y2="Births per woman"),
    )


def combo_tourism():
    arrivals = pick(singstat("M550001"), "Total International Visitor Arrivals By Place Of Residence", "M550001")
    occupancy = pick(singstat("M550081"), "Average Hotel Occupancy Rate", "M550081")
    return dict(
        title="Visitor arrivals & hotel occupancy, Singapore", short="Visitor arrivals & hotel occupancy",
        freq="Monthly", org="SingStat (STB)",
        source="Source: SingStat tables M550001 (International Visitor Arrivals) and M550081 (Hotel Statistics), "
               "data from the Singapore Tourism Board.",
        header=["Month", "Visitor arrivals", "Hotel occupancy (%)"],
        rows=month_rows(arrivals, occupancy, start=date(2019, 1, 1)), label_fmt="mmm yyyy", value_fmts=["#,##0", "0.0"],
        axes=dict(x="Month", y="Visitor arrivals", y2="Hotel occupancy (%)"),
    )


def combo_changi():
    s = singstat("M650601")
    passengers = pick(s, "Total Passengers", "M650601")
    cargo = pick(s, "Total Air Cargo Tonnage", "M650601")
    return dict(
        title="Changi Airport passengers & air cargo", short="Changi passengers & air cargo",
        freq="Yearly", org="SingStat (CAAS)",
        source="Source: SingStat table M650601, Civil Aircraft Arrivals And Departures, Passengers, Air Cargo "
               "Tonnage And Mail - Changi Airport (data from the Civil Aviation Authority of Singapore).",
        header=["Year", "Passengers", "Air cargo (tonnes)"],
        rows=year_rows(passengers, cargo), label_fmt="0", value_fmts=["#,##0", "#,##0"],
        axes=dict(x="Year", y="Passengers", y2="Air cargo (tonnes)"),
    )


def combo_electricity():
    gen = pick(singstat("M890831"), "Electricity Generation", "M890831")
    tariff = pick(singstat("M890991"), "Low Tension Supplies - Domestic", "M890991")
    return dict(
        title="Electricity generated & household tariff, Singapore", short="Electricity generated & tariff",
        freq="Monthly", org="SingStat (EMA, SP Group)",
        source="Source: SingStat tables M890831 (Electricity Generation, GWh) and M890991 (household low-tension "
               "tariff, cents per kWh before GST).",
        header=["Month", "Electricity generated (GWh)", "Household tariff (¢/kWh)"],
        rows=month_rows(gen, tariff, start=date(2019, 1, 1)), label_fmt="mmm yyyy", value_fmts=["#,##0.0", "0.00"],
        axes=dict(x="Month", y="Electricity generated (GWh)", y2="Household tariff (¢/kWh)"),
    )


def combo_coe():
    by_ex = {}
    for r in datagov("d_69b3380ad7e51aff3a7dcc84eba52b8a"):
        if r["vehicle_class"] == "Category A":
            by_ex[(r["month"], int(r["bidding_no"]))] = (num(r["quota"]), num(r["bids_received"]), num(r["premium"]))
    ordinal = {1: "1st", 2: "2nd", 3: "3rd"}
    rows = []
    for (m, n) in sorted(by_ex):
        y, mo = m.split("-")
        if int(y) >= 2024:
            label = date(int(y), int(mo), 1).strftime("%b %Y") + f" ({ordinal.get(n, str(n) + 'th')})"
            rows.append((label,) + by_ex[(m, n)])
    return dict(
        title="COE Category A: quota, bids & premium", short="COE Cat A quota, bids & premium",
        freq="Twice monthly", org="LTA",
        source="Source: LTA, COE Bidding Results / Prices (data.gov.sg d_69b3380ad7e51aff3a7dcc84eba52b8a). "
               "Category A: cars up to 1,600cc and 97kW (130bhp), or fully electric up to 110kW. One row per bidding exercise.",
        header=["Bidding exercise", "Quota", "Bids received", "Premium ($)"],
        rows=rows, label_fmt=None, value_fmts=["#,##0", "#,##0", '"$"#,##0'], widths={1: 20},
        axes=dict(x="Bidding exercise", y="Number of COEs", y2="Premium ($)"),
    )


COMBO_DATASETS = [("hdb", combo_hdb), ("births", combo_births), ("tourism", combo_tourism),
                  ("changi", combo_changi), ("electricity", combo_electricity), ("coe", combo_coe)]

# ---------------------------------------------------------------------------------------------
# Histogram samples (histogram/): first column = bins, then one column of counts per series.
# ---------------------------------------------------------------------------------------------

def census_total(table, drop=("Total",)):
    """A Census table's all-Singapore 'Total' row as [(bin, count)], in column order."""
    for no, name, values in singstat_rows(table):
        if name.strip().lower() == "total":
            return [(k, v) for k, v in values.items() if k not in drop and v is not None]
    raise RuntimeError(f"SingStat table {table} has no 'Total' row.")


def hist_hdb_prices():
    """Resale flats sold in the last 3 complete months, counted in $50k price bands."""
    rid = "d_8b84c4ee58e3cfc0ece0d773c8ca6abc"
    q = urllib.parse.urlencode({"resource_id": rid, "limit": 1, "sort": "month desc"})
    latest = get_json(f"https://data.gov.sg/api/action/datastore_search?{q}")["result"]["records"][0]["month"]
    y, m = map(int, latest.split("-"))
    today = date.today()
    if (y, m) == (today.year, today.month):  # this month is still being registered
        y, m = (y, m - 1) if m > 1 else (y - 1, 12)
    months = []
    for _ in range(3):
        months.insert(0, f"{y}-{m:02d}")
        y, m = (y, m - 1) if m > 1 else (y - 1, 12)
    prices = []
    for month in months:
        offset = 0
        while True:
            q = urllib.parse.urlencode({"resource_id": rid, "limit": 5000, "offset": offset,
                                        "filters": json.dumps({"month": month})})
            result = get_json(f"https://data.gov.sg/api/action/datastore_search?{q}")["result"]
            prices += [p for p in (num(r["resale_price"]) for r in result["records"]) if p is not None]
            offset += 5000
            if offset >= result["total"] or not result["records"]:
                break
    if not prices:
        raise RuntimeError("no resale transactions found")
    step, low, high = 50_000, 250_000, 1_500_000

    def money(v):
        return f"${v / 1e6:.2f}M".replace(".00M", "M") if v >= 1e6 else f"${v / 1e3:.0f}k"

    bins = [(f"Below {money(low)}", lambda p: p < low)]
    for start in range(low, high, step):
        bins.append((f"{money(start)}–{money(start + step)}", lambda p, s=start: s <= p < s + step))
    bins.append((f"{money(high)} & over", lambda p: p >= high))
    rows = [(label, sum(1 for p in prices if test(p))) for label, test in bins]
    first = date(int(months[0][:4]), int(months[0][5:]), 1).strftime("%b %Y")
    last = date(int(months[-1][:4]), int(months[-1][5:]), 1).strftime("%b %Y")
    return dict(
        title=f"HDB resale prices, {first} – {last}", short="HDB resale prices",
        summary=f"{len(prices):,} flats sold · {first} – {last} · HDB", freq="", org="HDB",
        source="Source: HDB, Resale flat prices based on registration date from Jan 2017 onwards "
               f"(data.gov.sg {rid}). Resale flats registered in {first} – {last}, counted in $50k price bands.",
        header=["Resale price", "Flats sold"],
        rows=rows, label_fmt=None, value_fmts=["#,##0"], widths={1: 18},
        axes=dict(x="Resale price", y="Number of flats sold"),
    )


def hist_age():
    """Residents by 5-year age group and sex, latest year."""
    rows = singstat_rows("M810011")
    group = {"2": {}, "3": {}}  # seriesNo 2.x = males, 3.x = females
    totals = {}
    for no, name, values in rows:
        no = str(no or "")
        if no in ("2", "3"):
            totals[no] = values
        top = no.split(".")[0]
        if top in group and "." in no and (re.fullmatch(r"\d+ - \d+ Years", name) or name == "90 Years & Over"):
            group[top][name] = values
    years = [y for y in totals["2"] if all(group[s].get(b, {}).get(y) is not None for s in group for b in group["2"])]
    if not years:
        raise RuntimeError("no year with every age group")
    year = max(years, key=int)
    bands = list(group["2"])  # table order: 0 - 4 Years ... 85 - 89 Years, 90 Years & Over
    for s in group:  # the bands must add up to the table's own totals
        if abs(sum(group[s][b][year] for b in bands) - totals[s][year]) > 1:
            raise RuntimeError(f"age groups don't add up to the total for {year}")
    label = lambda b: b.replace(" Years & Over", "+").replace(" Years", "").replace(" - ", "–")
    return dict(
        title=f"Singapore residents by age group, {year}", short="Residents by age & sex",
        summary=f"At end June {year} · by 5-year age group · SingStat", freq="", org="SingStat",
        source="Source: SingStat table M810011, Singapore Residents By Age Group, Ethnic Group And Sex, At End June. "
               "Residents = citizens and permanent residents.",
        header=["Age group", "Males", "Females"],
        rows=[(label(b), group["2"][b][year], group["3"][b][year]) for b in bands],
        label_fmt=None, value_fmts=["#,##0", "#,##0"],
        axes=dict(x="Age (years)", y="Residents"),
    )


def hist_household_income():
    return dict(
        title="Households by monthly household income from work, Singapore", short="Household income",
        summary="Census 2020 · resident households · SingStat", freq="", org="SingStat",
        source="Source: SingStat, Census of Population 2020, table 17779 (all-Singapore total). "
               "Excludes households with no employed person.",
        header=["Monthly household income from work", "Households"],
        rows=census_total("17779", drop=("Total", "No Employed Person")),
        label_fmt=None, value_fmts=["#,##0"], widths={1: 24},
        axes=dict(x="Monthly household income from work", y="Resident households"),
    )


def hist_income():
    return dict(
        title="Employed residents by gross monthly income from work, Singapore", short="Income from work",
        summary="Census 2020 · employed residents aged 15+ · SingStat", freq="", org="SingStat",
        source="Source: SingStat, Census of Population 2020, table 17571 (all-Singapore total). "
               "Employed residents aged 15 years and over.",
        header=["Gross monthly income from work", "Employed residents"],
        rows=census_total("17571"), label_fmt=None, value_fmts=["#,##0"], widths={1: 22},
        axes=dict(x="Gross monthly income from work", y="Employed residents"),
    )


def hist_household_size():
    return dict(
        title="Resident households by household size, Singapore", short="Household size",
        summary="Census 2020 · resident households · SingStat", freq="", org="SingStat",
        source="Source: SingStat, Census of Population 2020, table 17778 (all-Singapore total).",
        header=["Household size", "Households"],
        rows=census_total("17778"), label_fmt=None, value_fmts=["#,##0"], widths={1: 18},
        axes=dict(x="People in the household", y="Resident households"),
    )


HISTOGRAM_DATASETS = [("hdb_prices", hist_hdb_prices), ("age", hist_age), ("household_income", hist_household_income),
                      ("income", hist_income), ("household_size", hist_household_size)]

# Each group: its datasets, the folder its files go in, the page whose sample list is updated,
# and where the files are hosted.
GROUPS = [
    dict(name="Graphite", datasets=DATASETS, folder="samples", page="index.html",
         cdn="https://cdn.zhengjia.dev/graphite/samples/"),
    dict(name="Graphite Combo", datasets=COMBO_DATASETS, folder=os.path.join("combo", "samples"),
         page=os.path.join("combo", "index.html"), cdn="https://cdn.zhengjia.dev/graphite/combo/samples/"),
    dict(name="Graphite Histogram", datasets=HISTOGRAM_DATASETS, folder=os.path.join("histogram", "samples"),
         page=os.path.join("histogram", "index.html"), cdn="https://cdn.zhengjia.dev/graphite/histogram/samples/"),
]


# ---------------------------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------------------------

def period(v):
    return v.strftime("%b %Y") if isinstance(v, date) else re.sub(r" \(\w+\)$", "", str(v))


def summary(d):
    if d.get("summary"):
        return d["summary"]  # histograms describe their snapshot, not a date range
    return f"{d['freq']} · {period(d['rows'][0][0])} – {period(d['rows'][-1][0])} · {d['org']}"


def fingerprint(d):
    """Hash of the header + data, used to skip unchanged files and as the cache-busting tag."""
    payload = json.dumps([d["header"], [[period(c) if isinstance(c, date) else c for c in r] for r in d["rows"]]],
                         default=str, ensure_ascii=False)
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()[:8]


def existing_fingerprint(path):
    """Fingerprint stored in the existing file (cell A4 is unused, so it's kept in the workbook properties)."""
    if not os.path.exists(path):
        return ""
    try:
        wb = load_workbook(path, read_only=True)
    except Exception:
        return ""
    try:
        return wb.properties.keywords or ""
    finally:
        wb.close()  # read-only mode keeps the file open; Windows then refuses to replace it


def write_xlsx(d, path, fp):
    wb = Workbook()
    ws = wb.active
    ws.title = d["short"][:31]
    wb.properties.keywords = fp
    wb.properties.title = d["title"]
    ws["A1"] = d["title"]
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = summary(d)
    ws["A2"].font = Font(color="444444")
    ws["A3"] = d["source"]
    ws["A3"].font = Font(italic=True, color="666666")
    ws.append([])
    ws.append(d["header"])
    for c in ws[5]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    fmts = d.get("value_fmts") or [d["value_fmt"]] * (len(d["header"]) - 1)
    for r in d["rows"]:
        ws.append(list(r))
        row = ws.max_row
        if d["label_fmt"]:
            ws.cell(row, 1).number_format = d["label_fmt"]
        for col in range(2, len(d["header"]) + 1):
            ws.cell(row, col).number_format = fmts[col - 2]
    ws.freeze_panes = "B6"
    widths = d.get("widths") or {}
    for i in range(1, len(d["header"]) + 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(i, 16 if i > 1 else 14)
    tmp = path + ".tmp"
    wb.save(tmp)
    try:
        os.replace(tmp, path)  # never leave a half-written file behind
    except PermissionError:
        os.remove(tmp)
        raise RuntimeError(f"{os.path.basename(path)} is open in another program (Excel?). Close it and run again.")


def js_str(s):
    return json.dumps(s, ensure_ascii=False)


def entry_line(e):
    parts = [f"file: {js_str(e['file'])}", f"v: {js_str(e['v'])}", f"title: {js_str(e['title'])}",
             f"summary: {js_str(e['summary'])}"]
    if e.get("axes"):
        # x, y (and y2 for combo charts)
        axes = ", ".join(f"{k}: {js_str(e['axes'][k])}" for k in ("x", "y", "y2") if k in e["axes"])
        parts.append(f"axes: {{ {axes} }}")
    return "{ " + ", ".join(parts) + " }"


def sample_block(page):
    """(html, start, end, indent, {file: existing line}) for the SAMPLES:BEGIN/END list in a page."""
    with open(page, encoding="utf-8") as f:
        html = f.read()
    begin, end = "// SAMPLES:BEGIN", "// SAMPLES:END"
    if begin not in html or end not in html:
        raise RuntimeError(f"Couldn't find the SAMPLES:BEGIN / SAMPLES:END markers in {page}.")
    a = html.index(begin) + len(begin)
    b = html.index(end)
    indent = html[html.rindex("\n", 0, b) + 1:b]  # whitespace before // SAMPLES:END
    existing = {}
    for line in html[a:b].splitlines():
        m = re.search(r'\{ file: "([^"]+)"', line)
        if m:
            existing[m.group(1)] = line.strip().rstrip(",")
    return html, a, b, indent, existing


def update_index(page, lines):
    html, a, b, indent, _ = sample_block(page)
    new_html = html[:a] + "\n" + ",\n".join(indent + line for line in lines) + "\n" + indent + html[b:]
    if new_html != html:
        with open(page, "w", encoding="utf-8", newline="") as f:
            f.write(new_html)
        return True
    return False


def run_group(group):
    """Refresh one group's files and page. Returns (changed file paths, page changed?, failed names)."""
    folder = os.path.join(ROOT, group["folder"])
    page = os.path.join(ROOT, group["page"])
    os.makedirs(folder, exist_ok=True)
    existing = sample_block(page)[4]
    lines, changed, failed = [], [], []
    print(f"=== {group['name']} ({group['folder']})")
    for name, build in group["datasets"]:
        path = os.path.join(folder, name + ".xlsx")
        print(f"- {name}.xlsx")
        try:
            d = build()
            if not d["rows"]:
                raise RuntimeError("no rows returned")
        except Exception as err:
            print(f"    FAILED, keeping the existing file: {err}")
            failed.append(name)
            if name in existing:
                lines.append(existing[name])  # keep its current entry in the page
            continue
        fp = fingerprint(d)
        if fp == existing_fingerprint(path):
            print(f"    unchanged ({summary(d)})")
        else:
            try:
                write_xlsx(d, path, fp)
            except Exception as err:
                print(f"    FAILED, keeping the existing file: {err}")
                failed.append(name)
                if name in existing:
                    lines.append(existing[name])
                continue
            changed.append(os.path.join(group["folder"], name + ".xlsx"))
            print(f"    UPDATED -> {summary(d)}")
        lines.append(entry_line(dict(file=name, v=fp, title=d["short"], summary=summary(d), axes=d.get("axes"))))
    page_changed = update_index(page, lines)
    print()
    return changed, page_changed, failed


def coverage_problems():
    """Anything with samples that this script would NOT update: pages with a sample list that aren't
    in GROUPS, and .xlsx files in a sample folder that no dataset builds."""
    problems = []
    known_pages = {os.path.normcase(os.path.join(ROOT, g["page"])) for g in GROUPS}
    for dirpath, dirnames, filenames in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if not d.startswith(".") and d not in ("tools", "node_modules")]
        if "index.html" in filenames:
            page = os.path.join(dirpath, "index.html")
            with open(page, encoding="utf-8", errors="replace") as f:
                has_list = "// SAMPLES:BEGIN" in f.read()
            if has_list and os.path.normcase(page) not in known_pages:
                problems.append(f"{os.path.relpath(page, ROOT)} has a sample list, but no group in GROUPS updates it")
    for g in GROUPS:
        folder = os.path.join(ROOT, g["folder"])
        built = {name + ".xlsx" for name, _ in g["datasets"]}
        if os.path.isdir(folder):
            for f in sorted(os.listdir(folder)):
                if f.lower().endswith(".xlsx") and f not in built:
                    problems.append(f"{os.path.join(g['folder'], f)} isn't built by any dataset, so it will never be updated")
    return problems


def main():
    problems = coverage_problems()
    if problems:
        print("Not every sample is covered by this script:")
        for p in problems:
            print("    " + p)
        print("Add the page/dataset to GROUPS (or remove the stray file), then run again.")
        return 1

    any_failed = False
    report = []
    for group in GROUPS:
        try:
            changed, page_changed, failed = run_group(group)
        except Exception as err:
            print(f"    {group['name']} FAILED: {err}\n")
            any_failed = True
            continue
        report.append((group, changed, page_changed))
        any_failed = any_failed or bool(failed)
        if failed:
            print(f"Could not update ({group['name']}): {', '.join(failed)} - their existing files were left as they were.\n")

    for group, changed, page_changed in report:
        print(f"{group['name']}:")
        if changed:
            print(f"    upload to {group['cdn']} :")
            for path in changed:
                print(f"        {path}")
        else:
            print("    all sample files already up to date")
        if page_changed:
            print(f"    {group['page']} was updated (date ranges / version tags) - upload it too, after the files above")
    total = sum(len(g["datasets"]) for g in GROUPS)
    print(f"\nChecked all {total} sample files across {len(GROUPS)} pages ({', '.join(g['name'] for g in GROUPS)}).")
    return 1 if any_failed else 0


if __name__ == "__main__":
    sys.exit(main())
