"""
PediaZ v.01 - Pediatric Z-score Calculator based on WHO growth standards (0-5 y) and the WHO 2007 reference (5-19 y).

Main features
  * Smart column mapping: if the headers of the input file differ from the defaults
    (sex, age_months, weight_kg, height_cm) the program suggests a match for each
    variable in four combo boxes; the user confirms or corrects them.
  * WHO restricted-LMS adjustment for |z| > 3 (weight-for-age, weight-for-height, BMI-for-age).
  * Indicator-specific WHO plausibility (flag) limits, applied to z-scores rounded to two decimals.
  * 24-month boundary handled as in WHO software (< 24 months = length-based tables).
  * Input validation report (sheet 'input_checks'); sex coding; .xlsx/.xls/.csv/.txt input.
"""

import os
import re
import sys
import subprocess
import difflib
from datetime import datetime

import numpy as np
import pandas as pd

VERSION = "v.01"
APP = f"PediaZ {VERSION}"

BASE_DIR = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
WHO = os.path.join(BASE_DIR, "WHO_LMS")

STANDARD_COLS = ["sex", "age_months", "weight_kg", "height_cm"]
MAX_AGE_MONTHS = 228  # end of the WHO 2007 reference tables (19 y)

# WHO plausibility (flag) limits, in z-score units: (lower, upper); strict inequalities
FLAG_LIMITS = {
    "HFA": (-6.0, 6.0),
    "WFA": (-6.0, 5.0),
    "WFH": (-5.0, 5.0),
    "BFA": (-5.0, 5.0),
}
RESTRICTED = {"HFA": False, "WFA": True, "WFH": True, "BFA": True}


############################################
# COLUMN DETECTION (pure logic, no GUI)
############################################

SYNONYMS = {
    "sex": ["sex", "gender", "gen", "sexe", "sexo", "jens", "jensiat", "riagendr", "male", "female",
            "boygirl", "mf", "جنس", "جنسیت"],
    "age_months": ["agemonths", "agemonth", "agemon", "agem", "ageinmonths", "months", "month", "mo", "mos",
                   "ridagemn", "ridexagm", "ridageex", "age", "سن", "سنماه", "سنبهماه", "ماه"],
    "weight_kg": ["weightkg", "weight", "wt", "wgt", "wtkg", "bodyweight", "mass", "kg", "bmxwt", "وزن"],
    "height_cm": ["heightcm", "height", "ht", "hgt", "htcm", "length", "lengthcm", "stature", "lenhei", "cm",
                  "bmxht", "bmxrecum", "قد", "طول"],
}


def _norm(s):
    return re.sub(r"[^0-9a-zA-Z؀-ۿ]", "", str(s).lower())


def _score(header, target):
    h = _norm(header)
    if not h:
        return 0.0
    if h == _norm(target):
        return 100.0
    best = 0.0
    for syn in SYNONYMS[target]:
        syn = _norm(syn)
        if h == syn:
            best = max(best, 90.0)
        elif len(syn) >= 3 and syn in h:
            best = max(best, 60.0 + 10.0 * len(syn) / len(h))
        elif len(h) >= 3 and h in syn:
            best = max(best, 50.0)
        else:
            r = difflib.SequenceMatcher(None, h, syn).ratio()
            if r >= 0.8:
                best = max(best, r * 50.0)
    return best


def detect_columns(headers, threshold=30.0):
    """Return {standard_name: header or None}; every header is used at most once."""
    pairs = []
    for t in STANDARD_COLS:
        for h in headers:
            sc = _score(h, t)
            if sc >= threshold:
                pairs.append((sc, t, h))
    pairs.sort(key=lambda x: -x[0])
    result = {t: None for t in STANDARD_COLS}
    used = set()
    for sc, t, h in pairs:
        if result[t] is None and h not in used:
            result[t] = h
            used.add(h)
    return result


############################################
# GUI (Tkinter)
############################################

def ask_column_mapping(root, headers, preview_df, guess):
    """Show four combo boxes (sex, age, weight, height). Returns dict or None if cancelled."""
    import tkinter as tk
    from tkinter import ttk

    labels = {
        "sex": "Sex (m/f)",
        "age_months": "Age (months)",
        "weight_kg": "Weight (kg)",
        "height_cm": "Height / length (cm)",
    }

    win = root
    win.deiconify()
    win.title(f"{APP} - Confirm variables")
    win.resizable(False, False)
    result = {"value": None}

    frm = ttk.Frame(win, padding=14)
    frm.grid(row=0, column=0)

    ttk.Label(
        frm,
        text="PediaZ has matched the columns of your file to the four required variables.\n"
             "Please check each choice and correct it if needed.",
        justify="left",
    ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

    ttk.Label(frm, text="Variable", font=("TkDefaultFont", 9, "bold")).grid(row=1, column=0, sticky="w")
    ttk.Label(frm, text="Column in your file", font=("TkDefaultFont", 9, "bold")).grid(row=1, column=1, sticky="w", padx=8)
    ttk.Label(frm, text="Example values", font=("TkDefaultFont", 9, "bold")).grid(row=1, column=2, sticky="w")

    vars_, combos, examples = {}, {}, {}
    NONE = "-- select a column --"
    options = [NONE] + [str(h) for h in headers]

    def sample_text(col):
        if col in (None, NONE) or col not in preview_df.columns:
            return ""
        vals = preview_df[col].dropna().astype(str).head(3).tolist()
        return ", ".join(vals)

    def refresh(*_):
        chosen = [vars_[k].get() for k in STANDARD_COLS if vars_[k].get() != NONE]
        dup = len(chosen) != len(set(chosen))
        for k in STANDARD_COLS:
            examples[k].config(text=sample_text(vars_[k].get()))
        ok = (len(chosen) == 4) and not dup
        btn_ok.config(state="normal" if ok else "disabled")
        if dup:
            msg.config(text="The same column is selected for more than one variable.", foreground="#b00020")
        elif len(chosen) < 4:
            msg.config(text="Please select a column for every variable.", foreground="#b00020")
        else:
            msg.config(text="", foreground="#000000")

    for i, k in enumerate(STANDARD_COLS):
        ttk.Label(frm, text=labels[k]).grid(row=2 + i, column=0, sticky="w", pady=4)
        v = tk.StringVar(value=guess.get(k) if guess.get(k) is not None else NONE)
        cb = ttk.Combobox(frm, textvariable=v, values=options, state="readonly", width=32)
        cb.grid(row=2 + i, column=1, padx=8, pady=4)
        cb.bind("<<ComboboxSelected>>", refresh)
        ex = ttk.Label(frm, text="", foreground="#555555", width=28)
        ex.grid(row=2 + i, column=2, sticky="w")
        vars_[k], combos[k], examples[k] = v, cb, ex

    msg = ttk.Label(frm, text="")
    msg.grid(row=6, column=0, columnspan=3, sticky="w", pady=(8, 0))

    btns = ttk.Frame(frm)
    btns.grid(row=7, column=0, columnspan=3, sticky="e", pady=(10, 0))

    def on_ok():
        result["value"] = {k: vars_[k].get() for k in STANDARD_COLS}
        win.quit()

    def on_cancel():
        win.quit()

    btn_ok = ttk.Button(btns, text="Confirm and continue", command=on_ok)
    btn_ok.grid(row=0, column=0, padx=6)
    ttk.Button(btns, text="Cancel", command=on_cancel).grid(row=0, column=1)
    win.protocol("WM_DELETE_WINDOW", on_cancel)

    refresh()
    win.update_idletasks()
    w, h = win.winfo_reqwidth(), win.winfo_reqheight()
    win.geometry(f"+{(win.winfo_screenwidth() - w) // 2}+{(win.winfo_screenheight() - h) // 3}")
    win.lift()
    win.attributes("-topmost", True)
    win.after(200, lambda: win.attributes("-topmost", False))
    win.mainloop()
    frm.destroy()
    win.withdraw()
    return result["value"]


############################################
# LMS TABLES
############################################

_cache = {}


def load_table(file, xcol):
    path = os.path.join(WHO, file)
    if path in _cache:
        return _cache[path]
    df = pd.read_excel(path)
    df.columns = df.columns.astype(str).str.strip()
    t = df[[xcol, "L", "M", "S"]].astype(float).dropna().sort_values(xcol).values
    _cache[path] = t
    return t


def lms_lookup(t, x):
    """Linear interpolation of L, M, S at x. NaN outside the table range."""
    kw = dict(left=np.nan, right=np.nan)
    return (np.interp(x, t[:, 0], t[:, 1], **kw),
            np.interp(x, t[:, 0], t[:, 2], **kw),
            np.interp(x, t[:, 0], t[:, 3], **kw))


def z_from_lms(x, L, M, S, restricted):
    with np.errstate(all="ignore"):
        z = np.where(L == 0, np.log(x / M) / S, ((x / M) ** L - 1) / (L * S))
        if restricted:
            def sd(k):
                return np.where(L == 0, M * np.exp(S * k), M * (1 + L * S * k) ** (1 / L))
            sd3, sd2, sd2n, sd3n = sd(3), sd(2), sd(-2), sd(-3)
            z = np.where(z > 3, 3 + (x - sd3) / (sd3 - sd2), z)
            z = np.where(z < -3, -3 + (x - sd3n) / (sd2n - sd3n), z)
    return z


def compute_indicator(df, ind):
    """Return raw z-scores (restricted LMS where WHO applies it), NaN when not computable."""
    z = np.full(len(df), np.nan)
    restricted = RESTRICTED[ind]
    sexes = {"m": "boys", "f": "girls"}

    for key, label in sexes.items():
        sel = (df["sex"] == key).values
        if not sel.any():
            continue
        idx = np.where(sel)[0]
        age = df["age_months"].values[idx].astype(float)
        fl = np.floor(age)
        ht = df["height_cm"].values[idx].astype(float)
        wt = df["weight_kg"].values[idx].astype(float)
        bmi = df["BMI"].values[idx].astype(float)

        def apply(mask, file, xcol, meas, ref, clip=False):
            mask = mask & ~np.isnan(meas) & ~np.isnan(ref)
            if not mask.any():
                return
            t = load_table(file, xcol)
            r = ref[mask]
            if clip:
                r = np.clip(r, t[0, 0], t[-1, 0])
            L, M, S = lms_lookup(t, r)
            z[idx[mask]] = z_from_lms(meas[mask], L, M, S, restricted)

        if ind == "HFA":
            apply(fl < 24, f"lhfa_{label}_0-to-2-years_zscores.xlsx", "Month", ht, age, True)
            apply((fl >= 24) & (fl <= 60), f"lhfa_{label}_2-to-5-years_zscores.xlsx", "Month", ht, age, True)
            apply((fl >= 61) & (fl <= MAX_AGE_MONTHS), f"hfa-{label}-z-who-2007-exp.xlsx", "Month", ht, age, True)
        elif ind == "BFA":
            apply(fl < 24, f"bmi_{label}_0-to-2-years_zscores.xlsx", "Month", bmi, age, True)
            apply((fl >= 24) & (fl <= 60), f"bmi_{label}_2-to-5-years_zscores.xlsx", "Month", bmi, age, True)
            apply((fl >= 61) & (fl <= MAX_AGE_MONTHS), f"bmi-{label}-z-who-2007-exp.xlsx", "Month", bmi, age, True)
        elif ind == "WFA":
            apply(fl <= 60, f"wfa_{label}_0-to-5-years_zscores.xlsx", "Month", wt, age, True)
            apply((fl >= 61) & (fl <= 119), f"wfa-{label}-z-who-2007-exp.xlsx", "Month", wt, age, True)
        elif ind == "WFH":
            apply(fl < 24, f"wfl_{label}_0-to-2-years_zscores.xlsx", "Length", wt, ht)
            apply((fl >= 24) & (fl <= 60), f"wfh_{label}_2-to-5-years_zscores.xlsx", "Height", wt, ht)
    return z


############################################
# CLASSIFICATION
############################################

def interpret(z, ind):
    """Classify a z-score. As in WHO software, the z-score is rounded to two decimals before
    plausibility limits and cut-offs are applied. Categories are right-closed intervals:
    (-inf,-3], (-3,-2], (-2,1], (1,2], (2,3], (3,inf)."""
    if pd.isna(z):
        return "missing"
    z = round(float(z), 2)
    lo, hi = FLAG_LIMITS[ind]
    if z < lo:
        return "lowflag"
    if z > hi:
        return "highflag"
    if z <= -3:
        return "< -3 z"
    if z <= -2:
        return "-3 to -2 z"
    if z <= 1:
        return "-2 to +1 z"
    if z <= 2:
        return "+1 to +2 z"
    if z <= 3:
        return "+2 to +3 z"
    return "> +3 z"


############################################
# PREVALENCE
############################################

Z_ORDER = ["< -3 z", "-3 to -2 z", "-2 to +1 z", "+1 to +2 z", "+2 to +3 z", "> +3 z",
           "< -2 z", "> +1 z", "> +2 z", "lowflag", "highflag", "missing"]
EXCLUDE = ["missing", "lowflag", "highflag"]


def wilson_ci(n, total, z=1.96):
    if total == 0:
        return np.nan, np.nan
    p = n / total
    denom = 1 + z ** 2 / total
    center = (p + z ** 2 / (2 * total)) / denom
    margin = (z * np.sqrt((p * (1 - p) / total) + z ** 2 / (4 * total ** 2))) / denom
    return (center - margin) * 100, (center + margin) * 100


def prevalence_table(df, col):
    counts = df[col].value_counts(dropna=False)
    total = len(df)
    rows = []
    for cat in Z_ORDER:
        if cat == "< -2 z":
            n = int(counts.get("< -3 z", 0) + counts.get("-3 to -2 z", 0))
        elif cat == "> +1 z":
            n = int(counts.get("+1 to +2 z", 0) + counts.get("+2 to +3 z", 0) + counts.get("> +3 z", 0))
        elif cat == "> +2 z":
            n = int(counts.get("+2 to +3 z", 0) + counts.get("> +3 z", 0))
        else:
            n = int(counts.get(cat, 0))

        if total == 0:
            rows.append([cat, 0, np.nan, np.nan, np.nan, np.nan, np.nan])
            continue
        pct = n / total * 100
        if n == 0:
            rows.append([cat, 0, round(pct, 2), np.nan, np.nan, np.nan, np.nan])
            continue
        se = np.sqrt((pct / 100) * (1 - pct / 100) / total)
        wl, wh = (pct / 100 - 1.96 * se) * 100, (pct / 100 + 1.96 * se) * 100
        sl, sh = wilson_ci(n, total)
        rows.append([cat, n, round(pct, 2), round(wl, 2), round(wh, 2), round(sl, 2), round(sh, 2)])
    return pd.DataFrame(rows, columns=["Category", "N", "Percent", "CI_low_Wald", "CI_high_Wald",
                                       "CI_low_Wilson", "CI_high_Wilson"])


def prevalence_table_clean(df, col):
    return prevalence_table(df[~df[col].isin(EXCLUDE)], col)


def effective_sample_size(df, col):
    total = len(df)
    clean = int((~df[col].isin(EXCLUDE)).sum())
    return total, clean, (clean / total * 100 if total > 0 else 0)


def clean_descriptives(df, interp_col, z_col):
    clean_df = df[~df[interp_col].isin(EXCLUDE)]
    vals = pd.to_numeric(clean_df[z_col], errors="coerce").dropna()
    names = ["Mean", "SD", "Median", "IQR", "Min", "Max"]
    if len(vals) == 0:
        return pd.DataFrame({"Stat for CLEAN sample": names, "Value": [np.nan] * 6})
    q1, q3 = vals.quantile(0.25), vals.quantile(0.75)
    return pd.DataFrame({"Stat for CLEAN sample": names,
                         "Value": [round(vals.mean(), 2), round(vals.std(), 2), round(vals.median(), 2),
                                   round(q3 - q1, 2), round(vals.min(), 2), round(vals.max(), 2)]})


def auto_width(ws, df):
    if df.empty:
        return
    for i, col in enumerate(df.columns):
        width = max(df[col].fillna("").astype(str).map(len).max(), len(str(col))) + 2
        ws.set_column(i, i, min(width, 80))


############################################
# DATA PREPARATION
############################################

SEX_MAP = {
    "m": "m", "male": "m", "boy": "m", "boys": "m", "man": "m", "1": "m", "1.0": "m",
    "f": "f", "female": "f", "girl": "f", "girls": "f", "woman": "f", "2": "f", "2.0": "f",
    "پسر": "m", "مرد": "m", "مذکر": "m", "دختر": "f", "زن": "f", "مونث": "f",
}


def read_input(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".txt", ".tsv"):
        df = pd.read_csv(path, sep=None, engine="python")
    else:
        df = pd.read_excel(path)
    df.columns = df.columns.astype(str).str.strip()
    return df


def prepare(df_raw, colmap):
    """Rename the user's columns to the standard names and validate. Returns (df, checks)."""
    df = df_raw.copy()
    # if a standard name already exists in the file but is not the chosen column, keep it under a new name
    for std in STANDARD_COLS:
        chosen = colmap[std]
        if std in df.columns and chosen != std:
            df = df.rename(columns={std: f"{std}_original_unused"})
    df = df.rename(columns={colmap[k]: k for k in STANDARD_COLS})

    checks = []
    n = len(df)

    sx = df["sex"].astype(str).str.lower().str.strip().map(SEX_MAP)
    bad_sex = int(sx.isna().sum())
    df["sex"] = sx
    checks.append(["Sex not recognized (set to missing)", bad_sex])

    for c in ["age_months", "weight_kg", "height_cm"]:
        raw = df[c]
        num = pd.to_numeric(raw, errors="coerce")
        checks.append([f"{c}: non-numeric entries (set to missing)", int((num.isna() & raw.notna()).sum())])
        nonpos = int((num <= 0).sum())
        checks.append([f"{c}: zero or negative values (set to missing)", nonpos])
        num = num.where(num > 0)
        df[c] = num
    checks.append([f"age_months: outside 0-{MAX_AGE_MONTHS} months (no z-score)",
                   int(((df["age_months"] < 0) | (df["age_months"] > MAX_AGE_MONTHS + 0.999)).sum())])

    df["BMI"] = df["weight_kg"] / ((df["height_cm"] / 100) ** 2)
    checks.append(["Total records", n])
    return df, pd.DataFrame(checks, columns=["Check", "N"])


def unit_warnings(df):
    """Soft warnings about probable unit mix-ups."""
    w = []
    a = df["age_months"].dropna()
    if len(a) and a.max() <= 25 and a.median() < 20:
        w.append("age")
    h = df["height_cm"].dropna()
    if len(h) and h.median() < 3:
        w.append("height")
    return w


############################################
# ANALYSIS + REPORT

############################################
# Statistical comparison with the WHO standard, N(0,1) (no SciPy needed)
############################################
import math

MIN_N_TEST = 30
NOTES_TEXT = (
    "HOW TO READ THIS BLOCK\n"
    "- Bars: observed distribution of z-scores in the CLEAN sample. Line: expected distribution under the WHO standard, N(0,1). "
    "End bars hold all values below -4 or above +4.\n"
    "- Tests are one-sample tests against N(0,1): t test (mean), chi-square test (SD), Kolmogorov-Smirnov test (overall shape) "
    "and exact binomial tests (tail percentages). 'p (Holm)' adjusts for the number of blocks tested in this sheet.\n"
    "- Assumption: simple random sampling. In clustered or weighted surveys (e.g. NHANES) these p-values are too small; "
    "use survey-based methods for inference.\n"
    "- With large n, trivial differences become significant: judge the size of the difference, not only p.\n"
    "- A difference from N(0,1) describes the population; it does not mean a calculation error. The chi-square (SD) test is "
    "sensitive to non-normality, and for WFA, WFH and BFA the WHO adjustment beyond +/-3 makes the mean, SD and shape tests approximate.\n"
    "- Blocks overlap (total, sex, age), so tests are not independent. Tests are not run when n < 30."
)


def _betacf(a, b, x):
    tiny = 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    d = 1.0 / (d if abs(d) > tiny else tiny)
    h = d
    for m in range(1, 500):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d; d = d if abs(d) > tiny else tiny
        c = 1.0 + aa / c; c = c if abs(c) > tiny else tiny
        d = 1.0 / d; h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d; d = d if abs(d) > tiny else tiny
        c = 1.0 + aa / c; c = c if abs(c) > tiny else tiny
        d = 1.0 / d; delta = d * c; h *= delta
        if abs(delta - 1.0) < 1e-14:
            break
    return h


def betainc(a, b, x):
    """Regularized incomplete beta function I_x(a, b)."""
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    lbeta = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
    front = math.exp(lbeta + a * math.log(x) + b * math.log1p(-x))
    if x < (a + 1.0) / (a + b + 2.0):
        return front * _betacf(a, b, x) / a
    return 1.0 - front * _betacf(b, a, 1.0 - x) / b


def gammainc_lower(a, x):
    """Regularized lower incomplete gamma function P(a, x)."""
    if x <= 0.0:
        return 0.0
    if x < a + 1.0:
        term = total = 1.0 / a
        ap = a
        for _ in range(2000):
            ap += 1.0
            term *= x / ap
            total += term
            if abs(term) < abs(total) * 1e-15:
                break
        return total * math.exp(-x + a * math.log(x) - math.lgamma(a))
    tiny = 1e-300
    b = x + 1.0 - a
    c = 1.0 / tiny
    d = 1.0 / b
    h = d
    for i in range(1, 2000):
        an = -i * (i - a)
        b += 2.0
        d = an * d + b; d = d if abs(d) > tiny else tiny
        c = b + an / c; c = c if abs(c) > tiny else tiny
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < 1e-15:
            break
    return 1.0 - math.exp(-x + a * math.log(x) - math.lgamma(a)) * h


def norm_cdf(x):
    return 0.5 * math.erfc(-x / math.sqrt(2.0))


def p_mean_t(z):
    n = len(z)
    sd = z.std(ddof=1)
    if sd == 0:
        return np.nan
    t = z.mean() / (sd / math.sqrt(n))
    df = n - 1
    return betainc(df / 2.0, 0.5, df / (df + t * t))


def p_sd_chi2(z):
    n = len(z)
    stat = (n - 1) * z.var(ddof=1)
    cdf = gammainc_lower((n - 1) / 2.0, stat / 2.0)
    return min(1.0, 2.0 * min(cdf, 1.0 - cdf))


def ks_normal(z):
    """KS statistic against N(0,1) and asymptotic p (Stephens correction); conservative for tied (rounded) values."""
    x = np.sort(z)
    n = len(x)
    u, first = np.unique(x, return_index=True)
    cdf = np.array([norm_cdf(v) for v in u])
    after = (np.append(first[1:], n)) / n
    before = first / n
    d = float(max(np.max(np.abs(after - cdf)), np.max(np.abs(before - cdf))))
    lam = (math.sqrt(n) + 0.12 + 0.11 / math.sqrt(n)) * d
    p = 0.0
    for k in range(1, 101):
        t = 2.0 * (-1) ** (k - 1) * math.exp(-2.0 * k * k * lam * lam)
        p += t
        if abs(t) < 1e-12:
            break
    return d, min(1.0, max(0.0, p))


def p_binom_two_sided(k, n, p0):
    """Exact two-sided binomial p-value (doubling the smaller tail)."""
    upper = 1.0 if k <= 0 else betainc(k, n - k + 1, p0)          # P(X >= k)
    lower = 1.0 if k >= n else betainc(n - k, k + 1, 1.0 - p0)    # P(X <= k)
    return min(1.0, 2.0 * min(upper, lower))


def holm(pvals):
    """Holm step-down adjustment; NaN values are ignored."""
    idx = [i for i, p in enumerate(pvals) if p is not None and not pd.isna(p)]
    out = [np.nan] * len(pvals)
    order = sorted(idx, key=lambda i: pvals[i])
    m, run = len(order), 0.0
    for rank, i in enumerate(order):
        run = max(run, min(1.0, (m - rank) * pvals[i]))
        out[i] = run
    return out


TEST_NAMES = ["Mean z (t test)", "SD of z (chi-square test)", "Overall shape (Kolmogorov-Smirnov D)",
              "% < -3 z (exact binomial)", "% < -2 z (exact binomial)", "% > +1 z (exact binomial)",
              "% > +2 z (exact binomial)", "% > +3 z (exact binomial)"]


def compare_with_who(sub, interp_col, z_col):
    """Return None if n < MIN_N_TEST, otherwise (n, rows); each row = [name, observed, expected, difference, raw p]."""
    keep = ~sub[interp_col].isin(["missing", "lowflag", "highflag"])
    z = sub.loc[keep, z_col].astype(float).values
    n = len(z)
    if n < MIN_N_TEST:
        return None
    rows = []
    mean, sd = float(z.mean()), float(z.std(ddof=1))
    rows.append([TEST_NAMES[0], round(mean, 3), 0.0, round(mean, 3), p_mean_t(z)])
    rows.append([TEST_NAMES[1], round(sd, 3), 1.0, round(sd - 1.0, 3), p_sd_chi2(z)])
    d, pks = ks_normal(z)
    rows.append([TEST_NAMES[2], round(d, 3), "-", "-", pks])
    specs = [(TEST_NAMES[3], int((z <= -3).sum()), norm_cdf(-3)), (TEST_NAMES[4], int((z <= -2).sum()), norm_cdf(-2)),
             (TEST_NAMES[5], int((z > 1).sum()), 1 - norm_cdf(1)), (TEST_NAMES[6], int((z > 2).sum()), 1 - norm_cdf(2)),
             (TEST_NAMES[7], int((z > 3).sum()), 1 - norm_cdf(3))]
    for name, k, p0 in specs:
        rows.append([name, round(k / n * 100, 2), round(p0 * 100, 2), round(k / n * 100 - p0 * 100, 2),
                     p_binom_two_sided(k, n, p0)])
    return n, rows, z


# Layout: the chart spans columns I:N (right edge = right edge of the p (Holm) column) and runs from
# row r+11 down to the bottom of the 'highflag' row of the clean-sample table (r+31): 21 rows x 20 px.
CMP_COL_WIDTHS = (38, 12, 14, 12, 10, 10)
CHART_W_PX = sum(int(w_ * 7 + 5) for w_ in CMP_COL_WIDTHS)
CHART_H_PX = 21 * 20

CHART_BLOCKS = {"OVERALL - BOTH SEXES", "BOYS", "GIRLS", "AGE < 61 MONTHS - BOTH", "AGE < 61 MONTHS - BOYS",
                "AGE < 61 MONTHS - GIRLS", "AGE >= 61 MONTHS - BOTH", "AGE >= 61 MONTHS - BOYS",
                "AGE >= 61 MONTHS - GIRLS"}
EDGES = np.arange(-4.0, 4.0001, 0.5)


def hist_vs_normal(z):
    """Observed vs expected percentage per 0.5-z bin (right-closed); open-ended bins at both ends."""
    n = len(z)
    cuts = list(EDGES)
    labels = ["< -4"] + [f"{a:+.1f} to {b:+.1f}" for a, b in zip(cuts[:-1], cuts[1:])] + ["> +4"]
    obs, exp = [], []
    lo = -np.inf
    for hi in cuts + [np.inf]:
        obs.append(float(((z > lo) & (z <= hi)).sum()) / n * 100)
        exp.append(((norm_cdf(hi) if np.isfinite(hi) else 1.0) - (norm_cdf(lo) if np.isfinite(lo) else 0.0)) * 100)
        lo = hi
    return labels, obs, exp

############################################

METRICS = ["HFA", "WFA", "WFH", "BFA"]


def run_analysis(df, checks, colmap, output, input_name=""):
    for m in METRICS:
        z_raw = compute_indicator(df, m)
        df[f"{m}_z"] = np.round(z_raw, 2)
        df[f"{m}_interp"] = [interpret(v, m) for v in z_raw]
    df["BMI"] = df["BMI"].round(2)  # z-scores above were computed from the unrounded BMI

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        header = wb.add_format({"bold": True, "bg_color": "#FFF2CC"})
        pfmt = wb.add_format({"num_format": "0.0000", "align": "right"})
        ptxt = wb.add_format({"align": "right"})

        df.to_excel(writer, sheet_name="z_scores", index=False)
        ws = writer.sheets["z_scores"]
        ws.freeze_panes(1, 0)
        ws.set_row(0, None, header)
        auto_width(ws, df)

        for m in METRICS:
            name = f"{m}_prevalence"
            interp_col, z_col = f"{m}_interp", f"{m}_z"
            state = {"row": 0, "ws": None, "widths": [0] * 7}

            def track(tbl):
                for i, c in enumerate(tbl.columns[:7]):
                    w_ = max(tbl[c].fillna("").astype(str).map(len).max(), len(str(c))) + 2
                    state["widths"][i] = max(state["widths"][i], w_)

            pending = []

            def write_block(title, sub):
                pending.append((title, sub))

            def emit_block(title, sub, cmp_res, holm_p):
                r = state["row"]
                p_total = prevalence_table(sub, interp_col)
                p_total.to_excel(writer, sheet_name=name, startrow=r + 3, index=False)
                track(p_total)
                state["ws"] = writer.sheets[name]
                w = state["ws"]
                w.write(r, 0, title, header)
                w.write(r + 1, 0, "Based on TOTAL sample")
                total, clean, pct = effective_sample_size(sub, interp_col)
                w.write(r + 2, 0, f"Effective Sample Size: {clean}/{total} ({pct:.1f}%)")
                p_clean = prevalence_table_clean(sub, interp_col)
                clean_start = r + len(p_total) + 8
                p_clean.to_excel(writer, sheet_name=name, startrow=clean_start, index=False)
                track(p_clean)
                w.write(r + len(p_total) + 6, 0, f"Based on CLEAN sample (without missing and flags), n = {clean}")
                desc = clean_descriptives(sub, interp_col, z_col)
                desc.to_excel(writer, sheet_name=name, startrow=clean_start + len(p_clean) + 2, index=False)
                track(desc)
                # ---- right-hand side: comparison with the WHO standard
                w.write(r, 8, f"{title}: comparison with the WHO standard, N(0,1)", header)
                if cmp_res is None:
                    w.write(r + 1, 8, f"Not tested (clean sample n = {clean}; minimum n = {MIN_N_TEST}).")
                else:
                    n_cl, rows_c, zc = cmp_res
                    for j, h_ in enumerate(["Test (clean sample, n = %d)" % n_cl, "Observed", "Expected (WHO)",
                                            "Difference", "p (raw)", "p (Holm)"]):
                        w.write(r + 1, 8 + j, h_, header)
                    for i_, row_c in enumerate(rows_c):
                        for j, v in enumerate(row_c[:4]):
                            w.write(r + 2 + i_, 8 + j, v)
                        for j, v in ((4, row_c[4]), (5, holm_p[i_])):
                            if pd.isna(v):
                                w.write(r + 2 + i_, 8 + j, "NA", ptxt)
                            elif v < 0.0001:
                                w.write(r + 2 + i_, 8 + j, "<0.0001", ptxt)
                            else:
                                w.write_number(r + 2 + i_, 8 + j, round(v, 4), pfmt)
                    if title in CHART_BLOCKS:
                        labels, obs, exp = hist_vs_normal(zc)
                        for i_, (lb, o_, e_) in enumerate(zip(labels, obs, exp)):
                            w.write(r + i_, 26, lb)
                            w.write_number(r + i_, 27, round(o_, 3))
                            w.write_number(r + i_, 28, round(e_, 3))
                        k_ = len(labels)
                        col = wb.add_chart({"type": "column"})
                        col.add_series({"name": f"Observed (n = {n_cl})", "categories": [name, r, 26, r + k_ - 1, 26],
                                        "values": [name, r, 27, r + k_ - 1, 27], "gap": 30,
                                        "fill": {"color": "#8FAADC"}, "border": {"none": True}})
                        ln = wb.add_chart({"type": "line"})
                        ln.add_series({"name": "WHO standard, N(0,1)", "categories": [name, r, 26, r + k_ - 1, 26],
                                       "values": [name, r, 28, r + k_ - 1, 28],
                                       "line": {"color": "#C00000", "width": 2.25}})
                        col.combine(ln)
                        col.set_title({"name": title, "name_font": {"size": 10}})
                        col.set_x_axis({"name": "z-score bin (right-closed)", "num_font": {"size": 7, "rotation": -60}})
                        col.set_y_axis({"name": "% of children", "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}}})
                        col.set_legend({"position": "bottom"})
                        col.set_size({"width": CHART_W_PX, "height": CHART_H_PX})
                        w.insert_chart(r + 11, 8, col)
                        w.insert_textbox(r + 11, 14, NOTES_TEXT, {"width": 430, "height": CHART_H_PX, "x_offset": 12,
                                         "font": {"size": 8}, "align": {"vertical": "top"},
                                         "fill": {"color": "#FFFBE6"}, "line": {"color": "#BF9000"}})
                    else:
                        w.write(r + 10, 8, "Notes: see the notes beside the charts (e.g. OVERALL - BOTH SEXES) and the readme sheet.")
                state["row"] += len(p_total) * 2 + len(desc) + 18

            boys, girls = df["sex"] == "m", df["sex"] == "f"
            age = df["age_months"]
            write_block("OVERALL - BOTH SEXES", df)
            write_block("BOYS", df[boys])
            write_block("GIRLS", df[girls])
            lt, ge = age < 61, age >= 61
            write_block("AGE < 61 MONTHS - BOTH", df[lt])
            write_block("AGE < 61 MONTHS - BOYS", df[lt & boys])
            write_block("AGE < 61 MONTHS - GIRLS", df[lt & girls])
            write_block("AGE >= 61 MONTHS - BOTH", df[ge])
            write_block("AGE >= 61 MONTHS - BOYS", df[ge & boys])
            write_block("AGE >= 61 MONTHS - GIRLS", df[ge & girls])
            band = df[age <= 6]
            write_block("AGE <= 6 MONTHS - BOTH", band)
            write_block("AGE <= 6 MONTHS - BOYS", band[band["sex"] == "m"])
            write_block("AGE <= 6 MONTHS - GIRLS", band[band["sex"] == "f"])
            bins = list(range(0, MAX_AGE_MONTHS + 1, 12))
            for a, b in zip(bins[:-1], bins[1:]):
                band = df[(age >= a) & (age < b)]
                write_block(f"AGE {a}-{b - 1} MONTHS - BOTH", band)
                write_block(f"AGE {a}-{b - 1} MONTHS - BOYS", band[band["sex"] == "m"])
                write_block(f"AGE {a}-{b - 1} MONTHS - GIRLS", band[band["sex"] == "f"])
            # tests first (Holm needs all blocks), then write everything
            cmp_all = [compare_with_who(sub, interp_col, z_col) for _, sub in pending]
            holm_cols = []
            for t_i in range(len(TEST_NAMES)):
                holm_cols.append(holm([c[1][t_i][4] if c is not None else np.nan for c in cmp_all]))
            for b_i, (title, sub) in enumerate(pending):
                emit_block(title, sub, cmp_all[b_i], [holm_cols[t_i][b_i] for t_i in range(len(TEST_NAMES))])
            if state["ws"] is not None:
                for i, wd in enumerate(state["widths"]):
                    state["ws"].set_column(i, i, min(wd, 80))
                for i_, wd_ in zip(range(8, 14), CMP_COL_WIDTHS):
                    state["ws"].set_column(i_, i_, wd_)
                state["ws"].set_column(26, 28, 14)

        # data quality
        total = len(df)
        rows = []
        for m in METRICS:
            col = f"{m}_interp"
            miss = int((df[col] == "missing").sum())
            flags = int(df[col].isin(["lowflag", "highflag"]).sum())
            valid = total - miss - flags
            d = total if total else 1
            rows.append([m, total, miss, round(miss / d * 100, 3), flags, round(flags / d * 100, 3),
                         valid, round(valid / d * 100, 3)])
        dq = pd.DataFrame(rows, columns=["Metric", "Total", "Missing_N", "Missing_%", "Flags_N", "Flags_%",
                                         "Valid_N", "Valid_%"])
        dq.to_excel(writer, sheet_name="data_quality", index=False)
        ws = writer.sheets["data_quality"]
        ws.freeze_panes(1, 0)
        ws.set_row(0, None, header)
        auto_width(ws, dq)

        checks.to_excel(writer, sheet_name="input_checks", index=False)
        ws = writer.sheets["input_checks"]
        ws.set_row(0, None, header)
        auto_width(ws, checks)

        mapping = "; ".join(f"{k} <- '{colmap[k]}'" for k in STANDARD_COLS)
        flag_txt = "; ".join(f"{m}: z < {lo:g} or z > {hi:g}" for m, (lo, hi) in FLAG_LIMITS.items())
        readme = pd.DataFrame({"Information": [
            "Software name: PediaZ, an abbreviation of Pediatric Z-score Calculator.",
            f"Version: {VERSION}",
            f"Analysis date: {datetime.now().strftime('%Y-%m-%d')}",
            f"Input file: {os.path.basename(input_name)}" if input_name else "Input file: -",
            f"Column mapping used: {mapping}",
            "Z-scores were calculated using the WHO Child Growth Standards (0-5 y) and the WHO 2007 growth reference (5-19 y).",
            "Weight-based indicators (WFA, WFH, BFA) use the WHO restricted-LMS adjustment for |z| > 3.",
            f"Flag definition (WHO plausibility limits, applied to z-scores rounded to two decimals): {flag_txt}.",
            "Weight-for-age is calculated up to 119 months (younger than 10 years); weight-for-height up to 60 months.",
            "All 95% confidence intervals (95% CIs) were calculated using both the Wald and Wilson methods.",
            "Prevalence tables include both the total sample and the clean analytical sample, which excludes missing data and flags.",
            "The effective sample size (clean/total) is reported at the top of each prevalence block.",
            "To the right of each prevalence block, the clean-sample z-score distribution is compared with the WHO standard, N(0,1): t test (mean), chi-square test (SD), Kolmogorov-Smirnov test (shape) and exact binomial tests (percentages below -3, below -2, above +1, +2 and +3 z), with raw and Holm-adjusted p-values (Holm across the blocks of each sheet). Charts are drawn for the overall, sex-specific, <61-month and >=61-month blocks.",
            "These p-values assume simple random sampling; for clustered or weighted surveys they are too small and should be read descriptively. With large samples trivial differences are significant, so the size of the difference matters. Tests are not run when the clean sample has fewer than 30 records.",
            "Use of this software is free; citation is recommended.",
            "Developer: Dr. Amin Sharifi, Ph.D. in Nutritional Sciences (Minor in Epidemiology); ORCID iD: http://orcid.org/0000-0002-3164-0479",
            "Email: a.sharifi1983@gmail.com",
        ]})
        readme.to_excel(writer, sheet_name="readme", index=False)
        auto_width(writer.sheets["readme"], readme)
    return df


############################################
# MAIN (GUI workflow)
############################################

def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()

    if not messagebox.askyesno(
        APP,
        "Please prepare a data file (Excel, CSV or tab-delimited text) containing:\n\n"
        "sex, age (months), weight (kg), height (cm)\n\n"
        "Column names may differ from the defaults; PediaZ will ask you to confirm them.\n"
        "Z-scores will be calculated based on WHO standards.\n\nDo you want to continue?"):
        return

    path = filedialog.askopenfilename(
        title=f"{APP} - Select INPUT file",
        filetypes=[("Data files", "*.xlsx *.xls *.csv *.txt *.tsv"), ("All files", "*.*")])
    if not path:
        messagebox.showerror(APP, "No input file selected.")
        return

    if not os.path.isdir(WHO):
        messagebox.showerror(APP, "WHO_LMS folder not found.")
        return

    try:
        raw = read_input(path)
    except Exception as e:
        messagebox.showerror(APP, f"The file could not be read:\n{e}")
        return
    if raw.empty:
        messagebox.showerror(APP, "The file contains no data rows.")
        return

    guess = detect_columns(list(raw.columns))
    root.update()
    colmap = ask_column_mapping(root, list(raw.columns), raw, guess)
    if colmap is None:
        return

    output = filedialog.asksaveasfilename(
        title=f"{APP} - Save OUTPUT file", defaultextension=".xlsx",
        initialfile="PediaZ_output.xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output:
        return

    try:
        df, checks = prepare(raw, colmap)

        warn = unit_warnings(df)
        if "age" in warn and messagebox.askyesno(
                APP, "The age values look like YEARS (maximum <= 25).\nConvert age from years to months (x 12)?"):
            df["age_months"] = df["age_months"] * 12
        if "height" in warn and messagebox.askyesno(
                APP, "The height values look like METERS (median < 3).\nConvert height from meters to centimeters (x 100)?"):
            df["height_cm"] = df["height_cm"] * 100
            df["BMI"] = df["weight_kg"] / ((df["height_cm"] / 100) ** 2)

        run_analysis(df, checks, colmap, output, path)
    except PermissionError:
        messagebox.showerror(APP, "The output file could not be written. Close it in Excel and try again.")
        return
    except Exception as e:
        messagebox.showerror(APP, str(e))
        return

    if messagebox.askyesno(APP, f"Z-scores were calculated following WHO standards and saved to:\n{output}\n\n"
                                "Do you want to open the output file?"):
        if os.name == "nt":
            os.startfile(output)
        elif sys.platform == "darwin":
            subprocess.call(["open", output])
        else:
            subprocess.call(["xdg-open", output])


if __name__ == "__main__":
    main()
