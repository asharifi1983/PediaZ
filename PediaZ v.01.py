import pandas as pd
import numpy as np
from scipy.interpolate import interp1d
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import subprocess
from datetime import datetime


############################################
# ROOT
############################################

root = tk.Tk()
root.withdraw()

############################################
# START MESSAGE
############################################

continue_app = messagebox.askyesno(
    "PediaZ v.01",
    "Please prepare an Excel file containing the following columns and headers:\n\n" 
    "sex, age_months, weight_kg, height_cm\n\n"
    "Z-scores will be calculated based on WHO standards.\n\n"
    "Do you want to continue?"
)

if not continue_app:
    sys.exit()

############################################
# FILE DIALOG
############################################

INPUT = filedialog.askopenfilename(
    title="PediaZ v.01 - Select INPUT Excel file",
    filetypes=[("Excel files","*.xlsx")]
)

if not INPUT:
    messagebox.showerror("PediaZ v.01","No Excel file selected.")
    sys.exit()

OUTPUT = filedialog.asksaveasfilename(
    title="PediaZ v.01 - Save OUTPUT file",
    defaultextension=".xlsx",
    initialfile="PediaZ_output.xlsx",
    filetypes=[("Excel files","*.xlsx")]
)

if not OUTPUT:
    sys.exit()


############################################
# WHO PATH
############################################

BASE_DIR = sys._MEIPASS if getattr(sys,'frozen',False) else os.path.dirname(os.path.abspath(__file__))
WHO = os.path.join(BASE_DIR,"WHO_LMS")

if not os.path.exists(WHO):
    messagebox.showerror("PediaZ v.01","WHO_LMS folder not found.")
    sys.exit()


############################################
# AUTO WIDTH
############################################

def auto_width(ws,df):

    if df.empty:
        return

    for i,col in enumerate(df.columns):

        max_len = max(
            df[col].fillna("").astype(str).map(len).max(),
            len(str(col))
        ) + 2

        ws.set_column(i,i,max_len)


############################################
# LMS
############################################

cache={}

def load_table(file,xcol):

    path=os.path.join(WHO,file)

    if path in cache:
        return cache[path]

    df=pd.read_excel(path)
    df.columns=df.columns.str.strip()

    fL=interp1d(df[xcol],df['L'],bounds_error=False,fill_value=np.nan)
    fM=interp1d(df[xcol],df['M'],bounds_error=False,fill_value=np.nan)
    fS=interp1d(df[xcol],df['S'],bounds_error=False,fill_value=np.nan)

    cache[path]=(fL,fM,fS)
    return cache[path]


############################################
# VECTOR LMS Z
############################################

def lms_z_vector(x,L,M,S):

    x=np.array(x, dtype=float)
    L=np.array(L, dtype=float)
    M=np.array(M, dtype=float)
    S=np.array(S, dtype=float)

    z=np.full_like(x, np.nan, dtype=float)

    valid = ~(np.isnan(x)|np.isnan(L)|np.isnan(M)|np.isnan(S))

    if not np.any(valid):
        return z

    idx0 = valid & (L==0)
    idx1 = valid & (L!=0)

    z[idx0] = np.log(x[idx0]/M[idx0]) / S[idx0]
    z[idx1] = ((x[idx1]/M[idx1])**L[idx1] - 1) / (L[idx1]*S[idx1])

    z = np.clip(z, -6, 6)
    z = np.round(z, 2)

    return z


############################################
# INTERPRETATION
############################################

def interpret(z, metric=None):

    if pd.isna(z):
        return "missing"

    if metric=="HFA":
        if z < -6: return "lowflag"
        if z > 6: return "highflag"
    elif metric=="WFA":
        if z < -6: return "lowflag"
        if z > 5: return "highflag"
    elif metric=="WFH":
        if z < -5: return "lowflag"
        if z > 5: return "highflag"
    elif metric=="BFA":
        if z < -5: return "lowflag"
        if z > 5: return "highflag"

    if z < -3:
        return "< -3 z"
    if z < -2:
        return "-3 to -2 z"
    if z <= 1:
        return "-2 to +1 z"
    if z <= 2:
        return "+1 to +2 z"
    if z <= 3:
        return "+2 to +3 z"
    return "> +3 z"


############################################
# AGE GROUP
############################################

def age_group(m):

    if pd.isna(m):
        return "missing"

    m = int(m)
    y=m/12

    if y<5:
        return "0-4 y"
    elif y<10:
        return "5-9 y"
    elif y<15:
        return "10-14 y"
    elif y<=19:
        return "15-19 y"
    else:
        return "out_of_scope"


############################################
# PREVALENCE
############################################

Z_ORDER = [
    "< -3 z",
    "-3 to -2 z",
    "< -2 z",
    "-2 to +1 z",  
    "> +1 z",
    "+1 to +2 z", 
    "> +2 z",
    "+2 to +3 z",
    "> +3 z", 
    "lowflag",
    "highflag",
    "missing"
]

def wilson_ci(n, total, z=1.96):

    if total == 0:
        return np.nan, np.nan

    p = n/total
    denom = 1 + z**2/total
    center = (p + z**2/(2*total)) / denom
    margin = (z * np.sqrt((p*(1-p)/total) + z**2/(4*total**2))) / denom

    return (center-margin)*100, (center+margin)*100


def prevalence_table(df,col):

    counts = df[col].value_counts(dropna=False)
    total=len(df)

    rows=[]

    for cat in Z_ORDER:

        if cat == "< -2 z":
            n = int(counts.get("< -3 z",0) + counts.get("-3 to -2 z",0))
        elif cat == "> +1 z":
            n = int(counts.get("+1 to +2 z",0) + counts.get("+2 to +3 z",0) + counts.get("> +3 z",0))
        elif cat == "> +2 z":
            n = int(counts.get("+2 to +3 z",0) + counts.get("> +3 z",0))
        else:
            n=int(counts.get(cat,0))

        if total==0:
            rows.append([cat,0,np.nan,np.nan,np.nan,np.nan,np.nan])
            continue

        pct=n/total*100

        if n==0:
            rows.append([cat,0,round(pct,2),np.nan,np.nan,np.nan,np.nan])
            continue

        if cat in ["lowflag","highflag","missing"]:
            wald_low = wald_high = wil_low = wil_high = np.nan
        else:
            se=np.sqrt((pct/100)*(1-pct/100)/total)
            wald_low=(pct/100-1.96*se)*100
            wald_high=(pct/100+1.96*se)*100
            wil_low, wil_high = wilson_ci(n,total)

        rows.append([
            cat,
            n,
            round(pct,2),
            round(wald_low,2) if not pd.isna(wald_low) else np.nan,
            round(wald_high,2) if not pd.isna(wald_high) else np.nan,
            round(wil_low,2) if not pd.isna(wil_low) else np.nan,
            round(wil_high,2) if not pd.isna(wil_high) else np.nan
        ])

    return pd.DataFrame(rows,columns=[
        "Category","N","Percent",
        "CI_low_Wald","CI_high_Wald",
        "CI_low_Wilson","CI_high_Wilson"
    ])


def prevalence_table_clean(df,col):
    clean_df = df[~df[col].isin(["missing","lowflag","highflag"])]
    return prevalence_table(clean_df,col)


def effective_sample_size(df,col):
    total=len(df)
    clean=len(df[~df[col].isin(["missing","lowflag","highflag"])])
    pct=(clean/total*100) if total>0 else 0
    return total,clean,pct


def clean_descriptives(df, interp_col, z_col):
    clean_df = df[~df[interp_col].isin(["missing","lowflag","highflag"])]
    vals = pd.to_numeric(clean_df[z_col], errors='coerce').dropna()

    if len(vals)==0:
        return pd.DataFrame({
            "Stat for CLEAN sample":["Mean","SD","Median","IQR","Min","Max"],
            "Value":[np.nan]*6
        })

    q1 = vals.quantile(0.25)
    q3 = vals.quantile(0.75)

    return pd.DataFrame({
        "Stat for CLEAN sample":["Mean","SD","Median","IQR","Min","Max"],
        "Value":[
            round(vals.mean(),2),
            round(vals.std(),2),
            round(vals.median(),2),
            round(q3-q1,2),
            round(vals.min(),2),
            round(vals.max(),2)
        ]
    })


############################################
# VECTOR LMS FETCH
############################################

def compute_metric_vectorized(df, measure):

    z = np.full(len(df), np.nan)

    sexes = {'m':'boys','f':'girls'}

    for sex_key,sex_label in sexes.items():

        mask_sex = df['sex']==sex_key
        sub = df[mask_sex]

        if sub.empty:
            continue

        ages = np.floor(sub['age_months'].values).astype(float)
        heights = sub['height_cm'].values
        weights = sub['weight_kg'].values
        bmis = sub['BMI'].values

        idx = sub.index

        def apply_file(mask, file, xcol, measure_values, ref_values):
            if not np.any(mask):
                return

            fL,fM,fS = load_table(file,xcol)

            ref = ref_values[mask]

            L = fL(ref)
            M = fM(ref)
            S = fS(ref)

            z_vals = lms_z_vector(measure_values[mask],L,M,S)
            z[idx[mask]] = z_vals

        if measure=="HFA":

            apply_file(ages<=24,
                       f"lhfa_{sex_label}_0-to-2-years_zscores.xlsx",
                       'Month',
                       heights,
                       ages)

            apply_file((ages>24)&(ages<=60),
                       f"lhfa_{sex_label}_2-to-5-years_zscores.xlsx",
                       'Month',
                       heights,
                       ages)

            apply_file((ages>=61)&(ages<=216),
                       f"hfa-{sex_label}-z-who-2007-exp.xlsx",
                       'Month',
                       heights,
                       ages)

        elif measure=="BFA":

            apply_file(ages<=24,
                       f"bmi_{sex_label}_0-to-2-years_zscores.xlsx",
                       'Month',
                       bmis,
                       ages)

            apply_file((ages>24)&(ages<=60),
                       f"bmi_{sex_label}_2-to-5-years_zscores.xlsx",
                       'Month',
                       bmis,
                       ages)

            apply_file((ages>=61)&(ages<=216),
                       f"bmi-{sex_label}-z-who-2007-exp.xlsx",
                       'Month',
                       bmis,
                       ages)

        elif measure=="WFA":

            apply_file(ages<=120,
                       f"wfa_{sex_label}_0-to-5-years_zscores.xlsx",
                       'Month',
                       weights,
                       ages)

        elif measure=="WFH":

            apply_file(ages<=24,
                       f"wfl_{sex_label}_0-to-2-years_zscores.xlsx",
                       'Length',
                       weights,
                       heights)

            apply_file((ages>24)&(ages<=60),
                       f"wfh_{sex_label}_2-to-5-years_zscores.xlsx",
                       'Height',
                       weights,
                       heights)

    return z


############################################
# LOAD DATA
############################################

try:

    df=pd.read_excel(INPUT)
    df.columns=df.columns.str.strip()

    required=['sex','age_months','height_cm','weight_kg']

    for col in required:
        if col not in df.columns:
            messagebox.showerror("PediaZ v.01",f"Column '{col}' not found.")
            sys.exit()

    df['sex']=df['sex'].astype(str).str.lower().str.strip()
    df['age_months']=pd.to_numeric(df['age_months'],errors='coerce')
    df['height_cm']=pd.to_numeric(df['height_cm'],errors='coerce')
    df['weight_kg']=pd.to_numeric(df['weight_kg'],errors='coerce')

    df['BMI']=round(df['weight_kg']/((df['height_cm']/100)**2),2)

    metrics=['HFA','WFA','WFH','BFA']

    for m in metrics:
        df[f"{m}_z"] = compute_metric_vectorized(df,m)
        df[f"{m}_interp"] = df.apply(lambda row: interpret(row[f"{m}_z"], m), axis=1)

    df['age_group']=df['age_months'].apply(age_group)

    with pd.ExcelWriter(OUTPUT,engine='xlsxwriter') as writer:

        wb=writer.book
        header=wb.add_format({'bold':True,'bg_color':'#FFF2CC'})

        z_df = df.drop(columns=['age_group'])
        z_df.to_excel(writer,sheet_name='z_scores',index=False)

        ws=writer.sheets['z_scores']
        ws.freeze_panes(1,0)
        ws.set_row(0,None,header)
        auto_width(ws,z_df)

        ############################################
        # EXPANDED PREVALENCE
        ############################################

        for m in metrics:

            name=f"{m}_prevalence"
            interp_col=f"{m}_interp"
            z_col=f"{m}_z"

            ws_holder=[None]
            startrow=[0]
            last_table=[None]

            def write_block(title, sub_df):

                p_total = prevalence_table(sub_df,interp_col)

                if ws_holder[0] is None:
                    p_total.to_excel(writer,sheet_name=name,startrow=startrow[0]+3,index=False)
                    ws_holder[0] = writer.sheets[name]
                else:
                    p_total.to_excel(writer,sheet_name=name,startrow=startrow[0]+3,index=False)

                ws_holder[0].write(startrow[0],0,title,header)
                total, clean, pct = effective_sample_size(sub_df, interp_col)
                ws_holder[0].write(startrow[0]+1,0,f"Based on TOTAL sample: {total}")
                
                total, clean, pct = effective_sample_size(sub_df,interp_col)
                ws_holder[0].write(startrow[0]+2,0,f"Effective Sample Size including missing & flags: {clean}/{total} ({pct:.1f}%)")

                p_clean = prevalence_table_clean(sub_df,interp_col)
                clean_start = startrow[0]+len(p_total)+8
                p_clean.to_excel(writer,sheet_name=name,startrow=clean_start,index=False)
                ws_holder[0].write(startrow[0]+len(p_total)+6,0,f"Based on CLEAN sample (without missing and flags): {clean}")

                desc = clean_descriptives(sub_df, interp_col, z_col)
                desc.to_excel(writer,sheet_name=name,startrow=clean_start+len(p_clean)+2,index=False)

                startrow[0] += len(p_total)*2 + len(desc) + 18
                last_table[0] = p_total

            write_block("OVERALL — BOTH SEXES", df)
            write_block("BOYS", df[df['sex']=="m"])
            write_block("GIRLS", df[df['sex']=="f"])

            write_block("AGE < 61 MONTHS — BOTH", df[df['age_months']<61])
            write_block("AGE < 61 MONTHS — BOYS", df[(df['age_months']<61)&(df['sex']=="m")])
            write_block("AGE < 61 MONTHS — GIRLS", df[(df['age_months']<61)&(df['sex']=="f")])

            write_block("AGE ≥ 61 MONTHS — BOTH", df[df['age_months']>=61])
            write_block("AGE ≥ 61 MONTHS — BOYS", df[(df['age_months']>=61)&(df['sex']=="m")])
            write_block("AGE ≥ 61 MONTHS — GIRLS", df[(df['age_months']>=61)&(df['sex']=="f")])

            band = df[df['age_months']<=6]
            write_block("AGE ≤ 6 MONTHS — BOTH", band)
            write_block("AGE ≤ 6 MONTHS — BOYS", band[band['sex']=="m"])
            write_block("AGE ≤ 6 MONTHS — GIRLS", band[band['sex']=="f"])

            bins = list(range(0,217,12))

            for i in range(len(bins)-1):

                a,b=bins[i],bins[i+1]

                band = df[(df['age_months']>=a)&(df['age_months']<b)]

                write_block(f"AGE {a}-{b-1} MONTHS — BOTH", band)
                write_block(f"AGE {a}-{b-1} MONTHS — BOYS", band[band['sex']=="m"])
                write_block(f"AGE {a}-{b-1} MONTHS — GIRLS", band[band['sex']=="f"])

            if ws_holder[0] is not None:
                if last_table[0] is not None:
                    auto_width(ws_holder[0],last_table[0])

        ############################################
        # DATA QUALITY
        ############################################

        dq_rows=[]
        total=len(df)

        for m in metrics:

            interp_col=f"{m}_interp"

            miss=(df[interp_col]=="missing").sum()
            flags=((df[interp_col]=="lowflag")|(df[interp_col]=="highflag")).sum()
            valid=total-miss-flags

            dq_rows.append([
                m,total,
                miss,round(miss/total*100,1),
                flags,round(flags/total*100,1),
                valid,round(valid/total*100,1)
            ])

        dq=pd.DataFrame(dq_rows,columns=[
            "Metric","Total",
            "Missing_N","Missing_%",
            "Flags_N","Flags_%",
            "Valid_N","Valid_%"
        ])

        dq.to_excel(writer,sheet_name='data_quality',index=False)

        ws=writer.sheets['data_quality']
        ws.freeze_panes(1,0)
        ws.set_row(0,None,header)
        auto_width(ws,dq)

        ############################################
        # README
        ############################################

        readme=pd.DataFrame({
            "Information":[
                "Software name: PediaZ, an Abbreviation of Pediatric Z-score Calculator.",
                "Version: v.01",
                f"Analysis date: {datetime.now().strftime('%Y-%m-%d')}",
                "Z-scores were calculated using the World Health Organization (WHO) Child Growth Standards and the WHO 2007 references.",
                "All 95% confidence intervals (95% CIs) were calculated using both the Wald and Wilson methods.",
                "Flag definition: Based on WHO recommendations.",
                "Prevalence tables include both the total sample and the clean analytical sample, which excludes any missing data and flags.",
                "The effective sample size, calculated as the clean sample size divided by the total sample size, is reported at the top of each prevalence sheet.",
                "Use of this software is free; citation is recommended.",
                "Developer: Dr. Amin Sharifi, Ph.D. in Nutritional Sciences (Minor in Epidemiology); ORCID iD: http://orcid.org/0000-0002-3164-0479",
                "Email: a.sharifi1983@gmail.com"
            ]
        })

        readme.to_excel(writer,sheet_name='readme',index=False)
        auto_width(writer.sheets['readme'],readme)

    open_file = messagebox.askyesno(
        "PediaZ v.01",
        "Z-scores were calculated following WHO standards and saved in the same folder as the data.\n\nDo you want to open the output file?"
    )

    if open_file:

        if os.name=='nt':
            os.startfile(OUTPUT)
        elif sys.platform=='darwin':
            subprocess.call(['open',OUTPUT])
        else:
            subprocess.call(['xdg-open',OUTPUT])

except Exception as e:

    messagebox.showerror("PediaZ v.01",str(e))
