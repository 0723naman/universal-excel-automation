
import argparse, pandas as pd, numpy as np, os
from pathlib import Path

def detect_column_types(df):
    types = {}
    n = len(df)
    for col in df.columns:
        ser = df[col]

        # Numeric detection
        num = pd.to_numeric(ser, errors='coerce')
        num_ratio = num.notna().sum() / max(1,n)

        # Date detection
        dt = pd.to_datetime(ser, errors='coerce')
        dt_ratio = dt.notna().sum() / max(1,n)

        # Possible ID detection
        unique_ratio = ser.nunique(dropna=True) / max(1,n)

        if dt_ratio >= 0.6:
            types[col] = "date"
        elif num_ratio >= 0.6:
            types[col] = "numeric"
        elif unique_ratio > 0.85 and n > 20:
            types[col] = "id"
        else:
            types[col] = "categorical"
    return types

def clean_df(df):
    for c in df.select_dtypes(include=["object"]).columns:
        df[c] = df[c].astype(str).str.strip().replace({"nan":None})
    return df

def numeric_summary(df, numeric_cols):
    out=[]
    for c in numeric_cols:
        ser = pd.to_numeric(df[c], errors='coerce')
        out.append({
            "Column":c,
            "Count":int(ser.count()),
            "Missing":int(ser.isna().sum()),
            "Sum":float(ser.sum(skipna=True)),
            "Mean":float(ser.mean(skipna=True)) if ser.count()>0 else None,
            "Median":float(ser.median(skipna=True)) if ser.count()>0 else None,
            "Std":float(ser.std(skipna=True)) if ser.count()>1 else None,
            "Min":float(ser.min(skipna=True)) if ser.count()>0 else None,
            "Max":float(ser.max(skipna=True)) if ser.count()>0 else None
        })
    return pd.DataFrame(out)

def categorical_summary(df, cat_cols):
    result={}
    for c in cat_cols:
        vc=df[c].fillna("Missing").astype(str).value_counts().head(10)
        result[c]=vc.reset_index().rename(columns={"index":"Value",c:"Count"})
    return result

def date_summary(df, date_cols, numeric_cols):
    """
    Build monthly/year-month summaries for each detected date column.
    - If numeric_cols exists: sum numeric columns per month.
    - If no numeric cols: produce a simple count of rows per month.
    Safely handles edge cases to avoid 'No objects to concatenate'.
    """
    frames = {}
    for d in date_cols:
        dt = pd.to_datetime(df[d], errors='coerce', infer_datetime_format=True)
        if dt.isna().all():
            # nothing to summarize for this column
            continue

        df['_tmp_month'] = dt.dt.to_period('M').astype(str)

        # if we have numeric columns, try to aggregate them by sum
        if numeric_cols:
            # build per-column aggregation functions safely
            agg_map = {}
            for col in numeric_cols:
                # use lambda with default argument to avoid late-binding problem
                agg_map[col] = (lambda x, col=col: pd.to_numeric(x, errors='coerce').sum())

            try:
                monthly = df.groupby('_tmp_month').agg(agg_map)
                monthly = monthly.reset_index().rename(columns={'_tmp_month': 'YearMonth'})
                frames[d] = monthly
            except Exception:
                # fallback: produce counts per month if numeric aggregation fails
                monthly = df.groupby('_tmp_month').size().reset_index(name='Rows').rename(columns={'_tmp_month': 'YearMonth'})
                frames[d] = monthly
        else:
            # no numeric columns available -> produce simple counts per month
            monthly = df.groupby('_tmp_month').size().reset_index(name='Rows').rename(columns={'_tmp_month': 'YearMonth'})
            frames[d] = monthly

    # cleanup temporary column if present
    if '_tmp_month' in df.columns:
        try:
            df.drop(columns=['_tmp_month'], inplace=True)
        except Exception:
            pass
    return frames


def find_outliers(df,numeric_cols):
    out={}
    for c in numeric_cols:
        ser=pd.to_numeric(df[c],errors="coerce").dropna()
        if ser.empty: continue
        q1,q3=ser.quantile(0.25),ser.quantile(0.75)
        iqr=q3-q1
        lower,upper=q1-1.5*iqr,q3+1.5*iqr
        outs=ser[(ser<lower)|(ser>upper)]
        out[c]={"count":outs.count(),"examples":outs.head(5).tolist()}
    return out

def candidate_ids(df):
    out=[]
    n=len(df)
    for c in df.columns:
        unique=df[c].nunique(dropna=True)/max(1,n)
        non_null=df[c].notna().sum()
        if unique>0.85 and non_null>max(10,n*0.5):
            out.append(c)
    return out

def generate_report(input_path, output_path):
    df=pd.read_excel(input_path, sheet_name=0)
    df_clean=df.copy()

    df_clean=clean_df(df_clean)
    types=detect_column_types(df_clean)

    numeric_cols=[c for c,t in types.items() if t=="numeric"]
    cat_cols=[c for c,t in types.items() if t=="categorical"]
    date_cols=[c for c,t in types.items() if t=="date"]
    id_cols=[c for c,t in types.items() if t=="id"] + candidate_ids(df_clean)

    num_sum=numeric_summary(df_clean,numeric_cols) if numeric_cols else pd.DataFrame()
    cat_sum=categorical_summary(df_clean,cat_cols)
    date_sum=date_summary(df_clean,date_cols,numeric_cols)
    outliers=find_outliers(df_clean,numeric_cols)

    missing=df_clean.isna().sum().reset_index().rename(columns={"index":"Column",0:"MissingCount"})

    with pd.ExcelWriter(output_path,engine="openpyxl") as writer:
        df.to_excel(writer,"RawData",index=False)
        num_sum.to_excel(writer,"NumericSummary",index=False)
        missing.to_excel(writer,"MissingValues",index=False)

        for c,table in cat_sum.items():
            sheet=c[:28]+"_Top"
            table.to_excel(writer,sheet_name=sheet,index=False)

        for c,table in date_sum.items():
            sheet=c[:28]+"_Monthly"
            table.to_excel(writer,sheet_name=sheet,index=False)

        pd.DataFrame([{"Column":k,"Outliers":v["count"],"Examples":str(v["examples"])} for k,v in outliers.items()]).to_excel(writer,"Outliers",index=False)
        pd.DataFrame({"CandidateID":id_cols}).to_excel(writer,"ID_Candidates",index=False)

    print("Report Generated:", output_path)

def main():
    import argparse
    p=argparse.ArgumentParser()
    p.add_argument("--input","-i")
    p.add_argument("--output","-o")
    args=p.parse_args()

    os.makedirs("reports",exist_ok=True)

    if args.input:
        inp=Path(args.input)
        out=args.output if args.output else f"reports/{inp.stem}_report.xlsx"
        generate_report(inp,out)
    else:
        for fp in Path("data").glob("*.xlsx"):
            generate_report(fp,f"reports/{fp.stem}_report.xlsx")

if __name__=="__main__":
    main()
