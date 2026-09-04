import argparse, json, os, sys, traceback
import pandas as pd
import core

def out(value): print(json.dumps(value, ensure_ascii=False), flush=True)
def row_dict(row): return {str(k): "" if v is None else str(v) for k,v in row.to_dict().items()}
def inspect_data(req):
    path=req["data_path"]
    ext=os.path.splitext(path)[1].lower()
    if ext==".csv":
        raw=None
        for enc in ("utf-8-sig","utf-8","cp932"):
            try: raw=pd.read_csv(path,dtype=str,encoding=enc); break
            except UnicodeDecodeError: pass
        if raw is None: raise ValueError("CSVの文字コードを判定できませんでした。")
        raw=raw.dropna(axis=1,how="all")
        raw.columns=core.deduplicate_columns([str(c).strip() for c in raw.columns])
    else:
        engine="openpyxl" if ext in (".xlsx",".xlsm") else "xlrd"
        raw=pd.read_excel(path,sheet_name=0,header=0,dtype=str,engine=engine).dropna(axis=1,how="all")
        raw.columns=core.deduplicate_columns([core.normalize_excel_column_name(c) for c in raw.columns])
    raw=raw.loc[:,[c for c in raw.columns if str(c).strip()]]
    raw=core.clean_dataframe_values(raw)
    raw=core.format_bank_account_columns(raw)
    raw=core.format_amount_columns(raw,enabled=req.get("format_amount_with_comma",True))
    original_count=len(raw); first=raw.columns[0]
    example_mask=raw[first].astype(str).str.contains("例",na=False)
    example_count=int(example_mask.sum()); candidates=raw[~example_mask].copy()
    reasons=[]; excluded=[]; included=[]
    mode=req.get("row_exclude_mode",core.DEFAULT_ROW_EXCLUDE_MODE); target=req.get("row_exclude_target_column_number",1)
    for _,row in candidates.iterrows():
        if core.should_exclude_row_by_mode(row,mode,target):
            values=list(row.values); reason="除外条件に一致"
            if mode=="selected_column_number_empty":
                idx=int(target)-1; reason=f"{row.index[idx]}が空欄" if 0<=idx<len(row.index) else "指定列が範囲外"
            else:
                empty=[str(row.index[i]) for i,v in enumerate(values) if core.is_empty_cell_for_row_exclusion(v) and (mode!="any_empty_except_first" or i>0)]
                if empty: reason="、".join(empty[:3])+("ほか" if len(empty)>3 else "")+"が空欄"
            excluded.append({"row":row_dict(row),"reason":reason})
        else: included.append(row_dict(row))
    return {"original_count":original_count,"example_count":example_count,"excluded_count":len(excluded),"included_count":len(included),"columns":[str(c) for c in raw.columns],"included_rows":included[:500],"excluded_rows":excluded[:500]}
def generate(req):
    df=core.prepare_dataframe(req["data_path"],format_amount_with_comma=req.get("format_amount_with_comma",True),row_exclude_mode=req.get("row_exclude_mode",core.DEFAULT_ROW_EXCLUDE_MODE),row_exclude_target_column_number=req.get("row_exclude_target_column_number",1))
    core.generate_documents_from_template(req["template_path"],df,req.get("filename_keys",[]),req["output_path"],req.get("add_serial_number",False),req.get("serial_digits",2))
    return {"success":True,"output_path":req["output_path"],"generated_count":len(df)}
def main():
    p=argparse.ArgumentParser(); p.add_argument("command",choices=["inspect","generate"]); args=p.parse_args()
    try:
        req=json.load(sys.stdin); out(inspect_data(req) if args.command=="inspect" else generate(req))
    except Exception as e:
        print(json.dumps({"error":str(e),"traceback":traceback.format_exc()},ensure_ascii=False),file=sys.stderr,flush=True); sys.exit(1)
if __name__=="__main__": main()
