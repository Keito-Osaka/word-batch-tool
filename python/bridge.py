import argparse, json, os, sys, traceback
import pandas as pd
import core

def write_json(stream, value):
    data=(json.dumps(value,ensure_ascii=False)+"\n").encode("utf-8")
    stream.buffer.write(data); stream.buffer.flush()

def read_request():
    raw=sys.stdin.buffer.read()
    if not raw: raise ValueError("入力データが渡されていません。")
    return json.loads(raw.decode("utf-8"))

def row_dict(row):
    return {str(k): "" if v is None else str(v) for k,v in row.to_dict().items()}

def validate_file(path,label):
    if not path: raise ValueError(f"{label}が指定されていません。")
    if not os.path.isfile(path): raise FileNotFoundError(f"{label}が見つかりません。\n\n{path}")

def load_raw(path):
    validate_file(path,"置換データファイル")
    ext=os.path.splitext(path)[1].lower()
    if ext==".csv":
        last=None
        for enc in ("utf-8-sig","utf-8","cp932"):
            try: df=pd.read_csv(path,dtype=str,encoding=enc); break
            except UnicodeDecodeError as e: last=e
        else: raise ValueError("CSVの文字コードを判定できませんでした。") from last
        df=df.dropna(axis=1,how="all")
        df.columns=core.deduplicate_columns([str(c).strip() for c in df.columns])
    elif ext in (".xlsx",".xlsm",".xls"):
        engine="openpyxl" if ext in (".xlsx",".xlsm") else "xlrd"
        df=pd.read_excel(path,sheet_name=0,header=0,dtype=str,engine=engine).dropna(axis=1,how="all")
        df.columns=core.deduplicate_columns([core.normalize_excel_column_name(c) for c in df.columns])
    else: raise ValueError("対応していない形式です。CSVまたはExcelを指定してください。")
    df=df.loc[:,[c for c in df.columns if str(c).strip()]]
    if len(df.columns)==0: raise ValueError("使用できる列がありません。")
    return df

def exclusion_reason(row,mode,target):
    if mode=="selected_column_number_empty":
        idx=int(target)-1
        return f"「{row.index[idx]}」が空欄" if 0<=idx<len(row.index) else "指定列が範囲外"
    empty=[]
    for i,v in enumerate(row.values):
        if mode=="any_empty_except_first" and i==0: continue
        if core.is_empty_cell_for_row_exclusion(v): empty.append(str(row.index[i]))
    if not empty: return "現在の除外条件に一致"
    names="、".join(f"「{x}」" for x in empty[:3])+("ほか" if len(empty)>3 else "")
    return names+("がすべて空欄" if mode=="all_empty_except_first" else "が空欄")

def inspect_data(req):
    df=load_raw(req.get("data_path",""))
    df=core.clean_dataframe_values(df)
    df=core.format_bank_account_columns(df)
    df=core.format_amount_columns(df,enabled=req.get("format_amount_with_comma",True),include_keywords=req.get("amount_include_keywords"),exclude_keywords=req.get("amount_exclude_keywords"))
    original=len(df); first=df.columns[0]
    example_mask=df[first].astype(str).str.contains("例",na=False)
    examples=int(example_mask.sum()); candidates=df[~example_mask].copy()
    mode=req.get("row_exclude_mode",core.DEFAULT_ROW_EXCLUDE_MODE)
    target=req.get("row_exclude_target_column_number",core.DEFAULT_ROW_EXCLUDE_TARGET_COLUMN_NUMBER)
    included=[]; excluded=[]
    for idx,row in candidates.iterrows():
        item=row_dict(row)
        if core.should_exclude_row_by_mode(row,mode,target):
            excluded.append({"source_row_number":int(idx)+2,"row":item,"reason":exclusion_reason(row,mode,target)})
        else: included.append(item)
    return {"original_count":original,"example_count":examples,"excluded_count":len(excluded),"included_count":len(included),"columns":[str(c) for c in df.columns],"included_rows":included[:500],"excluded_rows":excluded[:500]}

def generate(req):
    template=req.get("template_path",""); data=req.get("data_path",""); output=req.get("output_path","")
    validate_file(template,"テンプレートWord"); validate_file(data,"置換データファイル")
    if not output: raise ValueError("出力先が指定されていません。")
    os.makedirs(output,exist_ok=True)
    df=core.prepare_dataframe(data,format_amount_with_comma=req.get("format_amount_with_comma",True),amount_include_keywords=req.get("amount_include_keywords"),amount_exclude_keywords=req.get("amount_exclude_keywords"),row_exclude_mode=req.get("row_exclude_mode",core.DEFAULT_ROW_EXCLUDE_MODE),row_exclude_target_column_number=req.get("row_exclude_target_column_number",core.DEFAULT_ROW_EXCLUDE_TARGET_COLUMN_NUMBER))
    core.generate_documents_from_template(template,df,req.get("filename_keys",[]),output,req.get("add_serial_number",False),req.get("serial_digits",2))
    return {"success":True,"output_path":output,"generated_count":len(df)}

def main():
    parser=argparse.ArgumentParser(); parser.add_argument("command",choices=["inspect","generate"]); args=parser.parse_args()
    try:
        req=read_request(); write_json(sys.stdout,inspect_data(req) if args.command=="inspect" else generate(req))
    except Exception as e:
        write_json(sys.stderr,{"error":str(e),"traceback":traceback.format_exc()}); sys.exit(1)
if __name__=="__main__": main()
