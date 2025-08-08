# --- ampl_neos_xlwings.py ---
import xlwings as xw
import numpy as np
import re, random, string, xmlrpc.client as xmlrpclib
import time
import base64, gzip
from io import BytesIO


import re

def scan_model_keywords(model_text):
    sets, params, vars, displays = set(), dict(), dict(), dict()
    
    param_pattern = re.compile(r"^param\s+([a-zA-Z_][a-zA-Z0-9_]*)(\s*\{[^}]+\})?\s*([^\s=]?=)?[^;=]*;") 
    set_pattern = re.compile(r"^set\s+([a-zA-Z_][a-zA-Z0-9_]*)(\s*\{[^}]+\})?")
    var_pattern = re.compile( r"^var\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\{([^}]*)\}")

    for line in model_text.splitlines():
        line = line.strip()

        # param
        m = param_pattern.match(line)
        if m:
            if m.group(3) and m.group(3)!=">=" and m.group(3)!="<=":
                continue
            name = m.group(1)
            index = m.group(2)
            if index:
                index_sets = [s.strip() for s in index.strip("{} ").split(",")]
                params[name] = index_sets
                vars[name] = index_sets
            else:
                params[name] = []
                vars[name] = []
            continue

        # set
        m = set_pattern.match(line)
        if m:
            sets.add(m.group(1))
            continue

        m = var_pattern.match(line)
        if m:
            name = m.group(1)
            index_sets =  [match.group(1) for match in re.finditer(r"(?:^|,)\s*(?:[^\s]+?\s+in\s+)?([a-zA-Z_][a-zA-Z0-9_]*)",m.group(2))]
            vars[name] = index_sets
            continue

        if line.startswith("_display") or line.startswith("display") :
            for v in line.replace("_display", "").replace("display", "").split(","):
                v = v.strip().strip(";")
                if v in vars:
                    displays[v] = vars[v]
                else:
                    displays[v] = []

    return sets, params,  displays

def n2s(n):
    """Convert a number to a string with 4 decimal places, removing trailing zeros."""
    if isinstance(n, (int, float)):
        return f"{n:.6g}"#.rstrip("0").rstrip(".")
    return str(n)


def generate_ampl_data_from_excel(sheet, sets, params):
    dat = "data;\n\n"
    set_d = {}
    for name in sets:
        try:
            values = sheet.range(name).value
            values = np.array(values).flatten()
            values = [n2s(v) for v in values if v is not None]
            set_d[name] = values
            dat += f"set {name} := {' '.join(values)};\n\n"
        except:
            print(f"[WARN] set '{name}' not found in sheet")

    for name, index in params.items():
        try:
            if not index:
                value = sheet.range(name).value
                dat += f"param {name} := {n2s(value)};\n\n"
            elif len(index) == 1:
                idx_set = set_d.get(index[0], [])
                values = np.array(sheet.range(name).value).flatten()
                if len(values) != len(idx_set):
                    raise ValueError(f"Length mismatch for param {name} and set {index[0]}")
                dat += f"param {name} :=\n"
                for i, val in zip(idx_set, values):
                    dat += f"  {i} {n2s(val)}\n"
                dat += ";\n\n"
            elif len(index) == 2:
                row_set, col_set = index
                row_vals = set_d.get(row_set, [])
                col_vals = set_d.get(col_set, [])
                values = np.array(sheet.range(name).value)
                if values.shape != (len(row_vals), len(col_vals)):
                    raise ValueError(f"Shape mismatch for param {name} and sets {row_set}, {col_set}")
                dat += f"param {name} : {' '.join(col_vals)} :=\n"
                # dat += f"param {name} :=\n"
                
                for i, row in zip(row_vals, values):
                    # for j, val in zip(col_vals, row):
                    #     if val is not None:
                    #         dat += f"  [{i}, {j}] {n2s(val)}\n" 
                    dat += f"{i} {' '.join(map(n2s, row))}\n"
                dat += ";\n\n"
            else:
                print(f"[WARN] Unsupported param dimension > 2 for {name}")
        except Exception as e:
            print(f"[WARN] param '{name}' not found or failed to process: {e}")

    return dat


NEOS_HOST = "neos-server.org"
NEOS_PORT = 3333

def neos():
    return xmlrpclib.ServerProxy(f"https://{NEOS_HOST}:{NEOS_PORT}")


def encode_gzip(text: str) -> str:
    buf = BytesIO()
    with gzip.GzipFile(fileobj=buf, mode="wb") as gz:
        gz.write(text.encode("utf-8"))
    compressed_bytes = buf.getvalue()
    return base64.b64encode(compressed_bytes).decode("utf-8")

def wrap_string(s, width=80):
    return "\n".join(s[i:i+width] for i in range(0, len(s), width))

def submit_ampl_job(email, model_text, category, solver, data_text):
    email = (
        "".join(random.choice(string.ascii_letters) for _ in range(5))
        + "@"
        + "".join(random.choice(string.ascii_letters) for _ in range(5))
        + ".com"
    ) if email=='rndom' else email

    xml = f"""<document>
  <category>{category}</category>
  <solver>{solver}</solver>
  <inputMethod>AMPL</inputMethod>
  <client>XNEOS</client>
  <priority>long</priority>
  <email>{email}</email>
    <model><![CDATA[
{model_text}
end;
    ]]></model>
    <data><base64>
{wrap_string(encode_gzip(data_text))}
    </base64></data>
<commands><![CDATA[
# Nothing here; commands are in model file
]]></commands>

<comments><![CDATA[]]></comments>
</document>
"""
    job_number, password = neos().submitJob(xml)
    return job_number, password

def neo_job_done(job_id, password):
    return neos().getJobStatus(job_id, password) == "Done"

def neos_update(sheet_name, model_text, job_id, password):
    sheet = xw.Book.caller().sheets[sheet_name]
    job_id=int(job_id)
    if len(model_text)<30:
        model_text=sheet.range(model_text).value
    try:
        _, _,  displays = scan_model_keywords(model_text)
    except Exception:
        print( f"Failed to parse model {job_id} ")
        return False
    if neos().getJobStatus(job_id, password) != "Done":
        print(f"Job {job_id} is not done yet")
        return False
    result = neos().getFinalResults(job_id, password)
    result_text = result.data.decode("utf-8")

    segments = re.split(r"_display.*", result_text)
    if len(segments) < 2 or len(segments[0])<20 or "Objective" not in segments[0]:
        # if len(result_text.splitlines())>10:
            print (f"No results found {job_id}\n" + result_text)
            return False
    
    index_sets={}
    def get_set_map(set_name):
        try:
            return index_sets[set_name]
        except KeyError:
            values={}
            try:
                # values= [n2s(x) for x in sheet.range(set_name).value ]
                values={n2s(k): i for i, k in enumerate(sheet.range(set_name).value)}
            except:
                pass
        index_sets[set_name] = values
        return values

    def write_back(var, val):
        try:
            if sheet.range(var).columns.count==1 and sheet.range(var).rows.count==len(val) and len(val)>1 and not isinstance(val[0], (list, np.ndarray)):
                val = np.array(val).reshape(-1, 1)
            sheet.range(var).value = val
        except:
            print(f"[WARN] Cannot write to {var} in sheet")

    for display_text in segments[1:]:
        lines = [l.strip() for l in display_text.strip().splitlines() if l.strip()]
        if not lines:
            continue
        var_name = lines[0] 
        dim_sets = displays.get(var_name,[])
        if not dim_sets:
            write_back(var_name, lines[1])
            continue

        try:
            values = sheet.range(var_name).value
        except:
            print(f"[WARN] Cannot find {var_name} from sheet")
            continue

        index_maps = [get_set_map(s) for s in dim_sets]

        if len(values)!= len(index_maps[0]):
            index_maps = index_maps[::-1]
    
        for i,line in enumerate(lines[1:]):
            try:
                new_value = line.split(',')
                if len(new_value) == 1:
                    values[i] = float(new_value[0])
                elif len(new_value) == 2:
                    values[index_maps[0][new_value[0]]]= float(new_value[1])
                elif len(new_value) == 3:
                    values[index_maps[0][new_value[0]]][index_maps[1][new_value[1]]] = float(new_value[2])
                else:
                    print(f"[WARN] Unexpected line format for {var_name}: {line}")
            except Exception as e:
                print(f"[WARN] Failed to parse line for {var_name}: {line} - {e}")

        write_back(var_name, values)

    # print( f" written  {job_id}")
    return True

def neos_kill(job_number: int, password: str) -> bool:
    try:
        result = neos().killJob(int(job_number), password, '')
        return result
    except Exception as e:
        return False
def neos_check(job_id, password, max_wait):
    
    start_time = time.time()
    try:
        while True:
            time.sleep(1)
            result = neos().getJobStatus(job_id, password)
            if result == "Done" or result == "Failed":
                return f'{result} {time.strftime("%H:%M:%S")}'
            if (time.time() - start_time) > max_wait+120:
                neos_kill(job_id, password)
                if(time.time() - start_time) > max_wait+180:
                    return f"❌ Timeout after {max_wait} seconds, job killed"
    except Exception as e:
        return f"❌ {e}"


def submit_and_monitor(sheet,  email, model, category, solver):
    sets, params,_ = scan_model_keywords(model)
    data = generate_ampl_data_from_excel(sheet, sets, params)
    job_id, password = submit_ampl_job( email, model, category, solver, data)
    return job_id, password


# if __name__ == "__main__":
#     xw.serve()
