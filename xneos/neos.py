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
    
    param_pattern = re.compile(r"^param\s+([a-zA-Z_][a-zA-Z0-9_]*)(\s*\{[^}]+\})?\s*(?![:=])[^;]*;?")
    set_pattern = re.compile(r"^set\s+([a-zA-Z_][a-zA-Z0-9_]*)(\s*\{[^}]+\})?")
    var_pattern = re.compile( r"^var\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\{([^}]*)\}")

    for line in model_text.splitlines():
        line = line.strip()

        # param
        m = param_pattern.match(line)
        if m:
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
        return f"{n:.4f}".rstrip("0").rstrip(".")
    return str(n)


def generate_ampl_data_from_excel(sheet, sets, params):
    dat = ""
    set_d = {}
    for name in sets:
        try:
            values = sheet.range(name).value
            values = np.array(values).flatten()
            values = [n2s(v) for v in values if v is not None]
            set_d[name] = values
            dat += f"set {name} := {' '.join(values)};\n"
        except:
            print(f"[WARN] set '{name}' not found in sheet")

    for name, index in params.items():
        try:
            if not index:
                value = sheet.range(name).value
                dat += f"param {name} := {n2s(value)};\n"
            elif len(index) == 1:
                idx_set = set_d.get(index[0], [])
                values = np.array(sheet.range(name).value).flatten()
                if len(values) != len(idx_set):
                    raise ValueError(f"Length mismatch for param {name} and set {index[0]}")
                dat += f"param {name} :=\n"
                for i, val in zip(idx_set, values):
                    dat += f"{i} {n2s(val)}\n"
                dat += ";\n"
            elif len(index) == 2:
                row_set, col_set = index
                row_vals = set_d.get(row_set, [])
                col_vals = set_d.get(col_set, [])
                values = np.array(sheet.range(name).value)
                if values.shape != (len(row_vals), len(col_vals)):
                    raise ValueError(f"Shape mismatch for param {name} and sets {row_set}, {col_set}")
                dat += f"param {name} : {' '.join(col_vals)} :=\n"
                for i, row in zip(row_vals, values):
                    dat += f"{i} {' '.join(map(n2s, row))}\n"
                dat += ";\n"
            else:
                print(f"[WARN] Unsupported param dimension > 2 for {name}")
        except:
            print(f"[WARN] param '{name}' not found or failed to process")

    return dat


NEOS_HOST = "neos-server.org"
NEOS_PORT = 3333
neos = xmlrpclib.ServerProxy(f"https://{NEOS_HOST}:{NEOS_PORT}")


def encode_gzip(text: str) -> str:
    buf = BytesIO()
    with gzip.GzipFile(fileobj=buf, mode="wb") as gz:
        gz.write(text.encode("utf-8"))
    compressed_bytes = buf.getvalue()
    return base64.b64encode(compressed_bytes).decode("utf-8")


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
{encode_gzip(data_text)}
    </base64></data>
<commands><![CDATA[# Nothing here; commands are in model file

]]></commands>

<comments><![CDATA[]]></comments>
</document>
"""
    job_number, password = neos.submitJob(xml)
    return job_number, password

def neos_update(sheet_name, model_text):
    sheet = xw.Book.caller().sheets[sheet_name]
    try:
        _, _,  displays = scan_model_keywords(sheet.range(model_text).value)
    except Exception:
        sheet.range("status").value = "⚠️ Failed to parse model"
        return
    job_id = int(sheet.range("jobid").value)
    password = sheet.range("pwd").value
    result = neos.getFinalResults(job_id, password)
    result_text = result.data.decode("utf-8")

    index_sets={}
    def get_set_map(set_name):
        try:
            return index_sets[set_name]
        except KeyError:
            values={}
            try:
                values= [str(x).strip() for x in sheet.range(set_name).value if x]
                values={k: i for i, k in enumerate(values)}
            except:
                pass
        index_sets[set_name] = values
        return values

    def write_back(var, val):
        try:
            sheet.range(var).value = val
        except:
            print(f"[WARN] Cannot write to {var} in sheet")

    segments = re.split(r"_display.*", result_text)
    if len(segments) < 2:
        sheet.range("status").value = "❌ No results found\n" + result_text
        return

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

    sheet.range("status").value = f"🟢 written"

def neos_check(job_id, password):
    try:
        while True:
            time.sleep(1)
            result = neos.getJobStatus(job_id, password)
            if result == "Done" or result == "Failed":
                return f'{time.strftime("%H:%M:%S")} {result}'
    except Exception as e:
        return f"❌ {e}"


def submit_and_monitor(sheet,  email, model, category, solver):
    model = sheet.range(model).value
    sets, params,_ = scan_model_keywords(model)
    data = generate_ampl_data_from_excel(sheet, sets, params)
    job_id, password = submit_ampl_job( email, model, category, solver, data)
    stime = time.strftime("%H:%M:%S")
    sheet.range("status").value = f"{stime} submitted"
    sheet.range("jobid").value = job_id
    sheet.range("pwd").value = password


# if __name__ == "__main__":
#     xw.serve()
