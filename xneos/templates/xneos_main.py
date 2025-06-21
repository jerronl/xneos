import xlwings as xw
from xneos import neos_check, submit_and_monitor, neos_update


@xw.func(async_mode="threading")
def check_neos(job_id, password):
    return neos_check(job_id, password)


@xw.sub
def solve(sht_name, email, model="model_text", category="milp", solver="CPLEX"):
    submit_and_monitor(xw.Book.caller().sheets[sht_name], email, model, category, solver)


@xw.sub
def update_neos_result(sheet_name, model_text="model_text"):
    neos_update(sheet_name, model_text)


if __name__ == "__main__":
    xw.serve()
