import xlwings as xw
from xneos import neos_check, submit_and_monitor, neos_update, neos_kill, neo_job_done  


@xw.func(async_mode="threading")
def check_neos(job_id, password, max_wait=600):
    return neos_check(job_id, password, max_wait=max_wait)


@xw.sub
def kill_neos(job_id, password):
    return neos_kill(job_id, password)


@xw.func
def solve(sht_name, email, model, category="milp", solver="CPLEX"):
    return submit_and_monitor(
        xw.Book.caller().sheets[sht_name], email, model, category, solver
    )


@xw.func
def update_neos_result(sheet_name, model_text, job_id, password):
    return neos_update(sheet_name, model_text, job_id, password)

@xw.func
def job_done( job_id, password):
    return neo_job_done(job_id, password)


if __name__ == "__main__":
    xw.serve()
