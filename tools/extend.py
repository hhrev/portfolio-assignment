"""Oil, implementation, integrity checks, Summary and figure additions."""
import impl
import oil
import reporting


def run(wb, C, perf):
    oilref = oil.run(wb)
    I = impl.inputs(wb)
    tc, plan, frank, eqd = impl.portfolio(wb, I)
    integ = impl.validation(wb)
    chk = impl.checks(wb, C, perf, oilref, tc, plan, frank, eqd, integ)
    reporting.summary(wb, C, perf, oilref, tc, plan, frank, eqd, chk)
    return reporting.figures(wb, C, perf)
