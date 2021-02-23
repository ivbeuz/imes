# -*- coding: utf-8 -*-
"""
Created on Thu Mar  7 16:15:29 2019

@author: JanGr
"""

from pyomo.environ import SolverFactory
from pyomo.opt.parallel import SolverManagerFactory
import pandas as pd


def RunningLocalServer(model, solver_name, time_limit):
    solver = SolverFactory(solver_name)
    solver.options['timelimit'] = time_limit
    solver.options['mipgap'] = 0.001  # Set optimality gap to 2.5%
    #kept at 7% as Julie did, to align with her results; now trying 2,5%

# =============================================================================
#     analysis = pd.DataFrame()
#     ub = pd.DataFrame()
#     lb = pd.DataFrame()
#     time = pd.DataFrame()
#     gap = pd.DataFrame()
#
#     time = time.append(pd.Series(solver.objective_bounds.data),ignore_index=True)
#     lb = lb.append(pd.Series(solver.s),ignore_index=True)
#     ub = ub.append(pd.Series(solver.Incumbent), ignore_index=True)
#     gap = gap.append(pd.Series(solver.Gap), ignore_index=True)
#     analysis = [time,lb,ub,gap]
# =============================================================================

    output = solver.solve(model, tee=True)

    return output  # , analysis


def RunningOnlineServer(model, server_name, solver_name):
    solver_manager = SolverManagerFactory(server_name)
    return solver_manager.solve(model, opt=solver_name)


def PrintResult(results, objective):
    results.write()

    # Print the status of the solved problem
    print("Status = %s" % results.solver.termination_condition)

    # Print the cost of the objective
    print("Objective = %f" % objective(objective))
