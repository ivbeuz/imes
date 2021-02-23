# -*- coding: utf-8 -*-
"""
Created on Thu Nov  7 08:42:18 2019

@author: IvB, latest update 23-2-2021
"""
import pandas as pd
# import numpy as np #for writing files, not needed if using:
import xlsxwriter

from math_prog_imes import Model
from pyomo_helper_imes import RunningLocalServer, PrintResult

###################IMPORT DATA#################################################
# Retrieving dataframes from MS access DB and creating dictionaries
db_file = './data/IMES_21node_case_data.xlsx'
db_file = pd.ExcelFile(db_file)

# Read data from the excel datafiles
# This only needs to occur once for the entire analysis if the data are...
# ... changed manually inside these tables.
print("Connecting to database")
df_conversion_efficiencies = db_file.parse('ConversionEfficiencies')
df_conversion_units = db_file.parse('ConversionUnits')
df_demand = db_file.parse('Demand')
df_locations = db_file.parse('Locations')
df_max_converted = db_file.parse('MaxConverted')
df_max_flow = db_file.parse('MaxFlow')
df_network = db_file.parse('Network')
df_storage_units = db_file.parse('StorageUnits')
df_supply = db_file.parse('Supply')
df_supply_units = db_file.parse('SupplyUnits')
# 2-year time periods
df_time_periods = db_file.parse('TimePeriods')
# 6-year time periods
#df_time_periods = RetrieveTableFromMSAccessDB(db_file, "TPNEW")

# =============================================================================
# Preparing design decision vector  ### Model
# =============================================================================
# Create empty dataframes for the investments made, to create the ...
# ...decision vector
decision_vector = pd.DataFrame()
supply_investments = pd.DataFrame()
supply_costs = pd.DataFrame()
converter_investments = pd.DataFrame()
converter_costs = pd.DataFrame()
line_investments = pd.DataFrame()
line_costs = pd.DataFrame()
storage_investments = pd.DataFrame()
storage_costs = pd.DataFrame()

# =============================================================================
# Making the SetsLists
# =============================================================================
locations = list(df_locations.Locations)
energy_carriers = list(["Electricity", "Gas", "Heat"])
energy_converters = list(["CHP", "HP", "P2G"])
supply_types = list(["Solar", "Wind"])  # add gas here? not needed?
edges = []
arcs = []
time_periods = list(df_time_periods.TimeSlot)

# Making the adges and arce from the network costsedges
network_costs_t0 = df_network.set_index(['Type', 'LocationFrom', 'LocationTo'])\
    .to_dict()['Costs']

for x in network_costs_t0:
    if(x[0] == 'Electricity'):
        edge = (x[1], x[2])
        edges.append(edge)
        arc1 = (x[1], x[2])
        arc2 = (x[2], x[1])
        arcs.append(arc1)
        arcs.append(arc2)

# =============================================================================
# Making Parameters
# =============================================================================
# Many of these parameters change according to the dataframe values, so need...
# ... to be calculated each run within the model

# Network Parameters
# The costs are also updated taking into account the development rate...
# ...and the discount rate(which is equal to 0.045).
# There is not always a development rate; discount_rate = 0.045
discount_rate = 0.045
# updated to specify all standard development rates, network dvt rates all 0.
dev_rate_conv_CHP = 0.0
dev_rate_conv_HP = 0.01
dev_rate_conv_P2G = 0.079
dev_rate_storage_elec = 0.05
dev_rate_storage_gas = 0.0
dev_rate_storage_heat = 0.016
dev_rate_supply_solar = 0.05
dev_rate_supply_wind = 0.022

starting_time_period = 2018  # t0
network_costs = {}
for time_period in time_periods:
    for location1 in locations:
        for location2 in locations:
            network_costs['Electricity', location1, location2, time_period] = \
                (network_costs_t0['Electricity', location1, location2]) /\
                (1+discount_rate)**(int(time_period)-starting_time_period)
            network_costs['Gas', location1, location2, time_period] = \
                (network_costs_t0['Gas', location1, location2]) /\
                (1+discount_rate)**(int(time_period)-starting_time_period)
            network_costs['Heat', location1, location2, time_period] = \
                ((network_costs_t0['Heat', location1, location2])) /\
                (1+discount_rate)**(int(time_period)-starting_time_period)

# opening hier om brownfield te definieren, zou dan wel veel uitgebreider moeten, zie code Noortje
earlier_line_investment_made = {}
for location1 in locations:
    for location2 in locations:
        earlier_line_investment_made["Electricity", location1, location2] = 0
        earlier_line_investment_made["Gas", location1, location2] = 0
        earlier_line_investment_made["Heat", location1, location2] = 0

# Maximum flow on a line
max_flow_line = df_max_flow.set_index(['Type'])\
    .to_dict()['MaxFlowLine']

loss_factor = df_max_flow.set_index(['Type']).to_dict()['LossFactor']

# Supply Parameters
# supply investement costs
supply_investment_costs_t0 = df_supply_units.set_index(['SupplyType'])\
    .to_dict()['Costs']
supply_investment_costs = {}
for time_period in time_periods:
    for x in supply_investment_costs_t0:
        if(x == 'Solar'):
            supply_investment_costs[x, time_period] = \
                (supply_investment_costs_t0[x])/(1+discount_rate +
                                                 dev_rate_supply_solar)**(int(time_period)-starting_time_period)
        else:  # Wind
            supply_investment_costs[x, time_period] = \
                (supply_investment_costs_t0[x])/(1+discount_rate +
                                                 dev_rate_supply_wind)**(int(time_period)-starting_time_period)

# Supply external factor added below, with other scenarios for gas. 

# Maximum energy supplied
max_energy_supplied = df_supply_units.set_index(['Type', 'SupplyType'])\
    .to_dict()['MaxSupply']

# Converter Parameters
# Conversion unit costs
converter_investment_costs_t0 = df_conversion_units.set_index(['Conversion'])\
    .to_dict()['Costs']
converter_investment_costs = {}
for time_period in time_periods:
    for x in converter_investment_costs_t0:
        if(x == 'P2G'):
            converter_investment_costs[x, time_period] = \
                (converter_investment_costs_t0[x])/(1+discount_rate +
                                                    dev_rate_conv_P2G)**(int(time_period)-starting_time_period)
        elif(x == 'HP'):
            converter_investment_costs[x, time_period] = \
                (converter_investment_costs_t0[x])/(1+discount_rate +
                                                    dev_rate_conv_HP)**(int(time_period)-starting_time_period)
        else:
            converter_investment_costs[x, time_period] = \
                (converter_investment_costs_t0[x])/(1+discount_rate)\
                ** (int(time_period)-starting_time_period)

# Maximum energy converted
max_converted = df_max_converted.set_index(['Type', 'ConversionUnit'])\
    .to_dict()['MaxConverted']

# conversion efficiencies
conversion_efficiencies = \
    df_conversion_efficiencies.set_index(['Type1',
                                          'Type2',
                                          'ConversionType'])\
    .to_dict()['Efficiency']

# Storage Parameters
storage_investment_costs_t0 = df_storage_units.set_index(['StorageType'])\
    .to_dict()['Costs']
storage_investment_costs = {}
for time_period in time_periods:
    for x in storage_investment_costs_t0:
        if(x == 'Electricity'):
            storage_investment_costs[x, time_period] = \
                (storage_investment_costs_t0[x])/(1+discount_rate +
                                                  dev_rate_storage_elec)**(int(time_period)-starting_time_period)
        if(x == 'Heat'):
            storage_investment_costs[x, time_period] = \
                (storage_investment_costs_t0[x])/(1+discount_rate +
                                                  dev_rate_storage_heat)**(int(time_period)-starting_time_period)
        else:
            storage_investment_costs[x, time_period] = \
                (storage_investment_costs_t0[x])/(1+discount_rate)\
                ** (int(time_period)-starting_time_period)

# Storage Losses
storage_losses = df_storage_units.set_index(['StorageType']).\
    to_dict()['StockDrain']

# Minimum storage level
min_stored = df_storage_units.set_index(['StorageType']).\
    to_dict()['MinStorageLevel']

# Maximum storage level
max_stored = df_storage_units.set_index(['StorageType']).\
    to_dict()['MaxStorageLevel']

# Demand and given supply at each location
# The demand of each location and energy type
# Each location has a certain demand percentage of the total demand. \
# This total demand changes over the years so the demand at each location...\
# ... changes over the years as well.

demand = df_demand.set_index(['Type', 'Location', 'TimePeriod']).\
    to_dict()['Demand']

# Amount of supply at the nodes - adjust CO2 reduction scenarios here
supply_scenario_read = df_supply.set_index(['iProduct',
                                            'iLocation',
                                            'iTimeSlot'])\
    .to_dict()['95%_red']

amount_given_only_gas = { key:value for (key,value) in supply_scenario_read.items() if key[0] == 'Gas'}

amount_given = {}
for x in amount_given_only_gas:
    amount_given[x] = amount_given_only_gas[x]
    amount_given['Electricity', x[1], x[2]] = 0
    amount_given['Heat', x[1], x[2]] = 0

# Amount of PV or wind supply at the nodes - adjust weather scenarios here - FIRST ATTEMPT, NOT WORKING YET 05082020
#supply_external_factor - currently limited to weather, bc gas supply adjustments already captured in other value
supply_external_factor = { key:value for (key,value) in supply_scenario_read.items() if key[0] in ('Solar', 'Wind')}


# If there are earlier investements made. In this way you can also take into account what already has been built (this is actually used for the problem where each year is solved seperately but it does not have any effect on the solution when you want to solve all periods together)
earlier_supply_investment_made = {}
earlier_converter_investment_made = {}
earlier_storage_investment_made = {}
for location in locations:
    earlier_supply_investment_made["Solar", location] = 0
    earlier_supply_investment_made["Wind", location] = 0
    earlier_converter_investment_made["CHP", location] = 0
    earlier_converter_investment_made["P2G", location] = 0
    earlier_converter_investment_made["HP", location] = 0
    earlier_storage_investment_made["Electricity", location] = 0
    earlier_storage_investment_made["Gas", location] = 0
    earlier_storage_investment_made["Heat", location] = 0

# =============================================================================
# Solve Model
# =============================================================================
print("Initializing model")
# Initialize sets, parameters, and variables
model = Model()
model.InitializeSets(locations, energy_carriers, energy_converters,
                     supply_types, edges, arcs, time_periods)
model.CreateParametersFromDictionaries(network_costs, max_flow_line,
                                       supply_investment_costs, supply_external_factor,
                                       max_energy_supplied,
                                       converter_investment_costs,
                                       max_converted, conversion_efficiencies,
                                       storage_investment_costs,
                                       storage_losses, min_stored, max_stored,
                                       loss_factor,
                                       demand, amount_given, 0, 0, 0, 0)
model.InitializeVariables()
model.InitializeObjective()
model.InitializeConstraints("With")

# Connect to solver and optimize
print("Running solver")
# You can set timelimit or other limits (eg/ optimallity gap) in pyomo_helper
# returns: solver.solve(model, tee = True)

results = RunningLocalServer(model.model, 'gurobi', 100)
#results = yes['output']

PrintResult(results, model.model.Cost)

#n_RES = 0
# for x in model.model.SupplyInvestmentMade:
#    n_RES = n_RES + model.model.SupplyInvestmentMade[x].value
#    if(model.model.SupplyInvestmentMade[x].value > 0):
#        print(x, model.model.SupplyInvestmentMade[x].value)
# print(n_RES)
#
#n_networks = 0
# for x in model.model.LineInvestmentMade:
#    n_networks = n_networks + model.model.LineInvestmentMade[x].value
#    if(model.model.LineInvestmentMade[x].value > 0):
#        print(x, model.model.LineInvestmentMade[x].value)
# print(n_networks)
#
# for time_period in time_periods:
#    elec_demand = 0
#    for location in locations:
#        elec_demand = elec_demand + demand['Electricity', location, time_period]
#    print(time_period, elec_demand)

#np.savetxt('investments.csv', (x, model.model.SupplyInvestmentMade[x].value), delimiter=',')

# variable names, all start with model.variable
#LineInvestmentMade, SupplyInvestmentMade, ConverterInvestmentMade, StorageInvestmentMade
#AmountSupplied, AmountFlow, AmountConverted, AmountStored_In, AmountStored_Out
# StorageStartPeriod,StorageEndPeriod

workbook = xlsxwriter.Workbook('testresults.xlsx')

worksheet1 = workbook.add_worksheet('TotalInvestments')
worksheet2 = workbook.add_worksheet('ConversionUnits')
worksheet3 = workbook.add_worksheet('Networks')
worksheet4 = workbook.add_worksheet('StorageUnits')
worksheet5 = workbook.add_worksheet('Supply')
worksheet6 = workbook.add_worksheet('Demand')
worksheet7 = workbook.add_worksheet('Cost data')

# Sheet 1 - Investment summary
worksheet1.write_string('A1', 'Investment type')
worksheet1.write_string('B1', '# investments')

n_conversion = 0
for x in model.model.ConverterInvestmentMade:
    n_conversion += model.model.ConverterInvestmentMade[x].value
worksheet1.write_string('A2', 'Conversion Units')
worksheet1.write('B2', n_conversion)

n_networks = 0
for x in model.model.LineInvestmentMade:
    n_networks += model.model.LineInvestmentMade[x].value
worksheet1.write_string('A3', 'Networks')
worksheet1.write('B3', n_networks)

n_storage = 0
for x in model.model.StorageInvestmentMade:
    n_storage += model.model.StorageInvestmentMade[x].value
worksheet1.write_string('A4', 'Storage Units')
worksheet1.write('B4', n_storage)

n_RES = 0
for x in model.model.SupplyInvestmentMade:
    n_RES = n_RES + model.model.SupplyInvestmentMade[x].value
worksheet1.write_string('A5', 'RES')
worksheet1.write('B5', n_RES)

# Sheet 2 - Conversion Investments & Amount Converted
conversion_titles = ['Conversion type', 'Location', 'Time period',
                     '# investments', 'Total costs', '',
                     'Energy type', 'Location from', 'Location to', 'Time period',
                     'Energy converted [PJ]']
worksheet2.write_row('A1', conversion_titles)

linenum = 1
for x in model.model.ConverterInvestmentMade:
    if(model.model.ConverterInvestmentMade[x].value > 0):
        linenum += 1
        worksheet2.write_string('A' + str(linenum), str(x[0]))
        worksheet2.write_string('B' + str(linenum), str(x[1]))
        worksheet2.write_string('C' + str(linenum), str(x[2]))
        worksheet2.write_number(
            'D' + str(linenum), model.model.ConverterInvestmentMade[x].value)

linenum = 1
for x in model.model.AmountConverted:
    if(model.model.AmountConverted[x].value > 0):
        linenum += 1
        worksheet2.write_row('G' + str(linenum), x)
#        worksheet2.write_string('H' + str(linenum), str(x[1]))
#        worksheet2.write_string('I' + str(linenum), str(x[2]))
#        worksheet2.write_string('J' + str(linenum), str(x[3]))
        worksheet2.write_number(
            'K' + str(linenum), model.model.AmountConverted[x].value)


# Sheet 3 - Line Investments & Amount Flow
network_titles = ['Energy type', 'Location from', 'Location to', 'Time period',
                  '# investments', 'Total costs', '', 'Energy type',
                  'Location from', 'Location to',
                  'Time period', 'Energy flow [PJ]']
worksheet3.write_row('A1', network_titles)

linenum = 1
for x in model.model.LineInvestmentMade:
    if(model.model.LineInvestmentMade[x].value > 0):
        linenum += 1
        worksheet3.write_row('A' + str(linenum), x)
#        worksheet3.write_string('A' + str(linenum), str(x[0]))
#        worksheet3.write_string('B' + str(linenum), str(x[1]))
#        worksheet3.write_string('C' + str(linenum), str(x[2]))
#        worksheet3.write_string('D' + str(linenum), str(x[3]))
        worksheet3.write_number(
            'E' + str(linenum), model.model.LineInvestmentMade[x].value)

linenum = 1
for x in model.model.AmountFlow:
    if(model.model.AmountFlow[x].value > 0):
        linenum += 1
        worksheet3.write_row('H' + str(linenum), x)
#        worksheet3.write_string('H' + str(linenum), str(x[1]))
#        worksheet3.write_string('I' + str(linenum), str(x[2]))
#        worksheet3.write_string('J' + str(linenum), str(x[3]))
        worksheet3.write_number(
            'L' + str(linenum), model.model.AmountFlow[x].value)

# Sheet 4 - Storage Investments & Amount Stored/Unstored & Storage End & Begin Period
storage_titles = ['Energy type', 'Location', 'Time period', '# investments', 'Total costs', '',
                  'Energy type', 'Location', 'Time period', 'Energy stored IN [PJ]', '',
                  'Energy type', 'Location', 'Time period', 'Energy stored OUT [PJ]']
                  #'Energy type', 'Location', 'Time period', 'Storage end [PJ]']
worksheet4.write_row('A1', storage_titles)

linenum = 1
for x in model.model.StorageInvestmentMade:
    if(model.model.StorageInvestmentMade[x].value > 0):
        linenum += 1
        worksheet4.write_row('A' + str(linenum), x)
#        worksheet4.write_string('B' + str(linenum), str(x[1]))
#        worksheet4.write_string('C' + str(linenum), str(x[2]))
        worksheet4.write_number(
            'D' + str(linenum), model.model.StorageInvestmentMade[x].value)

linenum = 1
for x in model.model.AmountStored_In:
    if(model.model.AmountStored_In[x].value > 0):
        linenum += 1
        worksheet4.write_string('G' + str(linenum), str(x[0]))
        worksheet4.write_string('H' + str(linenum), str(x[1]))
        worksheet4.write_string('I' + str(linenum), str(x[2]))
        worksheet4.write_number(
            'J' + str(linenum), model.model.AmountStored_In[x].value)

linenum = 1
for x in model.model.AmountStored_Out:
    if(model.model.AmountStored_Out[x].value > 0):
        linenum += 1
        worksheet4.write_string('L' + str(linenum), str(x[0]))
        worksheet4.write_string('M' + str(linenum), str(x[1]))
        worksheet4.write_string('N' + str(linenum), str(x[2]))
        worksheet4.write_number(
            'O' + str(linenum), model.model.AmountStored_Out[x].value)

#linenum = 1
# for x in model.model.StorageStartPeriod:
#    linenum += 1
#    worksheet4.write_string('K' + str(linenum), str(x[0]))
#    worksheet4.write_string('L' + str(linenum), str(x[1]))
#    worksheet4.write_string('M' + str(linenum), str(x[2]))
#    worksheet4.write_number('N' + str(linenum), model.model.StorageStartPeriod[x].value)
#
#linenum = 1
# for x in model.model.StorageEndPeriod:
#    linenum += 1
#    worksheet4.write_string('P' + str(linenum), str(x[0]))
#    worksheet4.write_string('Q' + str(linenum), str(x[1]))
#    worksheet4.write_string('R' + str(linenum), str(x[2]))
#    worksheet4.write_number('S' + str(linenum), model.model.StorageEndPeriod[x].value)

# Sheet 5 - Supply Investments & Amount Supplied & original demand
supply_titles = ['Supply type', 'Location', 'Time period', '# investments', 'Total costs', '',
                 'Energy type', 'Location', 'Time period', 'Energy supplied [PJ]',
                 '', 'Energy type', 'Location', 'Time period', 'Max gas supplied [PJ]']
worksheet5.write_row('A1', supply_titles)

linenum = 1
for x in model.model.SupplyInvestmentMade:
    if(model.model.SupplyInvestmentMade[x].value > 0):
        linenum += 1
        worksheet5.write_row('A' + str(linenum), x)
#        worksheet5.write_string('B' + str(linenum), str(x[1]))
#        worksheet5.write_string('C' + str(linenum), str(x[2]))
        worksheet5.write_number(
            'D' + str(linenum), model.model.SupplyInvestmentMade[x].value)

linenum = 1
for x in model.model.AmountSupplied:
    if(model.model.AmountSupplied[x].value > 0):
        linenum += 1
        worksheet5.write_row('G' + str(linenum), x)
#        worksheet5.write_string('H' + str(linenum), str(x[1]))
#        worksheet5.write_string('I' + str(linenum), str(x[2]))
        worksheet5.write_number(
            'J' + str(linenum), model.model.AmountSupplied[x].value)

linenum = 1
for key, value in amount_given.items():
    if(value > 0):
        linenum += 1
        worksheet5.write_row('L' + str(linenum), key)
        worksheet5.write('O' + str(linenum), value)
# linenum = 1 - DOESNT WORK BECAUSE THIS IS INDEXEDPARAM, DONT KNOW HOW TO FIX
# for x in model.model.AmountGiven:
#    if(value(model.model.AmountGiven[x]) > 0):
#        linenum += 1
#        worksheet5.write_string('K' + str(linenum), str(x[0]))
#        worksheet5.write_string('L' + str(linenum), str(x[1]))
#        worksheet5.write_string('M' + str(linenum), str(x[2]))
#        worksheet5.write_number('N' + str(linenum), value(model.model.AmountGiven[x]))

# Sheet 6 - Demand
demand_titles = ['Demand type', 'Location', 'Time period', 'Amount [PJ]']
worksheet6.write_row('A1', demand_titles)

linenum = 1
for key, value in demand.items():
    linenum += 1
    worksheet6.write_row('A' + str(linenum), key)
    worksheet6.write('D' + str(linenum), value)

# Sheet 7 - Cost data
costdata_titles = ['Conversion type', 'Time period', 'Cost', '',
                   'Network type', 'Location from', 'Location to',
                   'Time period', 'Cost', '',
                   'Storage type', 'Time period', 'Cost', '',
                   'Supply type', 'Time period', 'Cost']
worksheet7.write_row('A1', costdata_titles)

linenum = 1
for key, value in converter_investment_costs.items():
    if(value > 0):
        linenum += 1
        worksheet7.write_row('A' + str(linenum), key)
        worksheet7.write('C' + str(linenum), value)

linenum = 1
for key, value in network_costs.items():
    if(value > 0):
        linenum += 1
        worksheet7.write_row('E' + str(linenum), key)
        worksheet7.write('I' + str(linenum), value)

linenum = 1
for key, value in storage_investment_costs.items():
    if(value > 0):
        linenum += 1
        worksheet7.write_row('K' + str(linenum), key)
        worksheet7.write('M' + str(linenum), value)

linenum = 1
for key, value in supply_investment_costs.items():
    if(value > 0):
        linenum += 1
        worksheet7.write_row('O' + str(linenum), key)
        worksheet7.write('Q' + str(linenum), value)

workbook.close()

