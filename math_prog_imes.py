# -*- coding: utf-8 -*-
"""
Created on Fri Nov 29 14:42:08 2019, updated Wed May 13, 2020

@author: IrisvB
"""
# Create mathematical model
from pyomo.environ import (ConcreteModel, Var, minimize, Objective, Constraint,
                           Set, Param, NonNegativeIntegers, NonNegativeReals,
                           Binary)


class Model:

    model = None

    def __init__(self):
        self.model = ConcreteModel()

    # Initialization of all the sets
    def InitializeSets(self, list_of_locations, list_of_energy_carriers,
                       list_of_energy_converters, list_of_supply_types,
                       list_of_edges, list_of_arcs, list_of_time_periods):
        # Set of all the locations (L)
        self.model.Locations = Set(initialize=list_of_locations)
        # Set of all the energy carriers (EC)
        self.model.EnergyCarriers = Set(initialize=list_of_energy_carriers)
        # Set of all energy converters (MT)
        self.model.EnergyConverters = Set(initialize=list_of_energy_converters)
        # Set of all the different supply types (Wind and solar, whch are used for electricity supply)
        self.model.SupplyTypes = Set(initialize=list_of_supply_types)
        # Set of edges (E), which define the possible link invewstments that can be made
        self.model.Edges = Set(initialize=list_of_edges)
        # Set of arcs (A), whcih determine the between which locations there can be flow
        self.model.Arcs = Set(initialize=list_of_arcs)
        self.model.TimePeriods = Set(
            initialize=list_of_time_periods)  # Set of time periods (T)

    # Other parameters

    def CreateParametersFromDictionaries(self, network_costs, max_flow_line,
                                         supply_investment_costs, supply_external_factor,
                                         max_energy_supplied,
                                         converter_investment_costs,
                                         max_converted,
                                         conversion_efficiencies, storage_costs,
                                         storage_losses, min_stored, max_stored,
                                         loss_factor, demand, amount_given,
                                         earlier_line_investment_made,
                                         earlier_supply_investment_made,
                                         earlier_converter_investment_made,
                                         earlier_storage_investment_made):
        # Network Parameter
        self.model.NetworkCosts = Param(self.model.EnergyCarriers, self.model.Locations, self.model.Locations,
                                        self.model.TimePeriods, initialize=network_costs)  # Network Costs (c^F)
        # Maximum flow over a line Gamma^F, which depends on the energy carrier
        self.model.MaxFlowLine = Param(
            self.model.EnergyCarriers, initialize=max_flow_line)
        self.model.LossFactor = Param(
            self.model.EnergyCarriers, initialize=loss_factor)
        # Supply Parameters
        self.model.SupplyInvestmentCosts = Param(
            self.model.SupplyTypes, self.model.TimePeriods, initialize=supply_investment_costs)  # Supply Investment Costs (c^S)
        self.model.SupplyExternalFactor = Param(
            self.model.SupplyTypes, self.model.Locations, self.model.TimePeriods, initialize=supply_external_factor) #20200805 attempt to add external factor for supply for weather scenarios
        # Maximum amount of energy supplied (Gamma^S)
        self.model.MaxEnergySupplied = Param(
            self.model.EnergyCarriers, self.model.SupplyTypes, initialize=max_energy_supplied)
        # Converter Parameters
        self.model.ConverterInvestmentCosts = Param(
            self.model.EnergyConverters, self.model.TimePeriods, initialize=converter_investment_costs)  # Converter Investment Costs (c^M)
        # Maximum amount of energy that can be converted on the specific converter (Gamma^M)
        self.model.MaxConverted = Param(
            self.model.EnergyCarriers, self.model.EnergyConverters, initialize=max_converted)
        self.model.ConversionEfficiencies = Param(self.model.EnergyCarriers, self.model.EnergyCarriers, self.model.EnergyConverters,
                                                  initialize=conversion_efficiencies)  # Energy efficiencies in conversion units (eta^MT_{E,V})
        # Storage Parameters
        self.model.StorageCosts = Param(
            self.model.EnergyCarriers, self.model.TimePeriods, initialize=storage_costs)  # Storage Costs
        self.model.StorageLosses = Param(
            self.model.EnergyCarriers, initialize=storage_losses)  # Storage Losses
        # Minimum amount of energy stored
        self.model.MinStorage = Param(
            self.model.EnergyCarriers, initialize=min_stored)
        # Maximum amount of energy stored (Gamma^W)
        self.model.MaxStorage = Param(
            self.model.EnergyCarriers, initialize=max_stored)
        # Demand and amount of supply already existing
        # Demand of every location and energy carrier (D)
        self.model.Demand = Param(
            self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, initialize=demand)
        # Already existing supply (there is already some gas supply without any costs)
        self.model.AmountGiven = Param(
            self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, initialize=amount_given)
        # Parameters for the amount of investments that are already made on the "existing" infrastructure. (for now this is 2014 and run the remaining years (which is 2016-2050))
        self.model.EarlierLineInvestmentMade = Param(
            self.model.EnergyCarriers, self.model.Locations, self.model.Locations, initialize=earlier_line_investment_made, default=0)
        self.model.EarlierSupplyInvestmentMade = Param(
            self.model.SupplyTypes, self.model.Locations, initialize=earlier_supply_investment_made, default=0)
        self.model.EarlierConverterInvestmentMade = Param(
            self.model.EnergyConverters, self.model.Locations, initialize=earlier_converter_investment_made, default=0)
        self.model.EarlierStorageInvestmentMade = Param(
            self.model.EnergyCarriers, self.model.Locations, initialize=earlier_storage_investment_made, default=0)

    # Variables Initialization
    def InitializeVariables(self):
        # The pipeline investment variables, restricted to integer number of investments (B^F)
        self.model.LineInvestmentMade = Var(self.model.EnergyCarriers, self.model.Locations,
                                            self.model.Locations, self.model.TimePeriods, within=NonNegativeIntegers)
        # The supply investment variables, restricted to integer number of investments (B^S)
        self.model.SupplyInvestmentMade = Var(
            self.model.SupplyTypes, self.model.Locations, self.model.TimePeriods, within=NonNegativeIntegers)
        # The converter investment variables, restricted to integer number of investments (B^M)
        self.model.ConverterInvestmentMade = Var(
            self.model.EnergyConverters, self.model.Locations, self.model.TimePeriods, within=NonNegativeIntegers)
        # The storage investment variables, restricted to the maximum number of investments (B^W)
        self.model.StorageInvestmentMade = Var(
            self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, within=NonNegativeIntegers)
        self.model.AmountSupplied = Var(self.model.EnergyCarriers, self.model.Locations,
                                        self.model.TimePeriods, within=NonNegativeReals)  # amount of supply inside a node (S)
        self.model.AmountFlow = Var(self.model.EnergyCarriers, self.model.Locations, self.model.Locations,
                                    self.model.TimePeriods, within=NonNegativeReals)  # amount of flow to other nodes (F)
        self.model.AmountConverted = Var(self.model.EnergyCarriers, self.model.EnergyConverters, self.model.Locations,
                                         self.model.TimePeriods, within=NonNegativeReals)  # amount of energy converted on particular conversion unit(M)
        # amount of energy stored to a storage unit (Wstored)
        self.model.AmountStored_In = Var(
            self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, within=NonNegativeReals)
        # amount of energy taken out of a storage unit (Wstored)
        self.model.AmountStored_Out = Var(
            self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, within=NonNegativeReals)
#        self.model.StorageStartPeriod = Var(self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, within = NonNegativeReals) #amount of energy in the storage at the start of a period (Wstart)
#        self.model.StorageEndPeriod = Var(self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, within = NonNegativeReals) #amount of energy in the storage at the end os a period (Wend)

    # Objective function
    def InitializeObjective(self):
        self.model.Cost = Objective(rule=ConstructionRules.totalCosts,
                                    sense=minimize)  # Constraint (1)

    # Constraints
    def InitializeConstraints(self, with_or_without_storage):
        if(with_or_without_storage == "With"):
            self.model.MassBalanceConstraint = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.balanceConstraint)  # Constraint (2)
#            self.model.StartStorageConstraint = Constraint(self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule = ConstructionRules.startStorageConstraint) #Constraint (6) + (8)
#            self.model.EndStorageConstraint = Constraint(self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule = ConstructionRules.endStorageConstraint) #Constraint(7)
            self.model.MaxFlowConstraint = Constraint(
                self.model.EnergyCarriers, self.model.Edges, self.model.TimePeriods, rule=ConstructionRules.maxFlowConstraint)  # Constraint(9-11)
            self.model.MaxConvertedConstraint = Constraint(self.model.EnergyCarriers, self.model.Locations, self.model.EnergyConverters,
                                                           self.model.TimePeriods, rule=ConstructionRules.maxConvertedConstraint)  # Constraint(12-14)
#            self.model.MinimumStoredConstraint = Constraint(self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule = ConstructionRules.minimumStoredConstraint) #Constraint(15-17)
#            self.model.MaximumStoredConstraint = Constraint(self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule = ConstructionRules.maximumStoredConstraint) #Constraint(15,18,19)
            self.model.MaxSupplyConstraint = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.maxSupplyConstraint)  # Constraint(20-22)
            self.model.MaxFlowInvestmentMade = Constraint(
                self.model.EnergyCarriers, self.model.Edges, self.model.TimePeriods, rule=ConstructionRules.maxFlowInvestmentMade)  # Constraint(23)
            self.model.MaxSupplyInvestmentMade = Constraint(
                self.model.SupplyTypes, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.maxSupplyInvestmentMade)  # Constraint(24)
            self.model.MaxStorageInvestmentMade = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.maxStorageInvestmentMade)  # Constraint(25)
            self.model.MaxConverterInvestmentMade = Constraint(
                self.model.EnergyConverters, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.maxConverterInvestmentMade)  # Constraint(26)
            # Additional constraint which states that there can be no link places between locations if there is no link between two locations. This is done such that there can only be links build between locations once ( for example: Link between Node_1 and Node_2 is allowed, but a link between Node_2 and Node_1 not. It is still alllowed to let flow go both ways)
            self.model.NoFlowInvestmentMade = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.noFlowInvestmentMade)
            # Additional constraint. If there is not a possibility to place a link between two locations (so both ways not possible) there is no flow between these two locations. This only happens if the number of arcs is limited
            self.model.NoAmountFlow = Constraint(self.model.EnergyCarriers, self.model.Locations,
                                                 self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.noAmountFlow)
            self.model.MaxAmountStoredOut = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.maxAmountStoredOut)
            self.model.MaxAmountStoredIn = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.maxAmountStoredIn)
#            self.model.StartConstraint = Constraint(self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule = ConstructionRules.startConstraint)

        else:
            self.model.MassBalanceConstraint = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.balanceConstraint)  # Constraint (2)
            self.model.MaxFlowConstraint = Constraint(
                self.model.EnergyCarriers, self.model.Edges, self.model.TimePeriods, rule=ConstructionRules.maxFlowConstraint)  # Constraint(9-11)
            self.model.MaxConvertedConstraint = Constraint(self.model.EnergyCarriers, self.model.Locations, self.model.EnergyConverters,
                                                           self.model.TimePeriods, rule=ConstructionRules.maxConvertedConstraint)  # Constraint(12-14)
            self.model.MaxSupplyConstraint = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.maxSupplyConstraint)  # Constraint(20-22)
            self.model.MaxFlowInvestmentMade = Constraint(
                self.model.EnergyCarriers, self.model.Edges, self.model.TimePeriods, rule=ConstructionRules.maxFlowInvestmentMade)  # Constraint(23)
            self.model.MaxSupplyInvestmentMade = Constraint(
                self.model.SupplyTypes, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.maxSupplyInvestmentMade)  # Constraint(24)
            self.model.MaxConverterInvestmentMade = Constraint(
                self.model.EnergyConverters, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.maxConverterInvestmentMade)  # Constraint(26)
            self.model.NoFlowInvestmentMade = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.noFlowInvestmentMade)
            self.model.NoAmountFlow = Constraint(self.model.EnergyCarriers, self.model.Locations,
                                                 self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.noAmountFlow)
            self.model.NoStorage = Constraint(
                self.model.EnergyCarriers, self.model.Locations, self.model.TimePeriods, rule=ConstructionRules.noStorage)
# ------------------------------------------------------------------------------


class ConstructionRules:

    # Minimize LineInvestmentCosts + SupplyInvestmentCosts + ConverterInvestmentCosts
    # Costs for the whole period
    @staticmethod
    def totalCosts(model):
        return \
            sum(sum(sum(model.NetworkCosts[energy_type, edge[0], edge[1], time_period] * model.LineInvestmentMade[energy_type,  edge[0], edge[1], time_period]
                        for edge in model.Edges)
                    for energy_type in model.EnergyCarriers)
                for time_period in model.TimePeriods) + \
            sum(sum(sum(model.SupplyInvestmentCosts[supply_type, time_period] * model.SupplyInvestmentMade[supply_type, location, time_period]
                        for supply_type in model.SupplyTypes)
                    for location in model.Locations)
                for time_period in model.TimePeriods) + \
            sum(sum(sum(model.ConverterInvestmentCosts[energy_converter, time_period] * model.ConverterInvestmentMade[energy_converter, location, time_period]
                        for location in model.Locations)
                    for energy_converter in model.EnergyConverters)
                for time_period in model.TimePeriods) + \
            sum(sum(sum(model.StorageInvestmentMade[energy_type, location, time_period]*model.StorageCosts[energy_type, time_period]
                        for energy_type in model.EnergyCarriers)
                    for location in model.Locations)
                for time_period in model.TimePeriods)

    # removed from begin: model.AmountGiven[energy_type, location, time_period] + \
    @staticmethod
    def balanceConstraint(model, energy_type, location, time_period):
        return \
            model.Demand[energy_type, location, time_period] <=\
            model.AmountSupplied[energy_type, location, time_period] + \
            ((1-model.LossFactor[energy_type]) *
             sum(model.AmountFlow[energy_type, location_from, location, time_period]
                 for location_from in model.Locations if((location_from, location) in model.Arcs))) - \
            sum(model.AmountFlow[energy_type, location, location_to, time_period]
                for location_to in model.Locations if((location, location_to) in model.Arcs)) + \
            sum(sum(model.AmountConverted[energy_type_2, energy_converter, location, time_period] * model.ConversionEfficiencies[energy_type, energy_type_2, energy_converter]
                    for energy_converter in model.EnergyConverters)
                for energy_type_2 in model.EnergyCarriers) - \
            model.AmountStored_In[energy_type, location, time_period] + \
            model.AmountStored_Out[energy_type, location, time_period] #- \
            #model.Demand[energy_type, location, time_period] #== 0   #testing inequality
                         # before, demand was on the RHS, adjusted to what Julie used

    @staticmethod
    def maxSupplyConstraint(model, energy_type, location, time_period):
        if(energy_type == 'Electricity'):
            return model.AmountSupplied[energy_type, location, time_period] <=\
                sum(sum(model.SupplyInvestmentMade[supply_type, location,
                    time_period2]*model.MaxEnergySupplied[energy_type, supply_type] *
                    model.SupplyExternalFactor[supply_type, location, time_period] #attempted to add supply external factor, just for weather scenarios
                        for supply_type in model.SupplyTypes) for time_period2 in
                    model.TimePeriods if(int(time_period2) <= int(time_period))) + \
                sum(model.EarlierSupplyInvestmentMade[supply_type, location] *
                    model.MaxEnergySupplied[energy_type, supply_type]
                    for supply_type in model.SupplyTypes)
        # new attempt to add gas in this supply constraint, used to only have the if for electricity and an else with 0
        elif(energy_type == 'Gas'):
            return model.AmountSupplied[energy_type, location, time_period] <=\
                model.AmountGiven[energy_type, location, time_period]
        elif(energy_type == 'Heat'):
            return model.AmountSupplied[energy_type, location, time_period] == 0

    @staticmethod
    def maxFlowConstraint(model, energy_type, edge0, edge1, time_period):
        return \
            model.AmountFlow[energy_type, edge0, edge1, time_period] + model.AmountFlow[energy_type, edge1, edge0, time_period] <= \
            (sum(model.LineInvestmentMade[energy_type, edge0, edge1, time_period2] for time_period2 in model.TimePeriods
                 if(int(time_period2) <= int(time_period))) + model.EarlierLineInvestmentMade[energy_type, edge0, edge1])*model.MaxFlowLine[energy_type]

    @staticmethod
    def maxConvertedConstraint(model, energy_type, location, energy_converter, time_period):
        return \
            (sum(model.ConverterInvestmentMade[energy_converter, location,
                                               time_period2] for time_period2 in model.TimePeriods
                 if(int(time_period2) <= int(time_period))) +
             model.EarlierConverterInvestmentMade[energy_converter, location]) *\
            model.MaxConverted[energy_type, energy_converter] >= \
            model.AmountConverted[energy_type,
                                  energy_converter, location, time_period]

    
    @staticmethod
    def maxAmountStoredIn(model, energy_type, location, time_period):
        return model.AmountStored_In[energy_type, location, time_period] <= \
            sum(model.StorageInvestmentMade[energy_type, location,
                                            time_period2]*model.MaxStorage[energy_type]
                for time_period2 in model.TimePeriods
                if(int(time_period2) <= int(time_period))) - \
            sum(model.AmountStored_In[energy_type, location, time_period2] *
                model.StorageLosses[energy_type]**(time_period-time_period2)
                for time_period2 in model.TimePeriods
                if(int(time_period2) < int(time_period))) +\
            sum(model.AmountStored_Out[energy_type, location, time_period2]
                for time_period2 in model.TimePeriods
                if(int(time_period2) < int(time_period)))

    @staticmethod
    def maxAmountStoredOut(model, energy_type, location, time_period):
        return model.AmountStored_Out[energy_type, location, time_period] <= \
            sum(model.AmountStored_In[energy_type, location, time_period2] *
                model.StorageLosses[energy_type]**(time_period-time_period2)
                for time_period2 in model.TimePeriods
                if(int(time_period2) < int(time_period))) -\
            sum(model.AmountStored_Out[energy_type, location, time_period2]
                for time_period2 in model.TimePeriods
                if(int(time_period2) < int(time_period)))

    @staticmethod #adjusted slightly higher than with Julie, originally 0, 0, 1, 2, 3, 4; first whole run: 0,4,6,5,7,5,10
    #20200825 scenario 5 is infeasible with 0,4,4,5,5,5. Doubling solar?
    def maxSupplyInvestmentMade(model, supply_type, location, time_period):
        if((supply_type == 'Wind') and ((location == 'Node_1') or (location == 'Node_5') or (location == 'Node_6'))):
            return model.SupplyInvestmentMade[supply_type, location, time_period] == 0
        elif((supply_type == 'Wind') and (int(time_period) == 2018)):
            return model.SupplyInvestmentMade[supply_type, location, time_period] <= 4
        elif((supply_type == 'Solar') and (int(time_period) == 2018)):
            return model.SupplyInvestmentMade[supply_type, location, time_period] <= 6
        elif((supply_type == 'Wind') and (int(time_period) == 2020)):
            return model.SupplyInvestmentMade[supply_type, location, time_period] <= 5
        elif((supply_type == 'Solar') and (int(time_period) == 2020)):
            return model.SupplyInvestmentMade[supply_type, location, time_period] <= 7
#        elif((supply_type == 'Wind') and (time_period == '2022') or 
#             (time_period == '2024') or (time_period == '2026') or
#             (time_period == '2028') or (time_period == '2030')):
#            return model.SupplyInvestmentMade[supply_type, location, time_period] <= 5
#        elif((supply_type == 'Solar') and (int(time_period) == 2022) | 
#             (int(time_period) == 2024) | (int(time_period) == 2026) | 
#             (int(time_period) == 2028) | (int(time_period) == 2030)):
#            return model.SupplyInvestmentMade[supply_type, location, time_period] <= 9
        elif(supply_type == 'Wind'):
            return model.SupplyInvestmentMade[supply_type, location, time_period] <= 5
        #elif(supply_type == 'Solar'):
            #return model.SupplyInvestmentMade[supply_type, location, time_period] <= 10
        else:
            return model.SupplyInvestmentMade[supply_type, location, time_period] <= 8

    @staticmethod
    def maxFlowInvestmentMade(model, energy_type, edge0, edge1, time_period):
        if(energy_type == 'Heat'):
            return model.LineInvestmentMade[energy_type, edge0, edge1, time_period] <= 5
        else:
            # could be set to one for the other energy types
            return model.LineInvestmentMade[energy_type, edge0, edge1, time_period] <= 5

    @staticmethod
    def maxConverterInvestmentMade(model, energy_converter, location, time_period):
        if(energy_converter == 'P2G'):
            return model.ConverterInvestmentMade[energy_converter, location, time_period] <= 5
        elif(energy_converter == 'HP'):
            return model.ConverterInvestmentMade[energy_converter, location, time_period] <= 5
        else:
            return model.ConverterInvestmentMade[energy_converter, location, time_period] <= 5

    @staticmethod
    def maxStorageInvestmentMade(model, energy_carrier, location, time_period):
        return model.StorageInvestmentMade[energy_carrier, location, time_period] <= 5

# ==========================
    @staticmethod
    def noFlowInvestmentMade(model, energy_carrier, location_from, location_to, time_period):
        if((location_from, location_to) in model.Edges):
            return model.LineInvestmentMade[energy_carrier, location_from, location_to, time_period] >= 0
        else:
            return model.LineInvestmentMade[energy_carrier, location_from, location_to, time_period] == 0
  
    @staticmethod
    def noAmountFlow(model, energy_carrier, location_from, location_to, time_period):
        if((location_from, location_to) in model.Arcs):
            return model.AmountFlow[energy_carrier, location_from, location_to, time_period] >= 0
        else:
            return model.AmountFlow[energy_carrier, location_from, location_to, time_period] == 0

    @staticmethod
    def noStorage(model, energy_carrier, location, time_period):
        return model.AmountStored_In[energy_carrier, location, time_period] == 0
