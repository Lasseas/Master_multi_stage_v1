#write code here

import numpy as np
import sys
import pandas as pd
import pyomo.environ as pyo
from pyomo.opt import SolverFactory
import time
import os 
import matplotlib.pyplot as plt
import platform
import psutil
from pyomo.environ import *

##################################################################################
############################### READING EXCEL FILE ###############################
##################################################################################

# Function to read all sheets in an Excel file and save each as a .tab file in the current directory
def read_all_sheets(excel):
    # Load the Excel file
    input_excel = pd.ExcelFile(excel)
    
    # Loop over each sheet in the workbook
    for sheet in input_excel.sheet_names:
        # Read the current sheet, skipping the first two rows
        input_sheet = pd.read_excel(excel, sheet_name=sheet, skiprows=2)

        # Drop only fully empty rows (optional)
        data_nonempty = input_sheet.dropna(how='all')

        # Replace spaces in column names with underscores
        data_nonempty.columns = data_nonempty.columns.astype(str).str.replace(' ', '_')

        # Fill missing values with an empty string
        data_nonempty = data_nonempty.fillna('')

        # Convert all columns to strings before replacing whitespace characters in values
        data_nonempty = data_nonempty.applymap(lambda x: str(x) if pd.notnull(x) else "")
        
        # Save as a .tab file using only the sheet name as the file namec

        output_filename = f"{sheet}.tab"
        data_nonempty.to_csv(output_filename, header=True, index=False, sep='\t')
        print(f"Saved file: {output_filename}")

# Call the function with your Excel file
read_all_sheets('Input-data_CaseStudy_Flex_Shift_aFRR.xlsx')

####################################################################
######################### MODEL SPECIFICATIONS #####################
####################################################################

model = pyo.AbstractModel()

data = pyo.DataPortal() #Loading the data from a data soruce in a uniform manner (Excel)

"""
SETS 
"""
#sets defined here
TechToEC = {""}

#Defining Sets

#Declaring Sets
model.Time = pyo.Set(ordered=True) #Set of time periods (hours)
model.LoadShiftingIntervals = pyo.Set(ordered=True)
model.TimeLoadShift = pyo.Set(dimen = 2, ordered = True) #Subset of time periods for load shifting
model.Month = pyo.Set(ordered = True) #Set of months
model.TimeInMonth = pyo.Set(within = model.Time) #Subset of time periods in month m
model.Technology = pyo.Set(ordered = True) #Set of technologies
model.EnergyCarrier = pyo.Set(ordered = True)
#model.EnergyCarrierToTechnology = pyo.Set(model.EnergyCarrier, within = model.Technology, initialize=tab_to_dict("Subset_TechFromEC.tab")) #I^in -> Subset of technologies that can receive fuel from carreir e
#model.TechnologyToEnergyCarrier = pyo.Set(model.EnergyCarrier, within = model.Technology, initialize=tab_to_dict("Subset_TechToEC.tab")) #I^out -> Subset of technolgoes that can deliver to fuel carrier e
model.TechnologyToEnergyCarrier = pyo.Set(dimen=2, ordered = True)
model.EnergyCarrierToTechnology = pyo.Set(dimen=2, ordered = True)
model.FlexibleLoad = pyo.Set(ordered=True) #Set of flexible loads (batteries)
model.FlexibleLoadForEnergyCarrier = pyo.Set(dimen = 2, ordered = True)
#model.UpShiftForEnergyCarrier = pyo.Set(dimen = 2, ordered = True, within = model.FlexibleLoadForEnergyCarrier)
#model.DownShiftForEnergyCarrier = pyo.Set(dimen = 2, ordered = True, within = model.FlexibleLoadForEnergyCarrier)
model.ShiftableLoadForEnergyCarrier = pyo.Set(dimen = 2, ordered = True, within = model.FlexibleLoadForEnergyCarrier)
#model.FlexibleLoadForEnergyCarrier = pyo.Set(model.EnergyCarrier, within=model.FlexibleLoad, initialize=tab_to_dict("Set_of_FlexibleLoadForEC.tab"))
#model.ShiftableLoadForEnergyCarrier = pyo.Set(model.EnergyCarrier, within=model.FlexibleLoadForEnergyCarrier, initialize=tab_to_dict("Subset_ShiftableLoads.tab")) #Subset of shiftable loads (shift-able batteries)
model.Nodes = pyo.Set(ordered=True) #Set of Nodess
model.Nodes_DA = pyo.Set(ordered=True, within = model.Nodes) #Set of Nodess
model.Nodes_ID = pyo.Set(ordered=True, within = model.Nodes) #Set of Nodess
model.Nodes_RT = pyo.Set(ordered=True, within = model.Nodes) #Set of Nodess
model.Parent_Node = pyo.Set(dimen = 2, ordered = True)
model.Mode_of_operation = pyo.Set(ordered = True)


#Reading the Sets, and loading the data

data = pyo.DataPortal() #Loading the data from a data source in a uniform manner (Excel)

data.load(filename="Set_of_TimeSteps.tab", format="set", set=model.Time)
data.load(filename="Subset_LoadShiftWindow.tab", format="set", set=model.TimeLoadShift)
data.load(filename="Set_of_Month.tab", format = "set", set=model.Month)
data.load(filename="Subset_of_TimeStepsInMonth.tab", format = "set", set=model.TimeInMonth)
data.load(filename="Set_of_Technology.tab", format = "set", set=model.Technology)
data.load(filename="Set_of_FlexibleLoad.tab", format="set", set=model.FlexibleLoad)
data.load(filename="Set_of_Nodes.tab", format="set", set=model.Nodes)
data.load(filename="Set_of_EnergyCarrier.tab", format="set", set=model.EnergyCarrier)
data.load(filename="Subset_TechToEC.tab", format="set", set=model.TechnologyToEnergyCarrier)
data.load(filename="Subset_ECToTech.tab", format="set", set=model.EnergyCarrierToTechnology)
data.load(filename="Set_of_FlexibleLoadForEC.tab", format="set", set=model.FlexibleLoadForEnergyCarrier)
#data.load(filename="Subset_UpShiftForEC.tab", format="set", set=model.UpShiftForEnergyCarrier)
#data.load(filename="Subset_DownShiftForEC.tab", format="set", set=model.DownShiftForEnergyCarrier)
data.load(filename="Subset_ShiftableLoadForEC.tab", format="set", set=model.ShiftableLoadForEnergyCarrier)
data.load(filename="Set_of_LoadShiftingInterval.tab", format = "set", set = model.LoadShiftingIntervals)
data.load(filename="Set_of_Nodes_DA.tab", format = "set", set = model.Nodes_DA)
data.load(filename="Set_of_Nodes_ID.tab", format = "set", set = model.Nodes_ID)
data.load(filename="Set_of_Nodes_RT.tab", format = "set", set = model.Nodes_RT)
data.load(filename="Set_of_Parent_Node.tab", format = "set", set = model.Parent_Node)
data.load(filename="Set_of_Mode_of_Operation.tab", format = "set", set = model.Mode_of_operation)



"""
PARAMETERS
"""


#Declaring Parameters
model.Cost_Energy = pyo.Param(model.Nodes, model.Time, model.Technology, model.Mode_of_operation, default=0.0)  # Cost of using energy source i at time t
model.Cost_Battery = pyo.Param(model.Nodes, model.Time, model.FlexibleLoad)
model.Cost_Export = pyo.Param(model.Nodes, model.Time, model.EnergyCarrier)  # Income from exporting energy to the grid at time t
model.Cost_Expansion_Tec = pyo.Param(model.Technology) #Capacity expansion cost
model.Cost_Expansion_Bat = pyo.Param(model.FlexibleLoad) #Capacity expansion cost
model.Cost_Emission = pyo.Param() #Carbon price
model.Cost_Grid = pyo.Param() #Grid tariff
model.Cost_Imbal = pyo.Param(model.Nodes, model. Time)
model.aFRR_Up_Capacity_Price = pyo.Param(model.Nodes, model.Time, default=0.0)  # Capacity Price for aFRR up regulation 
model.aFRR_Dwn_Capacity_Price = pyo.Param(model.Nodes, model.Time, default=0.0)  # Capcaity Price for aFRR down regulation
model.aFRR_Up_Activation_Price = pyo.Param(model.Nodes, model.Time, default=0.0)  # Activation Price for aFRR up regulation 
model.aFRR_Dwn_Activation_Price = pyo.Param(model.Nodes, model.Time, default=0.0)  # Activatioin Price for aFRR down regulation 
model.Spot_Price = pyo.Param(model.Nodes, model.Time, default=0.0)
model.Intraday_Price = pyo.Param(model.Nodes, model.Time, default=0.0)
model.RK_Up_Price = pyo.Param(model.Nodes, model.Time, default=0.0)
model.RK_Dwn_Price = pyo.Param(model.Nodes, model.Time, default=0.0)
model.Demand = pyo.Param(model.Nodes, model.Time, model.EnergyCarrier, default = 0)  # Energy demand 
model.Max_charge_discharge_rate = pyo.Param(model.FlexibleLoad, default = 1) # Maximum symmetric charge and discharge rate
model.Charge_Efficiency = pyo.Param(model.FlexibleLoad)  # Efficiency of charging flexible load b [-]
model.Discharge_Efficiency = pyo.Param(model.FlexibleLoad)  # Efficiency of discharging flexible load b [-]
model.Technology_To_EnergyCarrier_Efficiency = pyo.Param(model.Technology, model.Mode_of_operation) #Efficiency of technology i when supplying fuel e
model.EnergyCarrier_To_Technlogy_Efficiency = pyo.Param(model.Technology, model.Mode_of_operation) #Efficiency of technology i when consuming fuel e
model.Max_Storage_Capacity = pyo.Param(model.FlexibleLoad)  # Maximum energy storage capacity of flexible load b [MWh]
model.Self_Discharge = pyo.Param(model.FlexibleLoad)  # Self-discharge rate of flexible load b [%]
model.Initial_SOC = pyo.Param(model.FlexibleLoad)  # Initial state of charge for flexible load b [-]
model.Node_Probability = pyo.Param(model.Nodes)  # Probability of Nodes s [-]
model.Max_Cable_Capacity = pyo.Param()  # Maximum capacity of power cable for import/export [MW]
model.Up_Shift_Max = pyo.Param()  # Maximum allowable up-shifting in load shifting periods as a percentage of demand [% of demand]
model.Down_Shift_Max = pyo.Param()  # Maximum allowable down-shifting in load shifting periods as a percentage of demand [% of demand]
model.Initial_Installed_Capacity = pyo.Param(model.Technology) #Initial installed capacity at site for technology i
model.Ramping_Factor = pyo.Param(model.Technology)
model.Availability_Factor = pyo.Param(model.Nodes, model.Time, model.Technology) #Availability factor for technology delivering to energy carrier 
model.Carbon_Intensity = pyo.Param(model.Technology, model.EnergyCarrier) #Carbon intensity when supplying from technology to fuel
model.Max_Export = pyo.Param() #Maximum allowable export per year, if no concession is given
model.Activation_Factor_UP_Regulation = pyo.Param(model.Nodes, model.Time, default = 0) # Activation factor determining the duration of up regulation
model.Activation_Factor_DWN_Regulation = pyo.Param(model.Nodes, model.Time, default = 0) # Activation factor determining the duration of dwn regulation
model.Available_Excess_Heat = pyo.Param() #Fraction of the total available excess heat at usable temperature level to \\& be used an energy source for the heat pump.
model.Energy2Power_Ratio = pyo.Param(model.FlexibleLoad)
model.Max_CAPEX_tech = pyo.Param(model.Technology)
model.Max_CAPEX_flex = pyo.Param(model.FlexibleLoad)
#model.Shiftable_Window = pyo.Param() #Time window if flexible time window is used

#Reading the Parameters, and loading the data

data.load(filename="Par_EnergyCost.tab", param=model.Cost_Energy, format = "table")
data.load(filename="Par_BatteryCost.tab", param=model.Cost_Battery, format = "table")
data.load(filename="Par_ExportCost.tab", param=model.Cost_Export, format = "table")
data.load(filename="Par_CostExpansion_Tec.tab", param=model.Cost_Expansion_Tec, format = "table")
data.load(filename="Par_CostExpansion_Bat.tab", param=model.Cost_Expansion_Bat, format = "table")
data.load(filename="Par_CostEmission.tab", param=model.Cost_Emission, format = "table")
data.load(filename="Par_CostGridTariff.tab", param=model.Cost_Grid, format = "table")
data.load(filename="Par_CostImbalance.tab", param=model.Cost_Imbal, format = "table")
data.load(filename="Par_aFRR_UP_CAP_price.tab", param=model.aFRR_Up_Capacity_Price, format = "table")
data.load(filename="Par_aFRR_DWN_CAP_price.tab", param=model.aFRR_Dwn_Capacity_Price, format = "table")
data.load(filename="Par_aFRR_UP_ACT_price.tab", param=model.aFRR_Up_Activation_Price, format = "table")
data.load(filename="Par_aFRR_DWN_ACT_price.tab", param=model.aFRR_Dwn_Activation_Price, format = "table")
data.load(filename="Par_SpotPrice.tab", param=model.Spot_Price, format = "table")
data.load(filename="Par_IntradayPrice.tab", param=model.Intraday_Price, format = "table")
data.load(filename="Par_RK_UpPrice.tab", param=model.RK_Up_Price, format = "table")
data.load(filename="Par_RK_DwnPrice.tab", param=model.RK_Dwn_Price, format = "table")
data.load(filename="Par_EnergyDemand.tab", param=model.Demand, format = "table")
data.load(filename="Par_MaxChargeDischargeRate.tab", param=model.Max_charge_discharge_rate, format = "table")
data.load(filename="Par_ChargeEfficiency.tab", param=model.Charge_Efficiency, format = "table")
data.load(filename="Par_DischargeEfficiency.tab", param=model.Discharge_Efficiency, format = "table")
data.load(filename="Par_TechToEC_Efficiency.tab", param=model.Technology_To_EnergyCarrier_Efficiency, format = "table")
data.load(filename="Par_ECToTech_Efficiency.tab", param=model.EnergyCarrier_To_Technlogy_Efficiency, format = "table")
data.load(filename="Par_MaxStorageCapacity.tab", param=model.Max_Storage_Capacity, format = "table")
data.load(filename="Par_SelfDischarge.tab", param=model.Self_Discharge, format = "table")
data.load(filename="Par_InitialSoC.tab", param=model.Initial_SOC, format = "table")
data.load(filename="Par_NodesProbability.tab", param=model.Node_Probability, format = "table")
data.load(filename="Par_MaxCableCapacity.tab", param=model.Max_Cable_Capacity, format = "table")
data.load(filename="Par_MaxUpShift.tab", param=model.Up_Shift_Max, format = "table")
data.load(filename="Par_MaxDwnShift.tab", param=model.Down_Shift_Max, format = "table")
data.load(filename="Par_InitialCapacityInstalled.tab", param=model.Initial_Installed_Capacity, format = "table")
data.load(filename="Par_AvailabilityFactor.tab", param=model.Availability_Factor, format = "table")
data.load(filename="Par_CarbonIntensity.tab", param=model.Carbon_Intensity, format = "table")
data.load(filename="Par_MaxExport.tab", param=model.Max_Export, format = "table")
data.load(filename="Par_ActivationFactor_Up_Reg.tab", param=model.Activation_Factor_UP_Regulation, format = "table")
data.load(filename="Par_ActivationFactor_Dwn_Reg.tab", param=model.Activation_Factor_DWN_Regulation, format = "table")
data.load(filename="Par_AvailableExcessHeat.tab", param=model.Available_Excess_Heat, format = "table")
data.load(filename="Par_Energy2Power_ratio.tab", param=model.Energy2Power_Ratio, format = "table")
#data.load(filename="Filename.tab", param=model.Shiftable_Window, format = "table")

"""
VARIABLES
"""


#Defining Variables

#Declaring Variables
model.x_UP = pyo.Var(model.Nodes, model.Time, model.FlexibleLoad, domain= pyo.NonNegativeReals)
model.x_DWN = pyo.Var(model.Nodes, model.Time, model.FlexibleLoad, domain= pyo.NonNegativeReals)
model.x_DA = pyo.Var(model.Nodes, model.Time, domain= pyo.NonNegativeReals)
model.x_ID_Up = pyo.Var(model.Nodes, model.Time, domain= pyo.NonNegativeReals)
model.x_ID_Dwn = pyo.Var(model.Nodes, model.Time, domain= pyo.NonNegativeReals)
model.x_RT_Up = pyo.Var(model.Nodes, model.Time, domain= pyo.NonNegativeReals)
model.x_RT_Dwn = pyo.Var(model.Nodes, model.Time, domain= pyo.NonNegativeReals)
model.y_out = pyo.Var(model.Nodes, model.Time, model.Technology, model.EnergyCarrier, model.Mode_of_operation, domain = pyo.NonNegativeReals)
model.y_in = pyo.Var(model.Nodes, model.Time, model.Technology, model.EnergyCarrier,  model.Mode_of_operation, domain = pyo.NonNegativeReals)
model.y_activity = pyo.Var(model.Nodes, model.Time, model.Technology, model.Mode_of_operation, domain = pyo.NonNegativeReals)
model.z_export = pyo.Var(model.Nodes, model.Time, model.EnergyCarrier, domain = pyo.NonNegativeReals, bounds = (0, 0))
model.q_charge = pyo.Var(model.Nodes, model.Time, model.FlexibleLoad, domain= pyo.NonNegativeReals)
model.q_discharge = pyo.Var(model.Nodes, model.Time, model.FlexibleLoad, domain= pyo.NonNegativeReals)
model.q_SoC = pyo.Var(model.Nodes, model.Time, model.FlexibleLoad, domain= pyo.NonNegativeReals)
model.v_new_tech = pyo.Var(model.Technology, domain = pyo.NonNegativeReals, bounds = (0, 0)) 
model.v_new_bat = pyo.Var(model.FlexibleLoad, domain = pyo.NonNegativeReals, bounds = (0, 0))
model.y_max = pyo.Var(model.Nodes, model.Month, domain = pyo.NonNegativeReals)
model.d_flex = pyo.Var(model.Nodes, model.Time, model.EnergyCarrier, domain = pyo.NonNegativeReals)


#import pdb; pdb.set_trace()
"""
STATING THE MATHEMATICAL MODEL
"""

"""
OBJECTIVE FUNCTION
"""
def objective(model):
    return(
        #Investment cost
        sum(model.Cost_Expansion_tec[i] * model.v_new_tech[i] for i in model.Technology)
        + sum(model.Cost_Expansion_Bat[b] * model.v_new_bat[b] for b in model.FlexibleLoad)

        #Day-ahead and Capacity reservation
        + sum(
            sum(model.Node_Probability[n]*(
                model.Spot_Price[n, t] * model.x_DA[n, t]
                - sum(model.aFRR_Up_Capacity_Price[n, t] * model.x_UP[n, t, b] + 
                      model.aFRR_Dwn_Capacity_Price[n, t] * model.x_DWN[n, t, b] for b in model.FlexibleLoad) 
            )for n in model.Nodes_DA)

            #Intraday market
            + sum(model.Node_Probability[n] * (model.Intraday_Price[n, t] * (model.x_ID_Up[n, t] - model.x_ID_Down[n, t])) for n in model.Nodes_ID)

            #Cost/Compensatioin activation
            + sum(model.Node_Probability[n]*(
                sum(-model.Activation_Factor_UP_Regulation[n, t] * model.aFRR_Up__Activation_Price[n, t] * model.x_UP[n, t, b]
                + model.Activation_Factor_DWN_Regulation[n, t] * model.aFRR_Dwn_Activatioin_Price[n, t] * model.x_DWN[n, t, b]
                for b in model.FlexibleLoad)

                #Cost of Consumption and Carbon emissions
                + sum(
                    sum(
                        sum(
                            model.y_activity[n, t, i , o] * (model.Cost_Energy[n, t, i, o] + model.Carbon_Intensity[i, e])
                            for o in model.Mode_of_operation
                        )
                        for i in model.Technology if (i, e) in model.TechnologyToEnergyCarrier
                    )
                    - model.Cost_Export[n, t, e] * model.z_export[n, t, e] for e in model.EnergyCarrier
                )

                #Real-time adjustment compensation/cost + imbalance cost
                + model.RK_Up_Price[n, t] * model.x_RT_Up[n, t] - model.RK_Dwn_Price[n, t] * model.x_RT_Dwn[n, t]
                + model.Cost_Imbal[n, t] * (model.x_RT_Up[n, t] + model.x_RT_Dwn[n, t])

                #Grid tariff
                + sum(model.Cost_Grid * model.y_max[n, m] for m in model.Month)

                #Variable storage costs
                + sum(model.Cost_Battery[n, t, b] * model.q_discharge[n, t, b] for b in model.FlexibleLoad)
            )
            for n in model.Nodes_RT
                
            )
            for t in model.Time
        )
    )

model.Objective = pyo.Objective(rule=objective, sense=pyo.minimize)

"""
CONSTRAINTS
"""  
###########################################################
############## aFRR Total, ikke med i LaTeX ###############
###########################################################
"""
def aFRR_up_total(model, t, e):
    if e == 'Electricity':
        return model.x_aFRR_UP[t] == sum(model.x_aFRR_UP_ind[t, b] for b in model.FlexibleLoad if (b,e) in model.FlexibleLoadForEnergyCarrier)
    else:
        return pyo.Constraint.Skip
    
model.aFRRUpTotal = pyo.Constraint(model.Time, model.EnergyCarrier, rule=aFRR_up_total)

def aFRR_dwn_total(model, t, e):
    if e == 'Electricity':
        return model.x_aFRR_DWN[t] == sum(model.x_aFRR_DWN_ind[t, b] for b in model.FlexibleLoad if (b,e) in model.FlexibleLoadForEnergyCarrier)
    else:
        return pyo.Constraint.Skip
    
model.aFRRDwnTotal = pyo.Constraint(model.Time, model.EnergyCarrier, rule=aFRR_dwn_total)
"""

###########################################################
############## ENERGY AND CONVERSION BALANCES #############
###########################################################
def energy_balance(model, n, t, e):
    if e == 'Electricity':
        return (
            model.Demand[n, t, e]
            + sum(model.Activation_Factor_UP_Regulation[n, t] * model.x_UP[n, t, b]
            - model.Activation_Factor_DWN_Regulation[n, t] * model.x_DWN[n, t, b] for b in model.FlexibleLoad if (b,e) in model.FlexibleLoadForEnergyCarrier)
            == sum(sum(model.y_out[n, t, i, e, o] for i in model.Technology if (i,e) in model.TechnologyToEnergyCarrier)
            - sum(model.y_in[n, t, i, e, o] for i in model.Technology if (i,e) in model.EnergyCarrierToTechnology) for o in model.Mode_of_operation)
            - model.z_export[n, t, e]
            - sum(
                model.Charge_Efficiency[b] * model.q_charge[n, t, b] - model.q_discharge[n, t, b]
                for b in model.FlexibleLoad if (b,e) in model.FlexibleLoadForEnergyCarrier
            )
        )
    else:
        return (
            model.Demand[n, t, e]
            == sum(sum(model.y_out[n, t, i, e, o] for i in model.Technology if (i,e) in model.TechnologyToEnergyCarrier)
            - sum(model.y_in[n, t, i, e, o] for i in model.Technology if(i,e) in model.EnergyCarrierToTechnology)) for o in model.Mode_of_operation
            - model.z_export[n, t, e]
            - sum(
                model.Charge_Efficiency[b] * model.q_charge[n, t, b] - model.q_discharge[n, t, b]
                for b in model.FlexibleLoad if (b,e) in model.FlexibleLoadForEnergyCarrier
            )
        )

model.EnergyBalance = pyo.Constraint(model.Nodes, model.Time, model.EnergyCarrier, rule=energy_balance)

#####################################################################################
########################### MARKET BALANCE DA/ID/RT #################################
#####################################################################################

def market_balance_ID(model, n, t, i, e, o, p):
    if (i, e) == ("Power_Grid", "Electricity"):
        return (model.y_out[n, t, i, e, o] == model.y_out[p, t, i, e, o] + model.x_ID_Up[n, t] - model.x_ID_Dwn[n, t])
    else:
        return pyo.Constraint.Skip
      
model.MarketBalanceID = pyo.Constraint(model.Nodes_ID, model.Time, model.Technology, model.EnergyCarrier, model.Mode_of_operation, model.Parent_Node, rule = market_balance_ID)


def market_balance_RT(model, n, t, i, e, o, p):
    if (i, e) == ("Power_Grid", "Electricity"):
        return (model.y_out[n, t, i, e, o] == model.y_out[p, t, i, e, o] + model.x_RT_Up[n, t] - model.x_RT_Dwn[n, t])
    else:
        return pyo.Constraint.Skip
      
model.MarketBalanceRT = pyo.Constraint(model.Nodes_RT, model.Time, model.Technology, model.EnergyCarrier, model.Mode_of_operation, model.Parent_Node, rule = market_balance_RT)

#####################################################################################
########################### CONVERSION BALANCE ######################################
#####################################################################################

def conversion_balance_out(model, n, t, i, e, o):
    if (i, e) in model.TechnologyToEnergyCarrier:
        return (model.y_out[n, t, i, e, o] == model.y_activity[n, t, i, o] * model.Technology_To_EnergyCarrier_Efficiency[i, o])
    else:
        return pyo.Constraint.Skip
      
model.ConversionBalanceOut = pyo.Constraint(model.Nodes, model.Time, model.Technology, model.EnergyCarrier, model.Mode_of_operation, rule = conversion_balance_out)


def conversion_balance_in(model, n, t, i, e, o):
    if (i, e) in model.EnergyCarrierToTechnology:
        return (model.y_in[n, t, i, e, o] == model.y_activity[n, t, i, o] * model.EnergyCarrier_To_Technlogy_Efficiency[i, o])
    else:
        return pyo.Constraint.Skip
           
model.ConversionBalanceIn = pyo.Constraint(model.Nodes, model.Time, model.Technology, model.EnergyCarrier, model.Mode_of_operation, rule = conversion_balance_in)


#####################################################################################
########################### TECHNOLOGY RAMPING CONSTRAINTS ##########################
#####################################################################################

def Ramping_Technology(model, n, t, i, e, o):
    if t > 1:
        model.y_out[n, t, i, e, o] - model.y_out[n, t-1, i, e, o] <= model.Ramping_Factor[i] * (model.Initial_Installed_Capacity[i] + model.v_new_tech[i])
    else:
        model.y_out[n, t, i, e, o] <= model.Ramping_Factor[i] * (model.Initial_Installed_Capacity[i] + model.v_new_tech[i])
           
model.RampingTechnology = pyo.Constraint(model.Nodes, model.Time, model.Technology, model.EnergyCarrier, model.Mode_of_operation, rule = Ramping_Technology)


#####################################################################################
############## HEAT PUMP LIMITATION - MÅ ENDRES I HENHOLD TIL INPUTDATA #############
#####################################################################################

def heat_pump_input_limitation_LT(model, n, t, o):
    return (
        model.y_out[n, t, 'HeatPump_LT', 'LT', o] - model.y_in[n, t, 'HeatPump_LT', 'Electricity', o]
        <= model.Available_Excess_Heat * (model.Demand[n, t, 'LT'])# + model.Demand[s, t, 'HT'])
    )

model.HeatPumpInputLimitationLT = pyo.Constraint(model.Nodes, model.Time, model.Mode_of_operation, rule=heat_pump_input_limitation_LT)

def heat_pump_input_limitation_MT(model, n, t, o):
    return (
        model.y_out[n, t, 'HeatPump_MT', 'MT', o] - model.y_in[n, t, 'HeatPump_MT', 'Electricity', o]
        <= model.Available_Excess_Heat * (model.Demand[n, t, 'MT'])# + model.Demand[s, t, 'HT'])
    )

model.HeatPumpInputLimitationMT = pyo.Constraint(model.Nodes, model.Time, model.Mode_of_operation, rule=heat_pump_input_limitation_MT)

######################################################
############## LOAD SHIFTING CONSTRAINTS #############
######################################################

def loads_shifting_time_window(model, n, i, b, e):
        return sum(model.q_charge[n,t,b] - model.q_discharge[n,t,b]/model.Discharge_Efficiency[b] for (interval, t) in model.TimeLoadShift if interval == i) == 0

#model.LoadShiftingWindow = pyo.Constraint(model.Nodes, model.TimeLoadShift, model.FlexibleLoad, model.EnergyCarrier, rule=loads_shifting_time_winidow)
model.LoadShiftingWindow = pyo.Constraint(model.Nodes, model.LoadShiftingIntervals, model.ShiftableLoadForEnergyCarrier, rule=loads_shifting_time_window)

def no_discharge_outside_load_shift(model, n, t, b, e):
    if not any(t == load_t[1] for load_t in model.TimeLoadShift):
        return model.q_discharge[n, t, b]  == 0
    else:
        return pyo.Constraint.Skip

model.NoDischargeOutsideLoadShift = pyo.Constraint(model.Nodes, model.Time, model.ShiftableLoadForEnergyCarrier, rule=no_discharge_outside_load_shift)

def no_charge_outside_load_shift(model, n, t, b, e):
    if not any(t == load_t[1] for load_t in model.TimeLoadShift):
        return model.q_charge[n, t, b]  == 0
    else:
        return pyo.Constraint.Skip

model.NochargeOutsideLoadShift = pyo.Constraint(model.Nodes, model.Time, model.ShiftableLoadForEnergyCarrier, rule=no_charge_outside_load_shift)


###########################################################
############## MAX ALLOWABLE UP/DOWN SHIFT ################
###########################################################

def Max_total_up_dwn_load_shift(model, n, i, t, b, e):
    #if e == 'Electricity':
        # Extract all time steps `tt` for the given interval `i`
    relevant_times = [tt for (j, tt) in model.TimeLoadShift if j == i]
        
    if t in relevant_times:
        return model.q_charge[n,t,b] + model.q_discharge[n,t,b]/model.Discharge_Efficiency[b] <= model.Up_Shift_Max * model.Demand[n,t,e]
            
    return pyo.Constraint.Skip

model.MaxTotalUpDwnLoadShift = pyo.Constraint(model.Nodes, model.TimeLoadShift, model.ShiftableLoadForEnergyCarrier, rule=Max_total_up_dwn_load_shift)

###########################################################
############## aFRR PARTICIPATION CONSTRAINTS #############
###########################################################

def aFRR_up_dwn_limit_demand_constraint(model, n, t, b, e):
    if e == 'Electricity':
        return model.x_DWN[n, t, b] + model.x_UP[n, t, b]/model.Discharge_Efficiency[b] <= model.Up_Shift_Max * model.Demand[n, t, e]

    return pyo.Constraint.Skip

model.aFRRUpDwnDemandLimitLoadShift = pyo.Constraint(model.Nodes, model.Time, model.ShiftableLoadForEnergyCarrier, rule=aFRR_up_dwn_limit_demand_constraint)

def no_aFRR_up_outside_load_shift(model, n, t, b, e):
    if not any(t == load_t[1] for load_t in model.TimeLoadShift):
        return model.x_UP[n, t, b]  == 0
    else:
        return pyo.Constraint.Skip

model.NoaFRRUpOutsideLoadShift = pyo.Constraint(model.Nodes, model.Time, model.ShiftableLoadForEnergyCarrier, rule=no_aFRR_up_outside_load_shift)

def no_aFRR_dwn_outside_load_shift(model, n, t, b, e):
    if not any(t == load_t[1] for load_t in model.TimeLoadShift):
        return model.x_DWN[n, t, b]  == 0
    else:
        return pyo.Constraint.Skip

model.NoaFRRDwnOutsideLoadShift = pyo.Constraint(model.Nodes, model.Time, model.ShiftableLoadForEnergyCarrier, rule=no_aFRR_dwn_outside_load_shift)

#################################################################################
############## CONNECTING SoC AND UP/DOWN-REGULATION FOR LOADSHIFT ##############
#################################################################################

def aFRR_up_limit_sum_constraint(model, n, i, t, b, e):
    if e == 'Electricity':
        # Extract all time steps `tt` for the given interval `i`
        relevant_times = [tt for (j, tt) in model.TimeLoadShift if j == i]
        if t in relevant_times and t == 1:
            soc_difference = model.Initial_SOC[b]*model.Max_Storage_Capacity[b] - model.q_SoC[n, relevant_times[-1] , b]
            charge_sum = sum(model.Up_Shift_Max * model.Demand[n, t, e] for k in relevant_times if k >= t)
            # Compute the discharge summation for times within the interval and before `t`
            return model.x_UP[n, t, b]/model.Discharge_Efficiency[b] <= soc_difference + charge_sum
        
        elif t in relevant_times and t > 1:
            # Get the minimum time in the interval
            soc_difference = model.q_SoC[n, t-1, b] - model.q_SoC[n, relevant_times[-1] , b]
            charge_sum = sum(model.Up_Shift_Max * model.Demand[n, t, e] for k in relevant_times if k >= t)
            # Compute the discharge summation for times within the interval and before `t`
            return model.x_UP[n, t, b]/model.Discharge_Efficiency[b] <= soc_difference + charge_sum
            
        return pyo.Constraint.Skip
    
    return pyo.Constraint.Skip

model.aFRRUpLimitLoadShift = pyo.Constraint(model.Nodes, model.TimeLoadShift, model.ShiftableLoadForEnergyCarrier, rule=aFRR_up_limit_sum_constraint)



def aFRR_dwn_limit_sum_constraint(model, n, i, t, b, e):
    if e == 'Electricity':
        # Extract all time steps `tt` for the given interval `i`
        relevant_times = [tt for (j, tt) in model.TimeLoadShift if j == i]
        if t in relevant_times and t == 1:
            # Get the minimum time in the interval
            soc_difference =  model.q_SoC[n, relevant_times[-1] , b] - model.Initial_SOC[b]*model.Max_Storage_Capacity[b] 
            charge_sum = sum(model.Up_Shift_Max * model.Demand[n, t, e] for k in relevant_times if k >= t)
            # Compute the discharge summation for times within the interval and before `t`
            return model.x_DWN[n, t, b] <= soc_difference + charge_sum
        
        elif t in relevant_times and t > 1:
            # Get the minimum time in the interval
            soc_difference =  model.q_SoC[n, relevant_times[-1] , b] - model.q_SoC[n, t-1, b] 
            charge_sum = sum(model.Up_Shift_Max * model.Demand[n, t, e] for k in relevant_times if k >= t)
            # Compute the discharge summation for times within the interval and before `t`
            return model.x_DWN[n, t, b] <= soc_difference + charge_sum
            
        return pyo.Constraint.Skip
    
    return pyo.Constraint.Skip

model.aFRRDownLimitLoadShift = pyo.Constraint(model.Nodes, model.TimeLoadShift, model.ShiftableLoadForEnergyCarrier, rule=aFRR_dwn_limit_sum_constraint)



def aFRR_limit(model, n, t, b, e):
    if e == 'Electricity' and (b,e) not in model.ShiftableLoadForEnergyCarrier:
        return model.x_DWN[n, t, b] + model.x_UP[n, t, b]/model.Discharge_Efficiency[b] <= model.Max_charge_discharge_rate[b] + model.Energy2Power_Ratio[b] * model.v_new_bat[b]
    else:
        return pyo.Constraint.Skip

model.aFRRLimit = pyo.Constraint(model.Nodes, model.Time, model.FlexibleLoadForEnergyCarrier, rule=aFRR_limit)

#################################################################################
############## ENSURE STORAGE CAPACITY UP/-DOWN REGULATION ######################
#################################################################################

def ensure_storage_capacity_up_regulation(model, n, t, b, e):
    if e == 'Electricity' and t > 1 and (b,e) not in model.ShiftableLoadForEnergyCarrier: 
        return model.q_SoC[n, t-1, b] - model.x_UP[n, t, b] >= 0
    elif e == 'Electricity' and t == 1 and (b,e) not in model.ShiftableLoadForEnergyCarrier:
        return model.Initial_SOC[b]*(model.Max_Storage_Capacity[b] + model.v_new_bat[b]) >= model.x_UP[n, t, b]
    else:
        return pyo.Constraint.Skip
    
model.EnsureStorageCapacityUpRegulation = pyo.Constraint(model.Nodes, model.Time, model.FlexibleLoadForEnergyCarrier, rule=ensure_storage_capacity_up_regulation)


def ensure_storage_capacity_down_regulation(model, n, t, b, e):
    if e == 'Electricity' and t > 1 and (b,e) not in model.ShiftableLoadForEnergyCarrier:  
        return model.q_SoC[n, t-1, b] - (model.Max_Storage_Capacity[b] + model.v_new_bat[b]) + model.x_DWN[n, t, b]  <= 0
    elif e == 'Electricity' and t == 1 and (b,e) not in model.ShiftableLoadForEnergyCarrier:
        return (model.Max_Storage_Capacity[b] + model.v_new_bat[b])-model.Initial_SOC[b]*(model.Max_Storage_Capacity[b] + model.v_new_bat[b]) >= model.x_DWN[n, t, b]
    else:
        return pyo.Constraint.Skip
    
model.EnsureStorageCapacityDownRegulation = pyo.Constraint(model.Nodes, model.Time, model.FlexibleLoadForEnergyCarrier, rule=ensure_storage_capacity_down_regulation)


###########################################################
############## RESERVE MARKET ACTIVATION ##################
###########################################################

def up_regulation_activation(model, n, t, b, e):
    if e == 'Electricity':
        return model.Activation_Factor_UP_Regulation[n, t] * model.x_UP[n, t, b] <= model.q_discharge[n, t, b]
    else:
        return pyo.Constraint.Skip

model.UpRegulationActivation = pyo.Constraint(model.Nodes, model.Time, model.FlexibleLoadForEnergyCarrier, rule=up_regulation_activation)


def down_regulation_activation(model, n, t, b, e):
    if e == 'Electricity':
        return model.Activation_Factor_DWN_Regulation[n, t] * model.x_DWN[n, t, b] <= model.Charge_Efficiency[b] * model.q_charge[n, t, b]
    else:
        return pyo.Constraint.Skip

model.DownRegulationActivation = pyo.Constraint(model.Nodes, model.Time, model.FlexibleLoadForEnergyCarrier, rule=down_regulation_activation)



########################################################################
############## FLEXIBLE ASSET CONSTRAINTS/STORAGE DYNAMICS #############
########################################################################
def flexible_asset_charge_discharge_limit(model, n, t, b, e):
    if (b,e) not in model.ShiftableLoadForEnergyCarrier:
        return (
            model.q_charge[n, t, b] 
            + model.q_discharge[n, t, b] / model.Discharge_Efficiency[b] 
            <= model.Max_charge_discharge_rate[b] + model.Energy2Power_Ratio[b] * model.v_new_bat[b]
        )
    else:
        return pyo.Constraint.Skip

model.FlexibleAssetChargeDischargeLimit = pyo.Constraint(model.Nodes, model.Time, model.FlexibleLoadForEnergyCarrier, rule=flexible_asset_charge_discharge_limit)


def state_of_charge(model, n, t, b, e):
    if t == 1:
        # Initialisation of flexible assets
        return (
            model.q_SoC[n, t, b]
            == model.Initial_SOC[b] * (model.Max_Storage_Capacity[b] + model.v_new_bat[b]) * (1 - model.Self_Discharge[b])
            + model.q_charge[n, t, b]
            - model.q_discharge[n, t, b] / model.Discharge_Efficiency[b]
        )
    else:
        # Storage Dynamics
        return (
            model.q_SoC[n, t, b]
            == model.q_SoC[n, t-1, b] * (1 - model.Self_Discharge[b])
            + model.q_charge[n, t, b]
            - model.q_discharge[n, t, b] / model.Discharge_Efficiency[b]
        )

model.StateOfCharge = pyo.Constraint(model.Nodes, model.Time, model.FlexibleLoadForEnergyCarrier, rule=state_of_charge)

def end_of_horizon_SoC(model, n, t, b, e):
    if t == model.Time.last():
        return model.q_SoC[n, t, b] == model.Initial_SOC[b] * (model.Max_Storage_Capacity[b] + model.v_new_bat[b])
    else:
        return pyo.Constraint.Skip
model.EndOfHorizonSoC = pyo.Constraint(model.Nodes, model.Time, model.FlexibleLoadForEnergyCarrier, rule = end_of_horizon_SoC)


def flexible_asset_energy_limit(model, n, t, b, e):
    return model.q_SoC[n, t, b] <= model.Max_Storage_Capacity[b] + model.v_new_bat[b]

model.FlexibleAssetEnergyLimits = pyo.Constraint(model.Nodes, model.Time, model.FlexibleLoadForEnergyCarrier, rule=flexible_asset_energy_limit)



####################################################
############## AVAILABILITY CONSTRAINT #############
####################################################

def supply_limitation(model, n, t, i, e):
    return (sum(sum(model.y_out[n, t, i, e, o] for o in model.Mode_of_operation) for e in model.EnergyCarrier if (i,e) in model.TechnologyToEnergyCarrier)  
            <= model.Availability_Factor[n, t, i] * (model.Initial_Installed_Capacity[i] + model.v_new_tech[i]))

model.SupplyLimitation = pyo.Constraint(model.Nodes, model.Time, model.TechnologyToEnergyCarrier, rule=supply_limitation)


##############################################################
############## EXPORT LIMITATION AND GRID TARIFF #############
##############################################################
def export_limitation(model, n, t, e):
    if e == 'Electricity':
        return model.z_export[n, t, e] <= model.Max_Export
    else:
        return pyo.Constraint.Skip

model.ExportLimitation = pyo.Constraint(model.Nodes, model.Time, model.EnergyCarrier, rule=export_limitation)

def peak_load(model, n, t, i, e, o, m):
    if i == 'Power_Grid' and e == 'Electricity':
        return (model.y_out[n, t, i, e, o] <= model.y_max[n, m])
    else:
        return pyo.Constraint.Skip

model.PeakLoad = pyo.Constraint(model.Nodes, model.TimeInMonth, model.TechnologyToEnergyCarrier, model.Mode_of_operation, model.Month, rule=peak_load)

##############################################################
##################### INVESTMENT LIMITATIONS #################
##############################################################
def CAPEX_technology_limitations(model, i):
    return (model.Cost_Expansion_Tec[i] * model.v_new_tech[i] <= model.Max_CAPEX_tech[i])

model.CAPEXTechnologyLim = pyo.Constraint(model.Technology, rule=CAPEX_technology_limitations)

def CAPEX_flexibleLoad_limitations(model, b):
    return (model.Cost_Expansion_Bat[b] * model.v_new_bat[b] <= model.Max_CAPEX_flex[b])

model.CAPEXFlexibleLoadLim = pyo.Constraint(model.FlexibleLoad, rule=CAPEX_flexibleLoad_limitations)

##############################################################
##################### CARBON EMISSION LIMIT ##################
##############################################################
def Carbon_Emission_Limit(model, n):
    return(
    sum(
        sum(
            sum(
                sum(model.y_activity[n, t, i, o] * model.Carbon_Intensity[i, e] 
                    for o in model.Mode_of_operation
                    )
                    for i in model.Technology if (i,e) in model.TechnologyToEnergyCarrier
            )
            for e in model.EnergyCarrier
        )
        for t in model.Time

    ) <= model.Max_Carbon_Emission
    )
    

model.CarbonEmissionLimit = pyo.Constraint(model.Nodes, rule=Carbon_Emission_Limit)

##############################################################
##################### NON-ANTICIPATIVITY #####################
##############################################################
def Day_ahead_to_Intraday(model, n, t, p):
    return (model.x_DA[n, t] == model.x_DA[p, t])

model.DayAheadToIntraday = pyo.Constraint(model.Nodes_ID, model.Time, model.Parent_node, rule=Day_ahead_to_Intraday)

def Day_ahead_to_Intraday_to_RealTime(model, n, t, p):
    return (model.x_DA[n, t] == model.x_DA[p, t])

model.DayAheadToIntradayToRealTime = pyo.Constraint(model.Nodes_RT, model.Time, model.Parent_node, rule=Day_ahead_to_Intraday_to_RealTime)

def Intraday_to_RealTime_Up(model, n, t, p):
    return (model.x_ID_Up[n, t] == model.x_ID_Up[p, t])

model.IntradayToRealTimeUp = pyo.Constraint(model.Nodes_RT, model.Time, model.Parent_node, rule=Intraday_to_RealTime_Up)

def Intraday_to_RealTime_Dwn(model, n, t, p):
    return (model.x_ID_Dwn[n, t] == model.x_ID_Dwn[p, t])

model.IntradayToRealTimeDown = pyo.Constraint(model.Nodes_RT, model.Time, model.Parent_node, rule=Intraday_to_RealTime_Dwn)

def Reserve_Capacity_Dwn_to_ID(model, n, t, b, p):
    return (model.x_DWN[n, t, b] == model.x_DWN[p, t, b])

model.DayAheadToIntraday = pyo.Constraint(model.Nodes_ID, model.Time, model.FlexibleLoad, model.Parent_Node, rule = Reserve_Capacity_Dwn_to_ID) 

def Reserve_Capacity_Dwn_to_RT(model, n, t, b, p):
    return (model.x_DWN[n, t, b] == model.x_DWN[p, t])

model.DayAheadToIntradayToRealTime = pyo.Constraint(model.Nodes_RT, model.Time, model.FlexibleLoad, model.Parent_node, rule=Reserve_Capacity_Dwn_to_RT)

def Reserve_Capacity_Up_to_ID(model, n, t, b, p):
    return (model.x_UP[n, t, b] == model.x_UP[p, t, b])

model.DayAheadToIntraday = pyo.Constraint(model.Nodes_ID, model.Time, model.FlexibleLoad, model.Parent_Node, rule = Reserve_Capacity_Dwn_to_ID) 

def Reserve_Capacity_Up_to_RT(model, n, t, b, p):
    return (model.x_UP[n, t, b] == model.x_UP[p, t])

model.DayAheadToIntradayToRealTime = pyo.Constraint(model.Nodes_RT, model.Time, model.FlexibleLoad, model.Parent_node, rule=Reserve_Capacity_Dwn_to_RT)

"""
# Start the timer
start_time = time.time()

# Solve the problem
opt = SolverFactory("gurobi", Verbose=True)
results = opt.solve(our_model)

# Stop the timer
end_time = time.time()

# Calculate and print the elapsed time
elapsed_time = end_time - start_time
print(f"Optimization problem solved in {elapsed_time:.2f} seconds")
"""


"""
MATCHING DATA FROM CASE WITH MATHEMATICAL MODEL AND PRINTING DATA
"""
our_model = model.create_instance(data)   
our_model.dual = pyo.Suffix(direction=pyo.Suffix.IMPORT) #Import dual values into solver results
#import pdb; pdb.set_trace()

#start the timer
start_time = time.time()
"""
SOLVING PROBLEM
"""
opt = SolverFactory("gurobi", Verbose=True)
#opt.options['LogFile'] = 'gurobi_log.txt'
results = opt.solve(our_model, tee=True)

#stop the timer
end_time = time.time()
running_time = end_time - start_time

"""
DISPLAY RESULTS??
"""

our_model.display('results.csv')
our_model.dual.display()
print("-" * 70)
print("Objective and running time:")
print(f"Objective value for this mongo model is: {round(pyo.value(our_model.Objective),2)}")
print(f"The instance was solved in {round(running_time, 4)} seconds🙂")
print("-" * 70)
print("Hardware details:")
print(f"Processor: {platform.processor()}")
print(f"Machine: {platform.machine()}")
print(f"System: {platform.system()} {platform.release()}")
print(f"CPU Cores: {psutil.cpu_count(logical=True)} (Logical), {psutil.cpu_count(logical=False)} (Physical)")
print(f"Total Memory: {psutil.virtual_memory().total / 1e9:.2f} GB")


print("-" * 70)
import pdb; pdb.set_trace()
"""
EXTRACT VALUE OF VARIABLES AND WRITE THEM INTO EXCEL FILE
"""
"""
def save_results_to_excel(model_instance):
    # Create an Excel writer object
    with pd.ExcelWriter('Variable_results.xlsx', engine='xlsxwriter') as writer:
        
        # Storing x_aFRR_UP results
        x_aFRR_UP_data = [(t, pyo.value(model_instance.x_aFRR_UP[t])) for t in model_instance.Time]
        x_aFRR_UP_df = pd.DataFrame(x_aFRR_UP_data, columns=['Time', 'x_aFRR_UP'])
        x_aFRR_UP_df.to_excel(writer, sheet_name='x_aFRR_UP', index=False)
        
        # Storing x_aFRR_DWN results
        x_aFRR_DWN_data = [(t, pyo.value(model_instance.x_aFRR_DWN[t])) for t in model_instance.Time]
        x_aFRR_DWN_df = pd.DataFrame(x_aFRR_DWN_data, columns=['Time', 'x_aFRR_DWN'])
        x_aFRR_DWN_df.to_excel(writer, sheet_name='x_aFRR_DWN', index=False)
        
        # Storing y_supply results
        y_supply_data = [(s, t, i, pyo.value(model_instance.y_supply[s, t, i])) for s in model_instance.Nodess for t in model_instance.Time for i in model_instance.ElectricitySources]
        y_supply_df = pd.DataFrame(y_supply_data, columns=['Nodes', 'Time', 'EnergySource', 'y_supply'])
        y_supply_df.to_excel(writer, sheet_name='y_supply', index=False)
        
        # Storing z_export results
        z_export_data = [(s, t, pyo.value(model_instance.z_export[s, t])) for s in model_instance.Nodess for t in model_instance.Time]
        z_export_df = pd.DataFrame(z_export_data, columns=['Nodes', 'Time', 'z_export'])
        z_export_df.to_excel(writer, sheet_name='z_export', index=False)
        
        # Storing q_charge results
        q_charge_data = [(s, t, b, pyo.value(model_instance.q_charge[s, t, b])) for s in model_instance.Nodess for t in model_instance.Time for b in model_instance.FlexibleLoads]
        q_charge_df = pd.DataFrame(q_charge_data, columns=['Nodes', 'Time', 'FlexibleLoad', 'q_charge'])
        q_charge_df.to_excel(writer, sheet_name='q_charge', index=False)
        
        # Storing q_discharge results
        q_discharge_data = [(s, t, b, pyo.value(model_instance.q_discharge[s, t, b])) for s in model_instance.Nodess for t in model_instance.Time for b in model_instance.FlexibleLoads]
        q_discharge_df = pd.DataFrame(q_discharge_data, columns=['Nodes', 'Time', 'FlexibleLoad', 'q_discharge'])
        q_discharge_df.to_excel(writer, sheet_name='q_discharge', index=False)
        
        # Storing e_stored results (for flexible loads)
        e_stored_data = [(t, b, pyo.value(model_instance.e_stored[t, b])) for t in model_instance.Time for b in model_instance.FlexibleLoads]
        e_stored_df = pd.DataFrame(e_stored_data, columns=['Time', 'FlexibleLoad', 'e_stored'])
        e_stored_df.to_excel(writer, sheet_name='e_stored', index=False)

# Call the function to save the results
#save_results_to_excel(our_model)


print("Results saved to Variable_Results.xlsx")
"""

def save_results_to_excel(model_instance, filename="Variable_Results.xlsx"):
    """
    Saves Pyomo variable results into an Excel file with filtered output.
    Only includes rows with non-zero or non-null values for variables.
    """
    import pandas as pd
    from pyomo.environ import value

    # Create an Excel writer object
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        # Loop over all active variables in the model
        for var in model_instance.component_objects(pyo.Var, active=True):
            var_name = var.name  # Get the variable name
            var_data = []
            
            # Collect variable data
            for index in var:
                try:
                    var_value = value(var[index])  # Safely get the variable value
                except ValueError:
                    var_value = 0  # If uninitialized, set to None
                if var_value:  # Include only non-zero and non-null values
                    var_data.append((index, var_value))
            
            # Transform data into a DataFrame
            if var_data:  # Only proceed if there is data
                df = pd.DataFrame(var_data, columns=["Index", var_name])
                
                # Dynamically unpack the indices into separate columns
                max_index_length = max(len(idx) if isinstance(idx, tuple) else 1 for idx, _ in var_data)
                unpacked_indices = pd.DataFrame(
                    [list(index) + [None] * (max_index_length - len(index)) if isinstance(index, tuple) else [index] for index, _ in var_data]
                )
                
                # Add unpacked indices to the DataFrame
                unpacked_indices.columns = [f"Index_{i+1}" for i in range(max_index_length)]
                df = pd.concat([unpacked_indices, df[var_name]], axis=1)
                
                # Write the filtered DataFrame to an Excel sheet
                df.to_excel(writer, sheet_name=var_name[:31], index=False)
    
    print(f"Variable results saved to {filename}")

# Usage after solving the model
save_results_to_excel(our_model, filename="Variable_Results.xlsx")

"""

def plot_results_from_excel(input_file, output_folder):
    
    os.makedirs(output_folder, exist_ok=True)  # Create folder if it doesn't exist

    # Define mapping for Index_1 to Nodes names
    Nodes_mapping = {1: "Nodes 1", 2: "Nodes 2", 3: "Nodes 3"}

    # Read the Excel file
    excel_file = pd.ExcelFile(input_file)

    for sheet_name in excel_file.sheet_names:
        # Read the sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        if sheet_name in ["x_aFRR_DWN", "x_aFRR_UP"]:
            # For x_aFRR_DWN and x_aFRR_UP, plot Index_1 vs. second column
            x_axis = df["Index_1"]
            y_axis = df.iloc[:, 1]  # Second column

            plt.figure(figsize=(12, 8))
            plt.plot(x_axis, y_axis, label=sheet_name, marker='o')

            # Add title and labels
            plt.title(f"{sheet_name}")
            plt.xlabel("Hours")
            plt.ylabel("Values")
            plt.legend(loc='best')  # Place legend inside the plot
            plt.grid(True)

            # Save the plot
            plot_filename = f"{sheet_name}.png"
            plt.savefig(os.path.join(output_folder, plot_filename))
            plt.tight_layout()
            plt.close()  # Close the plot to avoid overlapping

            print(f"Single plot saved for {sheet_name} in {output_folder}")

        elif sheet_name in ["x_aFRR_DWN_ind", "x_aFRR_UP_ind"]:
            # For x_aFRR_DWN_ind and x_aFRR_UP_ind, plot Index_1 vs. last column for each unique Index_2
            if "Index_1" in df.columns and "Index_2" in df.columns:
                plt.figure(figsize=(12, 8))

                x_axis = df["Index_1"]
                value_column = df.columns[-1]  # Last column contains the values
                unique_variables = df["Index_2"].unique()

                for variable in unique_variables:
                    variable_data = df[df["Index_2"] == variable]
                    plt.plot(
                        variable_data["Index_1"], variable_data[value_column],
                        label=variable, marker='o'
                    )

                # Add title and labels
                plt.title(f"{sheet_name}")
                plt.xlabel("Hours")
                plt.ylabel("Values")
                plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, title="Variables", borderaxespad=0.)
                plt.grid(True)

                # Save the plot
                plot_filename = f"{sheet_name}.png"
                plt.tight_layout()
                plt.savefig(os.path.join(output_folder, plot_filename))
                plt.close()  # Close the plot to avoid overlapping

                print(f"Plot saved for {sheet_name} in {output_folder}")

        else:
            # Default behavior for all other sheets
            if "Index_1" in df.columns and "Index_2" in df.columns:
                unique_index_1 = df["Index_1"].unique()

                # Create plots for each Index_1 value
                for index_1_value in unique_index_1:
                    filtered_df = df[df["Index_1"] == index_1_value]
                    x_axis = filtered_df["Index_2"]

                    plt.figure(figsize=(12, 8))

                    if "Index_3" in filtered_df.columns:
                        variable_column = "Index_3"
                        value_column = df.columns[-1]
                        unique_variables = filtered_df[variable_column].unique()

                        for variable in unique_variables:
                            variable_data = filtered_df[filtered_df[variable_column] == variable]
                            plt.plot(
                                variable_data["Index_2"], variable_data[value_column],
                                label=variable, marker='o'
                            )

                        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, title="Variables", borderaxespad=0.)
                    else:
                        value_column = df.columns[-1]
                        plt.plot(filtered_df["Index_2"], filtered_df[value_column], label=value_column, marker='o')

                    # Use Nodes mapping for title
                    Nodes_name = Nodes_mapping.get(index_1_value, f"Index_1 = {index_1_value}")
                    plt.title(f"{sheet_name} ({Nodes_name})")
                    plt.xlabel("Hours")
                    plt.ylabel("Values")
                    plt.grid(True)

                    plot_filename = f"{sheet_name}_{Nodes_name.replace(' ', '_')}.png"
                    plt.tight_layout()
                    plt.savefig(os.path.join(output_folder, plot_filename))
                    plt.close()

                    print(f"Plot saved for {sheet_name} ({Nodes_name}) in {output_folder}")



# Usage
if __name__ == "__main__":
    input_excel_file = "Variable_Results.xlsx"  # Path to the Excel file
    output_plots_folder = "plots"  # Folder to save the plots

    # Generate plots
    plot_results_from_excel(input_excel_file, output_plots_folder)

"""

import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.cm import get_cmap
import numpy as np


def generate_unique_colors(num_colors, colormap='tab20'):
    """Generate a list of unique colors."""
    cmap = get_cmap(colormap)
    return [cmap(i / num_colors) for i in range(num_colors)]


def plot_results_from_excel(input_file, output_folder):
    os.makedirs(output_folder, exist_ok=True)  # Create folder if it doesn't exist

    # Define mapping for Index_1 to Nodes names
    Nodes_mapping = {1: "Nodes 1", 2: "Nodes 2", 3: "Nodes 3"}

    # Read the Excel file
    excel_file = pd.ExcelFile(input_file)

    for sheet_name in excel_file.sheet_names:
        # Read the sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        if sheet_name in ["x_aFRR_DWN", "x_aFRR_UP"]:
            # For x_aFRR_DWN and x_aFRR_UP, plot Index_1 vs. second column
            x_axis = df["Index_1"]
            y_axis = df.iloc[:, 1]  # Second column

            plt.figure(figsize=(12, 8))
            plt.plot(x_axis, y_axis, label=sheet_name, marker='o', color='blue')

            plt.title(f"{sheet_name}")
            plt.xlabel("Hours")
            plt.ylabel("Values")
            plt.legend(loc='best')
            plt.grid(True)

            # Save the plot
            plot_filename = f"{sheet_name}.png"
            plt.tight_layout()
            plt.savefig(os.path.join(output_folder, plot_filename))
            plt.close()

        elif sheet_name in ["x_aFRR_DWN_ind", "x_aFRR_UP_ind"]:
            if "Index_1" in df.columns and "Index_2" in df.columns:
                plt.figure(figsize=(12, 8))

                x_axis = df["Index_1"]
                value_column = df.columns[-1]  # Last column contains the values
                unique_variables = df["Index_2"].unique()
                colors = generate_unique_colors(len(unique_variables))

                for variable, color in zip(unique_variables, colors):
                    variable_data = df[df["Index_2"] == variable]
                    plt.plot(
                        variable_data["Index_1"], variable_data[value_column],
                        label=variable, marker='o', color=color
                    )

                plt.title(f"{sheet_name}")
                plt.xlabel("Hours")
                plt.ylabel("Values")
                plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, title="Variables", borderaxespad=0.)
                plt.grid(True)

                # Save the plot
                plot_filename = f"{sheet_name}.png"
                plt.tight_layout()
                plt.savefig(os.path.join(output_folder, plot_filename))
                plt.close()

        else:
            if "Index_1" in df.columns and "Index_2" in df.columns:
                unique_index_1 = df["Index_1"].unique()

                for index_1_value in unique_index_1:
                    filtered_df = df[df["Index_1"] == index_1_value]
                    x_axis = filtered_df["Index_2"]

                    plt.figure(figsize=(12, 8))

                    if "Index_3" in filtered_df.columns:
                        variable_column = "Index_3"
                        value_column = df.columns[-1]
                        unique_variables = filtered_df[variable_column].unique()
                        colors = generate_unique_colors(len(unique_variables))

                        for variable, color in zip(unique_variables, colors):
                            variable_data = filtered_df[filtered_df[variable_column] == variable]
                            plt.plot(
                                variable_data["Index_2"], variable_data[value_column],
                                label=variable, marker='o', color=color
                            )

                        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, title="Variables", borderaxespad=0.)
                    else:
                        value_column = df.columns[-1]
                        plt.plot(filtered_df["Index_2"], filtered_df[value_column], label=value_column, marker='o', color='blue')

                    # Use Nodes mapping for title
                    Nodes_name = Nodes_mapping.get(index_1_value, f"Index_1 = {index_1_value}")
                    plt.title(f"{sheet_name} ({Nodes_name})")
                    plt.xlabel("Hours")
                    plt.ylabel("Values")
                    plt.grid(True)

                    plot_filename = f"{sheet_name}_{Nodes_name.replace(' ', '_')}.png"
                    plt.tight_layout()
                    plt.savefig(os.path.join(output_folder, plot_filename))
                    plt.close()
                    



# Usage
if __name__ == "__main__":
    input_excel_file = "Variable_Results.xlsx"  # Path to the Excel file
    output_plots_folder = "plots"  # Folder to save the plots

    # Generate plots
    plot_results_from_excel(input_excel_file, output_plots_folder)


def extract_demand_and_flex_demand(model):
    demand_data = []
    flex_demand_data = []

    for s in model.Nodes:
        for t in model.Time:
            for e in model.EnergyCarrier:
                if e == "Electricity":  # Restrict to Electricity for this use case
                    # Retrieve the values safely
                    demand_value = pyo.value(model.Demand[s, t, e])
                    flex_demand_value = (
                        demand_value
                        + sum(pyo.value(model.q_charge[s, t, b]) for b in model.FlexibleLoad if (b, e) in model.ShiftableLoadForEnergyCarrier)
                        - sum(pyo.value(model.q_discharge[s, t, b]) for b in model.FlexibleLoad if (b, e) in model.ShiftableLoadForEnergyCarrier)
                    )

                    # Append to lists
                    demand_data.append({'Nodes': s, 'Time': t, 'EnergyCarrier': e, 'Demand': demand_value})
                    flex_demand_data.append({'Nodes': s, 'Time': t, 'EnergyCarrier': e, 'flex_demand': flex_demand_value})

    # Convert to DataFrame
    demand_df = pd.DataFrame(demand_data)
    flex_demand_df = pd.DataFrame(flex_demand_data)

    return demand_df, flex_demand_df
 # Get the data
demand_df, flex_demand_df = extract_demand_and_flex_demand(our_model)

# Merge the DataFrames for unified plotting
merged_df = pd.merge(demand_df, flex_demand_df, on=['Nodes', 'Time', 'EnergyCarrier'])

# Plotting
plt.figure(figsize=(12, 6))
for Nodes in merged_df['Nodes'].unique():
    Nodes_data = merged_df[merged_df['Nodes'] == Nodes]
    plt.step(Nodes_data['Time'], Nodes_data['Demand'],label=f'Demand - Nodes {Nodes}')
    plt.step(Nodes_data['Time'], Nodes_data['flex_demand'], "--", label=f'Flex Demand - Nodes {Nodes}')

plt.xlabel('Time')
plt.ylabel('Demand (MW)')
plt.title('Demand and Flexible Demand Over Time')
plt.legend()
plt.grid(True)
plt.show()


"""

# Extract Demand and flex_demand from the model results
def extract_demand_and_flex_demand(model):
    demand_data = []
    flex_demand_data = []

    for s in model.Nodes:
        for t in model.Time:
            for e in model.EnergyCarrier:
                # Retrieve the values
                demand_value = pyo.value(model.Demand[s, t,"Electricity"])
                flex_demand_value = pyo.value(model.Demand[s, t, "Electricity"] + model.q_charge[s, t, "Shiftable_Electricity_Up"] + model.q_charge[s, t, "Shiftable_Electricity_Down"] - model.q_discharge[s, t, "Shiftable_Electricity_Up"] - model.q_discharge[s, t, "Shiftable_Electricity_Down"] )

                # Append to lists
                demand_data.append({'Nodes': 1, 'Time': t, 'EnergyCarrier': 'Electricity', 'Demand': demand_value})
                flex_demand_data.append({'Nodes': 1, 'Time': t, 'EnergyCarrier': 'Electricity', 'flex_demand': flex_demand_value})

    # Convert to DataFrame
    demand_df = pd.DataFrame(demand_data)
    flex_demand_df = pd.DataFrame(flex_demand_data)

    return demand_df, flex_demand_df
# Get the data
demand_df, flex_demand_df = extract_demand_and_flex_demand(our_model)

# Merge the DataFrames for plotting
merged_df = pd.merge(demand_df, flex_demand_df, on=['Nodes', 'Time', 'Electricity'], suffixes=('_Demand', '_flex'))

# Plotting
plt.figure(figsize=(12, 6))

# Group by Nodes and energy carrier for multiple plots


plt.plot('Time', 'Demand', label='Demand - {Nodes} - {"Electricity"}')
plt.plot(group['Time'], group['flex_demand'], '--', label=f'Flex Demand - {Nodes} - {"Electricity"}')

plt.xlabel('Time')
plt.ylabel('Demand')
plt.title('Demand and Flex Demand Over Time')
plt.legend()
plt.grid(True)
plt.show()
"""
