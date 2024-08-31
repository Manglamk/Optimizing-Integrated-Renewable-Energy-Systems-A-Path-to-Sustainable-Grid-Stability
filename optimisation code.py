import gurobipy as gp ## access all gurobi functions and classes through a gp. prefix
import pandas as pd # to read the excel file
from gurobipy import GRB ## everything in class GRB available without a prefix
import numpy as np
import matplotlib.pyplot as plt
# Sets
I = ["Q1", "Q2", "Q3", "Q4"] #QUARTERS
J = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66 ,67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109 , 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120 ] # 120 days in a quarter
K = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9,10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66 ,67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95] # 96 instances in a day ; 1 instance = 15 mins
M = ["I1", "I2"] # INDUSTRIES
N = ["R1", "R2", "R3", "R4", "R5"] # RESIDENCIES
# PARAMETRS
# READING EXCEL FILE
DRik = pd.read_excel('projectdata.xlsx', usecols=[4], sheet_name=3)#RESIDENT'S DEMAND
Q1R1 = DRik.head(96)#Q1R1
Q1R2 = DRik.iloc[96:192]#Q1R2
Q1R3 = DRik.iloc[192:288]#Q1R3
Q1R4 = DRik.iloc[288:384]#Q1R4
Q1R5 = DRik.iloc[384:480]#Q1R5
Q2R1 = DRik.iloc[480:576]#Q2R1
Q2R2 = DRik.iloc[576:672]#Q2R2
Q2R3 = DRik.iloc[672:768]#Q2R3
Q2R4 = DRik.iloc[768:864]#Q2R4
Q2R5 = DRik.iloc[864:960]#Q2R5
Q3R1 = DRik.iloc[960:1056]#Q3R1
Q3R2 = DRik.iloc[1056:1152]#Q3R2
Q3R3 = DRik.iloc[1152:1248]#Q3R3
Q3R4 = DRik.iloc[1248:1344]#Q3R4
Q3R5 = DRik.iloc[1344:1440]#Q3R5
Q4R1 = DRik.iloc[1440:1536]#Q4R1
Q4R2 = DRik.iloc[1536:1632]#Q4R2
Q4R3 = DRik.iloc[1632:1728]#Q4R3
Q4R4 = DRik.iloc[1728:1824]#Q4R4
Q4R5 = DRik.iloc[1824:1920]#Q4R5
Q1R1_array = Q1R1.values.flatten()
Q1R2_array = Q1R2.values.flatten()
Q1R3_array = Q1R3.values.flatten()
Q1R4_array = Q1R4.values.flatten()
Q1R5_array = Q1R5.values.flatten()
Q2R1_array = Q2R1.values.flatten()
Q2R2_array = Q2R2.values.flatten()
Q2R3_array = Q2R3.values.flatten()
Q2R4_array = Q2R4.values.flatten()
Q2R5_array = Q2R5.values.flatten()
Q3R1_array = Q3R1.values.flatten()
Q3R2_array = Q3R2.values.flatten()
Q3R3_array = Q3R3.values.flatten()
Q3R4_array = Q3R4.values.flatten()
Q3R5_array = Q3R5.values.flatten()
Q4R1_array = Q4R1.values.flatten()
Q4R2_array = Q4R2.values.flatten()
Q4R3_array = Q4R3.values.flatten()
Q4R4_array = Q4R4.values.flatten()
Q4R5_array = Q4R5.values.flatten()
gamma = {
    "Q1": {
        "R1": Q1R1_array.tolist(),
        "R2": Q1R2_array.tolist(),
        "R3": Q1R3_array.tolist(),
        "R4": Q1R4_array.tolist(),
        "R5": Q1R5_array.tolist()
    },
    "Q2": {
        "R1": Q2R1_array.tolist(),
        "R2": Q2R2_array.tolist(),
        "R3": Q2R3_array.tolist(),
        "R4": Q2R4_array.tolist(),
        "R5": Q2R5_array.tolist()
    },
    "Q3": {
        "R1": Q3R1_array.tolist(),
        "R2": Q3R2_array.tolist(),
        "R3": Q3R3_array.tolist(),
        "R4": Q3R4_array.tolist(),
        "R5": Q3R5_array.tolist()
    },
    "Q4": {
        "R1": Q4R1_array.tolist(),
        "R2": Q4R2_array.tolist(),
        "R3": Q4R3_array.tolist(),
        "R4": Q4R4_array.tolist(),
        "R5": Q4R5_array.tolist()
    }
}
DIEik = pd.read_excel('projectdata.xlsx', usecols=[4], sheet_name=2 )# INDUSTRY'S ELECTRICITY DEMAND
Q1I1=DIEik.iloc[0:96]#Q1I1
Q1I2=DIEik.iloc[96:192]#Q1I2
Q2I1=DIEik.iloc[192:288]#Q2I1
Q2I2 = DIEik.iloc[288:384]#Q2I2
Q3I1 = DIEik.iloc[384:480]#Q3I1
Q3I2 = DIEik.iloc[480:576]#Q3I2
Q4I1 = DIEik.iloc[576:672]#Q4I1
Q4I2 = DIEik.iloc[672:768]#Q4I2


Q1I1_array = Q1I1.values.flatten()
Q1I2_array = Q1I2.values.flatten()
Q2I1_array = Q2I1.values.flatten()
Q2I2_array = Q2I2.values.flatten()
Q3I1_array = Q3I1.values.flatten()
Q3I2_array = Q3I2.values.flatten()
Q4I1_array = Q4I1.values.flatten()
Q4I2_array = Q4I2.values.flatten()

epsilon = {
    "Q1": {
        "I1": Q1I1_array.tolist(),
        "I2": Q1I2_array.tolist()
    },
    "Q2": {
        "I1": Q2I1_array.tolist(),
        "I2": Q2I2_array.tolist()
    },
    "Q3": {
        "I1": Q3I1_array.tolist(),
        "I2": Q3I2_array.tolist()
    },
    "Q4": {
        "I1": Q4I1_array.tolist(),
        "I2": Q4I2_array.tolist()
    }
}

DIGik = pd.read_excel('projectdata.xlsx', usecols=[3], sheet_name=4 )#INDUSTRY'S GAS DEMAND
Q1IG=DIGik.iloc[0:96]#Q1IG
Q2IG=DIGik.iloc[96:192]#Q2IG
Q3IG=DIGik.iloc[192:288]#Q3IG
Q4IG = DIGik.iloc[288:384]#Q4IG
Q1IG_array = Q1IG.values.flatten()
Q2IG_array = Q2IG.values.flatten()
Q3IG_array = Q3IG.values.flatten()
Q4IG_array = Q4IG.values.flatten()
delta = {
    'Q1': Q1IG_array.tolist(),
    'Q2': Q2IG_array.tolist(),
    'Q3': Q3IG_array.tolist(),
    'Q4': Q4IG_array.tolist()
}

Sik = pd.read_excel('projectdata.xlsx', usecols=[3], sheet_name=0 )# SOLAR GENERATION
Q1S=Sik.iloc[0:96]#Q1S
Q2S=Sik.iloc[96:192]#Q2S
Q3S=Sik.iloc[192:288]#Q3S
Q4S = Sik.iloc[288:384]#Q4S
Q1S_array = Q1S.values.flatten()
Q2S_array = Q2S.values.flatten()
Q3S_array = Q3S.values.flatten()
Q4S_array = Q4S.values.flatten()
Q1S_list = Q1S_array.tolist()
Q2S_list = Q2S_array.tolist()
Q3S_list = Q3S_array.tolist()
Q4S_list = Q4S_array.tolist()
alpha = {
    'Q1': Q1S_list,
    'Q2': Q2S_list,
    'Q3': Q3S_list,
    'Q4': Q4S_list
}

Wik = pd.read_excel('projectdata.xlsx', usecols=[3], sheet_name=1 )# WIND GENERATION
Q1W=Wik.iloc[0:96]#Q1W
Q2W=Wik.iloc[96:192]#Q2W
Q3W=Wik.iloc[192:288]#Q3W
Q4W = Wik.iloc[288:384]#Q4W
Q1W_array = Q1W.values.flatten()
Q2W_array = Q2W.values.flatten()
Q3W_array = Q3W.values.flatten()
Q4W_array = Q4W.values.flatten()
Q1W_list = Q1W_array.tolist()
Q2W_list = Q2W_array.tolist()
Q3W_list = Q3W_array.tolist()
Q4W_list = Q4W_array.tolist()
beta = {
    'Q1': Q1W_list,
    'Q2': Q2W_list,
    'Q3': Q3W_list,
    'Q4': Q4W_list
}

lamda = 4000 # COST OF ONE SOLAR PANEL IN DOLLARS
mu = 3000000 # COST OF ONE WIND TURBINE IN DOLLARS
theta = 300 # COST OF STORAGE IN DOLLARS
Cs =  9 # CAPACITY OF ONE SOLAR PANEL IN MJ
Cw = 3600 # CAPACITY OF WIND TURBINE IN MJ

m = gp.Model("bijli")

# variables
BR = m.addVars(I, N, K, vtype=GRB.CONTINUOUS, name="BRijk") # ELECTRICITY SENT TO RESIDENTS IN INK
BIE = m.addVars(I, M, K, vtype=GRB.CONTINUOUS, name="BIEijk") # ELECTRICITY SENT TO INDUSTRIES IN IMK
BE = m.addVars(I, K, vtype=GRB.CONTINUOUS, name="BEijk") # ELECTRCITY SENT TO ELECTROLYZER IN IK
H = m.addVars(I, K, vtype=GRB.CONTINUOUS, name="Hijk") # HYDROGEN PRODUECD BY ELECTROLYZER IN IK
SI = m.addVars(I, J, K, vtype=GRB.CONTINUOUS, name="SIijk") # INTRADAY STORAGE OF HYDROGEN IN IJK
SL = m.addVars(I, J, K, vtype=GRB.CONTINUOUS, name="SLijk") # LIQUIFIED HYDROGEN STORAGE IN IJK
HFC = m.addVars(I, K, vtype=GRB.CONTINUOUS, name="HFCijk") # HYDROGEN SENT TO FUELCELL IN IK
EFC  = m.addVars(I, K, vtype=GRB.CONTINUOUS, name="EFCijk") # ELECTRICITY GENERATED BY FUELCELL IN IK
EFCR  = m.addVars(I, K, vtype=GRB.CONTINUOUS, name="EFCRijk") # ELECTRICITY SENT TO RESIDENT FROM FUELCELL IN IK
EFCIE  = m.addVars(I, K, vtype=GRB.CONTINUOUS, name="EFCIEijk") # ELECTRICITY SENT TO INDUSTRY FROM FUELCELL IN IK
max_Ns = m.addVar(vtype=GRB.CONTINUOUS, name="max_Ns") # NO. OF SOLAR PANELS NEEDED
max_Nw = m.addVar(vtype=GRB.CONTINUOUS, name="max_Nw")# NO. OF WIND TURBINES NEEDED

# DECISION VARIABLES

Ns = m.addVars(I, K, vtype=GRB.INTEGER, name="Ns")  # NO. OF SOLAR PANELS NEEDED IN IK
Nw = m.addVars(I, K, vtype=GRB.INTEGER, name="Nw")  # NO. OF WIND TURBINES NEEDED IN IK
Sh = m.addVar(vtype=GRB.INTEGER, name="Sh")      # HYDROGEN STORAGE

# OBJECTICE FUNCTION

m.setObjective((lamda * max_Ns) + (mu * max_Nw) + (theta * Sh), GRB.MINIMIZE)

#CONSTRAINTS

m.addConstrs((gp.quicksum(BR[i, n, k] for n in N) + gp.quicksum(BIE[i, m, k] for m in M) + BE[i, k] ==  alpha[i][k] * Ns[i, k] + beta[i][k] * Nw[i, k] #1
              for i in I
              for k in K),
              name="continuity of energy")


m.addConstrs(H[i, k]== 0.00388889 * BE[i, k] for i in I for k in K)#2 ELECTRICITY TO HYDROGEN CONVERSION
m.addConstrs(EFC[i, k] == 33 * 3.6 * HFC[i, k] for i in I for k in K)#3 HYDROGEN TO ELECTRCITY CONVERSION
m.addConstrs(H[i, k] - delta[i][k] - HFC[i, k] + SL[i, J[j_ind-1], K[94]] == SI[i, j, k] for i in I for j_ind, j in enumerate(J) for k in K if k == "1")#4 HYDROGEN STORAGE FOR FIRST INSTANCE
m.addConstrs((H[i, k] + SI[i, j, K[k_ind-1]] - delta[i][k] - HFC[i, k] + SL[i, j, K[k_ind-1]] - SI[i, j, k] == 0 
              for i in I 
              for j in J 
              for k_ind, k in enumerate(K) if k != 1), 
              name="constraint_name") #5 HYDROGEN STORAGE FOR REMANING INSTANCES
m.addConstrs(SI[i, J[j_ind-1], K[94]] == SL[i, j, K[0]] for i in I for j_ind, j in enumerate(J))#6 RE;ATION BETWEEN INTRADAY AND LIQUIFIED STORAGE
m.addConstrs((EFCR[i, k] + BR[i, n, k] == gamma[i][n][k] 
              for i in I
              for n in N
              for k in K ), 
              name="constraint_name") #7 RESIDENTIAL ELECTRICITY DEMAND 
m.addConstrs((EFCIE[i, k] + BIE[i, m, k] == epsilon[i][m][k]
              for i in I
              for m in M
              for k in K),
              name="constraint_name")#8 INDUSTRIAL ELECTRICITY DEMAND
m.addConstrs(SI[i, j, k] + SL[i, j , k] == Sh for i in I for j in J for k in K)#9 TOTAL STORAGE 
m.addConstrs(BR[i, n, k] >= 0 for i in I for k in K for n in N)#10
m.addConstrs(BIE[i, m, k] >= 0 for i in I for k in K for m in M)#11
m.addConstrs(BE[i, k] >= 0 for i in I for k in K)#12
m.addConstrs(H[i, k] >= 0 for i in I for k in K)#13
m.addConstrs(HFC[i, k] >= 0 for i in I for k in K)#14
m.addConstrs(EFC[i, k] >= 0 for i in I for k in K)#15
m.addConstrs(EFCR[i, k] >= 0 for i in I for k in K)#16
m.addConstrs(EFCIE[i, k] >= 0 for i in I for k in K)#17
m.addConstrs(SI[i, j, k] >= 0 for i in I for j in J for k in K)#18
m.addConstrs((Ns[i, k] >= 0 for i in I for k in K ), name="Ns_non_negative")#19
m.addConstrs(Cs * Ns[i, k] >= 0.1 * (epsilon[i][m][k] + gamma[i][n][k]) for i in I for m in M for n in N for k in K )#20 ATLEAST 10 PERCENT OF THE DEMAND SATISFIED FROM SOLAR PANELS
m.addConstrs((Ns[i, k] >= 0.01 * alpha[i][k] for i in I for k in K if alpha[i][k] != 0), name="Ns_non_zero") #21 If alpha[i][k] != 0, then Ns[i] != 0
m.addConstrs(SL[i, j, k] >= 0 for i in I for j in J for k in K)#22
m.addConstrs(alpha[i][k] * Ns[i, k] + beta[i][k] * Nw[i, k] + 33*3.6* SI[i, j, k] - gp.quicksum(epsilon[i][m][k] for m in M) - gp.quicksum(gamma[i][n][k] for n in N) >= 0 for i in I for k in K for m in M for n in N if i == 0 if k == 0 if n == 0 if m == 0)#23 DEMAND SATISFACTION FOR IJK == 0
m.addConstrs(alpha[i][k-1] * Ns[i, k] + beta[i][k-1] * Nw[i, k] + 33 * 3.6 * (SI[i, j, K[k_ind-1]] + SL[i, j, K[k_ind-1]]) - gp.quicksum(epsilon[i][m][k-1] for m in M) - gp.quicksum(gamma[i][n][k-1] for n in N) + alpha[i][k] * Ns[i, k] + beta[i][k] * Nw[i, k] + 33*3.6* SI[i, j, k] - epsilon[i][m][k]- gamma[i][n][k] >= 0 for i in I for j in J for n in N for m in M for k_ind, k in enumerate(K) if k != 0 if k_ind > 0 if k != 1)# 24 DEMAND SATISFACTION FOR IJK != 0
m.addConstrs((max_Ns >= Ns[i, k] for i in I for k in K), name="update_max_Ns")#25 
m.addConstrs((max_Nw >= Nw[i, k] for i in I for k in K), name="update_max_Nw") #26

m.optimize()
cost = lamda * max_Nw.x
print("TOTAL NUMBER OF SOLAR PANELS REQUIRED:",max_Ns.x)
print("COST OF TOTAL SOLAR PANEL INSTALLATION:", cost)
print("TOTAL NUMBER OF WIND TURBINES REQUIRED:",max_Nw.x)
print("COST OF TOTAL WIND TURBINES INSTALLATION:", mu * max_Nw.x)
print("TOTAL HYDROGEN STORAGE IN Kg:", Sh.x)
print("TOTAL COST OF HYDROGEN STORAGE IN Kg:", theta * Sh.x)
# VALUE OF OBJECTIVE FUNC
if m.status == gp.GRB.OPTIMAL:
    # Retrieve the objective function value
    obj_value = m.objVal
    print("Optimal Objective Function Value:", obj_value)
else:
    print("The model may not be solved to optimality.")



#PLOT OS Ns WITH INSTANCES
Ns_values = {}  # Dictionary to store Ns values for each quarter
for i in I:
    Ns_values[i] = [Ns[i, k].x for k in K]

# Plotting Ns values against instances for each quarter
plt.figure(figsize=(10, 6))
for i in I:
    plt.plot(range(len(Ns_values[i])), Ns_values[i], label=f'Quarter {i}')

plt.title('Number of Solar Panels (Ns) Over Instances')
plt.xlabel('Instance')
plt.ylabel('Number of Solar Panels')
plt.legend()
plt.grid(True)
plt.show()


#Nw PLOTS
for i in I:
    # Collect Nw values for the current quarter
    Nw_values = [Nw[i, k].x for k in K]

    # Create a new figure for each quarter
    plt.figure(figsize=(8, 5))

    # Plot Nw values
    plt.plot(range(len(Nw_values)), Nw_values, marker='o', linestyle='-')

    plt.title(f'Number of Wind Turbines (Nw) Over Instances - Quarter {i}')
    plt.xlabel('Instance')
    plt.ylabel('Number of Wind Turbines')
    plt.grid(True)
    plt.show()



# Collect Ns and Nw values for each quarter
Ns_values = [sum([Ns[i, k].x for k in K]) for i in I]  # Total number of solar panels for each quarter
Nw_values = [sum([Nw[i, k].x for k in K]) for i in I]  # Total number of wind turbines for each quarter




#BAR GRAPH REPRESENTATION OF Ns AND Nw
# Plotting Ns and Nw values against quarters
plt.figure(figsize=(10, 6))
plt.bar([i for i in I], Ns_values, color='skyblue', label='Number of Solar Panels')
plt.bar([i for i in I], Nw_values, color='orange', label='Number of Wind Turbines', alpha=0.6)
plt.title('Number of Solar Panels (Ns) and Wind Turbines (Nw) vs. Quarters')
plt.xlabel('Quarter')
plt.ylabel('Number')
plt.legend()
plt.grid(True)
plt.show()


#PLOT FOR VARIABLES

BR_values = {i: [sum([BR[i, n, k].x for n in N]) for k in K] for i in I}  # Electricity sent to residences for each quarter
BIE_values = {i: [sum([BIE[i, m, k].x for m in M]) for k in K] for i in I}  # Electricity sent to industries for each quarter
H_values = {i: [H[i, k].x for k in K] for i in I}  # Hydrogen produced by electrolyzer for each quarter

# PLOT OF BR 
plt.figure(figsize=(10, 6))
colors = ['b', 'g', 'r', 'c']  # Define colors for each quarter
for i, color in zip(I, colors):
    plt.plot(range(len(BR_values[i])), BR_values[i], marker='o', linestyle='-', color=color, label=f'Quarter {i}')
plt.title('Electricity Sent to Residences (BR) Over Instances')
plt.xlabel('Instance')
plt.ylabel('Value')
plt.legend()
plt.grid(True)
plt.tight_layout()  # Adjust layout to prevent overlap
plt.show()

# PLOT OF BIE
plt.figure(figsize=(10, 6))
for i, color in zip(I, colors):
    plt.plot(range(len(BIE_values[i])), BIE_values[i], marker='o', linestyle='-', color=color, label=f'Quarter {i}')
plt.title('Electricity Sent to Industries (BIE) Over Instances')
plt.xlabel('Instance')
plt.ylabel('Value')
plt.legend()
plt.grid(True)
plt.tight_layout()  # Adjust layout to prevent overlap
plt.show()

# PLOT OF H
plt.figure(figsize=(10, 6))
for i, color in zip(I, colors):
    plt.plot(range(len(H_values[i])), H_values[i], marker='o', linestyle='-', color=color, label=f'Quarter {i}')
plt.title('Hydrogen Produced by Electrolyzer (H) Over Instances')
plt.xlabel('Instance')
plt.ylabel('Value')
plt.legend()
plt.grid(True)
plt.tight_layout()  # Adjust layout to prevent overlap
plt.show()


