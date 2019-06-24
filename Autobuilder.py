import os
import win32com.client
import openpyxl
import random
from openpyxl import *
import re
import time
import ReadExcel


#Read in the excel workbook
print("\nReading Excel spreadsheet...")
wb = load_workbook('SetupAB.xlsx')
ExcelIndex = ReadExcel.get_excel_indices(wb, 'A', 'B', 2)

Sections = ReadExcel.get_properties(wb,ExcelIndex,'Section')
Materials = ReadExcel.get_properties(wb,ExcelIndex,'Material')
Nodes = ReadExcel.get_node_info(wb,ExcelIndex,'Bracing')
Bracing = ReadExcel.get_floor_or_bracing(wb,ExcelIndex,'Bracing')
FloorPlans = ReadExcel.get_floor_or_bracing(wb,ExcelIndex,'Floor Plans')
FloorBracing = ReadExcel.get_floor_or_bracing(wb,ExcelIndex,'Floor Bracing')
AllTowers = ReadExcel.read_input_table(wb, ExcelIndex)


print('\nInitializing SAP2000 model...')
# create SAP2000 object
SapObject = win32com.client.Dispatch('SAP2000v15.SapObject')
# start SAP2000
SapObject.ApplicationStart()
# create SapModel Object
SapModel = SapObject.SapModel
# initialize model
SapModel.InitializeNewModel()
# create new blank model
ret = SapModel.File.NewBlank()

#Define new materials
print("\nDefining materials...")
N_m_C = 10
SapModel.SetPresentUnits(N_m_C)
for Material, MatProps in Materials.items():
    MatName = MatProps['Name']
    MatType = MatProps['Material type']
    MatWeight = MatProps['Weight per volume']
    MatE = MatProps['Elastic modulus']
    MatPois = MatProps['Poisson\'s ratio']
    MatTherm = MatProps['Thermal coefficient']
    #Create material type
    ret = SapModel.PropMaterial.SetMaterial(MatName, MatType)
    if ret != 0:
        print('ERROR creating material type')
    #Set isotropic material proprties
    ret = SapModel.PropMaterial.SetMPIsotropic(MatName, MatE, MatPois, MatTherm)
    if ret != 0:
        print('ERROR setting material properties')
    #Set unit weight
    ret = SapModel.PropMaterial.SetWeightAndMass(MatName, 1, MatWeight)
    if ret != 0:
        print('ERROR setting material unit weight')

#Define new sections
print('Defining sections...')
kip_in_F = 3
SapModel.SetPresentUnits(kip_in_F)
for Section, SecProps in Sections.items():
    SecName = SecProps['Name']
    SecArea = SecProps['Area']
    SecTors = SecProps['Torsional constant']
    SecIn3 = SecProps['Moment of inertia about 3 axis']
    SecIn2 = SecProps['Moment of inertia about 2 axis']
    SecSh2 = SecProps['Shear area in 2 direction']
    SecSh3 = SecProps['Shear area in 3 direction']
    SecMod3 = SecProps['Section modulus about 3 axis']
    SecMod2 = SecProps['Section modulus about 2 axis']
    SecPlMod3 = SecProps['Plastic modulus about 3 axis']
    SecPlMod2 = SecProps['Plastic modulus about 2 axis']
    SecRadGy3 = SecProps['Radius of gyration about 3 axis']
    SecRadGy2 = SecProps['Radius of gyration about 2 axis']
    SecMat = SecProps['Material']
    #Create section property
    ret = SapModel.PropFrame.SetGeneral(SecName, SecMat, 0.1, 0.1, SecArea, SecSh2, SecSh3, SecTors, SecIn2, SecIn3, SecMod2, SecMod3, SecPlMod2, SecPlMod3, SecRadGy2, SecRadGy3, -1)
    if ret != 0:
        print('ERROR creating section property ' + SecName)

def get_acc_and_drift(SapObject):
    #Run Analysis
    print('Computing accelaration and drift...')
    SapModel.Analyze.RunAnalysis()
    print('Finished computing.')
    #Get RELATIVE acceleration from node 
    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
    SapModel.Results.Setup.SetComboSelectedForOutput('DEAD + GM', True)
    #set type to envelope
    SapModel.Results.Setup.SetOptionModalHist(1)
    #Get joint acceleration
    #Set units to metres
    N_m_C = 10
    SapModel.SetPresentUnits(N_m_C)
    g = 9.81
    ret = SapModel.Results.JointAccAbs('0-0-0', 0)#for now
    max_and_min_acc = ret[7]
    max_pos_acc = max_and_min_acc[0]
    min_neg_acc = max_and_min_acc[1]
    if abs(max_pos_acc) >= abs(min_neg_acc):
        max_acc = abs(max_pos_acc)/g
    elif abs(min_neg_acc) >= abs(max_pos_acc):
        max_acc = abs(min_neg_acc)/g
    else:
        print('Could not find max acceleration')
    #Get joint displacement
    #Set units to millimetres
    N_mm_C = 9
    SapModel.SetPresentUnits(N_mm_C)
    ret = SapModel.Results.JointDispl('0-0-0', 0)#for now
    max_and_min_disp = ret[7]
    max_pos_disp = max_and_min_disp[0]
    min_neg_disp = max_and_min_disp[1]
    if abs(max_pos_disp) >= abs(min_neg_disp):
        max_drift = abs(max_pos_acc)
    elif abs(min_neg_disp) >= abs(max_pos_disp):
        max_drift = abs(min_neg_disp)
    else:
        print('Could not find max drift')
    #Close SAP2000
    SapObject.ApplicationExit(True)
    return max_acc, max_drift

def print_acc_and_drift(SapObject):
    print('\nAnalyze')
    print('----------------------------------')
    max_acc_and_drift = get_sap_results(SapObject)
    print('Max acceleration is: ' + str(max_acc_and_drift[0]) + ' g')
    print('Max drift is: ' + str(max_acc_and_drift[1]) + ' mm')
    return max_acc_and_drift

def get_weight(SapObject):
    #Run Analysis
    print('Computing weight...')
    SapModel.Analyze.RunAnalysis()
    print('Finished computing.')
    #Get base reactions
    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
    SapModel.Results.Setup.SetCaseSelectedForOutput('DEAD')
    #SapModel.Results.BaseReact(NumberResults, LoadCase, StepType, StepNum, Fx, Fy, Fz, Mx, My, Mz, gx, gy, gz)
    ret = SapModel.Results.BaseReact()
    base_react = ret[7]
    return base_react

def get_FABI(SAPObject):
    results = get_acc_and_drift(SapObject)
    footprint = 96 #inches squared
    weight = get_weight(SapObject) #lb
    design_life = 100 #years
    construction_cost = 2500000*(weight**2)+6*(10**6)
    land_cost = 35000 * footprint
    annual_building_cost = (land_cost + construction_cost) / design_life
    annual_revenue = 430300
    equipment_cost = 20000000
    return_period_1 = 50
    return_period_2 = 300
    max_disp = results[1] #mm
    apeak_1 = results[0] #g's
    xpeak_1 = 100*max_disp/1524 #% roof drift
    structural_damage_1 = scipy.stats.norm(1.5, 0.5).cdf(xpeak_1)
    equipment_damage_1 = scipy.stats.norm(1.75, 0.7).cdf(apeak_1)
    economic_loss_1 = structural_damage_1*construction_cost + equipment_damage_1*equipment_cost
    annual_economic_loss_1 = economic_loss_1/return_period_1
    structural_damage_2 = 0.5
    equipment_damage_2 = 0.5
    economic_loss_2 = structural_damage_2*construction_cost + equipment_damage_2*equipment_cost
    annual_economic_loss_2 = economic_loss_2/return_period_2
    annual_seismic_cost = annual_economic_loss_1 + annual_economic_loss_2
    fabi = annual_revenue - annual_building_cost - annual_seismic_cost
    return fabi

def write_to_excel(SapObject):
    wb = openpyxl.Workbook()
    ws = wb.active
    #ws = wb.create_sheet(title = "FABI")
    ws['A1'] = 'Tower #'
    ws['A2'] = 'FABI'
    for tower in AllTowers:
        col = get_column_letter(tower.number+1)
        ws[col +'1'] = tower.number
        ws[col +'2'] = fabi
    wb.save('C:\\Users\\shirl\\OneDrive - University of Toronto\\Desktop\\Seismic\\FABI.xlsx') 
