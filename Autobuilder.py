import os
import win32com.client
import openpyxl
import random
from openpyxl import *
import re
import time
import ReadExcel

def build_floor_plan_and_bracing(SapModel, tower, all_floor_plans, all_floor_bracing, floor_num, floor_elev):
    print('Building floor plan...')
    floor_plan_num = tower.floor_plans[floor_num-1]
    floor_plan = all_floor_plans[floor_plan_num-1]
    #Create members for floor plan
    for member in floor_plan.members:
        kip_in_F = 3
        SapModel.SetPresentUnits(kip_in_F)
        start_node = member.start_node
        end_node = member.end_node
        start_x = start_node[0]
        start_y = start_node[1]
        start_z = floor_elev
        end_x = end_node[0]
        end_y = end_node[1]
        end_z = floor_elev
        section_name = member.sec_prop
        [ret, name] = SapModel.FrameObj.AddByCoord(start_x, start_y, start_z, end_x, end_y, end_z, PropName=section_name)
        if ret != 0:
            print('ERROR creating floor plan member on floor ' + str(floor_num))
    #assign masses to mass nodes and create steel rod
    mass_node_1 = floor_plan.mass_nodes[0]
    mass_node_2 = floor_plan.mass_nodes[1]
    floor_mass = tower.floor_masses[floor_num-1]
    mass_per_node = floor_mass/2
    #Create the mass node point
    [ret, mass_name_1] = SapModel.PointObj.AddCartesian(mass_node_1[0],mass_node_1[1],floor_elev,MergeOff=False)
    if ret != 0:
        print('ERROR setting mass nodes on floor ' + str(floor_num))
    [ret, mass_name_2] = SapModel.PointObj.AddCartesian(mass_node_2[0],mass_node_2[1],floor_elev,MergeOff=False)
    if ret != 0:
        print('ERROR setting mass nodes on floor ' + str(floor_num))
    #Assign masses to the mass nodes
    #Shaking in the x direcion!
    N_m_C = 10
    SapModel.SetPresentUnits(N_m_C)
    ret = SapModel.PointObj.SetMass(mass_name_1, [mass_per_node,0,0,0,0,0],0,True,False)
    if ret[0] != 0:
        print('ERROR setting mass on floor ' + str(floor_num))
    ret = SapModel.PointObj.SetMass(mass_name_2, [mass_per_node,0,0,0,0,0])
    if ret[0] != 0:
        print('ERROR setting mass on floor ' + str(floor_num))
    #Create steel rod
    kip_in_F = 3
    SapModel.SetPresentUnits(kip_in_F)
    [ret, name1] = SapModel.FrameObj.AddByCoord(mass_node_1[0], mass_node_1[1], floor_elev, mass_node_2[0], mass_node_2[1], floor_elev, PropName='Steel rod')
    if ret !=0:
        print('ERROR creating steel rod on floor ' + str(floor_num))
    #Create floor load forces
    N_m_C = 10
    SapModel.SetPresentUnits(N_m_C)
    ret = SapModel.PointObj.SetLoadForce(mass_name_1, 'DEAD', [0, 0, mass_per_node*9.81, 0, 0, 0])
    ret = SapModel.PointObj.SetLoadForce(mass_name_2, 'DEAD', [0, 0, mass_per_node*9.81, 0, 0, 0])
    #create floor bracing
    floor_bracing_num = tower.floor_plans[floor_num-1]
    floor_bracing = all_floor_bracing[floor_bracing_num-1]
    #Finding x and y scaling factors:
    all_plan_nodes = []
    for member in floor_plan.members:
        all_plan_nodes.append(member.start_node)
        all_plan_nodes.append(member.end_node)
    #Find max and min x and y coordinates
    max_node_x = 0
    max_node_y = 0
    min_node_x = 0
    min_node_y = 0
    for node in all_plan_nodes:
        if max_node_x < node[0]:
            max_node_x = node[0]
        if max_node_y < node[1]:
            max_node_y = node[1]
        if min_node_x > node[0]:
            min_node_x = node[0]
        if min_node_y > node[1]:
            min_node_y = node[1]
    scaling_x = max_node_x - min_node_x
    scaling_y = max_node_y - min_node_y
    #Create floor bracing
    print('Building floor bracing...')
    for member in floor_bracing.members:
        kip_in_F = 3
        SapModel.SetPresentUnits(kip_in_F)
        start_node = member.start_node
        end_node = member.end_node
        start_x = start_node[0] * scaling_x
        start_y = start_node[1] * scaling_y
        start_z = floor_elev
        end_x = end_node[0] * scaling_x
        end_y = end_node[1] * scaling_y
        end_z = floor_elev
        section_name = member.sec_prop
        [ret, name] = SapModel.FrameObj.AddByCoord(start_x, start_y, start_z, end_x, end_y, end_z, PropName=section_name)
        if ret != 0:
            print('ERROR creating floor bracing member on floor ' + str(floor_num))
    return SapModel



#----START-----------------------------------------------------START----------------------------------------------------#



print('\n--------------------------------------------------------')
print('Autobuilder by University of Toronto Seismic Design Team')
print('--------------------------------------------------------\n')

#Read in the excel workbook
print("\nReading Excel spreadsheet...")
wb = load_workbook('SetupAB.xlsx')
ExcelIndex = ReadExcel.get_excel_indices(wb, 'A', 'B', 2)

Sections = ReadExcel.get_properties(wb,ExcelIndex,'Section')
Materials = ReadExcel.get_properties(wb,ExcelIndex,'Material')
Bracing = ReadExcel.get_bracing(wb,ExcelIndex,'Bracing')
FloorPlans = ReadExcel.get_floor_plans(wb,ExcelIndex)
FloorBracing = ReadExcel.get_bracing(wb,ExcelIndex,'Floor Bracing')
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

TowerNum = 1
for Tower in AllTowers:
    print('\nBuilding tower number ' + str(TowerNum))
    print('-------------------------')
    NumFloors = len(Tower.floor_plans)
    CurFloorNum = 1
    CurFloorElevation = 0
    while CurFloorNum <= NumFloors:
        print('Floor ' + str(CurFloorNum))
        build_floor_plan_and_bracing(SapModel, Tower, FloorPlans, FloorBracing, CurFloorNum, CurFloorElevation)
        #INSERT FUNCTION TO CREATE BRACING AT CURRENT FLOOR

        #INSERT FUNCTION TO CREATE COLUMNS AT CURRENT FLOOR

        CurFloorHeight = Tower.floor_heights[CurFloorNum - 1]
        CurFloorElevation = CurFloorElevation + CurFloorHeight
        CurFloorNum += 1

    TowerNum += 1