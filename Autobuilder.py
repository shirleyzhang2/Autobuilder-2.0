import os
import win32com.client
import openpyxl
import random
from openpyxl import *
import re
import time
import ReadExcel

print('Initializing SAP2000 model...')
# create SAP2000 object
SapObject = win32com.client.Dispatch('SAP2000v15.SapObject')
# start SAP2000
SapObject.ApplicationStart()
# create SapModel Object
SapModel = SapObject.SapModel
# initiaize model
SapModel.InitializeNewModel()
# create new blank model
ret = SapModel.File.NewBlank()

