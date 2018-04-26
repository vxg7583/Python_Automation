from sharepoint import SharePointSite
import urllib2
from ntlm3 import HTTPNtlmAuthHandler
import getpass
import subprocess
import pandas as pd
import re
import time
import math
import numpy as np

start_time = time.time()
server_url = 'http://teamrooms.hca.corpad.net/'
site_url = server_url + 'sites/Unit_of_Distinction_Program/'

user = 'hca\\' + getpass.getuser()
pw = getpass.getpass()

## Adding New Test for NTLM
passman = urllib2.HTTPPasswordMgrWithDefaultRealm()
passman.add_password(None, site_url, user, pw)

auth_ntlm = HTTPNtlmAuthHandler.HTTPNtlmAuthHandler(passman)
opener = urllib2.build_opener(auth_ntlm)


##### End



site = SharePointSite(site_url, opener)
sp_list = site.lists['2018 - 1st Qtr UoD Submission Form']

print(sp_list)
print ('Connected to SharePoint...')

data = []
for row in sp_list.rows:
    data.append((row.Created,row.Modified,row.Status,row.Facility,row.DeptNo,row.Department_x0020_Name,
                 row._x0023__x0020_of_x0020_Cases_x00,
                 row.Staffed_x0020_Operating_x0020_Ro,row.Trauma_x0020_Facility_x0020__x00,row.Trauma_x0020_Level,row.Average_x0020_Daily_x0020_Census,
                 row.RN_x002f_Patient_x0020_Ratio_x00,row._x0023__x0020_of_x0020_Beds,row.Primary_x0020_Diagnosis_x0020__x,
                 row.Specialty_x0020_Area,row._x0023__x0020_of_x0020_Pt_x0020_,row._x0023__x0020_of_x0020_Falls_x000,
                 row._x0023__x0020_of_x0020_Falls_x00,row._x0025__x0020_Compliance_x0020__,row._x0025__x0020_of_x0020_catheter_,
                 row._x0025__x0020_reason_x0020_to_x0,row._x0025__x0020_Daily_x0020_Docume,
                 row._x0025__x0020_Discharge_x0020__x,row.productivity_x0020_index_x0020__,row.monthly_x0020_antibiotic_x0020_m,
                 row.monthly_x0020_disposition_x0020_,row.nurse_x0020_leader_x0020_visit_x,row.Q1_x0020__x002d__x0020_courtesy_,
                 row.Q2_x0020__x002d__x0020_nurses_x0,row.Q3_x0020__x002d__x0020_nurses_x0,row.Q4_x0020__x002d__x0020_Call_x002,
                 row.Q11_x0020__x002d__x0020_Bathroom,row.Courtesy_x0020_of_x0020_Nurses_x,row.Nurses_x0020_Listen_x0020__x002d,row.Nurses_x0020_privacy_x0020_conce,
                 row.Attention_x0020_to_x0020_needs_x,row.Inform_x0020_of_x0020_treatment_,row.Information_x0020_needed_x0020_b,row.De_x002d_Escalation_x0020_Traini,
                 row.Supply_x0020_Scanning_x0020_Comp,row.Nursing_x0020_Leader_x0020_Email,row.CNO))#42

#print(data)

# Create Dataframe from SharePoint Data
cols = ['Created','Modified', 'Status', 'Facility','Department Number','Department Name','Number of Cases SS','Number of staffed operating rooms SS','Trauma Facility SS','Trauma Level',
		'Average Daily Census','RN Patient Ratio Days','Number of Beds','Primary Diagnosis SS','Speciality Area','Number of Pt falls Case Analysis','Number of Falls Injury DP',
		'Number of Falls Injury Sharp Report','% Compliance Catheter','% Catheter Utilization','% Reason to insert Catheter','% Daily Documentation Urinary Catheter Monitor',
		'% Discharge MS','Productivity Index ER','Monthly Antibiotic Med to Admin Avg Time ER','Monthly Disposition To Leave Time ER','Nurse Leader Visit MS',#26
		'Q1 Courtest and Respect MS','Q2 Nurse Listen Carefully MS','Q3 Nurse Explains MS','Q4 Call Button MS','Q11 Bathroom MS','Courtesy of Nurses ES',
		'Nurses Listen ES','Nurses Privacy Concern ES','Attetion To Needs ES','Inform of Treatment ES','Information Needed Before Procedure SS','De Escalation Training ES',
		'Supply Scanning Compliance Percentage First Quarter','Nursing Leader Email','CNO']#42



melted = pd.DataFrame(data, columns=cols)
melted['Score'] = ""
melted['Score'] = pd.to_numeric(melted['Score'], errors='coerce').fillna(0)


# SCORING LOGIC

for i in range(0,len(melted)):
    if(melted.iloc[i,14] == 'Medical Surgical'):

        if(melted.iloc[i,15] == melted.iloc[i,16] == melted.iloc[i,17]):
            melted.iloc[i,42] =  1.00
        elif(melted.iloc[i,17] > melted.iloc[i,15] and melted.iloc[i,17] > melted.iloc[i,16]):
            melted.iloc[i,42] =  0.00
        else:
            melted.iloc[i,42] =  0.00

        if(melted.iloc[i,18] > 90):
            melted.iloc[i,42] =  1.00

        if(melted.iloc[i,26]  == 100):
            melted.iloc[i,42] =  1.00

        if(melted.iloc[i,27] > 75):
            melted.iloc[i,42] =  0.50
        elif(melted.iloc[i,27] >= 50 and melted.iloc[i,27] <= 75):
            melted.iloc[i,42] =  0.25
        else:
            melted.iloc[i,42] =  0.00


        if(melted.iloc[i,28] > 75):
            melted.iloc[i,42] =  0.50
        elif(melted.iloc[i,28] >= 50 and melted.iloc[i,28] <= 75):
            melted.iloc[i,42] =  0.25
        else:
            melted.iloc[i,42] =  0.00

        if(melted.iloc[i,29] > 75):
            melted.iloc[i,42] =  0.50
        elif(melted.iloc[i,29] >= 50 and melted.iloc[i,29] <= 75):
            melted.iloc[i,42] =  0.25
        else:
            melted.iloc[i,42] =  0.00




        if(melted.iloc[i,30] > 75):
            melted.iloc[i,42] =  0.50
        elif(melted.iloc[i,30] >= 50 and melted.iloc[i,30] <= 75):
            melted.iloc[i,42] =  0.25
        else:
            melted.iloc[i,42] =  0.00


    if(melted.iloc[i,14] == 'Emergency Services'):
        #print('im in')
        if(melted.iloc[i,15] == melted.iloc[i,16] == melted.iloc[i,17]):
            melted.iloc[i,42] =  1.00
        elif(melted.iloc[i,17] > melted.iloc[i,15] and melted.iloc[i,17] > melted.iloc[i,16]):
            melted.iloc[i,42] =  0.00
        else:
            melted.iloc[i,42] =  0.00

        if(melted.iloc[i,23] > 98):
            melted.iloc[i,42] =  1.00

        if(melted.iloc[i,24]  < 30):
            melted.iloc[i,42] =  1.00

        if(melted.iloc[i,25]  < 25):
            melted.iloc[i,42] =  1.00

        if(melted.iloc[i,38]  > 50):
            melted.iloc[i,42] =  1.00

        if(melted.iloc[i,27] > 75):
            melted.iloc[i,42] =  0.50
        elif(melted.iloc[i,27] >= 50 and melted.iloc[i,27] <= 75):
            melted.iloc[i,42] =  0.25
        else:
            melted.iloc[i,42] =  0.00


        if(melted.iloc[i,28] > 75):
            melted.iloc[i,42] =  0.50
        elif(melted.iloc[i,28] >= 50 and melted.iloc[i,28] <= 75):
            melted.iloc[i,42] =  0.25
        else:
            melted.iloc[i,42] =  0.00

        if(melted.iloc[i,29] > 75):
            melted.iloc[i,42] =  0.50
        elif(melted.iloc[i,29] >= 50 and melted.iloc[i,29] <= 75):
            melted.iloc[i,42] =  0.25
        else:
            melted.iloc[i,42] =  0.00




        if(melted.iloc[i,30] > 75):
            melted.iloc[i,42] =  0.50
        elif(melted.iloc[i,30] >= 50 and melted.iloc[i,30] <= 75):
            melted.iloc[i,42] =  0.25
        else:
            melted.iloc[i,42] =  0.00


    if(melted.iloc[i,14] == 'ICU' or melted.iloc[i,14] == 'PCU' or melted.iloc[i,14].lower() == 'icu' or melted.iloc[i,14].lower() == 'pcu'):
        #print('im in')
        if(melted.iloc[i,15] == melted.iloc[i,16] == melted.iloc[i,17]):
            melted.iloc[i,42] =  1.00
        elif((melted.iloc[i,17] > melted.iloc[i,15]) and (melted.iloc[i,17] > melted.iloc[i,16])):
            melted.iloc[i,42] =  0.00
        else:
            melted.iloc[i,42] =  0.00

        if(melted.iloc[i,21] > 90):
            melted.iloc[i,42] =  0.25


    if(melted.iloc[i,14] == 'Surgical Services'):
        if(melted.iloc[i,15] == melted.iloc[i,16] == melted.iloc[i,17]):
            melted.iloc[i,42] =  1.0
        elif(melted.iloc[i,17] > melted.iloc[i,15] and melted.iloc[i,17] > melted.iloc[i,16]):
            melted.iloc[i,42] =  0.0
        else:
            melted.iloc[i,42] =  0.0





melted = pd.DataFrame(pd.melt(melted, id_vars=['Created','Modified', 'Status', 'Facility','Department Name','Nursing Leader Email','CNO','Department Number',
    'Primary Diagnosis SS','Speciality Area','Trauma Level']
                , var_name="Measure", value_name="Value"))

# LOGIC REGEX

pattern1 = r"\'.*\'"
pattern2 = r"\'\S+@\S+\'"

# # # # # Cleaning Data

# Converting to string --> optional
melted['Facility'] = melted['Facility'].astype(str)
melted['CNO'] = melted['CNO'].astype(str)
melted['Nursing Leader Email'] = melted['Nursing Leader Email'].astype(str)

# loops

for i in range(0,len(melted)):
    words = re.findall(pattern1, melted.iloc[i,3])
    if words:
        melted.iloc[i,3] = words[0]


for i in range(0,len(melted)):
    words1 = re.findall(pattern2, melted.iloc[i,6])
    if words1:
        melted.iloc[i,6] = words1[0]

for i in range(0,len(melted)):
    words2 = re.findall(pattern2, melted.iloc[i,5])
    if words2:
        melted.iloc[i,5] = words2[0]

# END

# Adding extra columns
melted['Measure Valid'] = ""
melted['Flag'] = ""

# # LOGIC FOR DATA VALIDATION
for i in range(0,len(melted)):
    if(melted['Speciality Area'][i] == "Medical Surgical" and (melted['Measure'][i] == "Productivity Index ER" or melted['Measure'][i] == "% Discharge MS"
                                                            or melted['Measure'][i] == 'Average Daily Census' or melted['Measure'][i] == 'RN Patient Ratio Days'
                                                            or melted['Measure'][i] == 'Number of Pt falls Case Analysis' or melted['Measure'][i] == 'Number of Falls Injury DP'
                                                            or melted['Measure'][i] == '% Compliance Catheter' or melted['Measure'][i] == '% Daily Documentation Urinary Catheter Monitor'
                                                            or melted['Measure'][i] == '% Reason to insert Catheter' or melted['Measure'][i] == 'Nurse Leader Visit MS'
                                                            or melted['Measure'][i] == 'Q1 Courtest and Respect MS' or melted['Measure'][i] == 'Q2 Nurse Listen Carefully MS'
                                                            or melted['Measure'][i] ==  'Q3 Nurse Explains MS' or melted['Measure'][i] == 'Q4 Call Button MS'
                                                            or melted['Measure'][i] == 'Q11 Bathroom MS' or melted['Measure'][i] == 'Supply Scanning Compliance Percentage First Quarter'
                                                            or melted['Measure'][i] == '% Catheter Utilization')):
                                                            melted['Measure Valid'][i] = "Yes"

    elif(melted['Speciality Area'][i] == "Emergency Services" and (melted['Measure'][i] == "Productivity Index ER" or melted['Measure'][i] == 'Number of Falls Injury Sharp Report'

                                                            or melted['Measure'][i] == 'Number of Pt falls Case Analysis' or melted['Measure'][i] == 'Number of Falls Injury DP'
                                                            or melted['Measure'][i] == 'Nurses Listen ES'
                                                            or melted['Measure'][i] == 'Inform of Treatment ES' or melted['Measure'][i] == 'Nurses Privacy Concern ES'
                                                            or melted['Measure'][i] == 'De Escalation Training ES' or melted['Measure'][i] == 'Courtesy of Nurses ES'
                                                            or melted['Measure'][i] ==  'Monthly Disposition To Leave Time ER' or melted['Measure'][i] == 'Attetion To Needs ES'
                                                            or melted['Measure'][i] == 'Monthly Antibiotic Med to Admin Avg Time ER' or melted['Measure'][i] == 'Number of Beds')):
                                                            melted['Measure Valid'][i] = "Yes"

    elif(melted['Speciality Area'][i] == "Surgical Services" and (melted['Measure'][i] == 'Average Daily Census' or melted['Measure'][i] == 'RN Patient Ratio Days'

                                                            or melted['Measure'][i] == 'Number of Cases SS' or melted['Measure'][i] == 'Number of staffed operating rooms SS'
                                                            or melted['Measure'][i] == 'Trauma Facility SS'
                                                            or melted['Measure'][i] == 'Information Needed Before Procedure SS')):
                                                            melted['Measure Valid'][i] = "Yes"

    elif((melted['Speciality Area'][i] == "PCU" or melted['Speciality Area'][i] == "ICU") and (melted['Measure'][i] == 'Average Daily Census' or melted['Measure'][i] == 'RN Patient Ratio Days'

                                                            or melted['Measure'][i] == 'Number of Pt falls Case Analysis' or melted['Measure'][i] == 'Number of Falls Injury Sharp Report'
                                                            or melted['Measure'][i] == '% Compliance Catheter'
                                                            or melted['Measure'][i] == '% Daily Documentation Urinary Catheter Monitor'
                                                            or melted['Measure'][i] == 'Supply Scanning Compliance Percentage First Quarter')):
                                                            melted['Measure Valid'][i] = "Yes"






    else:
        melted['Measure Valid'][i] = "No"


melted['Flag'] = np.where(((melted["Measure Valid"] == "Yes") & (melted["Value"] != "")), 1, 0)

# # # # To Csv on Shared Folder
melted.to_csv('//CorpDpt02/QASShare/CSG Nursing/Nursing Data Ecosystem/9. Code/donottouch.csv', header=True, index=False)


# Run batch file
filepath = 'C:/Users/JGE6931/Documents/runed.bat'
p = subprocess.Popen(filepath, shell=True, stdout = subprocess.PIPE)
stdout, stderr = p.communicate()
print p.returncode # is 0 if success
print ('If 0 then completed successfuly, else nope!')
print stdout
# EXECUTION TIME
print(100 * '#')
print("\n")
print("--- %s seconds ---" % (time.time() - start_time))
print("\n")
print(100 * '#')

