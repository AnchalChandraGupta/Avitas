from xlrd import open_workbook
import re
import json
import os
from collections import OrderedDict

"""
file_loc : folder containing sub folders where each subfolder contains json with name as    [folder name] + ".json"
excel_loc : folder containing excel files
Map each excel to a json
excel_dict : dict object containing [json name] --> [excel name] items
"""


file_loc = '/home/abcde/Desktop/forms-seed-data-master/import_forms/'
excel_loc = '/home/abcde/Desktop/Excel/'

p1 = re.compile("(\d+)-[V|v][G|g][S|s].*(\d+).xlsx")
excel_dict = dict()
for excel_name in next(os.walk(excel_loc))[2]:
    json_name = "VGS" + p1.match(excel_name).group(1) + "_rev" + p1.match(excel_name).group(2)
    excel_dict[json_name] = excel_name


"""
For each sub folder [dir] in [file_loc], 
    Create the file name [file_name] for file inside the folder and open the file
    Load the file content as json [input_json]
    Edit and change the input_json as required 
    Print the errors
    Write the modified json to an output file [output_file]
"""

for dir in next(os.walk(file_loc))[1]:

    file_name = file_loc + dir + "/" + (
        dir if (dir + ".json") in next(os.walk(file_loc + dir))[2] else dir.lower()) + ".json"
    fp = open(file_name, encoding='utf-8', errors='replace')
    file_name = fp.name.rsplit('/')[7]
    print(file_name)

    try:
        inputjson = json.load(fp, object_pairs_hook=OrderedDict)
    except json.decoder.JSONDecodeError as e:
        print("\tInvalid json:\n\t", e)
        continue
    try:
        headerDef = inputjson["inspectionDocumentTemplate"]["formHeaderTemplate"]["headerDef"]
    except KeyError as e:
        print("\tKeyError:\n\t", e)
        continue


    fields = [str(i["name"]).lower() for i in headerDef]

    if "department" not in fields:
        headerDef.append({
            "name": "department",
            "dataType": "String",
            "displayName": "Department",
            "widget": {
                "type": "dropdown",
                "possibleValues": {
                    "Vendor": "Vendor",
                    "GE Shop": "GE Shop",
                    "Field Service Engineer": "Field Service Engineer"
                }
            }
        })

    if "inspectionprocedure" not in fields:
        headerDef.append({
            "name": "inspectionProcedure",
            "dataType": "String",
            "displayName": "Inspection Procedure",
            "widget": {
                "type": "readonly",
                "possibleValues": {
                    "value": ""
                }
            }
        })

    for i in headerDef:
        if i["name"] == "assemblySerNum":
            headerDef.remove(i)

    if dir not in excel_dict.keys():
        print("Excel for", dir, "not present")
    else:
        excel_doc = excel_dict[dir]
        sheet_names = open_workbook(excel_loc + excel_doc).sheet_names()
        component_type_set = set()

        for i in inputjson["componentDefinitionList"]:
            ctemp = re.sub("[ ]*[\&|\-|\\\|/][ ]*", " ", i["componentType"].lower())
            flag = False
            for x in sheet_names:
                stemp = re.sub("[ ]*[\&|\-|\\\|/][ ]*", " ", x.lower())
                if len(ctemp) >= len(stemp):
                    if stemp in ctemp:
                        i["componentType"] = x
                        flag = True
                else:
                    if ctemp in stemp:
                        i["componentType"] = x
                        flag = True
            if flag == False:
                if ctemp not in component_type_set:
                    component_type_set.add(ctemp)
                    print("\t: ComponentType not updated: sheet for [", i["componentType"], "] not present.")

    for i in inputjson["componentDefinitionList"]:
        for j in i["assessmentItems"]:
            if j["assessmentType"] == "Measurement":
                for k in j["assessmentDef"]:
                    if k["dataType"] == "Numeric":
                        k["variance"] = True

            for k in j["assessmentDef"]:
                if "iterations" in k.keys():
                    if k["iterations"] == 1:
                        k["iterationLabels"] = [""]
                else:
                    print("\t: iterations field not present in", i["componentType"])

        if "Misc Components" in i["componentType"]:
            i["dataDefKeys"] = ["partNumber",
                                "partNumberRevision",
                                "serialNumber",
                                "compNameDesc"
                                ]
            i["instructions"] = "Make Any Comments About Discrepant Parts Below"
            inputjson["formTemplates"] = [
                OrderedDict({
                    "componentType": "Misc Components",
                    "instructions": "LIST OTHER MISCELLANEOUS COMPONENTS (Nuts, Bolts, Etc.)"
                })
            ]

    output_file = open("/home/abcde/Desktop/Generated/" + file_name[:-5] + "_NEW.json", mode='w+')
    output_file.write(json.dumps(inputjson, indent=2))
    output_file.close()
