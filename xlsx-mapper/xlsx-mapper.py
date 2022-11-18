import argparse
import openpyxl
import json
from EditShareAPI import EsAuth, FlowMetadata
import re
import os
import config
import pickle
import sys
from halo import Halo

osusername = os.getlogin()

parser = argparse.ArgumentParser(description="CLI for mapping Metadata from .xlsx to EditShare Metadata-Fields")
parser.add_argument("-c", "--config", action="store_true", help="Configure login")
parser.add_argument("-m", "--map", metavar=("path"), nargs=1, help="Maps Metadata from xlsx-file to Editshare Metadata Fields.", type= str)
args = parser.parse_args()

if args.config:
    config.configure()

try:
    // can we make this section interactive, so we don't need to store the pw in a file and don't need to run config first?
    // we can try to make this run-once and store the data in temp folder
    // if time == available: don't store pw in clear text
    configfile = open(f"C:/Users/{osusername}/.xlsx-mapper/config.p", "rb")
    login = pickle.load(configfile)
    EsAuth.login("192.168.0.221", login[0], login[1])
    print(f"--- Logged in as {login[0]} ---\n")
except:
    print('You have to run "xlsx-mapper -c" first')
    sys.exit()

if args.map:
    path = args.map
    fields = FlowMetadata.getCustomMetadataFields()
    try:
        xlsx = openpyxl.load_workbook(path[0], data_only=True)
    except PermissionError:
        print("Please close the Excel Document!")
        sys.exit()
    sheet = xlsx.active
    rows = sheet.rows
    columns = sheet.columns

    fieldsdict = dict()
    for field in fields:
        if re.match("[0-9][0-9][0-9]", field["name"][:3]):
            fieldsdict[field["name"]] = field["db_key"]
    mappingdict = dict()
    for ir, row in enumerate(rows):
        if ir == 0:
            for ic, cell in enumerate(row):
                try:
                    mappingdict[ic] = fieldsdict[cell.value]
                except KeyError:
                    pass

    sheet = xlsx.active
    rows = sheet.rows
    mappings = dict()
    #print(mappingdict)
    for ir, row in enumerate(rows):
        mapping = dict()
        for ic, cell in enumerate(row):
            if ir > 0:
                if ic == 0:
                    if re.match("[0-9][0-9][0-9][0-9][0-9][0-9]", str(cell.value)) and ";" not in str(cell.value):
                        essence_id = str(cell.value)
                        mappings[essence_id] = dict()
                try:  
                    mappings[essence_id][mappingdict[ic]] = cell.value
                except KeyError:
                    pass
    for mapping in mappings:
        data = {
                "combine": "MATCH_ALL",
                "filters": [
                    {
                        "field": {
                            "fixed_field": "CLIPNAME",
                            "group": "SEARCH_FILES"
                        },
                        "match": "EQUAL_TO",
                        "search": mapping
                    }
                ]
            }
        data = json.dumps(data)
        with Halo(text=f'Searching {mapping}', spinner='dots'):
            clips = FlowMetadata.searchAdvanced(data)
            if not clips:
                data = {
                    "combine": "MATCH_ALL",
                    "filters": [
                        {
                            "field": {
                                "custom_field": "field_248",
                                "fixed_field": "CUSTOM_field_248",
                                "group": "SEARCH_ASSETS"
                            },
                            "match": "EQUAL_TO",
                            "search": mapping
                        } 
                    ]
                }
                data = json.dumps(data)
                clips = FlowMetadata.searchAdvanced(data)
        for clip in clips:                
            if "clip_id" in clip.keys():
                metadata = FlowMetadata.getClipData(clip["clip_id"])
                asset_id = metadata.asset["asset_id"]
                metadata_id = metadata.metadata["metadata_id"]
                capture_id = metadata.capture["capture_id"]
                id = metadata.display_name
                try:
                    identifier = mappings[id]["field_50"]
                    name = mappings[id]["field_63"].replace(":", "").replace("\\", "-").replace("/", "-").replace(";","").replace(",", "").replace(".", "").replace("?","").replace("!","").replace("ß","sz").replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("'","").replace('"',"")
                    clipname = f"{identifier}__{name}"
                except KeyError:
                    id = metadata.asset["custom"]["field_52"]
                    clipname = metadata.display_name
                data = dict()
                data["custom"] = mappings[id]
                data["custom"]["field_248"] = id
                data["custom"]["field_231"] = clip["clip_id"]
                data["custom"]["field_233"] = asset_id
                data["custom"]["field_235"] = capture_id
                #data["custom"]["field_127"] = metadata.asset["uuid"]
                data = json.dumps(data)
                r = FlowMetadata.updateAsset(asset_id, data)
                if r.status_code == 403:
                    print(f"{login[0]} has no permission to write Metadata!")
                    sys.exit()
                print(f"Mapped {id} -> {r.text}")
                data = dict()
                data["clip_name"] = clipname
                data = json.dumps(data)
                r = FlowMetadata.updateMetadata(metadata_id, data)
                if r["code"] == 403:
                    error = r["details"]
                    print(f"Renaming failed: {error}")
                    sys.exit()
                print(f"Renamed {metadata.display_name} -> {clipname}")
                
          

