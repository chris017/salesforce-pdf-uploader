from encodings import utf_8
from importlib.resources import path
import json
import requests
import base64
import os
import pandas
os.system("")

class Main:
    def __init__(self) -> None:
        pass
        # Load Excel Data
        print("""
-------------------------------------------------------------------------------------------------------------------
|                                                                                                                 |
|                                               PDF Uploader                                                      |
|                                                                                                                 |
-------------------------------------------------------------------------------------------------------------------   
        """)
        OK = '\033[92m' #GREEN
        WARNING = '\033[93m' #YELLOW
        FAIL = '\033[91m' #RED
        RESET = '\033[0m' #RESET COLOR

        df = pandas.read_excel(r"Path to your excel list")
        id_nummber = df.columns[0]
        id_nummbers = []
        ExternalID__c = df.columns[1]
        ExternalID__c_all = []

        # Login credentials
        params = {
            "grant_type": "password",
            "client_id": "Your_comnsumer_key", # Consumer Key
            "client_secret": "Your_consumer_secret", # Consumer Secret
            "username": "Your_username", # The email you use to login
            "password": "Your_password" # Concat your password and your security token
        }

        r = requests.post("https://""""mysalsforce.com""""/services/oauth2/token", params=params)
        # if you connect to a Sandbox, use test.salesforce.com instead
        access_token = r.json().get("access_token")
        instance_url = r.json().get("instance_url")
        #print("Access Token:", access_token)
        #print("Instance URL", instance_url)

        def sf_api_call(action, parameters = {}, method = 'get', data = {}):
            """
            Helper function to make calls to Salesforce REST API.
            Parameters: action (the URL), URL params, method (get, post or patch), data for POST/PATCH.
            """
            headers = {
                'Content-type': 'application/json',
                'Accept-Encoding': 'gzip',
                'Authorization': 'Bearer %s' % access_token
            }
            if method == 'get':
                r = requests.request(method, instance_url+action, headers=headers, params=parameters, timeout=30)
            elif method in ['post', 'patch']:
                r = requests.request(method, instance_url+action, headers=headers, json=data, params=parameters, timeout=10)
            else:
                # other methods not implemented in this example
                raise ValueError('Method should be get or post or patch.')
            # print('Debug: API %s call: %s' % (method, r.url) )
            if r.status_code < 300:
                if method=='patch':
                    return None
                else:
                    return r.json()
            else:
                raise Exception('API error when calling %s : %s' % (r.url, r.content))

        # Loop trough Excel data and wirte in new list
        for index, row in df.iterrows():
            id_nummbers.append(row[id_nummber])
            ExternalID__c_all.append(row[ExternalID__c])
        
        warning_message = str(input(f'{WARNING}Be carefull Changes will be made to Migration Orders continue?: (y/n) {RESET}')).strip().upper()
        
        if warning_message == 'Y':
            end_of_uploads = int(input(f"{WARNING}Please enter the ammount of uploads: {RESET}"))
        else:
            exit()


        counter = 0
        num = counter
        # While loop for more than one opperation
        while counter < end_of_uploads:
            if counter == end_of_uploads:
                break
            file_index = []
            file_new = []
            file_count = 0
            inp = ExternalID__c_all[num]
            Opportunity_id = id_nummbers[num]
            thisdir = os.getcwd()
            for r, d, f in os.walk(r"Your Search Path"): # change the hard drive, if you want
                for file in f:
                    filepath = os.path.join(r, file)
                    if inp in file:
                        new_path = os.path.join(r, file)   
                        file_index.append(new_path)
                        file_new.append(file)
                        file_count += 1
                        #print(file)
            print(f"\n{OK}---Found {file_count} file/s---{RESET}")
            # 1) Create a ContentVersion
            #print(file)
            #print(new_path)
            for index_of_file, file_name in zip(file_index, file_new):
                with open(index_of_file, "rb") as f:
                    encoded_string = base64.b64encode(f.read())
            
                ContentVersion = sf_api_call('/services/data/v54.0/sobjects/ContentVersion', method = 'post', data={
                        'Title': file_name,
                        'PathOnClient': index_of_file,
                        'VersionData': str(encoded_string.decode())
                    })
                ContentVersion_id = ContentVersion.get('id')
                # 2) Get the ContentDocument id
                ContentVersion = sf_api_call('/services/data/v54.0/sobjects/ContentVersion/%s' % ContentVersion_id)
                ContentDocument_id = ContentVersion.get('ContentDocumentId')

                # 3) Create a ContentDocumentLink
                ContentDocumentLink = sf_api_call('/services/data/v40.0/sobjects/ContentDocumentLink', method = 'post', data={
                        'ContentDocumentId': ContentDocument_id,
                        'LinkedEntityId': Opportunity_id,
                        'ShareType': 'V'
                    })
                print(Opportunity_id)
                print(file_name)
                print(f"{WARNING}{index_of_file}{RESET}")
            counter += 1
            num +=1  
        print(f"{OK}Upload successful{RESET}")
Main()