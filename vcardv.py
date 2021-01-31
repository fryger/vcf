import msal
import requests
import json
import base64
import configparser


class Vcard:
    def __init__(self):
        self.config = self.__load_config()
        self.headers = self.__authorize()
        self.data = []
        self.data_formatted = []
        self.__request_data()
        self.csv_data = self.__generate_csv()

    def __load_config(self):
        config = configparser.ConfigParser()
        config.read('config.ini')
        if len(config.sections()) <=0:
            raise Exception("No, empty or wrong config loaded")
        else:
            return config

    def __authorize(self):
        scope = [self.config['application']['scope']]
        app = msal.ConfidentialClientApplication(
            self.config['auth']['appID'],
            authority=self.config['application']['authority'],
            client_credential=self.config['auth']['appSecret'])
        token = app.acquire_token_silent(scope, account=None)
        if not token:
            token = app.acquire_token_for_client(scopes=scope)
        return  {"Authorization": "Bearer " + token["access_token"]}

    def __pagination(self,res):
        if "@odata.nextLink" in res:
            res = requests.get(res["@odata.nextLink"], headers=self.headers).json()
            self.data.append(res)
            self.__pagination(res)

    def __picture(self,id):
        req = requests.get(f"https://graph.microsoft.com/v1.0/users/{id}/photo/$value", headers=self.headers)
        if str(req) == "<Response [200]>":
            pic = base64.b64encode(req.content)
            pic = str(pic)[2:]
            pic = pic[:len(pic) - 1]
            return pic

    def __request_data(self):
        result = requests.get(self.config['application']['query'], headers=self.headers).json()
        self.data.append(result)
        self.__pagination(result)
        self.__format_data()

    def __add_company(self,string):
        mail = string[string.find("@") + 1:]
        if mail == "ecol-unicon.com":
            company = "Ecol-Unicon Sp. z o.o."
        if mail == "ecol-group.com":
            company = "Ecol-Group Sp. z o.o."
        if mail == "retencja.pl":
            company = "RetencjaPL Sp. z o.o."
        if mail == "biopro.pl":
            company = "Biopro Sp. z o.o."
        return company

    def __format_data(self):
        for i in self.data:
            for j in i['value']:
                try:
                    user = j["mail"]
                    user = user[:user.find("@")]
                    if user.find(".") != -1:
                        if user in self.config['users']['excluded']:
                            print(f"Excluded user {j['displayName']}, {j['mail']}")
                        else:
                            j["picture"] = self.__picture(j["id"])
                            j["company"] = self.__add_company(j["mail"])
                            print(f"Appending user {j['displayName']}, {j['mail']}")
                            self.data_formatted.append(j)
                except:
                    print(f"Excluded user {j['displayName']}, {j['mail']}")

    def __output_to_file(self,name,data,type):
        with open(name, 'w', encoding='utf8') as outfile:
            if type == 'json':
                json.dump(data, outfile, ensure_ascii=False)
            if type == 'string':
                outfile.write(data)

    def __generate_csv(self):
        output = ''
        for i in range(len(self.data_formatted)):
            output += 'BEGIN:VCARD\n'
            output += 'VERSION:2.1\n'
            output += f'N:{self.data_formatted[i]["surname"]};{self.data_formatted[i]["givenName"]};;;\n'
            output += f'FN:{self.data_formatted[i]["displayName"]}\n'
            output += f'TEL;CELL:{str(self.data_formatted[i]["mobilePhone"]).replace(" ", "")}\n'
            output += f'EMAIL:{self.data_formatted[i]["mail"]}\n'
            output += f'ORG:{self.data_formatted[i]["company"]};{self.data_formatted[i]["department"]}\n'
            output += f'TITLE:{self.data_formatted[i]["jobTitle"]}\n'
            output += f'PHOTO;ENCODING=BASE64;JPEG:{self.data_formatted[i]["picture"]}\n'
            output += 'END:VCARD\n'
        self.__output_to_file("kontakty.vcf", output,"string")
        return output

    def upload(self):
        data = str.encode(self.csv_data)
        len_data = len(data)
        upload_range = f'bytes 0-{len_data - 1}/{len_data}'
        response = requests.post(
            self.config['application']['file_query'],
            headers=self.headers,
            json={
                '@microsoft.graph.conflictBehavior': 'replace',
                'description': 'A large test file',
                'fileSystemInfo': {'@odata.type': 'microsoft.graph.fileSystemInfo'},
                'name': 'kontakty.vcf'
            }
            )
        response_dict = json.loads(response.text)
        upload_response = requests.put(
            response_dict["uploadUrl"],
            headers={
                'Content-Length': str(len_data),
                'Content-Range': upload_range
            },
            data=data
        )

        print(upload_response.content)




Vcard.upload()
