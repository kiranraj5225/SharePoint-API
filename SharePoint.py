from requests import Session
import json
import base64

class ConnError(Exception):
    pass

class ConstructLoginURL:
    def __init__(self, http_type, host, tenant_id, suffix):
        self.url = http_type+"://"+host+"/"+tenant_id+"/"+suffix

class ConstructURL:
    def __init__(self, http_type, host, site, suffix):
        self.url = http_type+"://"+host+site+"/"+suffix

class Configuration:
    def __init__(self, origin_tenant_id, origin_sp_host, origin_sp_host_type, origin_sp_site, origin_sp_client_id, origin_sp_client_secret):
        self.origin_tenant_id = origin_tenant_id
        self.origin_sp_host = origin_sp_host
        self.origin_sp_host_type = origin_sp_host_type
        self.origin_sp_site = origin_sp_site
        self.origin_sp_client_id = origin_sp_client_id
        self.origin_sp_client_secret = origin_sp_client_secret

class SharePoint:

    URL_SUFFIX = "/_api/web/"

    def __init__(self, config):
        self.parsed_config = config
        self.session = Session()
        self.base_url = {}

    def login(self):
        """
            To authenticate and attach the access token in request headers so all the
            upcoming request headers will have the access token
        """
        try:
            configuration = Configuration(**self.parsed_config)
            login_url = ConstructLoginURL(http_type="https", host="accounts.accesscontrol.windows.net",
                                          tenant_id=configuration.origin_tenant_id, suffix="tokens/OAuth/2").url
            headers = {
                "Content-Type": "application/x-www-form-urlencoded"
            }
            body = {
                "grant_type": "client_credentials",
                "client_id": configuration.origin_sp_client_id+"@"+configuration.origin_tenant_id,
                "client_secret": configuration.origin_sp_client_secret,
                "resource": "00000003-0000-0ff1-ce00-000000000000/"+configuration.origin_sp_host+"@"
                            +configuration.origin_tenant_id
            }
            response = self.session.post(login_url, headers=headers, data=body)
            if response.status_code == 200:
                details = response.json()
                if isinstance(details, dict):
                    self.session.headers.update({
                                                 "Accept": "application/json;odata=verbose",
                                                 "Content-Type": "application/json;odata=verbose",
                                                 "Authorization": "Bearer {0}".format(details.get("access_token")),
                            })
                    print("Login Success!")
                    return True
            else:
                raise ConnError("Error! Unable to login.. {0}".format(response.text))
        except TypeError as te:
            raise te
        except Exception as e:
            raise e

    # To get the list metadata
    def get_list_info(self, site, list_name):
        try:
            configuration = Configuration(**self.parsed_config)
            url = ConstructURL(http_type="https", host=configuration.origin_sp_host, site=site, suffix="_api/web/lists/GetByTitle('{0}')".format(list_name)).url
            list_info_response = self.session.get(url)
            if list_info_response.status_code == 200:
                return list_info_response.json()
            elif list_info_response.status_code == 404:
                return "ListNotFound"
            else:
                return list_info_response.text
        except Exception as e:
            raise e

    # To get the items from the list
    def get_list_items(self, site, list_name):
        data = []
        try:
            configuration = Configuration(**self.parsed_config)
            url = ConstructURL(http_type="https", host=configuration.origin_sp_host, site=site,
                               suffix="_api/web/lists/GetByTitle('{0}')/items?$select=ID,Title,Attachments,Created,Modified,Editor/Title,File&$expand=Editor,File".format(list_name)).url
            list_info_response = self.session.get(url)
            if list_info_response.status_code == 200:
                response = list_info_response.json()
                data.extend(response["d"]["results"])
                if "__next" in response["d"]:
                    next_items = self.perform_next_item(nextURL=response["d"]["__next"])
                    data.extend(next_items)
                return data
            elif list_info_response.status_code == 404:
                return "ListNotFound"
            else:
                return list_info_response.text
        except Exception as e:
            raise e

    # To get teh list item by unique ID
    def get_list_item_by_id(self, site, list_name, id):
        try:
            configuration = Configuration(**self.parsed_config)
            url = ConstructURL(http_type="https", host=configuration.origin_sp_host, site=site,
                               suffix="_api/web/lists/GetByTitle('{0}')/items({1})?$select=ID,Title,Attachments,Editor/Title,File&$expand=Editor,File".format(list_name, id)).url
            list_info_response = self.session.get(url)
            if list_info_response.status_code == 200:
                response = list_info_response.json()
                return response["d"]
            elif list_info_response.status_code == 404:
                return "ListNotFound"
            else:
                return list_info_response.text
        except Exception as e:
            raise e

    # Function to handle the next pages and its results
    def perform_next_item(self, nextURL=""):
        consolidated_next_items = []
        try:
            while nextURL!="":
                next_items_response = self.session.get(nextURL)
                if next_items_response.status_code == 200:
                    next_items = next_items_response.json()
                    consolidated_next_items.extend(next_items["d"]["results"])
                    if "__next" not in next_items["d"]:
                        nextURL = ""
                    else:
                        nextURL = next_items["d"]["__next"]
            return consolidated_next_items
        except Exception as e:
            raise e

    # Update an item in SharePoint list
    def update_an_item(self, site, list_name, dataObj):
        try:
            list_info = self.get_list_info(site, list_name)
            if isinstance(list_info, dict):
                ListItemEntityTypeFullName = list_info["d"]["ListItemEntityTypeFullName"]
                configuration = Configuration(**self.parsed_config)
                url = ConstructURL(http_type="https", host=configuration.origin_sp_host, site=site,
                                   suffix="_api/web/lists/getByTitle('{0}')/items({1})".format(list_name,
                                                                                               dataObj["ID"])).url
                body = {
                        "__metadata": {"type": ListItemEntityTypeFullName},
                        **dataObj
                    }
                if "IF-MATCH" not in self.session.headers.keys():
                    self.session.headers.update({"IF-MATCH": "*"})
                if "X-HTTP-Method" not in self.session.headers:
                    self.session.headers.update({"X-HTTP-Method": "MERGE"})
                updated_item_response = self.session.post(url, data=json.dumps(body))
                print(updated_item_response.status_code)
                if updated_item_response.status_code == 201:
                    updated_item = updated_item_response.json()
                elif updated_item_response.status_code == 400:
                    return updated_item_response.text
                elif updated_item_response.status_code == 404:
                    return "ListNotFound"
                else:
                    return updated_item_response.text
        except Exception as e:
            raise e

    # Get list of files from a folder
    def get_files_from_folder(self, site, folderName):
        try:
            configuration = Configuration(**self.parsed_config)
            url = ConstructURL(http_type="https", host=configuration.origin_sp_host, site=site,
                               suffix="_api/web/GetFolderByServerRelativeUrl('{0}')/Files".format(folderName)).url
            files_info_response = self.session.get(url)
            if files_info_response.status_code == 200:
                response = files_info_response.json()
                return response["d"]["results"]
            elif files_info_response.status_code == 404:
                return "ListNotFound"
            else:
                return files_info_response.text
        except Exception as e:
            raise e

    # Get a file details from a folder
    def get_a_file_info_from_folder(self, site, folderName, fileName):
        try:
            configuration = Configuration(**self.parsed_config)
            url = ConstructURL(http_type="https", host=configuration.origin_sp_host, site=site,
                               suffix="_api/web/GetFolderByServerRelativeUrl('{0}')/Files('{1}')".format(folderName, fileName)).url
            file_info_response = self.session.get(url)
            if file_info_response.status_code == 200:
                response = file_info_response.json()
                return response["d"]["results"]
            elif file_info_response.status_code == 404:
                return "ListNotFound"
            else:
                return file_info_response.text
        except Exception as e:
            raise e

    # Get a file's Base64Encoded value from a folder
    def get_a_file_content_from_folder(self, site, folderName, fileName):
        try:
            configuration = Configuration(**self.parsed_config)
            # setting the header to support the Base64Encoded
            self.session.headers.update(
                {"Content-Type": "application/octet-stream", "content-disposition": "attachment"})
            url = ConstructURL(http_type="https", host=configuration.origin_sp_host, site=site,
                               suffix="_api/web/GetFolderByServerRelativeUrl('{0}')/Files('{1}')/$value".format(folderName, fileName)).url
            file_info_response = self.session.get(url)
            # Setting the header back to support for json
            self.session.headers.update({
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose"
            })
            if file_info_response.status_code == 200:
                response = file_info_response.json()
                return response["d"]["results"]
            elif file_info_response.status_code == 404:
                return "ListNotFound"
            else:
                return file_info_response.text
        except Exception as e:
            raise e

    def create_a_file_inside_folder(self, site, folderName, fileName, localFilePath):
        try:
            configuration = Configuration(**self.parsed_config)
            self.session.headers.update(
                {"Content-Type": "application/json", "Accept": "application/json"})
            url = ConstructURL(http_type="https", host=configuration.origin_sp_host, site=site,
                               suffix="_api/web/GetFolderByServerRelativeUrl('{0}')/Files/add(url='{1}',overwrite=true)".format(folderName, fileName)).url
            file_info_response = self.session.post(url, data=open(localFilePath, 'rb').read())
            # Setting the header back to support for json
            self.session.headers.update({
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose"
            })
            if file_info_response.status_code == 200:
                # response = file_info_response.json()
                return "Success"
            elif file_info_response.status_code == 404:
                return "ListNotFound"
            else:
                return file_info_response.text
        except Exception as e:
            raise e
