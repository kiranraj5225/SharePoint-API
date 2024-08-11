<h1 align="left">Access SharePoint REST API Using Python</h1>

###

<p align="left">You don't require a fancy library to connect with Microsoft SharePoint. You can use this program to generate custom report or deploy the program in gitlab/jenkins/any web api program to automate the tasks.<br><br>You would need a service principle, to access the SharePoint environment. You may use the service principle created from SharePoint App-Only approach in Azure AD.</p>

###

<h2 align="left">How to create service principle?.</h2>

###

<p align="left">✨ You need to create the service principle with App Registration in Azure AD.<br>✨ Add the required delegate SharePoint  - Selected Site permission.<br>✨ Grant admin consent<br>✨ For step by step instruction, You may go through this blogpost:
  https://www.c-sharpcorner.com/article/setting-up-sharepoint-app-only-principal-with-app-registration/
</p>

###

```python
from SharePoint import SharePoint

sp = SharePoint(config={
  "origin_tenant_id":"tenat_id",
  "origin_sp_host":"sharepoint_host_name",
  "origin_sp_host_type":"SP",
  "origin_sp_site":"/sites/sitename",
  "origin_sp_client_id":"client_id",
  "origin_sp_client_secret":"client_secret_value"})

if sp.login():
    # Function to create a file copy of local file in sharepoint folder
    # site: Sharepoint site or sub site path
    # Example: /sites/POCSite
    # folderName: Sharepoint folder path
    # Example: Shared Documents/Demo
    # fileName: File name with file extension
    # Example: filename.pdf
    # localFilePath: local path of the file

  sp.create_a_file_inside_folder(
    site="/sites/POCSite",
    folderName="Shared Documents/Demo",
    fileName="Sample.pdf",
    localFilePath=r"{0}".format(filePath)
  )
```
New functionalities will be added soon.
