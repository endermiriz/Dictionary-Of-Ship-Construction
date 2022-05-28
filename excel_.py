import sys

from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
import openpyxl
try:
    authcookie = Office365('https://pirireisedutr.sharepoint.com/', username='ender.miriz@pru.edu.tr', password='St7844?s').GetCookies()
    i = 1
except:
    i = 0

site = Site('https://pirireisedutr.sharepoint.com/sites/GeminaatSzlkProjesiGrubu', version=Version.v2016, authcookie=authcookie)
folder = site.Folder('Shared Documents/Beta Testing')

with open("dict_data_setting.xlsx", "rb") as file_obj:
    file_as_string = file_obj.read()
print('---')


folder.upload_file(file_as_string, 'dict_data.xlsx')