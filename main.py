# import pandas as pd
# data = pd.read_excel('Data/dict_data.xlsx')
# data2 = data['Kelimeler'].str.title()
# data3 = data2.sort_values()
# data4 = data3.drop_duplicates(keep='last')
# data4.to_excel('Data/dict_data.xlsx', sheet_name='1',index=False)
# kelimeler = []
#
# kelimeler.append(data4)
# print(kelimeler)
# -----------
# import openpyxl as xl
#
# wb = xl.load_workbook("Data/dict_data.xlsx")
#
#
# sheet = wb["1"]
#
#
# row_count = sheet.max_row
# column_count = sheet.max_column
#
# print(row_count)
# print(column_count)
# ----------

# import pandas as pd
# df = pd.read_excel('Data/dict_data.xlsx')
# df.Kelimeler = df.Kelimeler.str.capitalize()
# df.Anlamlar = df.Anlamlar.str.capitalize()
# df = df.sort_values(by=['Kelimeler','Anlamlar'])
# df = df.drop_duplicates(subset=['Kelimeler'],keep='last')
# df['Anlamlar'] = df["Kelimeler"] +"\n\n"+ df['Anlamlar']
# df.to_excel('fileName.xlsx',index=False)

# ---------
# df = pd.read_excel('Data/dict_data.xlsx')
# i=df[df["Kelimeler"]==item.text()]
# print(i)

# ----------
# kullanici_adi = ""
# sifre = ""
# def login():
#     global kullanici_adi,sifre
#     kullanici_adi = input("Kullanıcı adı:")
#     sifre = input("Şifre:")
#
# def seem():
#     print(kullanici_adi)
#     print(sifre)
# def giris():
#     if kullanici_adi == "ender" and sifre == "miriz":
#         print("Giriş Yapıldı.")
#
#     elif kullanici_adi == "ender" and sifre != "miriz":
#         print("Şifre hatalı.")
#     elif kullanici_adi != "ender" and sifre == "miriz":
#         print("Kullanıcı adı hatalı.")
#     else:
#         print("Kullanıcı adı ve şifre hatalı.")
# login()
# seem()
# giris()
# ---------------
# list_sen = ['Eggs on Sunday']
# list_wrd = ['Eggs', 'Fruits']
#
# if len(list_sen) == 1:
#     list_sen.clear()
# else:
#
#     res = [all([k in s for k in list_wrd]) for s in list_sen]
#
#     silincek = ([list_sen[i] for i in range(0, len(res)) if res[i]])
#     list_sen.remove(*silincek)
#     print(silincek)
# print(list_sen)
# ---------------
# list_sen = ['Eggs on Sunday', 'Mehmet Yılmaz', "Artık dur"]
# while 10:
#     try:
#         print("Silinmeden önce: ", list_sen)
#         list_wrd = input("Silmek istediğin kelime:")
#         if len(list_sen) == 1:
#             list_sen.clear()
#         else:
#             res = [all([k in s for k in list_wrd]) for s in list_sen]
#
#             silincek = ([list_sen[i] for i in range(0, len(res)) if res[i]])
#             list_sen.remove(*silincek)
#     except:
#         pass
#
#     print("Silindikten sonra: ", list_sen)
#
# -----------------
# arr = ['An American', 'Barack Obama', '4.7', '18979', 'An Indian', 'Mahatma Gandhi', '4.7', '18979',
#     'A Canadian', 'Stephen Harper', '4.6', '19234']
#
# inputStr = input("Enter String: ")
#
# for val in arr:
#     if inputStr in val:
#         print(val)
#         arr.remove(val)
# print("Last List :\n ",arr)
# ---------------

# import re
# sentences = ['a long brown fox','i never knew One fox long']
# words = [input("Silmek istediğiniz kelime: ")]
# indices = [ [ i for i,sentence in enumerate( sentences ) if re.search( '.+'.join( word.split()), sentence)] for word in words]
# string_indices = str(indices)
# s = string_indices.replace('[','')
# s = s.replace(']','')
# string_indices = int(s)
# del sentences[string_indices]
#
# print(sentences)
# --------------------
# anlamlar = ["Akta [Bakkal]\n\nBakkala gidip 10 Tl bozdur","Shop [Market]\n\nMarkete gidip 100 Tl bozdur","Car [Araba]\n\nArabayla gidip 100 Tl bozdur."]
# for anlam in anlamlar:
#     kelime = anlam.split('\n', 1)[0]
#     print(kelime)
#     print(anlam)
#     print("------------")
# ------------
# my_string="Arkas[Lojistik]tr:\n\nmerhaba python dünyası , ben bir yeniyim ing:\n\nhello python world , i'm a beginner "
# print (my_string.split("ing:\n\n",1)[1] )
# print("-------------------")
# start = "tr:\n\n"
# end = "ing:\n\n"
# my_string="Arkas[Lojistik]tr:\n\nmerhaba python dünyası , ben bir yeniyim ing:\n\nhello python world , i'm a beginner "
# print (my_string[my_string.find(start)+len(start):my_string.rfind(end)])

# -------------------
from datetime import datetime

# # datetime object containing current date and time
# now = datetime.now()
#
# # dd/mm/YY H:M:S
# username = "ender.miriz@pru.edu.tr"
# dt_string = now.strftime("%d.%m.%Y %H.%M")
# f= open(username+" "+dt_string+".txt","w+")
# print("date and time =", dt_string)

# ------------------

# from shareplum import Site
# from shareplum import Office365
# from shareplum.site import Version
# try:
#     authcookie = Office365('https://pirireisedutr.sharepoint.com/', username='ender.miriz@pru.edu.tr',
#                            password='St7844?s').GetCookies()
#     i = 1
# except:
#     i = 0
#
# site = Site('https://pirireisedutr.sharepoint.com/sites/GIN2006Listeler', version=Version.v2016,
#             authcookie=authcookie)
# folder = site.Folder('Shared Documents/list')
#
# with open("29.03.2022 13.56.txt", "rb") as file_obj:
#     file_as_string = file_obj.read()
# print('---')
#
# folder.upload_file(file_as_string, '29.03.2022 13.56.txt')
# ---------------------
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
import os
try:
    try:
        authcookie = Office365('https://pirireisedutr.sharepoint.com/', username="ender.miriz@pru.edu.tr",
                               password="St7844?s").GetCookies()
        i = 1
        print("Başarılı")
    except:
        i = 0
    site = Site('https://pirireisedutr.sharepoint.com/sites/GIN2006Listeler', version=Version.v2016,
                authcookie=authcookie)
    now = datetime.now()
    dt_string = now.strftime("%d.%m.%Y %H.%M")
    filename = mail + " " + dt_string + ".txt"
    print(filename)
    tringanlam = Window.tringanlam
    with open(filename, 'w') as file:
        for line in tringanlam:
            file.write(line)
            file.write('\n')
    folder = site.Folder('Shared Documents/list')
    with open(filename, "rb") as file_obj:
        file_as_string = file_obj.read()
    print('---')

    folder.upload_file(file_as_string, filename)
    try:
        if os.path.exists(filename):
            os.remove(filename)
        else:
            print("txt dosyası bulunamadı")
    except:
        pass
except:
    pass