import selenium
from  selenium import webdriver
import time
import requests
import datetime
from datetime import date
from datetime import timedelta
import pandas as pd
import glob
import os, zipfile


print("Start Date : YYYY-AA-GG\n")
start_date = str(input())

print("End Date : YYYY-AA-GG\n")
end_date = str(input())


def find_date_range(start_date,end_date):
    start_date = datetime.date(int(start_date[0:4]), int(start_date[5:7]), int(start_date[8:10]))
    end_date = datetime.date(int(end_date[0:4]), int(end_date[5:7]), int(end_date[8:10]))

    day_list = []

    delta = end_date - start_date

    for i in range(delta.days + 1):
        day = start_date + timedelta(days=i)
        day_list.append(str(day.year) + "/" + str(day.strftime('%m')) + "/thb" + str(day.year) + str(day.strftime('%m')) + str(day.strftime('%d')) + "1.zip")

    return day_list

def unzip_files():   #zip halindeki dosyaların çıkartılmasını sağlar
    dir_name = r"C:\Users\duzgun.ilaslan\Porsalan\BISTDAILY"
    extension = ".zip"

    os.chdir(dir_name) 

    for item in os.listdir(dir_name): 
        if item.endswith(extension): 
            file_name = os.path.abspath(item) 
            zip_ref = zipfile.ZipFile(file_name) 
            zip_ref.extractall(r"C:\Users\duzgun.ilaslan\Porsalan\BISTDAILY") #ziplerin çıkartılacağı dosya yolu
            zip_ref.close() 
            os.remove(file_name) 


def read_files():  #belirtilen klasöre indirilen csv dosyalarını okuyarak tek bir df'te birleştirir.
    path = r"C:\Users\duzgun.ilaslan\Porsalan\BISTDAILY"
    csv_files = glob.glob(path + "/*.csv")
    df_list = (pd.read_csv(file,sep=";") for file in csv_files)
    big_df   = pd.concat(df_list, ignore_index=True)

    return big_df

driver_path = r"C:\Users\duzgun.ilaslan\Porsalan\WebDriver\chromedriver.exe" #Selenium'un kullandığı chrome driver dizinini gösterir.
op = webdriver.ChromeOptions()
prefs = {'download.default_directory' : r'C:\Users\duzgun.ilaslan\Porsalan\BISTDAILY'}  #download edilecek dizini gösterir.
op.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(executable_path=driver_path, options=op)


for j in find_date_range(start_date,end_date):   #Bist bültenindeki url oluşturulur ve eğer o güne ait dosya varsa indirilir.
    url = "https://www.borsaistanbul.com/data/thb/" + str(j)
    r=requests.get(url)

    if r.status_code==200:
        driver.get(url)

    time.sleep(2)

unzip_files()
df = read_files()
print(df.iloc[[0]])
df.to_excel("deneme.xlsx")






