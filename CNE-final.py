
# coding: utf-8

# In[76]:

from IPython.display import clear_output
import bs4,time
import requests
import pandas as pd
from os.path import isfile

#from random import randint
from win32com.client import Dispatch
#from retrying import retry


# In[77]:

# this is the base_url
base_url = "http://www.cne.gob.ve/divulgacion_referendo_enmienda_2009/index.php?"


# In[78]:

# the bot pretends to be a standard Mozilla browser
hdrs = {"User-Agent": "Mozilla/5.0 (Windows NT 6.0; WOW64; rv:24.0) Gecko/20100101 Firefox/24.0"}


# In[79]:

# The pages are sorted in a hierarchy of geographical units.
# Centers with strange first letters are put in a page without a letter (the first one)

estado = "cod_estado="
estado_cod = [
              "01","02","03","04","05","06","07","08","09","10",
              "11","12","13","14","15","16","17","18","19","20",
              "21","22","23","24"]

municipio = "&cod_municipio="
municipio_cod = [
                 "01","02","03","04","05","06","07","08","09","10",
                 "11","12","13","14","15","16","17","18","19","20",
                 "21","22","23","24","25","26","27","28","29"]

parroquia = "&cod_parroquia="
parroquia_cod = [
                "01","02","03","04","05","06","07","08","09","10",
                "11","12","13","14","15","16","17","18","19","20",
                "21","22"]

centro = "&cod_centro="
centro_cod = [
            "001","002","003","004","005","006","007","008","009","010"
            "011","012","013","014","015","016","017","018","019","020"
            "021","022","023","024","025","026","027","028","029","030"
            "031","032","033","034","035","036","037","038","039","040"
            "041","042","043","044","045","046","047","048","049","050"
            "051","052","053","054","055","056","057","058","059","060"
            "061","062","063","064","065","066","067","068","069","070"
            "071","072","073","074","075","076","077","078","079","080"
            "081","082","083","084","085","086","087","088","089","090"
            "091","092","093","094","095","096","097","098","099","100"
            "101","102","103","104","105","106","107","108","109","110"
            "111","112","113","114","115","116","117","118","119","120"
            "121","122","123","124","125","126","127","128","129","130"
            "131","132","133","134","135","136","138","139","140","141"
            "142","143","144","145","146","147","148","150","151","152"
            "153","154","155","156","159","161","163","164","167","169"
            "170","171","172","173","174","175","176","177","179","182"
            "183","184","185","186","187","188","189","190","191","192"
            "193","196","206","207","208","209","210","211","212","214"
            "215","221","223","231","235","240","241","242","243","244"
            "245","247","249"]

mesa = "&num_mesa="
mesa_cod = [
            "1","2","3","4","5","6","7","8","9","10",
            "11","12","13","14","15","16","17","18","19","20",
            "21","22","23","24","25"]


# In[80]:

def get_itemlist(thesoup):
    try:
        big_resultados=[]
        resultados = []

        lotsofitems = thesoup.find_all("span",class_="tah12_2")

        resultados+= [lotsofitems[1].get_text()]
        resultados+= [lotsofitems[2].get_text()] #si_porc
        resultados+= [lotsofitems[4].get_text()] #no_votos
        resultados+= [lotsofitems[5].get_text()] #no_porc

        big_resultados+=resultados+[thepage]
        
        #skip = resultados == None

        return big_resultados

    except IndexError:
        print('Broken Link')


# In[81]:

skip = False
big_resultados=[]
if skip == False:
    for i in estado_cod:
        for j in municipio_cod:
            for k in parroquia_cod:

                thepage = base_url+str(estado)+i+str(municipio)+j+str(parroquia)+k
                # 1.call the url
                stuff = requests.get(thepage, headers=hdrs)

                # 2.transform to soup using html.parser parser
                soup = bs4.BeautifulSoup(stuff.text, "html.parser")

                # 3.extract the new reviews from this page
                resultados = get_itemlist(soup)
                big_resultados+=[resultados]

                skip = resultados == None

                # 4.print something to show how the process progresses
                print("URL:",thepage)
                
        break
print("listo muchachos")        


# In[7]:

clean=[x for x in big_resultados if x is not None]


# In[8]:

len(clean)


# In[9]:

pd.DataFrame(clean)
# save the data as a csv file
output = pd.DataFrame(clean)
output.to_csv("resultados.csv")


# In[10]:

xl = Dispatch("Excel.Application")
xl.Visible = True # otherwise excel is hidden

# newest excel does not accept forward slash in path
wb = xl.Workbooks.Open(r'C:\Users\Carlos\Python\Scraping the web\resultados.csv')
#wb.Close()
#xl.Quit()


# In[ ]:



