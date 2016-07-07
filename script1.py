from bs4 import BeautifulSoup
import re 
from openpyxl import Workbook
import requests
import urllib.request

counter=0

#~ wb = Workbook()
#~ ws = wb.active
#~ ws.append([1,2,3])
#~ wb.save("sample.xlsx")

wb = Workbook()
ws = wb.active

def tens(url):
    global counter
    print(url)
    s = requests.session() 
    r = s.get(url) 
    content = r.text
    soup = BeautifulSoup(content, "html.parser")
    site_id_list = []
    for a in soup.findAll('a', {"id": re.compile('^ctl00_ContentPlaceHolder1_dgResults_')}):
        s = a['href']
        id_list = re.match('.*?([0-9]+)$', s).group(1)
        print(id_list)
        site_id_list.append(id_list)
    for id_list in site_id_list:
        counter+=1
        print('Counter: {}'.format(counter))
        s = 'SiteServiceDetails.aspx?SiteID'+id_list
        sdigits = re.match('.*?([0-9]+)$', s).group(1)
        urlbase = 'http://humanservicesdirectory.vic.gov.au/'
        urlbase+= 'SiteServiceDetails.aspx?SiteID='+sdigits
        content = urllib.request.urlopen(urlbase).read()
        soup = BeautifulSoup(content, "html.parser")
        data = soup.findAll('a', {"id": "ctl00_ContentPlaceHolder1_dlServices_ctl00_SiteServiceSectionDetailsComponent_hlEmail"})
        #print(data[0].contents[0])
        try:
            Email=data[0].contents[0]
            #print('Email: {email}'.format(email=Email))
        except IndexError:
            Email='no email'
            #print('Email: {email}'.format(email=Email))
        
        data = soup.findAll('div', {"id": "ctl00_ContentPlaceHolder1_dlServices_ctl00_SiteServiceSectionDetailsComponent_gvHours"})
        #print(data)
        for item in data:
            napisy = item.findAll('td')
        #print(type(napisy))
        data_list=[]
        #print(napisy)
        for item in napisy:
            data_list.append(item.contents[0])
        #print(data_list)
        try:
            Saturday_index = data_list.index( 'Saturday' )
            Saturday=data_list[Saturday_index + 1]
        except ValueError:
            Saturday='-'
        try:
            Sunday_index = data_list.index( 'Sunday' )
            Sunday=data_list[Sunday_index + 1]
        except ValueError:
            Sunday='-'
        try:
            Weekday_index = data_list.index( 'Weekday' )
            Weekday=data_list[Weekday_index + 1]
            Monday=Weekday
            Tuesday=Weekday
            Wednesday=Weekday
            Thursday=Weekday
            Friday=Weekday        
        except ValueError:
            Weekday='-'
            Monday=data_list[data_list.index( 'Monday' ) + 1]
            Tuesday=data_list[data_list.index( 'Tuesday' ) + 1]
            Wednesday=data_list[data_list.index( 'Wednesday' ) + 1]
            Thursday=data_list[data_list.index( 'Thursday' ) + 1]
            Friday=data_list[data_list.index( 'Friday' ) + 1]
        
        #print('Saturday: {saturday}, Weekday : {weekday}, Sunday : {sunday}, Monday: {monday}'.format(saturday=Saturday,weekday=Weekday, sunday=Sunday, monday=Monday))
        data = soup.findAll('div', {"id": "ctl00_ContentPlaceHolder1_dlServices_ctl00_SiteServiceSectionDetailsComponent_gridViewContactNumbers"})
        for item in data:
            napisy = item.findAll('td')
            data_list=[]
            for item in napisy:
                data_list.append(item.contents[0])
                #print(data_list)
            Phone = data_list[1]+' '+data_list[2]
            #print('Phone: {phone}'.format(phone=Phone))

        s = 'SitePracDetails.aspx?SiteID='+id_list
        sdigits = re.match('.*?([0-9]+)$', s).group(1)
        urlbase = 'http://humanservicesdirectory.vic.gov.au/'
        urlbase+= 'SitePracDetails.aspx?SiteID='+sdigits
        content = urllib.request.urlopen(urlbase).read()
        soup = BeautifulSoup(content, "html.parser")
        data = soup.findAll('ul', id="ctl00_ContentPlaceHolder1_bulletListPractitioners")
        Practitioners_after_split = []
        if data:
            for link in data:
                links = link.findAll('a', {"href": re.compile('^SitePracDetails.aspx')})
                #print(links)
                Practitioners=[]
                for prac in links:
                    Practitioners.append(prac.contents[0])
                for item in Practitioners:
                    tmp_list = item.split()
                    Practitioners_after_split.append(tmp_list[1])
                    Practitioners_after_split.append(tmp_list[2]) 
        else:
            Practitioners = 'no practitioners listed'
        #print(id_list, Practitioners)   
        #print(id_list, Practitioners_after_split)   
        s = 'SiteDetails.aspx?SiteID'+id_list
        sdigits = re.match('.*?([0-9]+)$', s).group(1)
        urlbase = 'http://humanservicesdirectory.vic.gov.au/'
        urlbase+= 'SiteDetails.aspx?SiteID='+sdigits
        content = urllib.request.urlopen(urlbase).read()
        soup = BeautifulSoup(content, "html.parser")
        data = soup.findAll('span', {"id": "ctl00_ContentHeaderPlaceHolder_DisplaySiteHeader1_lblSuburb"})
        data_tuple=data[0].contents[0].split(', ')
        Suburb=data_tuple[0]
        Postcode=data_tuple[1]
        #print('suburb: {suburb}, code: {code}'.format(suburb=Suburb, code=Postcode))
        
        data = soup.findAll('span', {"id": "ctl00_ContentHeaderPlaceHolder_DisplaySiteHeader1_lblAgency"})
        Company = data[0].contents[0]
        #print('Company: {company}'.format(company=Company))
        data = soup.findAll('span', {"id": "ctl00_ContentPlaceHolder1_lblAddress"})
        data_list=[]
        data_list.append(data[0].contents[0])
        Address=', '.join(data_list)
        #print('Address: {adres}'.format(adres=Address))

        fields = [Company, 
        Address,
        Suburb,
        Postcode,
        Phone,
        Email,
        Monday,
        Tuesday,
        Wednesday,
        Thursday,
        Friday,
        Saturday,
        Sunday,
        ]
        fields.extend(Practitioners_after_split)
        ws.append(fields)


numlist = [1]
for num in numlist:
    url = "http://humanservicesdirectory.vic.gov.au/SearchResults.aspx?name=pharmacy&state=VIC&ActiveNum=" + str(num) 
    tens(url)  
    
wb.save("sample.xlsx")












