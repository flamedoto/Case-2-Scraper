import time
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import *
import re
import pandas as pd
import math
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut



geolocator = Nominatim(user_agent="geo")


def do_geocode(address, attempt=1, max_attempts=5):
    try:
        return geolocator.geocode(address)
    except GeocoderTimedOut:
        time.sleep(1)
        if attempt <= max_attempts:
            return do_geocode(address, attempt=attempt + 1)
        raise


class PublicCase():
	  # US Proxy IP PORT
    PROXY = "23.227.207.163:3128"



      # Main URl
    # URL = 'https://public.courts.in.gov/mycase/#/vw/CaseSummary/eyJ2Ijp7IkNhc2VUb2tlbiI6IkRhc0ZabFVBYUZBSExmd1RsY28tZ0ZwemFVMkRuREZWOXlzeG5qUGotZVkxIn19'
    # URL = 'https://www.google.com/webhp?hl=en&sa=X&ved=0ahUKEwiI7fzAgensAhWIXsAKHXCoAjQQPAgI'

    ## Defining options for chrome browser
    options = webdriver.ChromeOptions()
    # ssl certificate error ignore
    options.add_argument("--ignore-certificate-errors")
    # Adding proxy
    options.add_argument('--proxy-server=%s' % PROXY)
    Browser = webdriver.Chrome(executable_path="chromedriver", options=options)

    # Excel file declaration
    ExcelFile = pd.ExcelWriter('data.xlsx')


    # Global variable deceleration
    TotalCase = 0
    TotalCaseDone = 0
    # Total rows in excel file
    InvdividualSheetRows = 0
    InvdividualSheetLastCaseID = ""
    AttorneySheetLastCaseID = ""
    AttorneySheetRow = 0

    def ExcelColorGray(self, s):
        # reutrn excel row with gray of total len 24
        return ['background-color: gray'] * 28

    def ExcelColor(self, s):
        # reutrn excel row with yellow of total len 24
        return ['background-color: yellow'] * 28




    def addressfilter(self, addr):
        #		addr = """C/O Daniel L. Russello
        # McNevin & McInness, LLP
        # 5442 S. East Street, Suite C-14
        # Indianapolis, IN 46227"""

        # Spliting address by  new line so we can seperate all the variables
        addr = addr.split("\n")
        address = ""
        # print(addr)

        # Splitting string by new line to seperate address mailing name statezipcity
        addr1 = addr[-1].split(',')
        # First index will City
        city = addr1[0]
        # Split last index by space in which last index will be zip code and will index wil lbe state
        zipcode = addr1[-1].lstrip().split(' ')[-1]
        state = addr1[-1].lstrip().split(' ')[0]

        # removing city state zip line from the array
        addr.pop(len(addr) - 1)

        # iterating array to get address
        for i in range(len(addr)):
            # Geolocator geo code will return complete address if the address provided is correct that way we can find that the index of array is address of mailing name

            ##			location = self.geolocator.geocode(addr[i])
            try:
                location = do_geocode(addr[i])
                time.sleep(1)
            except:
                location = None
            # if provided value is not address it will raise an error if it does not we will store that address in adddress variable remove address index from array and break the loop
            try:
                demovar = location.address
                address = addr[i]
                addr.pop(i)
                # print(address)
                break
            except Exception as e:
                # print(str(e))
                pass
        # all the remaining indexes will mailing name
        mailingname = "".join(addr)

        return mailingname, address, city, state, zipcode

    def getinput(self):
        excel_data_df = pd.read_excel('input.xlsx', header=None)
        i = 0
        casenumbers = []
        for data in excel_data_df.values:
            # if its first iteration skip it, because its the header
            if i == 0:
                i += 1
                continue
            # Appending case number found in excel file to array
            casenumbers.append(data[0])
            i += 1
        # log
        print("Total Input Search queries found : ", len(casenumbers))

        # self.TotalCase = len(casenumbers)

        # return all the case number found in excel file
        return casenumbers

    def searchcase(self):

        # calling get input function, function will Extract all inputs from Input excel file
        casenumbers = self.getinput()
        # search query url
        ur = 'https://public.courts.in.gov/mycase/#/vw/Search'
        #ur = 'https://www.google.com'
        caselen = 0
        for case in casenumbers:
            self.TotalCaseDone = 0
            print("Searching for Case Number : ", case)
            self.Browser.get(ur)

            # Find the input text file of case number in the form
            casefield = WebDriverWait(self.Browser, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='SearchCaseNumber']")))
            #casefield = self.Browser.find_element_by_xpath("//input[@id='SearchCaseNumber']")
            # Clear Text Field
            casefield.clear()
            # Entering case number in the text field
            casefield.send_keys(case)

            time.sleep(3)
            # Find submit button
            submitbutton = self.Browser.find_element_by_xpath("//button[@class='btn btn-default']")

            # Submit the search query
            submitbutton.click()
            time.sleep(5)

            # search result function will calculate total result and iterate over all the found pages
            self.searchresults()

            caselen += 1
            # log
            print("Case queries done " + str(caselen) + " out of ", len(casenumbers))


    def searchresults(self):
        time.sleep(2)
        # Find total result found text i.e '1 to 20 of 577'
        totalresult = WebDriverWait(self.Browser, 10).until(EC.presence_of_element_located((By.XPATH,"//span[@data-bind='html: dpager.Showing']")))
        #totalresult = self.Browser.find_element_by_xpath("//span[@data-bind='html: dpager.Showing']").text
        # extract all numbers from '1 to 20 of 577' using regex
        totalresult = re.findall(r'\d+', totalresult.text)
        totalresult = [int(i) for i in totalresult]
        self.TotalCase = int(max(totalresult))
        # log
        print("Total Search result found : ", max(totalresult))
        # dividing the max number from regex output by total result per page
        totalresult = int(math.ceil(int(max(totalresult)) / 20))
        print("Total Pages :", totalresult)

        # loop till total pages
        for tot in range(totalresult):
            # Finding search result per page
            results = self.Browser.find_elements_by_xpath("//a[@class='result-title']")
            if len(results) < 1:
                self.Browser.refresh()
                results = self.Browser.find_elements_by_xpath("//a[@class='result-title']")
            # Calling function will take parameter of all search results , This function will click on each search result one by one and scrape data from it
            self.searchresultiterate(results)

            # Find and click on next page button
            nextbutton = self.Browser.find_element_by_xpath("//button[@title='Go to next result page']").click()
            time.sleep(3)
            # log
            print("Pages Done " + str(tot + 1) + " Out of ", totalresult)




    def searchresultiterate(self, results):
        # Iterating over all search result per page
        for i in range(len(results)):
            # Click on each search result if stale element exception find search result from page again
            try:
                results[i].click()
            except StaleElementReferenceException:
                try:
                    results = self.Browser.find_elements_by_xpath("//a[@class='result-title']")
                    results[i].click()
                except:
                    self.Browser.refresh()
                    time.sleep(4)
                    results = self.Browser.find_elements_by_xpath("//a[@class='result-title']")
                    results[i].click()


            # calling data extraction function, this function will extract all required data from the Case Page
            self.DataExtraction()

            # Previous page through js code
            self.Browser.execute_script("window.history.go(-1)")
            time.sleep(2)
            # log
            self.TotalCaseDone += 1
    ##            if i>1:
    ##                break
            print("Result(s) Scraped " + str(i + 1) + " Out of " + str(len(results)) + " Total Cases Scraped : " + str(self.TotalCaseDone) + " / ", self.TotalCase)

    def DataExtraction(self):
        #self.Browser.get("https://public.courts.in.gov/mycase/#/vw/CaseSummary/eyJ2Ijp7IkNhc2VUb2tlbiI6IjlCTmlYYzhMaVJaZ0l3UVpYVDl6Y3RULS1OcE9weW9hUXdRN21rR2R5ODgxIn19")
        time.sleep(4)

        # Finding first table in which Case Number is present (Case Detail Table)
        try:
            casetypevar = WebDriverWait(self.Browser, 10).until(EC.presence_of_all_elements_located((By.XPATH,"//div[@class='col-xs-12 col-sm-8 col-md-6']//table//tr")))
            #casetypevar = self.Browser.find_elements_by_xpath('//div[@class="col-xs-12 col-sm-8 col-md-6"]//table//tr')
        except NoSuchElementException:
            time.sleep(2)
            self.Browser.refresh()
            time.sleep(2)
            casetypevar = WebDriverWait(self.Browser, 10).until(EC.presence_of_all_elements_located((By.XPATH,"//div[@class='col-xs-12 col-sm-8 col-md-6']//table//tr")))

# Finding All Parties dropdowns
        partydetail = self.Browser.find_elements_by_xpath("//table[@class='ccs-parties table table-condensed table-hover']//span[@class='small glyphicon glyphicon-collapse-down']")
        # totlen variable is used for how many drop is being clicked
        totlen = 0
        uc = []
        # iterating over all the dropdowns found
        for pd in partydetail:
            # Click each and every one the of them if error means not clickable then skip it
            try:
                pd.click()
                totlen += 1
            except:
                # uc = index of unclickable divs
                totlen += 1
                uc.append(totlen)
                pass
        # Finding Table of parties
        pct = self.Browser.find_elements_by_xpath("//table[@class='ccs-parties table table-condensed table-hover']//tr")
        # Calling Fucntion partiescase takes parameter, Party Table,Case Detail Table,total len multiply by 2

        #print(pct)
        self.partiescase(pct, casetypevar, totlen * 2, uc)

        #print(self.casedetails(casetypevar))


    def partiescase(self, pct, casetypevar, totlen, uc):
        storearray = []
        rowtypearray =[]

        # Calling function Case details takes parameter , case detail table, This function will scrape all the required details from table i.e Case Number and will return 6 variables
        try:
            casenumber, court, type1, filed, status, statusdate = self.casedetails(casetypevar)
        except StaleElementReferenceException:
            casetypevar = self.Browser.find_elements_by_xpath('//div[@class="col-xs-12 col-sm-8 col-md-6"]//table//tr')
            casenumber, court, type1, filed, status, statusdate = self.casedetails(casetypevar)

        partyexists = self.checkotherpartiesexists(pct,totlen,uc)
        # This variable will use to skip one iteration after other
        skip = False
        countaddress,countattadress,countdod,countdesc,countatt,countattphone =0,0,0,0,0,0
        itercount = 0
        for i in range(totlen):
            dcGender,dcDOD,dcZip,dcState,dcCity,dcMailingName,dcAddress,dcName = "","","","","","","",""
            PR,PR_gender,PR_mailingname,PR_mailingaddres,PR_mailingcity,PR_mailingstate,PR_mailingzip = "","","","","","",""
            attorneyName,attorneyCompany,attorneyaddress,attorneycity,attorneystate,attorneyzipcode,attorneyphone = "","","","","","",""
            attorneysheet = False
            if skip == True:
                # skip False and skip iteration
                skip = False
                continue
            # if skip is False
            else:
                skip = True
            itercount += 1
            # if current index is the index of unclickable div then skip the iteration
            if itercount in uc:
                continue
            if 'Decedent' in pct[i].text:
                dcName = pct[i].text.replace('Decedent', '').lstrip()
                if 'DOD' in pct[i + 1].text:
                    dcDOD = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDOD']")[countdod].text
                else:
                    countdod -= 1
                if 'Address' in pct[i + 1].text:
                    addr = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[countaddress].text
                    dcMailingName, dcAddress, dcCity, dcState, dcZip = self.addressfilter(addr)
                else:
                    countaddress -= 1
                if 'Description' in pct[i + 1].text:
                    dcGender = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDesc']")[countdesc].text
                else:
                    countdesc -= 1
                countatt -= 1 
                countattadress -= 1 
                countattphone -= 1




            elif 'Executor' in pct[i].text:
                PR = pct[i].text.replace('Executor', '').lstrip()
                if 'Address' in pct[i + 1].text:
                    addr = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[countaddress].text
                    PR_mailingname, PR_mailingaddres, PR_mailingcity, PR_mailingstate, PR_mailingzip = self.addressfilter(addr)
                else:
                    countaddress -= 1
                if 'Description' in pct[i + 1].text:
                    PR_gender = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDesc']")[countdesc].text
                else:
                    countdesc -= 1


                if 'Attorney' in pct[i + 1].text:
  
                    try:
                        # Find all the attorney in the parties table take of text of current index attorney
                        attorneyName = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
                        # if attorney is not pro see
                        if 'Pro Se' not in attorneyName:
                            # Find all the attorney addresses in the parties table take of text of current index attorney
                            attorneyad = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadress].text

                            # Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
                            attorneyCompany, attorneyaddress, attorneycity, attorneystate, attorneyzipcode = self.addressfilter(attorneyad)

                            try:
                                attorneyphone = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyPhone']")[countattphone].text
                            except:
                                countattphone -= 1


                        else:
                            # if it is doesnt have the address

                            # else decrease one from attorney address counts varaible

                            countattadress -= 1
                            countattphone -= 1

                    except NoSuchElementException:
                        pass


                else:
                    countatt -= 1
                    countattadress -= 1
                    countattphone -= 1


                try:
                    if attorneyad == addr:
                        attorneysheet = True
                    elif addr == "":
                        attorneysheet = True
                    else:
                        attorneysheet = False
                except:
                    attorneysheet = True


                countdod -= 1














            elif 'Executrix' in pct[i].text:
                PR = pct[i].text.replace('Executrix', '').lstrip()
                if 'Address' in pct[i + 1].text:
                    addr = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[countaddress].text
                    PR_mailingname, PR_mailingaddres, PR_mailingcity, PR_mailingstate, PR_mailingzip = self.addressfilter(addr)
                else:
                    countaddress -= 1
                if 'Description' in pct[i + 1].text:
                    PR_gender = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDesc']")[countdesc].text
                else:
                    countdesc -= 1


                if 'Attorney' in pct[i + 1].text:
  
                    try:
                        # Find all the attorney in the parties table take of text of current index attorney
                        attorneyName = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
                        # if attorney is not pro see
                        if 'Pro Se' not in attorneyName:
                            # Find all the attorney addresses in the parties table take of text of current index attorney
                            attorneyad = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadress].text

                            # Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
                            attorneyCompany, attorneyaddress, attorneycity, attorneystate, attorneyzipcode = self.addressfilter(attorneyad)

                            try:
                                attorneyphone = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyPhone']")[countattphone].text
                            except:
                                countattphone -= 1


                        else:
                            # if it is doesnt have the address

                            # else decrease one from attorney address counts varaible

                            countattadress -= 1
                            countattphone -= 1

                    except NoSuchElementException:
                        pass


                else:
                    countatt -= 1
                    countattadress -= 1
                    countattphone -= 1



                try:
                    if attorneyad == addr:
                        attorneysheet = True
                    elif addr == "":
                        attorneysheet = True
                    else:
                        attorneysheet = False
                except:
                    attorneysheet = True


                countdod -= 1





            elif 'Special Administrator' in pct[i].text:
                PR = pct[i].text.replace('Special Administrator', '').lstrip()
                if 'Address' in pct[i + 1].text:
                    addr = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[countaddress].text
                    PR_mailingname, PR_mailingaddres, PR_mailingcity, PR_mailingstate, PR_mailingzip = self.addressfilter(addr)
                else:
                    countaddress -= 1
                if 'Description' in pct[i + 1].text:
                    PR_gender = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDesc']")[countdesc].text
                else:
                    countdesc -= 1


                if 'Attorney' in pct[i + 1].text:
  
                    try:
                        # Find all the attorney in the parties table take of text of current index attorney
                        attorneyName = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
                        # if attorney is not pro see
                        if 'Pro Se' not in attorneyName:
                            # Find all the attorney addresses in the parties table take of text of current index attorney
                            attorneyad = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadress].text

                            # Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
                            attorneyCompany, attorneyaddress, attorneycity, attorneystate, attorneyzipcode = self.addressfilter(attorneyad)

                            try:
                                attorneyphone = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyPhone']")[countattphone].text
                            except:
                                countattphone -= 1


                        else:
                            # if it is doesnt have the address

                            # else decrease one from attorney address counts varaible

                            countattadress -= 1
                            countattphone -= 1

                    except NoSuchElementException:
                        pass


                else:
                    countatt -= 1
                    countattadress -= 1
                    countattphone -= 1



                try:
                    if attorneyad == addr:
                        attorneysheet = True
                    elif addr == "":
                        attorneysheet = True
                    else:
                        attorneysheet = False
                except:
                    attorneysheet = True


                countdod -= 1






            elif 'Successor Personal Representative' in pct[i].text:
               # print("srp")
                PR = pct[i].text.replace('Successor Personal Representative', '').lstrip()
                if 'Address' in pct[i + 1].text:
                    addr = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[countaddress].text
                    PR_mailingname, PR_mailingaddres, PR_mailingcity, PR_mailingstate, PR_mailingzip = self.addressfilter(addr)
                else:
                    countaddress -= 1
                if 'Description' in pct[i + 1].text:
                    PR_gender = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDesc']")[countdesc].text
                else:
                    countdesc -= 1


                if 'Attorney' in pct[i + 1].text:
  
                    try:
                        # Find all the attorney in the parties table take of text of current index attorney
                        attorneyName = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
                        # if attorney is not pro see
                        if 'Pro Se' not in attorneyName:
                            # Find all the attorney addresses in the parties table take of text of current index attorney
                            attorneyad = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadress].text

                            # Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
                            attorneyCompany, attorneyaddress, attorneycity, attorneystate, attorneyzipcode = self.addressfilter(attorneyad)

                            try:
                                attorneyphone = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyPhone']")[countattphone].text
                            except:
                                countattphone -= 1


                        else:
                            # if it is doesnt have the address

                            # else decrease one from attorney address counts varaible

                            countattadress -= 1
                            countattphone -= 1

                    except NoSuchElementException:
                        pass


                else:
                    countatt -= 1
                    countattadress -= 1
                    countattphone -= 1


                try:
                    if attorneyad == addr:
                        attorneysheet = True
                    elif addr == "":
                        attorneysheet = True
                    else:
                        attorneysheet = False
                except:
                    attorneysheet = True


                countdod -= 1







            elif 'Co-Personal Representative' in pct[i].text:
                PR = pct[i].text.replace('Co-Personal Representative', '').lstrip()
                if 'Address' in pct[i + 1].text:
                    addr = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[countaddress].text
                    PR_mailingname, PR_mailingaddres, PR_mailingcity, PR_mailingstate, PR_mailingzip = self.addressfilter(addr)
                else:
                    countaddress -= 1
                if 'Description' in pct[i + 1].text:
                    PR_gender = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDesc']")[countdesc].text
                else:
                    countdesc -= 1


                if 'Attorney' in pct[i + 1].text:
  
                    try:
                        # Find all the attorney in the parties table take of text of current index attorney
                        attorneyName = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
                        # if attorney is not pro see
                        if 'Pro Se' not in attorneyName:
                            # Find all the attorney addresses in the parties table take of text of current index attorney
                            attorneyad = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadress].text

                            # Calling Function address filter takes parame    ter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
                            attorneyCompany, attorneyaddress, attorneycity, attorneystate, attorneyzipcode = self.addressfilter(attorneyad)

                            try:
                                attorneyphone = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyPhone']")[countattphone].text
                            except:
                                countattphone -= 1


                        else:
                            # if it is doesnt have the address

                            # else decrease one from attorney address counts varaible

                            countattadress -= 1
                            countattphone -= 1

                    except NoSuchElementException:
                        pass


                else:
                    countatt -= 1
                    countattadress -= 1
                    countattphone -= 1


                try:
                    if attorneyad == addr:
                        attorneysheet = True
                    elif addr == "":
                        attorneysheet = True
                    else:
                        attorneysheet = False
                except:
                    attorneysheet = True


                countdod -= 1




            elif 'Personal Representative' in pct[i].text:
                PR = pct[i].text.replace('Personal Representative', '').lstrip()
                if 'Address' in pct[i + 1].text:
                    addr = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[countaddress].text
                    PR_mailingname, PR_mailingaddres, PR_mailingcity, PR_mailingstate, PR_mailingzip = self.addressfilter(addr)
                else:
                    countaddress -= 1
                if 'Description' in pct[i + 1].text:
                    PR_gender = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDesc']")[countdesc].text
                else:
                    countdesc -= 1

#
                if 'Attorney' in pct[i + 1].text:
  
                    try:
                        # Find all the attorney in the parties table take of text of current index attorney
                        attorneyName = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
                        # if attorney is not pro see
                        if 'Pro Se' not in attorneyName:
                            # Find all the attorney addresses in the parties table take of text of current index attorney
                            attorneyad = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadress].text

                            # Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
                            attorneyCompany, attorneyaddress, attorneycity, attorneystate, attorneyzipcode = self.addressfilter(attorneyad)

                            try:
                                attorneyphone = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyPhone']")[countattphone].text
                            except:
                                countattphone -= 1


                        else:
                            # if it is doesnt have the address

                            # else decrease one from attorney address counts varaible

                            countattadress -= 1
                            countattphone -= 1

                    except NoSuchElementException:
                        pass


                else:
                    countatt -= 1
                    countattadress -= 1
                    countattphone -= 1



                try:
                    if attorneyad == addr:
                        attorneysheet = True
                    elif addr == "":
                        attorneysheet = True
                    else:
                        attorneysheet = False
                except:
                    attorneysheet = True


                countdod -= 1






            elif 'Other' in pct[i].text:
               # print("other")
                PR = pct[i].text.replace('Other', '').lstrip()
                if 'Address' in pct[i + 1].text:
                    addr = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[countaddress].text
                    PR_mailingname, PR_mailingaddres, PR_mailingcity, PR_mailingstate, PR_mailingzip = self.addressfilter(addr)
                else:
                    countaddress -= 1
                if 'Description' in pct[i + 1].text:
                    PR_gender = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDesc']")[countdesc].text
                else:
                    countdesc -= 1

                if 'Attorney' in pct[i + 1].text:
  
                    try:
                        # Find all the attorney in the parties table take of text of current index attorney
                        attorneyName = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
                        # if attorney is not pro see
                        if 'Pro Se' not in attorneyName:
                            # Find all the attorney addresses in the parties table take of text of current index attorney
                            attorneyad = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadress].text

                            # Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
                            attorneyCompany, attorneyaddress, attorneycity, attorneystate, attorneyzipcode = self.addressfilter(attorneyad)

                            try:
                                attorneyphone = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyPhone']")[countattphone].text
                            except:
                                countattphone -= 1


                        else:
                            # if it is doesnt have the address

                            # else decrease one from attorney address counts varaible

                            countattadress -= 1
                            countattphone -= 1

                    except NoSuchElementException:
                        pass


                else:
                    countatt -= 1
                    countattadress -= 1
                    countattphone -= 1



                try:
                    if attorneyad == addr:
                        attorneysheet = True
                    elif addr == "":
                        attorneysheet = True
                    else:
                        attorneysheet = False
                except:
                    attorneysheet = True


                countdod -= 1







            elif 'Petitioner' in pct[i].text:
               # print("pet")
                if partyexists == False:
                    PR = pct[i].text.replace('Petitioner', '').lstrip()
                    if 'Address' in pct[i + 1].text:
                        addr = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[countaddress].text
                        PR_mailingname, PR_mailingaddres, PR_mailingcity, PR_mailingstate, PR_mailingzip = self.addressfilter(addr)
                    else:
                        countaddress -= 1
                    if 'Description' in pct[i + 1].text:
                        PR_gender = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyDesc']")[countdesc].text
                    else:
                        countdesc -= 1


                    if 'Attorney' in pct[i + 1].text:
      
                        try:
                            # Find all the attorney in the parties table take of text of current index attorney
                            attorneyName = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
                            # if attorney is not pro see
                            if 'Pro Se' not in attorneyName:
                                # Find all the attorney addresses in the parties table take of text of current index attorney
                                attorneyad = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadress].text

                                # Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
                                attorneyCompany, attorneyaddress, attorneycity, attorneystate, attorneyzipcode = self.addressfilter(attorneyad)

                                try:
                                    attorneyphone = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyPhone']")[countattphone].text
                                except:
                                    countattphone -= 1


                            else:
                                # if it is doesnt have the address

                                # else decrease one from attorney address counts varaible

                                countattadress -= 1
                                countattphone -= 1

                        except NoSuchElementException:
                            pass


                    else:
                        countatt -= 1
                        countattadress -= 1
                        countattphone -= 1


                    try:
                        if attorneyad == addr:
                            attorneysheet = True
                        elif addr == "":
                            attorneysheet = True
                        else:
                            attorneysheet = False
                    except:
                        attorneysheet = True
                countdod -= 1
            else:
                if 'Address' not in pct[i + 1].text:
                    countaddress -= 1
                if 'Description' not in pct[i + 1].text:
                    countdesc -= 1

                if 'DOD' not in pct[i + 1].text:
                    countdod -= 1

                if 'Attorney' in pct[i + 1].text:
                 #   print("true")
                    try:
                        # Find all the attorney in the parties table take of text of current index attorney
                        attorneyName = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
                        # if attorney is not pro see
                        if 'Pro Se' not in attorneyName:
                            # Find all the attorney addresses in the parties table take of text of current index attorney
                            attorneyad = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadress].text

                          
                            try:
                                attorneyphone = pct[i + 1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyPhone']")[countattphone].text
                            except:
                                countattphone -= 1


                        else:
                            # if it is doesnt have the address

                            # else decrease one from attorney address counts varaible

                            countattadress -= 1
                            countattphone -= 1

                    except NoSuchElementException:
                        pass


                else:
                   # print("true1")
                    countatt -= 1
                    countattadress -= 1
                    countattphone -= 1







           # print("attorney end:   ",countatt)





            countattphone += 1
            countattadress += 1
            countatt += 1
            countdesc += 1
            countdod += 1
            countaddress += 1
            rowtypearray.append(attorneysheet)
            storearray.append([casenumber, court, type1, filed, status, statusdate, PR, PR_gender, PR_mailingname,PR_mailingaddres, PR_mailingcity,
                       PR_mailingstate, PR_mailingzip, dcName, dcGender, dcDOD, dcMailingName, dcAddress, dcCity,
                       dcState, dcZip,attorneyName,attorneyCompany,attorneyaddress, attorneycity, attorneystate, attorneyzipcode,
                      attorneyphone])
        if True in rowtypearray:
            for st in storearray:

            # Calling excel write function will take 20 parameters of all excel col required by user, this function will write data in excel and save it
                self.ExcelWriteAttorney(st[0], st[1], st[2], st[3], st[4], st[5], st[6], st[7], st[8], st[9],
                                st[10], st[11], st[12], st[13], st[14], st[15], st[16],
                                st[17], st[18], st[19], st[20], st[21], st[22],
                                st[23], st[24], st[25], st[26], st[27])
        else:
            for st in storearray:

            # Calling excel write function will take 20 parameters of all excel col required by user, this function will write data in excel and save it
                self.ExcelWriteIndiviual(st[0], st[1], st[2], st[3], st[4], st[5], st[6], st[7], st[8], st[9],
                                st[10], st[11], st[12], st[13], st[14], st[15], st[16],
                                st[17], st[18], st[19], st[20], st[21], st[22],
                                st[23], st[24], st[25], st[26], st[27])



    def ExcelWriteAttorney(self, casenumber, court, type1, filed, status, statusdate, PR, PR_gender, PR_mailingname,PR_mailingaddres, PR_mailingcity,
                        PR_mailingstate, PR_mailingzip, decendentName, decendentGender, decendentDOD, decendentMailingname, decendentMailingaddress, decendentMailingcity,
                        decendentMailingstate, decendentMailingzip,attorneyName,attorneyCompany,attorneyaddress, attorneycity, attorneystate, attorneyzipcode,
                        attorneyphone):

        sheetname = 'Attorney'

        df = pd.DataFrame({"Case Number": [casenumber], "Status": [status], "Township": [court], "Type": [type1],
                           "Filed Date": [filed.title()], "Status Date": [statusdate.title()],
                           "Personal Representative": [PR.title()],"Gender":[PR_gender], "Mailing Name": [PR_mailingname.title()],
                           "Mailing Address": [PR_mailingaddres.title()], "Mailing City": [PR_mailingcity.title()],
                           "Mailing State": [PR_mailingstate.upper()], "Mailing Zip": [PR_mailingzip.title()],
                           "Decedent Name": [decendentName.title()],"Decendent Gender":[decendentGender],"DOD":[decendentDOD], "Decedent Mailing Name": [decendentMailingname.title()],
                           "Decedent Address": [decendentMailingaddress.title()], "Decedent City": [decendentMailingcity.title()],
                           "Decedent State": [decendentMailingstate.upper()], "Decedent Zip": [decendentMailingzip.title()],
                           "Attorney Name": [attorneyName.title()], "Attorney Address": [attorneyaddress.title()],
                           "Company Name": [attorneyCompany.title()],"Attorney Phone":[attorneyphone],
                           "Attorney City": [attorneycity.title()], "Attorney State": [attorneystate.upper()],
                           "Attorney Zip": [attorneyzipcode.title()]})



        if self.AttorneySheetRow == 0:
            df.to_excel(self.ExcelFile, index=False, sheet_name=sheetname)
            self.AttorneySheetRow = self.ExcelFile.sheets[sheetname].max_row
            self.AttorneySheetLastCaseID = casenumber
        else:
            # if this is the new case add a new line to excel before adding case data to excel
            if self.AttorneySheetLastCaseID != casenumber:
                # creating empty dataframe of element len 24
                df1 = pd.DataFrame({"Case Number": [""], "Status": [""], "Township": [""], "Type": [""],
                           "Filed Date": [""], "Status Date": [""],
                           "Personal Representative": [""],"Gender":[""], "Mailing Name": [""],
                           "Mailing Address": [""], "Mailing City": [""],
                           "Mailing State": [""], "Mailing Zip": [""],
                           "Decedent Name": [""],"Decendent Gender":[""],"DOD":[""], "Decedent Mailing Name": [""],
                           "Decedent Address": [""], "Decedent City": [""],
                           "Decedent State": [""], "Decedent Zip": [""],
                           "Attorney Name": [""], "Attorney Address": [""],
                           "Company Name": [""],"Attorney Phone":[""],
                           "Attorney City": [""], "Attorney State": [""],
                           "Attorney Zip": [""]})

                # applying color to the row axis 1 = row
                df1 = df1.style.apply(self.ExcelColorGray, axis=1)
                # df1 = df1.style.set_properties(**{'height': '300px'})
                # writing colored row to excel
                df1.to_excel(self.ExcelFile, index=False, sheet_name=sheetname, header=False, startrow=self.AttorneySheetRow)
                self.AttorneySheetRow = self.ExcelFile.sheets[sheetname].max_row
                # then writing data
                df.to_excel(self.ExcelFile, index=False, sheet_name=sheetname, header=False, startrow=self.AttorneySheetRow)
                self.AttorneySheetRow = self.ExcelFile.sheets[sheetname].max_row
            else:
                df.to_excel(self.ExcelFile, index=False, sheet_name=sheetname, header=False, startrow=self.AttorneySheetRow)
                self.AttorneySheetRow = self.ExcelFile.sheets[sheetname].max_row
            self.AttorneySheetLastCaseID = casenumber

        self.ExcelFile.save()
    def ExcelWriteIndiviual(self, casenumber, court, type1, filed, status, statusdate, PR, PR_gender, PR_mailingname,PR_mailingaddres, PR_mailingcity,
                        PR_mailingstate, PR_mailingzip, decendentName, decendentGender, decendentDOD, decendentMailingname, decendentMailingaddress, decendentMailingcity,
                        decendentMailingstate, decendentMailingzip,attorneyName,attorneyCompany,attorneyaddress, attorneycity, attorneystate, attorneyzipcode,
                        attorneyphone):

        sheetname = 'Individual'

        df = pd.DataFrame({"Case Number": [casenumber], "Status": [status], "Township": [court], "Type": [type1],
                           "Filed Date": [filed.title()], "Status Date": [statusdate.title()],
                           "Personal Representative": [PR.title()],"Gender":[PR_gender], "Mailing Name": [PR_mailingname.title()],
                           "Mailing Address": [PR_mailingaddres.title()], "Mailing City": [PR_mailingcity.title()],
                           "Mailing State": [PR_mailingstate.upper()], "Mailing Zip": [PR_mailingzip.title()],
                           "Decedent Name": [decendentName.title()],"Decendent Gender":[decendentGender],"DOD":[decendentDOD], "Decedent Mailing Name": [decendentMailingname.title()],
                           "Decedent Address": [decendentMailingaddress.title()], "Decedent City": [decendentMailingcity.title()],
                           "Decedent State": [decendentMailingstate.upper()], "Decedent Zip": [decendentMailingzip.title()],
                           "Attorney Name": [attorneyName.title()], "Attorney Address": [attorneyaddress.title()],
                           "Company Name": [attorneyCompany.title()],"Attorney Phone":[attorneyphone],
                           "Attorney City": [attorneycity.title()], "Attorney State": [attorneystate.upper()],
                           "Attorney Zip": [attorneyzipcode.title()]})

        if self.InvdividualSheetRows == 0:
            df.to_excel(self.ExcelFile, index=False, sheet_name=sheetname)
            self.InvdividualSheetRows = self.ExcelFile.sheets[sheetname].max_row
            self.InvdividualSheetLastCaseID = casenumber
        else:
            # if this is the new case add a new line to excel before adding case data to excel
            if self.InvdividualSheetLastCaseID != casenumber:
                # creating empty dataframe of element len 24
                df1 = pd.DataFrame({"Case Number": [""], "Status": [""], "Township": [""], "Type": [""],
                           "Filed Date": [""], "Status Date": [""],
                           "Personal Representative": [""],"Gender":[""], "Mailing Name": [""],
                           "Mailing Address": [""], "Mailing City": [""],
                           "Mailing State": [""], "Mailing Zip": [""],
                           "Decedent Name": [""],"Decendent Gender":[""],"DOD":[""], "Decedent Mailing Name": [""],
                           "Decedent Address": [""], "Decedent City": [""],
                           "Decedent State": [""], "Decedent Zip": [""],
                           "Attorney Name": [""], "Attorney Address": [""],
                           "Company Name": [""],"Attorney Phone":[""],
                           "Attorney City": [""], "Attorney State": [""],
                           "Attorney Zip": [""]})

                # applying color to the row axis 1 = row
                df1 = df1.style.apply(self.ExcelColorGray, axis=1)
                # df1 = df1.style.set_properties(**{'height': '300px'})
                # writing colored row to excel
                df1.to_excel(self.ExcelFile, index=False, sheet_name=sheetname, header=False, startrow=self.InvdividualSheetRows)
                self.InvdividualSheetRows = self.ExcelFile.sheets[sheetname].max_row
                # then writing data
                df.to_excel(self.ExcelFile, index=False, sheet_name=sheetname, header=False, startrow=self.InvdividualSheetRows)
                self.InvdividualSheetRows = self.ExcelFile.sheets[sheetname].max_row
            else:
                df.to_excel(self.ExcelFile, index=False, sheet_name=sheetname, header=False, startrow=self.InvdividualSheetRows)
                self.InvdividualSheetRows = self.ExcelFile.sheets[sheetname].max_row
            self.InvdividualSheetLastCaseID = casenumber
        self.ExcelFile.save()

    def checkotherpartiesexists(self,pct,totlen,uc):
        skip = False
        itercount = 0
        for i in range(totlen):
            if skip == True:
                # skip False and skip iteration
                skip = False
                continue
            # if skip is False
            else:
                skip = True
            itercount += 1
            # if current index is the index of unclickable div then skip the iteration
            if itercount in uc:
                continue
            if 'Personal Representative' in pct[i].text:
                return True
            elif 'Co-Personal Representative' in pct[i].text:
                return True
            elif 'Special Administrator' in pct[i].text:
                return True
            elif 'Executor' in pct[i].text:
                return True
            elif 'Executrix' in pct[i].text:
                return True
            elif 'Successor Personal Representative' in pct[i].text:
                return True
            elif 'Other' in pct[i].text:
                return True

        return False





    def casedetails(self, casetypevar):
        # required Variables
        casenumber = ''
        court = ''
        type1 = ''
        filed = ''
        status = ''
        statusdate = ''

        # iterating table rows (tr) of table
        for cases in casetypevar:
            # if case number is present in it remove case number text from the string and add it to variable
            if 'case number' in cases.text.lower():
                casenumber = cases.text.replace(' ', '').strip('CaseNumber')
            # if court is present in it remove court text from the string and add it to variable
            elif 'court' in cases.text.lower():
                court = cases.text.strip('Court').lstrip()
            # if type is present in it remove type text from the string and add it to variable
            elif 'type' in cases.text.lower():
                type1 = cases.text.replace('Type', '')
            # if filed is present in it remove filed text from the string and add it to variable
            elif 'filed' in cases.text.lower():
                filed = cases.text.replace('Filed', '')
            # if status is present in it
            elif 'status' in cases.text.lower():
                # Split status by comma(,) last index will be status and first will be status date always
                t = cases.text.replace('Status', '').split(',')
                status = t[-1]
                statusdate = t[0]

        # returning all the required varialbes
        return casenumber.strip(), court.strip(), type1.strip(), filed.strip(), status.strip(), statusdate.strip()



a = PublicCase()
#a.DataExtraction()
a.searchcase()
