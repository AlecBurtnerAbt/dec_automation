# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""


from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os, sys, inspect
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
from easygui import multenterbox, fileopenbox, boolbox, diropenbox
from openpyxl import *




def generate_browser():
    #chromedriver = os.path.join(current_folder,"chromedriver.exe")
    chromedriver = "C:\\Users\\chromedriver.exe"
    driver = webdriver.Chrome(executable_path=chromedriver) 
    return driver



def retrieve_hin_validations(hins, deas, user, password):
    '''
    function to retrive validation information from IQvia
    website.
    
    Inputs
    validations: should be a list of ION HIN #s to look up
    
    Outputs:
    Excel Spreadsheet of HIN#s to organization specialty
    '''

    completed_hin_validations = None
    completed_dea_validations = None
    driver = generate_browser()
    driver.implicitly_wait(15)
    driver.get('https://onekeyweb.imshealth.com/onekeyweb/')
    wait = WebDriverWait(driver,15)
    #pass User ID
    user_id_input = driver.find_element_by_xpath('//input[@id="txtUserID"]')
    user_id_input.send_keys(user)
    remember_me_checkbox = driver.find_element_by_xpath('//input[@id="chkRememberMe"]')    
    remember_me_checkbox.click()
    continue_button = driver.find_element_by_xpath('//input[@id="btnValidate"]')
    continue_button.click()
    
    #wait until password box is available, then enter
    password_input_box = wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@id="txtPassword"]')))
    password_input_box.send_keys(password)    
    submit_button = driver.find_element_by_xpath('//input[@id="btnLogin"]')    
    submit_button.click()    
    
    #find organizations link and click it
    organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
    organization_link.click()    
    
    #the below block of codes does hins
    if hins != None: 
        completed_hin_validations = {key:list() for key in hins}
        for hin in hins:
            hin = hin.strip()
            try:
                source_file_lookup_box = wait.until(EC.presence_of_element_located((By.XPATH,'//input[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_OrganizationSearch_chkSourceFileSearch"]')))
                source_file_lookup_box.click()
                hin_checkbox = wait.until(EC.presence_of_element_located((By.XPATH,'//input[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_OrganizationSearch_rdosearchInHIN"]')))
                hin_checkbox.click()
                hin_input_box = driver.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_OrganizationSearch_txtHINNumber"]')
                hin_input_box.send_keys(hin)
                submit_button2 = driver.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_OrganizationSearch_btnSearch"]')
                submit_button2.click()
                
                #see if anything comes back
                try:
                    facility_labels = driver.find_elements_by_xpath('//table[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceList_grdHINOrganizations"]/tbody/tr[1]/th')
                    facility_labels = ['facility '+x.text.lower() for x in facility_labels]
                    facility_data = driver.find_elements_by_xpath('//table[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceList_grdHINOrganizations"]/tbody/tr[2]/td')
                    facility_data = [x.text for x in facility_data]
                    links = driver.find_elements_by_xpath(f'//a[contains(text(),"{hin}")]')
                    links[0].click()
                    wait.until(EC.staleness_of(links[0]))
                    try:
                        validation_labels = driver.find_elements_by_xpath('//table[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceCrossRef_grdHCOSCrossReference"]/tbody/tr[1]/th')
                        validation_labels = ['validation '+x.text.lower() for x in validation_labels]
                        validation_data = driver.find_elements_by_xpath('//table[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceCrossRef_grdHCOSCrossReference"]/tbody/tr[2]/td')
                        validation_data = [x.text for x in validation_data]                        
                        facility_labels.extend(validation_labels)
                        facility_data.extend(validation_data)
                        complete_data = dict(zip(facility_labels, facility_data))
                        completed_hin_validations.update({hin:complete_data})
                        return_to_search_button = driver.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceCrossRef_btnReturnToResult"]')
                        return_to_search_button.click()
                        home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                        home_link.click()
                        organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                        organization_link.click()
                    except NoSuchElementException as ex:
                        try:
                            print('c')
                            err = driver.find_element_by_xpath('//td[contains(text(),"No HCOS Cross References were found.")]')
                            err = err.text.strip()
                            err_info = {'error': err}
                            completed_hin_validations.update({hin:err_info})
                            home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                            home_link.click()
                            organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                            organization_link.click() 
                        except:
                            print('d')
                            print(f'Other error for {hin}')
                            err_info = {'error':'other error'}
                            completed_hin_validations.update({hin:err_info})
                            home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                            home_link.click()
                            organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                            organization_link.click() 
                            continue
                            
                except NoSuchElementException as ex:
                    print('e')
                    specialty = 'Could not find in OneKey'
                    err_info = {'error':specialty}
                    completed_validations.update({hin:err_info})
                    home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                    home_link.click()
                    organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                    organization_link.click()
            except:
                print('f')
                err_info = {'error':'other error'}
                completed_hin_validations.update({hin:err_info})
                home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                home_link.click()
                organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                organization_link.click()
                continue
            
    #This block of code does the dea numbers
    if deas != None:
        completed_dea_validations = {key: list() for key in deas}
        for dea in deas:
            dea = dea.strip()
            try:
                source_file_lookup_box = wait.until(EC.presence_of_element_located((By.XPATH,'//input[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_OrganizationSearch_chkSourceFileSearch"]')))
                source_file_lookup_box.click()
                dea_input_box = driver.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_OrganizationSearch_txtDEANumber"]')
                dea_input_box.send_keys(dea)
                submit_button2 = driver.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_OrganizationSearch_btnSearch"]')
                submit_button2.click()
                
                #see if anything comes back
                try:
                    facility_labels = driver.find_elements_by_xpath('//table[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceList_grdDEAOrganizations"]/tbody/tr[1]/th')
                    facility_labels = ['facility '+x.text.lower() for x in facility_labels]
                    facility_data = driver.find_elements_by_xpath('//table[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceList_grdDEAOrganizations"]/tbody/tr[2]/td')
                    facility_data = [x.text for x in facility_data]
                    links = driver.find_elements_by_xpath(f'//a[contains(text(),"{dea}")]')
                    links[0].click()
                    try:
                        validation_labels = driver.find_elements_by_xpath('//table[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceCrossRef_grdHCOSCrossReference"]/tbody/tr[1]/th')
                        validation_labels = ['validation '+x.text.lower() for x in validation_labels]
                        validation_data = driver.find_elements_by_xpath('//table[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceCrossRef_grdHCOSCrossReference"]/tbody/tr[2]/td')
                        validation_data = [x.text for x in validation_data]                        
                        facility_labels.extend(validation_labels)
                        facility_data.extend(validation_data)
                        complete_data = dict(zip(facility_labels, facility_data))
                        completed_dea_validations.update({dea:complete_data})
                        return_to_search_button = driver.find_element_by_xpath('//input[@id="ctl00_ContentPlaceHolder1_plcBody_HCOS_SourceCrossRef_btnReturnToResult"]')
                        return_to_search_button.click()
                        home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                        home_link.click()
                        organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                        organization_link.click() 
                    except NoSuchElementException as ex:
                        try:
                            err = driver.find_element_by_xpath('//td[contains(text(),"No HCOS Cross References were found.")]')
                            err = err.text.strip()
                            err_info = {'error':err}
                            completed_dea_validations.update({dea:err_info})
                            home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                            home_link.click()
                            organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                            organization_link.click() 
                        except:
                            print(f'Other error for {dea}')
                            err_info = {'error':'No HCOS Cross References were found'}
                            completed_dea_validations.update({dea:err_info})
                            home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                            home_link.click()
                            organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                            organization_link.click() 
                            continue
                            
                except NoSuchElementException as ex:
                    print('e')
                    specialty = 'Could not find in OneKey'
                    err_info = {'error':specialty}
                    completed_validations.update({dea:err_info})
                    home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                    home_link.click()
                    organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                    organization_link.click()
            except:
                print('f')
                err_info = {'error':'other error'}
                completed_dea_validations.update({dea:err_info})
                home_link = driver.find_element_by_xpath('//a[text()="Home"]')
                home_link.click()
                organization_link = driver.find_element_by_xpath('//a[text()="Organizations"]')
                organization_link.click()
                continue 
    driver.close()
    return completed_hin_validations, completed_dea_validations



#function to determine which identifiers to use
def scope():
    get_hins = boolbox(msg='Would you like to validate HIN numbers?', title='Validate HIN numbers?')
    get_deas = boolbox(msg='Would you like to validate DEA numbers?', title='Validate DEA numbers?')
    return get_hins, get_deas
     
def get_variables(get_hins, get_deas):
    hins=None
    deas=None
    if get_hins:
        hin_file = fileopenbox(msg='Please select the Excel file with the HIN list',title='Select HIN list')
        hins =list( pd.read_excel(hin_file,use_cols='A').iloc[:,0])
    if get_deas:
        dea_file = fileopenbox(msg='Please select the Excel file with the DEA list', title='Select DEA list')
        deas =list( pd.read_excel(dea_file,use_cols='A').iloc[:,0])
    msg = "Please enter your user id and password for OneKeyWeb"
    title = "Validator Program"
    fieldNames = ["User Name","Password"]
    fieldValues = []
    user, password = multenterbox(msg,title,fieldNames,fieldValues)
    path = diropenbox(msg="Where would you like the output file to land?", title="Select Directory")
    return hins, deas, user, password, path



def write_output(path,completed_hin_validations=None,completed_dea_validations=None):
    '''
    Turns dictionary of HIN:Specialty into an excel output
    '''
    df = pd.DataFrame()
    df2 = pd.DataFrame()
    if completed_hin_validations:
        df = pd.DataFrame.from_dict(completed_hin_validations, orient='index').reset_index().rename(mapper={'index':'HIN'}, axis=1)
        df.to_excel(path+'\\'+'Completed HIN Validations.xlsx')
    if completed_dea_validations:
        df2 = pd.DataFrame.from_dict(completed_dea_validations, orient='index').reset_index().rename(mapper={'index':'DEA'}, axis=1)
        df2.to_excel(path+'\\'+'Completed DEA Validations.xlsx')

    
    
def main():
    get_hins, get_deas = scope()
    hins, deas, user, password, path = get_variables(get_hins, get_deas)
    validated_hins, validated_deas = retrieve_hin_validations(hins=hins, deas=deas,user=user,password=password)
    write_output(path=path,completed_hin_validations=validated_hins, completed_dea_validations=validated_deas)

if __name__ == "__main__":
    main()
