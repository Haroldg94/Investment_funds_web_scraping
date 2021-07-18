import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import time
import re


def selecting_fund(select, value_to_select):
    '''function that selects a investment fund by value on a webpage's dropdown '''
    
    # select by value
    select.select_by_value(value_to_select)
    time.sleep(2)
    
    html = browser.page_source
    
    return html


def open_excel(excel_path = 'investment_funds.xlsx'):
    saved_data = pd.read_excel(excel_path)
    return saved_data

def main(browser):
        
    browser.get(site_url)
    time.sleep(2)
    default_html = browser.page_source
    
    soup = BeautifulSoup(default_html, 'lxml')
    
    # Find select tag
    select_tag = soup.find("select")
    
    # find all option tag inside select tag
    options = select_tag.find_all("option")
    options_dict = {}
    
    # Iterate through all option tags and get inside text
    for i,option in enumerate(options):
        # Using a regular expression to extract the value of the option
        option_value = re.findall('\"([0-9]{1,2})',str(options[i]))[0]
        option_text = option.text
        options_dict[option_value] = option_text
        print(option_value,option_text)
        
    element_dropdown = browser.find_element_by_name('nmSelectFondo')
    select = Select(element_dropdown)
    
    historical_df = open_excel()
    
    temp = pd.DataFrame(columns = historical_df.columns)
    
    for val in options_dict.keys():
        fund_to_select = val
        
        print('Extracting information for: '+ options_dict[val])
        
        html = selecting_fund(select, fund_to_select)
    
        soup = BeautifulSoup(html, 'lxml')
    
        df = pd.read_html(html)
        df = pd.DataFrame(df[0])
    
        general_info_df = df.loc[:4,:].T
        general_info_df = general_info_df.loc[~general_info_df.duplicated(keep = 'first'),:].T
        general_info_df.columns = ['parameter','value']
        general_info_df.value.fillna('No aplica',inplace = True)
    
        days_profitability_df = df.loc[7:8,:].T
        days_profitability_df.columns = ['parameter','value']
    
        years_profitability_df = df.loc[10:11,:].T
        years_profitability_df.columns = ['parameter','value']
    
        closing_date_df = df.loc[13:14,:].T
        closing_date_df = closing_date_df.loc[~closing_date_df.duplicated(keep = 'first'),:].T
        closing_date_df.columns = ['parameter','value']
    
        fund_info = pd.concat([general_info_df,days_profitability_df,years_profitability_df,closing_date_df])
    
        fund_info.value = fund_info.value.str.replace('$','')
        fund_info.value = fund_info.value.str.replace('%','')
    
        key_list = fund_info.parameter.tolist()
        key_list.append('Fondo de Inversion')
        value_list = fund_info.value.tolist()
        value_list.append(options_dict[fund_to_select])
        value_list = [[x] for x in value_list]
        fund_info_dict = dict(zip(key_list,value_list))
    
        #Columns to replace some character
        str_columns = ['Valor de la unidad','7 días','30 días','180 días','Año corrido',
                       'Último año','Últimos dos años','Últimos tres años']
    
        current_fund_info = pd.DataFrame(data = fund_info_dict)
        current_fund_info['Valor en Pesos'] = current_fund_info['Valor en Pesos'].str.replace(',','').astype('float')
        current_fund_info[str_columns] = current_fund_info[str_columns].apply(lambda x: x.str.replace(',','.'), axis = 0)
        
        current_fund_info[str_columns] = current_fund_info[str_columns].apply(lambda x : x.astype('float') 
                                                                          if ~x.str.contains('N/A').any() 
                                                                          else x.astype('object'), axis = 0)
        
        #current_fund_info = current_fund_info.astype({'Valor de la unidad':'float64',                          
        #                                              'Valor en Pesos':'float64',
        #                                              '7 días':'float64',
        #                                              '30 días':'float64',
        #                                              '180 días':'float64',
        #                                              'Año corrido':'float64',
        #                                              'Último año':'float64',
        #                                              'Últimos dos años':'float64',
        #                                              'Últimos tres años':'float64',
        #                                             })
    
        if current_fund_info['Fondo administrador por'].str.contains('Fiduciaria').any():
            current_fund_info['Valor de la unidad'] = current_fund_info['Valor de la unidad']*1000
    
        current_fund_info['Fecha de Cierre'] = pd.to_datetime(current_fund_info['Fecha de Cierre'], format = '%Y/%m/%d')
        current_fund_info['Fecha Extracción']= pd.to_datetime(time.strftime('%Y/%m/%d', time.localtime(time.time())))
    
        temp = pd.concat([temp,current_fund_info])
        
    browser.quit()
    
    # Appending new data to the historial dataset
    historical_df = pd.concat([historical_df, temp]).reset_index(drop = True)
    
    print('upgrading file...')
    historical_df.to_excel('investment_funds.xlsx',index = False)
    print('file successfully upgraded')
    
if __name__ == "__main__":
    site_url = 'https://www.grupobancolombia.com/personas/productos-servicios/inversiones/fondos-inversion-colectiva/aplicacion-fondos/'
    
    options = Options()
    options.add_argument("--headless") # To Avoid the navigator to open
    
    browser = webdriver.Chrome(options= options)
    
    main(browser)