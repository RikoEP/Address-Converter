# -*- coding: utf-8 -*-
"""
Created on Tue Apr  7 18:44:21 2020

@author: Riko EP
"""
from random import choice
from bs4 import BeautifulSoup
import requests as req
import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import re


DEKSTOP_AGENTS = ['Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 5.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36',
                  'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36']

USER_AGENTS = ['My Geocode',
               'Geocode 1'
               'Test',
               'Geocode Python',
               'Test Geocode',
               'Address Converter',
               'My Converter',
               'Address To Coordinate',
               'Reverse Geocode',
               'Python Geocode']


def random_headers():
    # fungsi untuk random user agent -> agar tidak dicurigai google
    return {'User-Agent': choice(DEKSTOP_AGENTS), 'Accept': 'text/html,application/xhtml+xml,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'}


def get_coordinate(address):
    # untuk mengambil koordinat long lat
    try:
        url = "https://www.google.com/maps/search/" + address
        resp = req.get(url, timeout=20, headers=random_headers())
        soup = BeautifulSoup(resp.text, 'html.parser')

        meta = soup.find('meta', {'property': 'og:image'})['content']
        coordinate_string = meta[meta.find('=') + 1:meta.find('&')]

        lat = float(coordinate_string[:coordinate_string.find('%')])
        long = float(coordinate_string[coordinate_string.find('%') + 3:])

        return lat, long

    except:
        return 0, 0
    
def get_address(driver, lat, long):
    try:
        driver.get('https://www.google.com/maps')
        time.sleep(5)
        driver.find_element_by_id('searchboxinput').send_keys(str(lat) + ' ' + str(long))
        time.sleep(1)
        driver.find_element_by_id('searchbox-searchbutton').click()
        time.sleep(5)
        address = driver.find_element_by_xpath('//*[@id="pane"]/div/div[1]/div/div/div[8]/div/div[1]/span[3]/span[3]').text

        return address

    except:
        return None


def get_details(acc_id, address, lat, long, driver):
    try:
        # untuk mengambil detail alamat
        if lat == 0 and long == 0:
            coordinate_list = np.array([acc_id, address, None, None, None, None, None, None])
    
            transposed = np.transpose(coordinate_list)
    
            print("Location Not Found")
    
            return transposed
        
        address_map = get_address(driver, lat, long)
        adress_split = address_map.split(',')
        
        municipality = adress_split[len(adress_split) - 4][1:]
        region =  adress_split[len(adress_split) - 3][1:]
        city = adress_split[len(adress_split) - 2][1:]
        province = adress_split[len(adress_split) - 1][1:]
        
        if bool(re.search(r'\d', province)):
            province = province[:province.rfind(' ')]
    
        item_list = [acc_id, address, lat, long, municipality, region, city, province]
        coordinate_list = np.array(item_list)
        transposed = np.transpose(coordinate_list)

        print(transposed)

        return transposed

    except:
        coordinate_list = np.array([acc_id, address, lat, long, None, None, None, None])

        transposed = np.transpose(coordinate_list)

        print("Post Code Not Found")

        return transposed


def collect_data():
    # membaca data, input file excel
    data = pd.read_excel('EXCEL_FILE')
    id_list = data['AccountID']
    street_list = data['Street']

    return id_list, street_list


def process_data(id_list, street_list):
    driver = None
    try:
        options = Options()
        options.add_argument("window-size=1400,600")
        user_agent = random_headers()
        options.add_argument(f'user-agent={user_agent}')
        driver = webdriver.Chrome(executable_path='PATH_TO_WEBDRIVER', chrome_options=options)
        
        # transformasi data dan cleaning data alamat
        result = []
        count = 1
        for acc_id, address in zip(id_list, street_list):
            print(count)
            count += 1
    
            lat, long = get_coordinate(address)
    
            if 'RT' in address:
                if 'KEL' in address:
                    str1 = address.split('RT')[0]
                    str2 = address.split('KEL')[1]
                    new_address = str1 + 'KEL' + str2
                else:
                    str1 = address.split('RT')[0]
                    new_address = str1
    
                lat, long = get_coordinate(new_address)
            else:
                lat, long = get_coordinate(address)
    
            res = get_details(acc_id, address, lat, long, driver)
            result.append(res)
    
        return result
    
    finally:
            if driver is not None:
                driver.close()
                print('Process Complete')


if __name__ == '__main__':
    id_list, street_list = collect_data()

    result = process_data(id_list, street_list)

    df = pd.DataFrame(result)
    df.columns = ['Account ID', 'Address', 'Latitude', 'Longitue', 'Desa/Kelurahan', 'Kecamatan', 'Kota/Kabupaten', 'Provinsi']
    
    df.to_excel('SAVE_PATH', float_format='%.7f', index=False)
