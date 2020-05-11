# -*- coding: utf-8 -*-
"""
Created on Tue Apr  7 18:44:21 2020

@author: Riko EP
"""
from random import choice
from bs4 import BeautifulSoup
from geopy.geocoders import Nominatim
import requests as req
import numpy as np
import pandas as pd


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


def get_details(acc_id, address, lat, long):
    # untuk mengambil detail alamat
    if lat == 0 and long == 0:
        coordinate_list = np.array([acc_id, address, None, None, 0, None])

        transposed = np.transpose(coordinate_list)

        print("Location Not Found")

          return transposed

    try:
        locator = Nominatim(user_agent=choice(USER_AGENTS))
        location = locator.reverse([lat, long], timeout=10, language='ID')

        raw_address = location.raw

        if 'village' in raw_address['address']:
            item_list = [acc_id, address, float(raw_address['lat']), float(raw_address['lon']), int(
                raw_address['address']['postcode']), raw_address['address']['village']]
        elif 'municipality' in raw_address['address']:
            item_list = [acc_id, address, float(raw_address['lat']), float(raw_address['lon']), int(
                raw_address['address']['postcode']), raw_address['address']['municipality']]
        else:
            item_list = [acc_id, address, float(raw_address['lat']), float(
                raw_address['lon']), int(raw_address['address']['postcode']), None]

        coordinate_list = np.array(item_list)
        transposed = np.transpose(coordinate_list)

        print(transposed)

        return transposed

    except:
        coordinate_list = np.array([acc_id, address, lat, long, 0, None])

        transposed = np.transpose(coordinate_list)

        print("Post Code Not Found")

        return transposed


def collect_data():
    # membaca data, input file excel
    data = pd.read_excel('xxx.xlsx')
    id_list = data['AccountID']
    street_list = data['GAB']

    return id_list, street_list


def process_data(id_list, street_list):
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

        res = get_details(acc_id, address, lat, long)
        result.append(res)

    return result


def join_data(result):
    # join data long lat dengan data kode pos untuk memperoleh detail alamat
    df1 = pd.DataFrame(result)
    df1.columns = ['AccountID', 'Alamat', 'Latitude',
                   'Longitude', 'PostCode', 'Kelurahan']
    df1['PostCode'] = df1['PostCode'].astype(int)

    df2 = pd.read_excel('Data Kodepos Indonesia.xlsx')
    df2 = df2.loc[:, ['PostCode', 'Kecamatan', 'Kabupaten', 'Propinsi']]
    df2['PostCode'] = df2['PostCode'].astype(int)

    join = pd.merge(left=df1, right=df2, how='left',
                    on='PostCode').drop_duplicates('AccountID')
    result_data = join[['AccountID', 'Alamat', 'Latitude',
                        'Longitude', 'Kelurahan', 'Kecamatan', 'Kabupaten', 'Propinsi']]
    result_data.columns = ['AccountID', 'Alamat', 'Latitude', 'Longitude',
                           'Desa/Kelurahan', 'Kecamatan', 'Kota/Kabupaten', 'Provinsi']

    return result_data


if __name__ == '__main__':
    id_list, street_list = collect_data()

    result = process_data(id_list, street_list)

    result_data = join_data(result)

    result_data.to_excel('result.xlsx',
                         float_format='%.7f', index=False)
