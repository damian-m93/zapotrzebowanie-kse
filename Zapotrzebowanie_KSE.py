from datetime import datetime as dt
import pandas as pd
from urllib import error as httperr

zmiany_czasu = []
lista_bledow = []

'''
Brakuje godzin przy zmianie czasu z zimowego na letni!!!
'''


# WPKD

def wpkd(data_od, data_do):
    nazwa_pliku = 'wpkd.xlsx'
    adres = 'https://www.pse.pl/getcsv/-/export/csv/WPKD/data/'
    tabela_wpkd = pd.DataFrame(columns=['Data', 'Godzina', 'WPKD'])
    okres = pd.date_range(data_od, data_do)
    for data in okres:
        plik = adres + data.strftime('%Y%m%d')
        try:
            wpkd_dzien = pd.read_csv(plik, encoding='ISO-8859-1', sep=';', usecols=[0, 1, 2], parse_dates=[0])
        except httperr.HTTPError as error:
            print('---------')
            print('Błąd')
            print(data.strftime('%Y-%m-%d'))
            print(error)
            lista_bledow.append('(WPKD) Błąd dla daty: ' + data.strftime('%Y-%m-%d') + ' - ' + str(error))
            continue
        wpkd_dzien.rename(index=str, columns={'Krajowe zapotrzebowanie na moc': 'WPKD'}, inplace=True)
        try:
            wpkd_dzien['Godzina'] = wpkd_dzien['Godzina'].astype(int)
        except ValueError:
            wpkd_dzien['Godzina'] = pd.to_numeric(wpkd_dzien['Godzina'], errors='coerce')
            wpkd_dzien.dropna(axis=0, how='any', inplace=True)
            wpkd_dzien['Godzina'] = wpkd_dzien['Godzina'].astype(int)
            print('---------')
            print('Zmiana czasu - usunięto godzinę 2A')
            zmiany_czasu.append('(WPKD) Zmiana czasu - usunięto godzinę 2A dla daty: ' + data.strftime('%Y-%m-%d'))
        try:
            tabela_wpkd = pd.merge(tabela_wpkd, wpkd_dzien, how='outer')
            print('(WPKD) ' + data.strftime('%Y-%m-%d'))
        except BaseException as error:
            print('-----------')
            print('Błąd')
            print(data.strftime('%Y-%m-%d'))
            print(error)
            lista_bledow.append('(WPKD) Błąd dla daty: ' + data.strftime('%Y-%m-%d') + ' - ' + str(error))
            continue
    tabela_wpkd.to_excel(nazwa_pliku)
    return tabela_wpkd


# PKD

def pkd(data_od, data_do):
    nazwa_pliku = 'pkd.xlsx'
    adres = 'https://www.pse.pl/getcsv/-/export/csv/PKD/data/'
    tabela_pkd = pd.DataFrame(columns=['Data', 'Godzina', 'PKD'])
    okres = pd.date_range(data_od, data_do)
    for data in okres:
        plik = adres + data.strftime('%Y%m%d')
        try:
            pkd_dzien = pd.read_csv(plik, encoding='ISO-8859-1', sep=';', usecols=[0, 1, 2], parse_dates=[0])
        except httperr.HTTPError as error:
            print('---------')
            print('Błąd')
            print(data.strftime('%Y-%m-%d'))
            print(error)
            lista_bledow.append('(PKD) Błąd dla daty: ' + data.strftime('%Y-%m-%d') + ' - ' + str(error))
            continue
        pkd_dzien.rename(index=str, columns={'Krajowe zapotrzebowanie na moc': 'PKD'}, inplace=True)
        try:
            pkd_dzien['Godzina'] = pkd_dzien['Godzina'].astype(int)
        except ValueError:
            pkd_dzien['Godzina'] = pd.to_numeric(pkd_dzien['Godzina'], errors='coerce')
            pkd_dzien.dropna(axis=0, how='any', inplace=True)
            pkd_dzien['Godzina'] = pkd_dzien['Godzina'].astype(int)
            print('---------')
            print('Zmiana czasu - usunięto godzinę 2A')
            zmiany_czasu.append('(PKD) Zmiana czasu - usunięto godzinę 2A dla daty: ' + data.strftime('%Y-%m-%d'))
        try:
            tabela_pkd = pd.merge(tabela_pkd, pkd_dzien, how='outer')
            print('(PKD) ' + data.strftime('%Y-%m-%d'))
        except BaseException as error:
            print('-----------')
            print('Błąd')
            print(data.strftime('%Y-%m-%d'))
            print(error)
            lista_bledow.append('(PKD) Błąd dla daty: ' + data.strftime('%Y-%m-%d') + ' - ' + str(error))
            continue
    tabela_pkd.to_excel(nazwa_pliku)
    return tabela_pkd


# BPKD

def bpkd(data_od, data_do):
    nazwa_pliku = 'bpkd.xlsx'
    adres = 'https://www.pse.pl/getcsv/-/export/csv/BPKD/data/'
    tabela_bpkd = pd.DataFrame(columns=['Data', 'Godzina', 'BPKD'])
    okres = pd.date_range(data_od, data_do)
    for data in okres:
        plik = adres + data.strftime('%Y%m%d')
        try:
            bpkd_dzien = pd.read_csv(plik, encoding='ISO-8859-1', sep=';', usecols=[0, 1, 2], parse_dates=[0])
        except httperr.HTTPError as error:
            print('---------')
            print('Błąd')
            print(data.strftime('%Y-%m-%d'))
            print(error)
            lista_bledow.append('(BPKD) Błąd dla daty: ' + data.strftime('%Y-%m-%d') + ' - ' + str(error))
            continue
        bpkd_dzien.rename(index=str, columns={'Krajowe zapotrzebowanie na moc': 'BPKD'}, inplace=True)
        try:
            bpkd_dzien['Godzina'] = bpkd_dzien['Godzina'].astype(int)
        except ValueError:
            bpkd_dzien['Godzina'] = pd.to_numeric(bpkd_dzien['Godzina'], errors='coerce')
            bpkd_dzien.dropna(axis=0, how='any', inplace=True)
            bpkd_dzien['Godzina'] = bpkd_dzien['Godzina'].astype(int)
            print('---------')
            print('Zmiana czasu - usunięto godzinę 2A')
            zmiany_czasu.append('(BPKD) Zmiana czasu - usunięto godzinę 2A dla daty: ' + data.strftime('%Y-%m-%d'))
        try:
            tabela_bpkd = pd.merge(tabela_bpkd, bpkd_dzien, how='outer')
            print('(BPKD) ' + data.strftime('%Y-%m-%d'))
        except BaseException as error:
            print('-----------')
            print('Błąd')
            print(data.strftime('%Y-%m-%d'))
            print(error)
            lista_bledow.append('(BPKD) Błąd dla daty: ' + data.strftime('%Y-%m-%d') + ' - ' + str(error))
            continue
    tabela_bpkd.to_excel(nazwa_pliku)
    return tabela_bpkd


# WYK KSE

def wyk(data_od, data_do):
    nazwa_pliku = 'wyk.xlsx'
    adres = 'https://www.pse.pl/getcsv/-/export/csv/WYK_KSE/data/'
    tabela_wyk = pd.DataFrame(columns=['Data', 'Godzina', 'Wykonanie KSE'])
    okres = pd.date_range(data_od, data_do)
    for data in okres:
        plik = adres + data.strftime('%Y%m%d')
        try:
            wyk_dzien = pd.read_csv(plik, encoding='ISO-8859-1', sep=';', usecols=[0, 1, 2], parse_dates=[0], decimal=',')
        except httperr.HTTPError as error:
            print('---------')
            print('Błąd')
            print(data.strftime('%Y-%m-%d'))
            print(error)
            lista_bledow.append('(WYK) Błąd dla daty: ' + data.strftime('%Y-%m-%d') + ' - ' + str(error))
            continue
        wyk_dzien.rename(index=str, columns={'Krajowe zapotrzebowanie na moc': 'Wykonanie KSE'}, inplace=True)
        try:
            wyk_dzien['Godzina'] = wyk_dzien['Godzina'].astype(int)
        except ValueError:
            wyk_dzien['Godzina'] = pd.to_numeric(wyk_dzien['Godzina'], errors='coerce')
            wyk_dzien.dropna(axis=0, how='any', inplace=True)
            wyk_dzien['Godzina'] = wyk_dzien['Godzina'].astype(int)
            print('---------')
            print('Zmiana czasu - usunięto godzinę 2A')
            zmiany_czasu.append('(WYK) Zmiana czasu - usunięto godzinę 2A dla daty: ' + data.strftime('%Y-%m-%d'))
        try:
            tabela_wyk = pd.merge(tabela_wyk, wyk_dzien, how='outer')
            print('(WYK) ' + data.strftime('%Y-%m-%d'))
        except BaseException as error:
            print('-----------')
            print('Błąd')
            print(data.strftime('%Y-%m-%d'))
            print(error)
            lista_bledow.append('(WYK) Błąd dla daty: ' + data.strftime('%Y-%m-%d') + ' - ' + str(error))
            continue
    tabela_wyk.to_excel(nazwa_pliku)
    return tabela_wyk


# Złącz do sumarycznego arkusza

def polacz_arkusze():
    sciezka_plikow = ''
    zapotrzebowanie = pd.DataFrame(columns=['Data', 'Godzina'])
    try:
        wpkd_dane = pd.read_excel(sciezka_plikow + 'wpkd.xlsx')
        zapotrzebowanie = pd.merge(zapotrzebowanie, wpkd_dane, how='outer')
    except FileNotFoundError:
        print('Pominięto WPKD - brak pliku!')
    try:
        pkd_dane = pd.read_excel(sciezka_plikow + 'pkd.xlsx')
        zapotrzebowanie = pd.merge(zapotrzebowanie, pkd_dane, how='outer')
    except FileNotFoundError:
        print('Pominięto PKD - brak pliku!')
    try:
        bpkd_dane = pd.read_excel(sciezka_plikow + 'bpkd.xlsx')
        zapotrzebowanie = pd.merge(zapotrzebowanie, bpkd_dane, how='outer')
    except FileNotFoundError:
        print('Pominięto BPKD - brak pliku!')
    try:
        wyk_dane = pd.read_excel(sciezka_plikow + 'wyk.xlsx')
        zapotrzebowanie = pd.merge(zapotrzebowanie, wyk_dane, how='outer')
    except FileNotFoundError:
        print('Pominięto WYK - brak pliku!')
    zapotrzebowanie.sort_values(by=['Data', 'Godzina'], ascending=True, inplace=True)
    zapotrzebowanie.to_excel('Zapotrzebowanie_Plan_Wykonanie.xlsx')
    return zapotrzebowanie


# Wykonanie kodu

# wpkd = wpkd(dt(2009, 1, 1), dt(2018, 12, 31))
# pkd = pkd(dt(2009, 1, 1), dt(2018, 12, 31))
# bpkd = bpkd(dt(2012, 1, 1), dt(2018, 12, 31))
# wyk = wyk(dt(2009, 1, 1), dt(2018, 12, 31))


polacz_arkusze()

print('Zmiany czasu: ')
for i in zmiany_czasu:
    print(i)

print('-------------')
print('Lista błędów: ')
for i in lista_bledow:
    print(i)

