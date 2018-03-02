# -*- coding: utf-8 -*-
import xlsxwriter
import urllib2
from bs4 import BeautifulSoup

organization_number = 0
employee_number = 1

org_name_max_width = 0
employee_name_max_width = 0
employee_type_max_width = 0
employee_title_max_width = 0
employee_address_max_width = 0
employee_city_max_width = 0
employee_state_max_width = 0
employee_zip_max_width = 0
employee_hoa_max_width = 0
employee_phone_max_width = 0
employee_contact_max_width = 0
employee_website_max_width = 0



workbook = xlsxwriter.Workbook('result.xlsx')
organization_sheet = workbook.add_worksheet('Organization')
employee_sheet = workbook.add_worksheet("Employee")

employee_sheet.write(0, 0, "Organization")
employee_sheet.write(0, 1, "NAME")
employee_sheet.write(0, 2, "TYPE")
employee_sheet.write(0, 3, "TITLE")
employee_sheet.write(0, 4, "ADDRESS")
employee_sheet.write(0, 5, "CITY")
employee_sheet.write(0, 6, "STATE")
employee_sheet.write(0, 7, "ZIP")
employee_sheet.write(0, 8, "HOA")
employee_sheet.write(0, 9, "PHONE")
employee_sheet.write(0, 10, "CONTACT")
employee_sheet.write(0, 11, "WEBSITE")

def get_employees(org_name, url):
    global employee_number, employee_name_max_width, employee_type_max_width, employee_phone_max_width, employee_address_max_width, employee_city_max_width, employee_contact_max_width
    global employee_state_max_width, employee_title_max_width, employee_sheet, employee_website_max_width, employee_zip_max_width, employee_hoa_max_width
    website_url = 'http://www.homeowners-associations-florida.com/' + url
    header = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36'
    }

    request = urllib2.Request(website_url, headers=header)

    try:
        web_page = urllib2.urlopen(request)
    except urllib2.HTTPError, e:
        # print( e.fp.read() )
        return

    page_content = BeautifulSoup(web_page, 'html.parser')
    data = page_content.find_all('table', {'class': 'rnr-vrecord'})
    for item in data:
        e_name = item.find('span', id = lambda x: x and x.endswith('_NAME')).get_text().strip()
        e_type = item.find('span', id = lambda x: x and x.endswith('_TYPE')).get_text().strip()
        e_title = item.find('span', id = lambda x: x and x.endswith('_TITLE')).get_text().strip()
        e_address = item.find('span', id = lambda x: x and x.endswith('_ADDRESS')).get_text().strip()
        e_city = item.find('span', id = lambda x: x and x.endswith('_CITY')).get_text().strip()
        e_state = item.find('span', id = lambda x: x and x.endswith('_STATE')).get_text().strip()
        e_zip = item.find('span', id = lambda x: x and x.endswith('_ZIP')).get_text().strip()
        e_hoa = item.find('span', id = lambda x: x and x.endswith('_FULLSUB')).get_text().strip()
        e_phone = item.find('span', id = lambda x: x and x.endswith('_PHONE')).get_text().strip()
        e_contact = item.find('span', id = lambda x: x and x.endswith('_CONTACT')).get_text().strip()
        e_website = item.find('span', id = lambda x: x and x.endswith('_URL')).get_text().strip()

        employee_name_max_width = employee_name_max_width if employee_name_max_width > len(e_name) else len(e_name)
        employee_type_max_width = employee_type_max_width if employee_type_max_width > len(e_type) else len(e_type)
        employee_title_max_width = employee_title_max_width if employee_title_max_width > len(e_title) else len(e_title)
        employee_address_max_width = employee_address_max_width if employee_address_max_width > len(e_address) else len(e_address)
        employee_city_max_width = employee_city_max_width if employee_city_max_width > len(e_city) else len(e_city)
        employee_state_max_width = employee_state_max_width if employee_state_max_width > len(e_state) else len(e_state)
        employee_zip_max_width = employee_zip_max_width if employee_zip_max_width > len(e_zip) else len(e_zip)
        employee_hoa_max_width = employee_hoa_max_width if employee_hoa_max_width > len(e_hoa) else len(e_hoa)
        employee_phone_max_width = employee_phone_max_width if employee_phone_max_width > len(e_phone) else len(e_phone)
        employee_contact_max_width = employee_contact_max_width if employee_contact_max_width > len(e_contact) else len(e_contact)
        employee_website_max_width = employee_website_max_width if employee_website_max_width > len(e_website) else len(e_website)
        
        employee_sheet.write(employee_number, 0, org_name)
        employee_sheet.write(employee_number, 1, e_name)
        employee_sheet.write(employee_number, 2, e_type)
        employee_sheet.write(employee_number, 3, e_title)
        employee_sheet.write(employee_number, 4, e_address)
        employee_sheet.write(employee_number, 5, e_city)
        employee_sheet.write(employee_number, 6, e_state)
        employee_sheet.write(employee_number, 7, e_zip)
        employee_sheet.write(employee_number, 8, e_hoa)
        employee_sheet.write(employee_number, 9, e_phone)
        employee_sheet.write(employee_number, 10, e_contact)
        employee_sheet.write(employee_number, 11, e_website)

        employee_number = employee_number + 1

for i in range(1, 106):
    # specify the url
    website_url = 'http://www.homeowners-associations-florida.com/florida_hoa_search_list.php?pagesize=500&goto=' + str(i)
    header = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36'
    }

    request = urllib2.Request(website_url, headers=header)

    try:
        web_page = urllib2.urlopen(request)
    except urllib2.HTTPError, e:
        # print(e.fp.read())
        continue

    page_content = BeautifulSoup(web_page, 'html.parser')
    data = page_content.find_all('tr', {'class': 'rnr-row'})

    for item in data:
        organization_name = item.find('td').get_text().strip()
        width = len(organization_name)
        org_name_max_width = org_name_max_width if org_name_max_width > width else width
        organization_sheet.write(organization_number, 0, organization_name)
        organization_number = organization_number + 1
        organization_url = item.find('a')['href']
        get_employees(organization_name, organization_url)

organization_sheet.set_column(0, 0, org_name_max_width + 5)
employee_sheet.set_column(0, 0, org_name_max_width + 5)
employee_sheet.set_column(1, 1, employee_name_max_width + 5)
employee_sheet.set_column(2, 2, employee_type_max_width + 5)
employee_sheet.set_column(3, 3, employee_title_max_width + 5)
employee_sheet.set_column(4, 4, employee_address_max_width + 5)
employee_sheet.set_column(5, 5, employee_city_max_width + 5)
employee_sheet.set_column(6, 6, employee_state_max_width + 5)
employee_sheet.set_column(7, 7, employee_zip_max_width + 5)
employee_sheet.set_column(8, 8, employee_hoa_max_width + 5)
employee_sheet.set_column(9, 9, employee_phone_max_width + 5)
employee_sheet.set_column(10, 10, employee_contact_max_width + 5)
employee_sheet.set_column(11, 11, employee_website_max_width + 5)

workbook.close()

