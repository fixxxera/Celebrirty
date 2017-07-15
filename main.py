import datetime
import math
import os

import requests
import xlsxwriter

session = requests.session()

headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "X-currency-code": "USD",
    "X-request-id": "7d3a9258ef1641:82f359fa14b5958d",
    "X-country-code": "BGR",
    "Refer": "https://www.celebritycruises.com/spa/",
    "Cookie": "celebrityWR=newsite; akaas_Prod_Celeb=2147483647~rv=69~id=ed4b99cbb23e8eb93c64506e8b8fd553; mbox=session#73d490d9e4a14340b9b1cd976eb1a1ef#1500142958|PC#73d490d9e4a14340b9b1cd976eb1a1ef.26_25#1563385898; AMCV_981337045329610C0A490D44%40AdobeOrg=283337926%7CMCMID%7C57376148977644089083001278996633580261%7CMCAAMLH-1500745854%7C6%7CMCAAMB-1500745854%7CNRX38WO0n5BH8Th-nqAG_A%7CMCAID%7CNONE; utag_main=v_id:015d475feb5400adbdaa7d09885000044003700900bd0$_sn:1$_ss:0$_st:1500142897854$ses_id:1500141054805%3Bexp-session$_pn:3%3Bexp-session$_prevpage:home%3Bexp-1500144691581$vapi_domain:celebritycruises.com; __zlcmid=hWgb56wg46706B; timeStamp=1500141062149; anonyToken=eyJhbGciOiJIUzM4NCIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE1MDAxNDEwOTM0MDksInRva2VuSWQiOiJkOWU2OTkxYi02M2Q5LTRmMjAtOGE0Zi04MjgwMzI5MTVhN2QifQ.-pa4tz3fvcwfILiS2KHHI7rn_6DkP7cMVUIz7o1xwcveopFei7tyu8HzWPRv1tGh; sessionID=1500141061477; sessionAgencyId=; authToken=eyJhbGciOiJIUzM4NCIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE1MDAxNDEwOTM0MDksInRva2VuSWQiOiJkOWU2OTkxYi02M2Q5LTRmMjAtOGE0Zi04MjgwMzI5MTVhN2QifQ.-pa4tz3fvcwfILiS2KHHI7rn_6DkP7cMVUIz7o1xwcveopFei7tyu8HzWPRv1tGh; timerAuth=1500141093568; s_getNewRepeat=1500141097250-New; gpv_pn=home; s_ppvl=home%2C19%2C13%2C729%2C1920%2C501%2C1920%2C1080%2C1%2CP; s_ppv=home%2C19%2C19%2C729%2C1920%2C501%2C1920%2C1080%2C1%2CP; s_cc=true; _ga=GA1.2.2000261071.1500141064; _gid=GA1.2.1893632878.1500141064; _gat_GA_USA_Traffic_NEWSITE=1; _gat_Whole_Site_For_Search=1; __qca=P0-1747121582-1500141063987; _gat_tealium_0=1; rr_rcs=eF4FwbkRgDAQA8DEEb1oRsInDjqgDT8JARlQP7tlub_nmtTKhEwqxKPGbmgDVN5xsg_lrEQ0G7EOIyMbup3WdOtVP2nHETc; mf_052a099e-1e93-47a8-854c-b8b99f949c57=-1; s_sq=celebritycruiseprod%3D%2526c.%2526a.%2526activitymap.%2526page%253Dhome%2526link%253DFIND%252520CRUISE%2526region%253Ditinerary-search-anchor%2526pageIDType%253D1%2526.activitymap%2526.a%2526.c%2526pid%253Dhome%2526pidt%253D1%2526oid%253DFIND%252520CRUISE%2526oidt%253D3%2526ot%253DSUBMIT; _uetsid=_uetad156ce5; lang=EN; locale=/int; office=IBR; currencyCode=USD; wuc=BGR; sailings-filter1=duration=2,3,4,5",
    "Authorization": 'token eyJhbGciOiJIUzM4NCIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE1MDAxNDEwOTM0MDksInRva2VuSWQiOiJkOWU2OTkxYi02M2Q5LTRmMjAtOGE0Zi04MjgwMzI5MTVhN2QifQ.-pa4tz3fvcwfILiS2KHHI7rn_6DkP7cMVUIz7o1xwcveopFei7tyu8HzWPRv1tGh',
}
session.headers.update(headers)
url = "https://www.celebritycruises.com/prd/cruises/?&office=IBR&bookingType=FIT&includeFacets=false&cruiseType=CO&includeResults=true&count=5&offset=0&groupBy=PACKAGE&sortBy=PRICE&sortOrder=ASCENDING"
page = session.get(url)
cruises = page.json()
start_row = 1
total_count = cruises['hits']
pages_count = int(math.ceil(total_count / 5))
all_cruises = []
itineraries = []
offset = 0
all_ports = []
while pages_count > 0:
    print("Downloading page", start_row)
    url = "https://www.celebritycruises.com/prd/cruises/?&office=IBR&bookingType=FIT&includeFacets=false&cruiseType=CO&includeResults=true&count=5&offset=" + str(
        offset) + "&groupBy=PACKAGE&sortBy=PRICE&sortOrder=ASCENDING"
    result_page = session.get(url).json()
    for p in result_page['packages']:
        itineraries.append(p)
    start_row += 1
    pages_count -= 1
    offset += 5


def convert_date(unformated):
    splitter = unformated.split("-")
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    if month == 'Jan':
        month = '1'
    elif month == 'Feb':
        month = '2'
    elif month == 'Mar':
        month = '3'
    elif month == 'Apr':
        month = '4'
    elif month == 'May':
        month = '5'
    elif month == 'Jun':
        month = '6'
    elif month == 'Jul':
        month = '7'
    elif month == 'Aug':
        month = '8'
    elif month == 'Sep':
        month = '9'
    elif month == 'Oct':
        month = '10'
    elif month == 'Nov':
        month = '11'
    elif month == 'Dec':
        month = '12'
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def calculate_days(date, duration):
    dateobj = datetime.datetime.strptime(date, "%m/%d/%Y")
    calculated = dateobj + datetime.timedelta(days=int(duration))
    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def match_by_meta(param):
    baltic = ['Petropavlovsk, Russia', 'Bergen, Norway', 'Flam, Norway', 'Geiranger, Norway', 'Alesund, Norway',
              'Stavanger, Norway', 'Skjolden, Norway', 'Stockholm, Sweden', 'Helsinki, Finland',
              'St. Petersburg, Russia', 'Tallinn, Estonia', 'Riga, Latvia', 'Warnemunde, Germany',
              'Copenhagen, Denmark', 'Kristiansand, Norway', 'Skagen, Denmark', 'Fredericia, Denmark',
              'Rostock (Berlin), Germany', 'Nynashamn, Sweden', 'Oslo, Norway', 'Amsterdam, Netherlands',
              'Reykjavik, Iceland',
              'Zeebrugge (Brussels), Belgium', 'Southampton, England']
    eastern_med = ['Athens (Piraeus), Greece', 'Katakolon, Greece', 'Dubrovnik, Croatia', 'Mykonos, Greece',
                   'Rhodes, Greece', 'Chania (Souda),Crete, Greece', 'Koper, Slovenia', 'Split, Croatia',
                   'Santorini, Greece', 'Zadar, Croatia', 'Corfu, Greece', 'Kotor, Montenegro']
    west_med = ['Catania,Sicily,Italy', 'Ajaccio, Corsica', 'Alicante, Spain', 'Barcelona, Spain', 'Bilbao, Spain',
                'Cadiz, Spain', 'Cannes, France', 'Cartagena, Spain', 'Florence / Pisa (Livorno),Italy',
                'Fuerteventura, Canary', 'Funchal (Madeira), Portugal', 'Genoa, Italy', 'Gibraltar, United Kingdom',
                'Ibiza, Spain', 'La Coruna, Spain', 'La Spezia, Italy', 'Lanzarote, Canary Islands',
                'Las Palmas, Gran Canaria', 'Lisbon, Portugal', 'Malaga, Spain', 'Marseille, France',
                'Messina (Sicily), Italy', 'Montecarlo, Monaco', 'Naples, Italy', 'Nice (Villefranche)',
                'Palma De Mallorca, Spain', 'Ponta Delgada, Azores', 'Portofino, Italy', 'Provence (Toulon), France',
                'Ravenna, Italy', 'Sete, France', 'St. Peter Port, Channel Isl', 'Tenerife, Canary Islands',
                'Valencia, Spain', 'Valletta, Malta', 'Venice, Italy', 'Vigo, Spain']
    europe = ['Rome (Civitavecchia), Italy', 'Le Havre (Paris), France', 'Akureyri, Iceland',
              'Belfast, Northern Ireland', 'Cherbourg, France', 'Cork (Cobh), Ireland', 'Dover, England',
              'Dublin, Ireland', 'Edinburgh, Scotland', 'Greenock (Glasgow), Scotland', 'Inverness/Loch Ness, Scotland',
              'Lerwick/Shetland, Scotland', 'Liverpool, England',
              'Waterford (Dunmore E.), Ireland']

    ports_visited = param

    ports_list = []
    for i in range(len(ports_visited)):

        if i == 0:
            pass
        else:
            ports_list.append(ports_visited[i])
    isBaltic = False
    isEMED = False
    isWMED = False
    isE = False
    for port in ports_list:
        if port in baltic:
            isBaltic = True
    for port in ports_list:
        if port in eastern_med:
            isEMED = True
            break
    for port in ports_list:
        if port in west_med:
            isWMED = True
            break
    for port in ports_list:
        if port in europe:
            isE = True
            break
    if isEMED:
        return ['Eastern Med', 'E']
    elif isWMED:
        return ['Western Med', 'E']
    elif isBaltic:
        return ['Baltic', 'E']
    else:
        if param[0] in baltic:
            return ['Baltic', 'E']
        else:
            return ['', 'E']


def get_vessel_id(ves_name):
    if ves_name == "Equinox":
        return "687"
    elif ves_name == "Solstice":
        return "579"
    elif ves_name == "Silhouette":
        return "737"
    elif ves_name == "Reflection":
        return "756"
    elif ves_name == "Eclipse":
        return "712"
    elif ves_name == "Xperience":
        return "1023"
    elif ves_name == "Xploration":
        return "1024"
    elif ves_name == "Constellation":
        return "403"
    elif ves_name == "Infinity":
        return "55"
    elif ves_name == "Millennium":
        return "58"
    elif ves_name == "Summit":
        return "60"
    elif ves_name == "Xpedition":
        return "438"
    else:
        return "000"


packages = set()
unique = set()


def get_destination(dc):
    if dc == 'CARIB':
        return ['C', 'Carib']
    elif dc == 'EUROP':
        return ['E', 'Europe']
    elif dc == 'T.ATL':
        return ['E', 'Europe']
    elif dc == 'FAR.E' or dc == "DUIND":
        return ['O', 'Exotics']
    elif dc == 'ALCAN':
        return ['A', 'Alaska']
    elif dc == 'PACIF':
        return ['PA', 'PACIF']
    elif dc == 'TPACI':
        return ['I', 'Transpacific']
    elif dc == 'HAWAI':
        return ['H', 'Hawaii']
    elif dc == 'AUSTL':
        return ['P', 'Australia/New Zealand']
    elif dc == 'BERMU':
        return ['BM', 'Bermuda']
    elif dc == 'ATLCO':
        return ['NN', 'Canada/New England']
    elif dc == 'BAHAM':
        return ['BH', 'Bahamas']
    elif dc == 'GALAP':
        return ['S', 'Galapagos']
    elif dc == 'SAMER':
        return ['S', 'South America']
    elif dc == 'T.PAN':
        return ['T', 'Panama Canal']
    pass


def split_australia(ports):
    p = ['Adelaide, Australia', 'Airlie Beach, Qld, Australia', 'Alotau',
         'Ben Boyd National Park (Scenic Cruising Port)', 'Broome', 'Burnie', 'Busselton', 'Cairns, Australia',
         'Cooktown', 'Eden', 'Esperance, Australia', 'Exmouth', 'Fiordland National Park (Scenic Cruising)',
         'Fraser Island', 'Perth (Fremantle), Australia', 'Geraldton', 'Gladstone', 'Hamilton Island',
         'Hobart, Tasmania', 'Kangaroo Island', 'Kimberley Coast (Scenic Cruising Port)', 'Mooloolaba - Sunshine Coast',
         'Mornington Peninsula', 'Napier, New Zealand', 'Picton, New Zealand', 'Port Lincoln', 'Portland, Maine',
         'Stewart Island', 'Sydney Harbour Mooring (Scenic Cruising Port)', 'Townsville',
         'White Island (Scenic Cruising Port)', 'Wilsons Promontory (Scenic Cruising Port)']

    o = ['Benoa', 'Ko Chang', 'Komodo Island', 'Krabi', 'Bangkok/Laem Chabang, Thailand', 'Langkawi', 'Lombok',
         'Makassar', 'Phuket, Thailand', 'Probolinggo', 'Sabang (Palau Weh)', 'Sihanoukville']

    ip = ['Apia, Samoa', 'Conflict Islands', "Dili - Timor L'Este", 'Dravuni Island', 'Gizo Island', 'Honiara',
          'Kawanasausau Strait & Milne Bay (Scenic Cruising Port)', 'Kiriwina Island', 'Kitava',
          'Lifou, Loyalty Island', 'Madang',
          'Mutiny on the Bounty (Scenic Cruising Port)', "Nuku 'alofa, Tonga", 'Rabaul', 'Santo',
          "Vavau (Neiafu), Tonga", 'Vitu Islands (Scenic Cruising Port)', 'Wewak', 'Lahaina (Maui), Hawaii']
    ports_list = []
    for i in range(len(ports)):

        if i == 0:
            pass
        else:
            ports_list.append(ports[i])
    result = []
    is_exotic = False
    is_pacific = False
    for element in o:
        if element in ports_list:
            is_exotic = True
    if not is_exotic:
        for element in ip:
            if element in ports_list:
                is_pacific = True
    if not is_pacific:
        for element in p:
            if element in ports_list:
                pass
    if is_exotic:
        result.append("Exotics")
        result.append("O")
        return result
    elif is_pacific:
        result.append("South Pacific -- All")
        result.append("I")
        return result
    else:
        result.append("Australia/New Zealand")
        result.append("P")
        return result


def split_carib(ports):
    cu = ['Santiago de Cuba', 'Cienfuegos', 'Havana']
    wc = ['Costa Maya, Mexico', 'Cozumel, Mexico', 'Falmouth, Jamaica', 'George Town, Grand Cayman',
          'Ocho Rios, Jamaica']

    ec = ['Basseterre, St. Kitts', 'Bridgetown, Barbados', 'Castries, St. Lucia', 'Charlotte Amalie, St. Thomas',
          'Fort De France', 'Kingstown, St. Vincent', 'Philipsburg, St. Maarten', 'Ponce, Puerto Rico',
          'Punta Cana, Dominican Rep', 'Roseau, Dominica', 'San Juan, Puerto Rico', 'St. Croix, U.S.V.I.',
          "St. George's, Grenada", "St. John's, Antigua", 'Tortola, B.V.I']

    bm = ['Kings Wharf, Bermuda']

    ports_list = []
    for i in range(len(ports)):

        if i == 0:
            pass
        else:
            ports_list.append(ports[i])
    result = []
    iscu = False
    isbm = False
    isec = False
    iswc = False
    for element in cu:
        if element in ports_list:
            iscu = True
    if not iscu:
        for element in bm:
            if element in ports_list:
                isbm = True
    if not isbm:
        for element in ec:
            if element in ports_list:
                isec = True
    if not isec:
        for element in wc:
            if element in ports_list:
                iswc = True
    if iscu:
        result.append("Cuba")
        result.append("C")
    elif isbm:
        result.append("Bermuda")
        result.append("BM")
        return result
    elif isec:
        result.append("East Carib")
        result.append("C")
        return result
    elif iswc:
        result.append("West Carib")
        result.append("C")
        return result
    else:
        result.append("Carib")
        result.append("C")
        return result


def parse_data(cruise):
    cruise_line_name = "Celebrity Cruises"
    cruise_id = "3"
    for c in cruise["results"]:
        code = c["destCode"]
        brochure_name = c["packageName"]
        vessel_name = c["shipNameSlug"].split("-")[1]
        vessel_id = get_vessel_id(vessel_name)
        number_of_nights = int(c["duration"])
        sailings = c['sailings']
        ports = c['itenaryports']
        days = c['days']
        days_set = set()
        package_id = c['packageID']
        if 'International Dateline (At Sea)' in ports:
            for i in days:
                days_set.add(i)
            if len(days) != len(days_set):
                number_of_nights -= 1
            else:
                number_of_nights += 1
        for s in sailings:
            sail_date = convert_date(s['startDate'])
            return_date = calculate_days(sail_date, number_of_nights)
            if "inside" in s:
                interior_bucket_price = s['inside']['price']
                if interior_bucket_price == "Sold Out":
                    interior_bucket_price = 'N/A'
                else:
                    interior_bucket_price = interior_bucket_price.split('.')[0].replace(',', '')
            else:
                interior_bucket_price = 'N/A'
            if "oceanView" in s:
                oceanview_bucket_price = s['oceanView']['price']
                if oceanview_bucket_price == "Sold Out":
                    oceanview_bucket_price = 'N/A'
                else:
                    oceanview_bucket_price = oceanview_bucket_price.split('.')[0].replace(',', '')
            else:
                oceanview_bucket_price = 'N/A'
            if "veranda" in s:
                balcony_bucket_price = s['veranda']['price']
                if balcony_bucket_price == "Sold Out":
                    balcony_bucket_price = 'N/A'
                else:
                    balcony_bucket_price = balcony_bucket_price.split('.')[0].replace(',', '')
            else:
                balcony_bucket_price = 'N/A'
            if "suite" in s:
                suite_bucket_price = s['suite']['price']
                if suite_bucket_price == "Sold Out":
                    suite_bucket_price = 'N/A'
                else:
                    suite_bucket_price = suite_bucket_price.split('.')[0].replace(',', '')
            else:
                suite_bucket_price = 'N/A'
            destination = get_destination(code)
            print(code)
            dest_code = destination[0]
            dest_name = destination[1]
            if dest_code == 'E':
                dest = match_by_meta(ports)
                dest_code = dest[1]
                dest_name = dest[0]
            if "Caribbean" in brochure_name:
                if "Western" in brochure_name or "West" in brochure_name:
                    dest_code = 'C'
                    dest_name = "West Carib"
                if "Eastern" in brochure_name:
                    dest_code = 'C'
                    dest_name = "East Carib"
            if dest_code == 'I':
                if "Japan" in brochure_name:
                    dest_code = "O"
                    dest_name = 'Exotics'
            if dest_name == 'Australia/New Zealand':
                dest = split_australia(ports)
                dest_code = dest[1]
                dest_name = dest[0]
            if dest_code == 'S':
                if 'Panama Canal, Panama' in ports:
                    dest_code = 'T'
                    dest_name = "Panama Canal"
            if dest_code == "C":
                for p in ports:
                    unique.add(p)
            if dest_code == "C" and dest_name == 'Carib':
                dest = split_carib(ports)
                dest_code = dest[1]
                dest_name = dest[0]
                if 'Oranjestad, Aruba' in ports:
                    dest_code = "C"
                    dest_name = 'East Carib'
            is_mexican = False
            if dest_name == "PACIF" and vessel_name == "Infinity":
                for p in ports:
                    if 'Ensenada' in p:
                        is_mexican = True
                        break
            if is_mexican:
                dest_code = "M"
                dest_name = "Mexico"
            else:
                if dest_name == "PACIF" and vessel_name == "Infinity":
                    dest_name = "Alaska"
                    dest_code = "A"
            temp = [dest_code, dest_name, vessel_id, vessel_name, cruise_id, cruise_line_name,
                    package_id,
                    brochure_name, number_of_nights, sail_date, return_date,
                    interior_bucket_price, oceanview_bucket_price, balcony_bucket_price, suite_bucket_price]
            print(temp)
            temp2 = [temp]
            all_cruises.append(temp2)


processed_cruises = 0


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Celebrity Cruises.xlsx'
    if not os.path.exists(userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(
            now.month) + '-' + str(now.day)):
        os.makedirs(
            userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    workbook = xlsxwriter.Workbook(path_to_file)

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 25)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:D", 25)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 30)
    worksheet.set_column("G:G", 20)
    worksheet.set_column("H:H", 50)
    worksheet.set_column("I:I", 20)
    worksheet.set_column("J:J", 20)
    worksheet.set_column("K:K", 20)
    worksheet.set_column("L:L", 20)
    worksheet.set_column("M:M", 25)
    worksheet.set_column("N:N", 20)
    worksheet.set_column("O:O", 20)
    worksheet.write('A1', 'DestinationCode', bold)
    worksheet.write('B1', 'DestinationName', bold)
    worksheet.write('C1', 'VesselID', bold)
    worksheet.write('D1', 'VesselName', bold)
    worksheet.write('E1', 'CruiseID', bold)
    worksheet.write('F1', 'CruiseLineName', bold)
    worksheet.write('G1', 'ItineraryID', bold)
    worksheet.write('H1', 'BrochureName', bold)
    worksheet.write('I1', 'NumberOfNights', bold)
    worksheet.write('J1', 'SailDate', bold)
    worksheet.write('K1', 'ReturnDate', bold)
    worksheet.write('L1', 'InteriorBucketPrice', bold)
    worksheet.write('M1', 'OceanViewBucketPrice', bold)
    worksheet.write('N1', 'BalconyBucketPrice', bold)
    worksheet.write('O1', 'SuiteBucketPrice', bold)
    row_count = 1
    money_format = workbook.add_format({'bold': True})
    ordinary_number = workbook.add_format({"num_format": '#,##0'})
    date_format = workbook.add_format({'num_format': 'm d yyyy'})
    centered = workbook.add_format({'bold': True})
    money_format.set_align("center")
    money_format.set_bold(True)
    date_format.set_bold(True)
    centered.set_bold(True)
    ordinary_number.set_bold(True)
    ordinary_number.set_align("center")
    date_format.set_align("center")
    centered.set_align("center")
    for ship in data_array:
        for l in ship:
            column_count = 0
            for r in l:
                try:
                    if column_count == 0:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 1:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 2:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 3:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 4:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 5:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 6:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 7:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 8:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 9:
                        date_time = datetime.datetime.strptime(str(r), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, centered)
                        column_count += 1
                    elif column_count == 10:
                        date_time = datetime.datetime.strptime(str(r), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, centered)
                        column_count += 1
                    elif column_count == 11:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 12:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 13:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 14:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                except ValueError:
                    worksheet.write_string(row_count, column_count, r, centered)
                    column_count += 1
            row_count += 1
    workbook.close()
    pass


for i in itineraries:
    number_of_days = i['duration']
    brochure_name = i['name']
    package_id = i['packageId']
    vessel_name = i['ship']['shortName'].title()
    vessel_id = get_vessel_id(vessel_name)
    cruise_line_name = "Celebrity Cruises"
    cruise_id = "3"
    if not i['justACruise']:
        continue
    destination = i['destinationCode']
    dest = get_destination(destination)
    destination_code = dest[0]
    destination_name = dest[1]
    ports = []
    days = set()
    for itin in i['itineraries']:
        ports.append(itin['port']['name'].title())
        days.add(itin['dayNumber'])
    if 'International Dateline' in ports:
        if len(days) != len(ports):
            number_of_days -= 1
        else:
            number_of_days += 1
    for s in i['sailings']:
        inside = 'N/A'
        outside = 'N/A'
        balcony = 'N/A'
        deluxe = 'N/A'
        suite = 'N/A'
        sail_date = convert_date(s['sailDate'])
        return_date = calculate_days(sail_date, number_of_days)
        for p in s['categoryPrices']:
            if p['name'] == 'interior':
                price = p['price']
                if price == "Sold Out" or price is None or price == '':
                    inside = 'N/A'
                else:
                    inside = price
            elif p['name'] == 'outside':
                price = p['price']
                if price == "Sold Out" or price is None or price == '':
                    outside = 'N/A'
                else:
                    outside = price
            elif p['name'] == 'balcony':
                price = p['price']
                if price == "Sold Out" or price is None or price == '':
                    balcony = 'N/A'
                else:
                    balcony = price
            elif p['name'] == 'deluxe':
                price = p['price']
                if price == "Sold Out" or price is None or price == '':
                    suite = 'N/A'
                else:
                    suite = price
        if destination_code == 'E':
            dest = match_by_meta(ports)
            destination_code = dest[1]
            destination_name = dest[0]
        if "Caribbean" in brochure_name:
            if "Western" in brochure_name or "West" in brochure_name:
                destination_code = 'C'
                destination_name = "West Carib"
            if "Eastern" in brochure_name:
                destination_code = 'C'
                destination_name = "East Carib"
        if destination_code == 'I':
            if "Japan" in brochure_name:
                destination_code = "O"
                destination_name = 'Exotics'
        if destination_name == 'Australia/New Zealand':
            dest = split_australia(ports)
            destination_code = dest[1]
            destination_name = dest[0]
        if destination_code == 'S':
            if 'Panama Canal, Panama' in ports:
                destination_code = 'T'
                destination_name = "Panama Canal"
        if destination_code == "C" and destination_name == 'Carib':
            dest = split_carib(ports)
            destination_code = dest[1]
            destination_name = dest[0]
            if 'Oranjestad, Aruba' in ports:
                destination_code = "C"
                destination_name = 'East Carib'
        is_mexican = False
        if destination_name == "PACIF" and vessel_name == "Infinity":
            for p in ports:
                if 'Ensenada' in p:
                    is_mexican = True
                    break
        if is_mexican:
            destination_code = "MX"
            destination_name = "Mexico"
        else:
            if destination_name == "PACIF" and vessel_name == "Infinity":
                destination_name = "Alaska"
                destination_code = "A"
        temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, cruise_line_name,
                package_id,
                brochure_name, number_of_days, sail_date, return_date,
                inside, outside, balcony, suite]
        print(temp)
        temp2 = [temp]
        all_cruises.append(temp2)
write_file_to_excell(all_cruises)
f = open("ports.txt", 'w')
for row in list(unique):
    f.write(row + '\n')
f.close()
