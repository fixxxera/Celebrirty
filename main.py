import datetime
import os

import requests
import xlsxwriter

session = requests.session()

headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive"
}
session.headers.update(headers)
url = "http://www.celebritycruises.com/cruise-search/searchResults?&dest=ANY&isWidget=false&sortBy=1&selectedInput" \
      "=&cruiseType=CO&cruisesOnly=Y&accessCabin=&includeAdjascentPorts=Y&state=&captain_id=&couponCodes=&isSenior" \
      "=&isMilitary=&isFireandPolice=&sailStartDate=ANY&sailEndDate=ANY&port=ANY&duration=ANY&port=ANY&ship=ANY" \
      "&startRow=0&_=1481022158300 "
page = session.get(url)
cruises = page.json()
start_row = 0
counter = round(int(cruises["totalPackages"]) / 10)
all_cruises = []
all_itineraries = int(cruises["totalPackages"])
current = 0


def make_request(rw):
    url_to_call = "http://www.celebritycruises.com/cruise-search/searchResults?&dest=ANY&isWidget=false&sortBy=7" \
                  "&selectedInput=&cruiseType=CO&cruisesOnly=Y&accessCabin=&includeAdjascentPorts=Y&state=&captain_id" \
                  "=&couponCodes=&isSenior=&isMilitary=&isFireandPolice=&sailStartDate=ANY&sailEndDate=ANY&port=ANY" \
                  "&duration=ANY&port=ANY&ship=ANY&startRow=" + str(rw) + "&_=1484662960933 "
    pages = session.get(url_to_call)
    return pages.json()


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
    bermuda = ['Kings Wharf, Bermuda']
    # alaska = ['Seward, Alaska', 'Hubbard Glacier, Alaska', 'Juneau, Alaska', 'Skagway, Alaska',
    #           'Icy Strait Point, Alaska', 'Ketchikan, Alaska', 'Inside Passage, Alaska', 'Vancouver, British Columbia',
    #           'Seattle, Washington', 'Astoria, Oregon', 'Sitka, Alaska', 'Tracy Arm Fjord, Alaska',
    #           'Victoria, British Columbia', 'Dutch Harbor, Alaska', 'Nanaimo, British Columbia']
    # trans_pacific = ['Colon, Panama', 'Panama Canal, Panama']
    # galapagos = ['Baltra, Galapagos', 'Daphne Island, Galapagos', 'Gardner Bay (Española), Galapagos',
    #              'Punta Suarez, Galapagos', 'Cormorant Point, Galapagos', 'Post Office, Galapagos',
    #              'Punta Espinoza, Galapagos', 'Dragon Hill, Galapagos', 'Puerto Ayora, Galapagos',
    #              'Puerto Egas, Galapagos', 'Rabida, Galapagos', 'Elizabeth Bay, Galapagos', 'Bartolome, Galapagos',
    #              'North Seymour, Galapagos', 'Santa Fe Island (Galapagos)', 'Mosquera Islet (Galapagos)',
    #              'Punta Moreno, Isabela', 'Urvina Bay, Isabela', 'Punta Vicente Roca, Isabela', 'Caleta Tagus, Isabela',
    #              'South Plaza, Santa Cruz', 'Sullivan Bay, Santiago', 'Las Bachas, (Santa Cruz)',
    #              'Puerto Baquerizo Moreno, San Cristobal', 'Punta Pitt (San Cristobal)', 'Cerro Brujo, San Cristobal',
    #              'Puerto Villamil, Isabella', 'Wall Of Tears (Isabella)', 'Espumilla Beach, Santiago',
    #              'Black Turtle Cove', 'El Barranco (Genovesa)', 'Chinese Hat Islet', 'Kicker Rock (San Cristobal)',
    #              'Los Lobos (San Cristobal)']
    # samerica = ['Valparaiso, Chile', 'Arica, Chile', 'Lima, Peru', 'Manta, Ecuador', 'Puerto Montt, Chile',
    #             'Chilean Fjords, South America', 'Strait Of Magellan, South America', 'Punta Arenas, Chile',
    #             'Ushuaia, Argentina', 'Cape Horn, Chile', 'Puerto Madryn, Argentina', 'Punta Del Este, Uruguay',
    #             'Montevideo, Uruguay', 'Buenos Aires, Argentina', 'Pisco, San Martin, Peru',
    #             'Sao Paulo (Santos), Brazil', 'Ilhabela, Brazil', 'Buzios, Brazil', 'Rio De Janeiro, Brazil',
    #             'Gerlache Straight, South America', 'Port Stanley, Falkland Islands', 'Puerto Quetzal, Guatemala',
    #             'Puntarenas, Costa Rica']
    # mexico = ['Ensenada, Mexico', 'Cabo San Lucas, Mexico', 'Puerto Vallarta, Mexico']
    # hawaii = ['Kailua Kona, Hawaii', 'Hilo, Hawaii', 'Lahaina (Maui), Hawaii']
    # china = ['Hong Kong, China', 'Shanghai (Baoshan), China', 'Tianjin, China']
    # exotics = ['Abu Dhabi, United Arab Emirates', 'Dubai, United Arab Emirates', 'Muscat, Oman', 'Khasab, Oman',
    #            'Aqaba, Jordan', 'Suez Canal (Passage)', 'Nha Trang, Vietnam', 'Hue / Danang (Chan May), Vietnam',
    #            'Manila, Philippines', 'Taipei (Keelung), Taiwan', 'Busan, South Korea', 'Tokyo (Yokohama), Japan',
    #            'Hakodate, Japan', 'Sapporo (Muroran), Japan', 'New Mangalore, India', 'Goa (Mormugao), India',
    #            'Bombay (Mumbai), India', 'Otaru, Japan', 'Cochin, India', 'Colombo, Sri Lanka',
    #            'Hanoi (Halong Bay), Vietnam', 'Aomori, Japan', 'Nagasaki, Japan', 'Jeju Island, South Korea',
    #            'Seoul (Incheon), South Korea', 'Mt Fuji (Shimizu), Japan', 'Kobe, Japan', 'Kochi, Japan',
    #            'Kaohsiung, Taiwan', 'Hualien, Taiwan', 'Boracay, Philippines', 'Kota Kinabalu, Malaysia',
    #            'Okinawa, Japan', 'Hiroshima, Japan', 'Kagoshima, Japan']
    # initd = ['International Date Line', 'Lautoka, Fiji']
    # austr = ['Dunedin, New Zealand', 'Dusky Sound, New Zealand', 'Doubtful Sound, New Zealand',
    #          'Milford Sound, New Zealand', 'Newcastle, Australia', 'Cairns, Australia', 'Isle Of Pines, New Caledonia',
    #          'Mystery Island, Vanuatu', 'Lifou, Loyalty Island', 'Lifou, Loyalty Island', 'Noumea, New Caledonia',
    #          'Sydney, Australia']
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
    # can_new_eng = ['Boston, Massachusetts', 'New York, New York', 'Newport, Rhode Island', 'Bar Harbor, Maine']
    # east_carib = ["St. John's, Antigua", 'Tortola, B.V.I', 'Bridgetown, Barbados', 'Roseau, Dominica', 'Punta Cana, Dominican Rep', "St. George's, Grenada", 'Labadee, Hispaniola',
    #               'San Juan, Puerto Rico',
    #               'Basseterre, St. Kitts', 'Castries, St. Lucia', 'Philipsburg, St. Maarten', 'Charlotte Amalie, St. Thomas', 'Kingstown, St. Vincent',
    #               'St. Croix, U.S.V.I.',
    #               'Fort De France']
    # west_carib = ['Belize City, Belize', 'Puerto Limon, Costa Rica', 'George Town, Grand Cayman', 'Roatan, Honduras', 'Falmouth, Jamaica', 'Costa Maya, Mexico',
    #               'Cozumel, Mexico']
    # carib = ['Oranjestad, Aruba', 'Kralendijk, Bonaire', 'Cartagena', 'Willemstad, Curacao', 'Fort Lauderdale, Florida',
    #          'Key West, Florida',
    #          'Miami, Florida', 'New Orleans, Louisiana']
    west_med = ['Catania,Sicily,Italy', 'Ajaccio, Corsica', 'Alicante, Spain', 'Barcelona, Spain', 'Bilbao, Spain',
                'Cadiz, Spain', 'Cannes, France', 'Cartagena, Spain', 'Florence / Pisa (Livorno),Italy',
                'Fuerteventura, Canary', 'Funchal (Madeira), Portugal', 'Genoa, Italy', 'Gibraltar, United Kingdom',
                'Ibiza, Spain', 'La Coruna, Spain', 'La Spezia, Italy', 'Lanzarote, Canary Islands',
                'Las Palmas, Gran Canaria', 'Lisbon, Portugal', 'Malaga, Spain', 'Marseille, France',
                'Messina (Sicily), Italy', 'Montecarlo, Monaco', 'Naples, Italy', 'Nice (Villefranche)',
                'Palma De Mallorca, Spain', 'Ponta Delgada, Azores', 'Portofino, Italy', 'Provence (Toulon), France',
                'Ravenna, Italy', 'Sete, France', 'St. Peter Port, Channel Isl', 'Tenerife, Canary Islands',
                'Valencia, Spain', 'Valletta, Malta', 'Venice, Italy', 'Vigo, Spain']
    # nn = ['Boston, Massachusetts', 'New York, New York', 'Newport, Rhode Island', 'Bar Harbor, Maine']
    europe = ['Rome (Civitavecchia), Italy', 'Le Havre (Paris), France', 'Akureyri, Iceland',
              'Belfast, Northern Ireland', 'Cherbourg, France', 'Cork (Cobh), Ireland', 'Dover, England',
              'Dublin, Ireland', 'Edinburgh, Scotland', 'Greenock (Glasgow), Scotland', 'Inverness/Loch Ness, Scotland',
              'Lerwick/Shetland, Scotland', 'Liverpool, England',
              'Waterford (Dunmore E.), Ireland']
    # bahamas = ['Cococay, Bahamas', 'Nassau, Bahamas', 'Grand Bahama Island']

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
    elif dc == 'FAR.E':
        return ['O', 'Exotics']
    elif dc == 'ALCAN':
        return ['A', 'Alaska']
    elif dc == 'PACIF':
        return ['A', 'Alaska']
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
    isbm = False
    isec = False
    iswc = False
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
    if isbm:
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


while counter > 0:
    cruises = make_request(start_row)
    start_row += 10
    counter -= 1
    parse_data(cruises)

write_file_to_excell(all_cruises)
f = open("ports.txt", 'w')
for row in list(unique):
    f.write(row + '\n')
f.close()
