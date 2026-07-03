"""
California Superior Court locations for all 58 counties.
Used to auto-populate court header fields on small claims forms.

Sources:
  - serve-now.com/resources/courts/california
  - Individual county court websites (lacourt.ca.gov, sdcourt.ca.gov,
    riverside.courts.ca.gov, sanbernardino.courts.ca.gov,
    ventura.courts.ca.gov, fresno.courts.ca.gov, santaclara.courts.ca.gov,
    saccourt.ca.gov, alameda.courts.ca.gov)
  - courts.ca.gov / selfhelp.courts.ca.gov

Notes:
  - For large counties, "courthouses" lists the locations that accept
    small claims filings. Always verify with the county court website.
"""

CALIFORNIA_COURTS: dict = {
    "Alameda": {
        "court_name": "Superior Court of California, County of Alameda",
        "courthouses": [
            # Alameda court accepts small claims at Oakland and Hayward only.
            {"name": "Rene C. Davidson Courthouse", "address": "1225 Fallon Street", "city": "Oakland", "zip": "94612"},
            {"name": "Hayward Hall of Justice", "address": "24405 Amador Street", "city": "Hayward", "zip": "94544"},
            {"name": "Fremont Hall of Justice", "address": "39439 Paseo Padre Parkway", "city": "Fremont", "zip": "94538"},
            {"name": "Wiley W. Manuel Courthouse", "address": "661 Washington Street", "city": "Oakland", "zip": "94607"},
            {"name": "Gale-Schenone Hall of Justice", "address": "5672 Stoneridge Drive", "city": "Pleasanton", "zip": "94588"},
        ],
    },
    "Alpine": {
        "court_name": "Superior Court of California, County of Alpine",
        "courthouses": [
            {"name": "Alpine County Courthouse", "address": "14777 State Route 89", "city": "Markleeville", "zip": "96120"},
        ],
    },
    "Amador": {
        "court_name": "Superior Court of California, County of Amador",
        "courthouses": [
            {"name": "Amador County Superior Court", "address": "108 Court Street", "city": "Jackson", "zip": "95642"},
        ],
    },
    "Butte": {
        "court_name": "Superior Court of California, County of Butte",
        "courthouses": [
            {"name": "Chico Courthouse", "address": "655 Oleander Avenue", "city": "Chico", "zip": "95926"},
            {"name": "Oroville Courthouse", "address": "1775 Concord Avenue", "city": "Oroville", "zip": "95965"},
            {"name": "Gridley Courthouse", "address": "239 Sycamore Street", "city": "Gridley", "zip": "95948"},
        ],
    },
    "Calaveras": {
        "court_name": "Superior Court of California, County of Calaveras",
        "courthouses": [
            {"name": "Calaveras County Courthouse", "address": "891 Mountain Ranch Road", "city": "San Andreas", "zip": "95249"},
        ],
    },
    "Colusa": {
        "court_name": "Superior Court of California, County of Colusa",
        "courthouses": [
            {"name": "Colusa County Courthouse", "address": "547 Market Street", "city": "Colusa", "zip": "95932"},
        ],
    },
    "Contra Costa": {
        "court_name": "Superior Court of California, County of Contra Costa",
        "courthouses": [
            {"name": "A. F. Bray Courthouse (Martinez)", "address": "1020 Ward Street", "city": "Martinez", "zip": "94553"},
            {"name": "Wakefield Taylor Courthouse", "address": "725 Court Street", "city": "Martinez", "zip": "94553"},
            {"name": "Concord Courthouse", "address": "2970 Willow Pass Road", "city": "Concord", "zip": "94519"},
            {"name": "Pittsburg Courthouse", "address": "45 Civic Avenue", "city": "Pittsburg", "zip": "94565"},
            {"name": "Richmond Courthouse", "address": "100 37th Street", "city": "Richmond", "zip": "94805"},
            {"name": "Walnut Creek Courthouse", "address": "640 Ygnacio Valley Road", "city": "Walnut Creek", "zip": "94596"},
        ],
    },
    "Del Norte": {
        "court_name": "Superior Court of California, County of Del Norte",
        "courthouses": [
            {"name": "Del Norte County Courthouse", "address": "450 H Street, Room 209", "city": "Crescent City", "zip": "95531"},
        ],
    },
    "El Dorado": {
        "court_name": "Superior Court of California, County of El Dorado",
        "courthouses": [
            {"name": "El Dorado Courthouse (Placerville)", "address": "495 Main Street", "city": "Placerville", "zip": "95667"},
            {"name": "South Lake Tahoe Courthouse", "address": "1354 Johnson Boulevard, Suite 2", "city": "South Lake Tahoe", "zip": "96150"},
            {"name": "Cameron Park Courthouse", "address": "3321 Cameron Park Drive", "city": "Cameron Park", "zip": "95682"},
        ],
    },
    "Fresno": {
        "court_name": "Superior Court of California, County of Fresno",
        "courthouses": [
            {"name": "B.F. Sisk Courthouse", "address": "1130 O Street", "city": "Fresno", "zip": "93721"},
            {"name": "Fresno Main Courthouse", "address": "1100 Van Ness Avenue", "city": "Fresno", "zip": "93724"},
        ],
    },
    "Glenn": {
        "court_name": "Superior Court of California, County of Glenn",
        "courthouses": [
            {"name": "Willows Courthouse", "address": "526 West Sycamore Street", "city": "Willows", "zip": "95988"},
            {"name": "Orland Courthouse", "address": "821 East South Street", "city": "Orland", "zip": "95963"},
        ],
    },
    "Humboldt": {
        "court_name": "Superior Court of California, County of Humboldt",
        "courthouses": [
            {"name": "Humboldt County Courthouse", "address": "825 Fifth Street", "city": "Eureka", "zip": "95501"},
        ],
    },
    "Imperial": {
        "court_name": "Superior Court of California, County of Imperial",
        "courthouses": [
            {"name": "El Centro Courthouse", "address": "939 Main Street", "city": "El Centro", "zip": "92243"},
            {"name": "Brawley Courthouse", "address": "220 Main Street", "city": "Brawley", "zip": "92227"},
            {"name": "Calexico Courthouse", "address": "415 East 4th Street", "city": "Calexico", "zip": "92231"},
        ],
    },
    "Inyo": {
        "court_name": "Superior Court of California, County of Inyo",
        "courthouses": [
            {"name": "Independence Courthouse", "address": "168 North Edwards Street", "city": "Independence", "zip": "93526"},
            {"name": "Bishop Courthouse", "address": "301 West Line Street", "city": "Bishop", "zip": "93514"},
        ],
    },
    "Kern": {
        "court_name": "Superior Court of California, County of Kern",
        "courthouses": [
            {"name": "Bakersfield Courthouse", "address": "1415 Truxtun Avenue", "city": "Bakersfield", "zip": "93301"},
            {"name": "Metro Division Courthouse", "address": "1215 Truxtun Avenue", "city": "Bakersfield", "zip": "93301"},
            {"name": "Delano Courthouse", "address": "1122 Jefferson Street", "city": "Delano", "zip": "93215"},
            {"name": "Mojave Courthouse", "address": "1773 Highway 58", "city": "Mojave", "zip": "93501"},
            {"name": "Ridgecrest Courthouse", "address": "132 East Coso Street", "city": "Ridgecrest", "zip": "93555"},
            {"name": "Taft Courthouse", "address": "311 Lincoln Street", "city": "Taft", "zip": "93268"},
        ],
    },
    "Kings": {
        "court_name": "Superior Court of California, County of Kings",
        "courthouses": [
            {"name": "Hanford Courthouse", "address": "1426 South Drive", "city": "Hanford", "zip": "93230"},
            {"name": "Corcoran Courthouse", "address": "1000 Chittenden Avenue", "city": "Corcoran", "zip": "93212"},
            {"name": "Avenal Courthouse", "address": "501 East Kings Street", "city": "Avenal", "zip": "93204"},
        ],
    },
    "Lake": {
        "court_name": "Superior Court of California, County of Lake",
        "courthouses": [
            {"name": "Lakeport Courthouse", "address": "255 North Forbes Street", "city": "Lakeport", "zip": "95453"},
            {"name": "Clearlake Courthouse", "address": "7000A South Center Drive", "city": "Clearlake", "zip": "95422"},
        ],
    },
    "Lassen": {
        "court_name": "Superior Court of California, County of Lassen",
        "courthouses": [
            {"name": "Lassen County Courthouse", "address": "220 South Lassen Street", "city": "Susanville", "zip": "96130"},
        ],
    },
    "Los Angeles": {
        "court_name": "Superior Court of California, County of Los Angeles",
        # Small claims is heard at the following hub courthouses (verified via lacourt.ca.gov):
        "courthouses": [
            {"name": "Stanley Mosk Courthouse", "address": "111 North Hill Street", "city": "Los Angeles", "zip": "90012"},
            {"name": "Bellflower Courthouse", "address": "10025 East Flower Street", "city": "Bellflower", "zip": "90706"},
            {"name": "Beverly Hills Courthouse", "address": "9355 Burton Way", "city": "Beverly Hills", "zip": "90210"},
            {"name": "Chatsworth Courthouse", "address": "9425 Penfield Avenue", "city": "Chatsworth", "zip": "91311"},
            {"name": "Compton Courthouse", "address": "200 West Compton Boulevard", "city": "Compton", "zip": "90220"},
            {"name": "Governor George Deukmejian Courthouse (Long Beach)", "address": "275 Magnolia Avenue", "city": "Long Beach", "zip": "90802"},
            {"name": "Inglewood Courthouse", "address": "One Regent Street", "city": "Inglewood", "zip": "90301"},
            {"name": "Michael Antonovich Antelope Valley Courthouse", "address": "42011 4th Street West", "city": "Lancaster", "zip": "93534"},
            {"name": "Pasadena Courthouse", "address": "300 East Walnut Street", "city": "Pasadena", "zip": "91101"},
            {"name": "Santa Monica Courthouse", "address": "1725 Main Street", "city": "Santa Monica", "zip": "90401"},
            {"name": "Van Nuys Courthouse East", "address": "6230 Sylmar Avenue", "city": "Van Nuys", "zip": "91401"},
            {"name": "West Covina Courthouse", "address": "1427 West Covina Parkway", "city": "West Covina", "zip": "91790"},
            {"name": "Torrance Courthouse", "address": "825 Maple Avenue", "city": "Torrance", "zip": "90503"},
            {"name": "El Monte Courthouse", "address": "11234 East Valley Boulevard", "city": "El Monte", "zip": "91731"},
            {"name": "Glendale Courthouse", "address": "600 East Broadway", "city": "Glendale", "zip": "91206"},
            {"name": "Burbank Courthouse", "address": "300 East Olive Avenue", "city": "Burbank", "zip": "91502"},
            {"name": "San Fernando Courthouse", "address": "900 Third Street", "city": "San Fernando", "zip": "91340"},
            {"name": "Pomona Courthouse South", "address": "400 Civic Center Plaza", "city": "Pomona", "zip": "91766"},
            {"name": "Norwalk Courthouse", "address": "12720 Norwalk Boulevard", "city": "Norwalk", "zip": "90650"},
            {"name": "Whittier Courthouse", "address": "7339 South Painter Avenue", "city": "Whittier", "zip": "90602"},
            {"name": "Santa Clarita Courthouse", "address": "23747 West Valencia Boulevard", "city": "Santa Clarita", "zip": "91355"},
            {"name": "Malibu Courthouse", "address": "23525 Civic Center Way", "city": "Malibu", "zip": "90265"},
            {"name": "Alhambra Courthouse", "address": "150 West Commonwealth Avenue", "city": "Alhambra", "zip": "91801"},
            {"name": "Redondo Beach Courthouse", "address": "117 West Torrance Boulevard", "city": "Redondo Beach", "zip": "90277"},
            {"name": "Sylmar Courthouse", "address": "16350 Filbert Street", "city": "Sylmar", "zip": "91342"},
        ],
    },
    "Madera": {
        "court_name": "Superior Court of California, County of Madera",
        "courthouses": [
            {"name": "Madera Courthouse", "address": "209 West Yosemite Avenue", "city": "Madera", "zip": "93637"},
            {"name": "Chowchilla Courthouse", "address": "141 South 2nd Street", "city": "Chowchilla", "zip": "93610"},
        ],
    },
    "Marin": {
        "court_name": "Superior Court of California, County of Marin",
        "courthouses": [
            {"name": "Marin County Civic Center Courthouse", "address": "3501 Civic Center Drive", "city": "San Rafael", "zip": "94903"},
        ],
    },
    "Mariposa": {
        "court_name": "Superior Court of California, County of Mariposa",
        "courthouses": [
            {"name": "Mariposa County Courthouse", "address": "5088 Bullion Street", "city": "Mariposa", "zip": "95338"},
        ],
    },
    "Mendocino": {
        "court_name": "Superior Court of California, County of Mendocino",
        "courthouses": [
            {"name": "Ukiah Courthouse", "address": "100 North State Street", "city": "Ukiah", "zip": "95482"},
            {"name": "Fort Bragg Courthouse", "address": "700 South Franklin Street", "city": "Fort Bragg", "zip": "95437"},
            {"name": "Willits Courthouse", "address": "125 East Commercial Street", "city": "Willits", "zip": "95490"},
            {"name": "Point Arena Courthouse", "address": "24000 South Highway One", "city": "Point Arena", "zip": "95468"},
        ],
    },
    "Merced": {
        "court_name": "Superior Court of California, County of Merced",
        "courthouses": [
            {"name": "Merced Courthouse", "address": "627 West 21st Street", "city": "Merced", "zip": "95340"},
        ],
    },
    "Modoc": {
        "court_name": "Superior Court of California, County of Modoc",
        "courthouses": [
            {"name": "Modoc County Courthouse", "address": "205 South East Street", "city": "Alturas", "zip": "96101"},
        ],
    },
    "Mono": {
        "court_name": "Superior Court of California, County of Mono",
        "courthouses": [
            {"name": "Bridgeport Courthouse", "address": "State Highway 395 North", "city": "Bridgeport", "zip": "93517"},
            {"name": "Mammoth Lakes Courthouse", "address": "452 Old Mammoth Road", "city": "Mammoth Lakes", "zip": "93546"},
        ],
    },
    "Monterey": {
        "court_name": "Superior Court of California, County of Monterey",
        "courthouses": [
            {"name": "Salinas Courthouse", "address": "1200 Aguajito Road", "city": "Monterey", "zip": "93940"},
            {"name": "Salinas (Main Courthouse)", "address": "240 Church Street", "city": "Salinas", "zip": "93901"},
            {"name": "King City Courthouse", "address": "250 Franciscan Way", "city": "King City", "zip": "93930"},
            {"name": "Marina Courthouse", "address": "3180 Del Monte Boulevard", "city": "Marina", "zip": "93933"},
        ],
    },
    "Napa": {
        "court_name": "Superior Court of California, County of Napa",
        "courthouses": [
            {"name": "Napa County Courthouse", "address": "1111 Third Street", "city": "Napa", "zip": "94559"},
        ],
    },
    "Nevada": {
        "court_name": "Superior Court of California, County of Nevada",
        "courthouses": [
            {"name": "Nevada City Courthouse", "address": "201 Church Street", "city": "Nevada City", "zip": "95959"},
            {"name": "Truckee Courthouse", "address": "10075 Levon Avenue", "city": "Truckee", "zip": "96161"},
        ],
    },
    "Orange": {
        "court_name": "Superior Court of California, County of Orange",
        "courthouses": [
            {"name": "Central Justice Center", "address": "700 Civic Center Drive West", "city": "Santa Ana", "zip": "92701"},
            {"name": "Civil Complex Center", "address": "751 West Santa Ana Boulevard", "city": "Santa Ana", "zip": "92701"},
            {"name": "Harbor Justice Center", "address": "4601 Jamboree Road", "city": "Newport Beach", "zip": "92660"},
            {"name": "Lamoreaux Justice Center", "address": "341 The City Drive South", "city": "Orange", "zip": "92868"},
            {"name": "North Justice Center", "address": "1275 North Berkeley Avenue", "city": "Fullerton", "zip": "92832"},
            {"name": "Stephen K. Tamura West Justice Center", "address": "8141 13th Street", "city": "Westminster", "zip": "92683"},
            {"name": "Costa Mesa Justice Complex", "address": "3390 Harbor Boulevard", "city": "Costa Mesa", "zip": "92626"},
        ],
    },
    "Placer": {
        "court_name": "Superior Court of California, County of Placer",
        "courthouses": [
            {"name": "Auburn Courthouse", "address": "101 Maple Street", "city": "Auburn", "zip": "95603"},
            {"name": "Roseville Courthouse", "address": "300 Taylor Street", "city": "Roseville", "zip": "95678"},
            {"name": "South Placer (Rocklin) Courthouse", "address": "1000 Sunset Boulevard", "city": "Rocklin", "zip": "95765"},
            {"name": "Tahoe City Courthouse", "address": "2501 North Lake Boulevard", "city": "Tahoe City", "zip": "96145"},
        ],
    },
    "Plumas": {
        "court_name": "Superior Court of California, County of Plumas",
        "courthouses": [
            {"name": "Quincy Courthouse", "address": "520 Main Street, Room 104", "city": "Quincy", "zip": "95971"},
            {"name": "Portola Courthouse", "address": "161 Nevada Street", "city": "Portola", "zip": "96122"},
        ],
    },
    "Riverside": {
        "court_name": "Superior Court of California, County of Riverside",
        "courthouses": [
            {"name": "Riverside Historic Courthouse", "address": "4050 Main Street", "city": "Riverside", "zip": "92501"},
            {"name": "Riverside Hall of Justice", "address": "4100 Main Street", "city": "Riverside", "zip": "92501"},
            {"name": "Robert Presley Hall of Justice", "address": "4175 Main Street", "city": "Riverside", "zip": "92501"},
            {"name": "Banning Justice Center", "address": "311 East Ramsey Street", "city": "Banning", "zip": "92220"},
            {"name": "Blythe Courthouse", "address": "265 North Broadway", "city": "Blythe", "zip": "92225"},
            {"name": "Corona Courthouse", "address": "505 South Buena Vista Avenue, Room 201", "city": "Corona", "zip": "92882"},
            {"name": "Hemet Courthouse", "address": "880 North State Street", "city": "Hemet", "zip": "92543"},
            {"name": "Indio Courthouse", "address": "46-200 Oasis Street", "city": "Indio", "zip": "92201"},
            {"name": "Moreno Valley Courthouse", "address": "13800 Heacock Street, Building D, Suite 201", "city": "Moreno Valley", "zip": "92553"},
            {"name": "Southwest Justice Center (Murrieta)", "address": "30755-D Auld Road", "city": "Murrieta", "zip": "92563"},
            {"name": "Temecula Courthouse", "address": "41002 County Center Drive, Suite 100", "city": "Temecula", "zip": "92591"},
        ],
    },
    "Sacramento": {
        "court_name": "Superior Court of California, County of Sacramento",
        "courthouses": [
            # Small claims is filed at Carol Miller Justice Center:
            {"name": "Carol Miller Justice Center", "address": "301 Bicentennial Circle", "city": "Sacramento", "zip": "95826"},
            {"name": "Gordon D. Schaber Courthouse", "address": "720 Ninth Street", "city": "Sacramento", "zip": "95814"},
            {"name": "Tani G. Cantil-Sakauye Courthouse", "address": "500 G Street", "city": "Sacramento", "zip": "95814"},
            {"name": "Lorenzo Patiño Hall of Justice", "address": "651 I Street", "city": "Sacramento", "zip": "95814"},
            {"name": "William R. Ridgeway Family Relations Courthouse", "address": "3341 Power Inn Road", "city": "Sacramento", "zip": "95826"},
        ],
    },
    "San Benito": {
        "court_name": "Superior Court of California, County of San Benito",
        "courthouses": [
            {"name": "San Benito County Courthouse", "address": "440 Fifth Street", "city": "Hollister", "zip": "95023"},
        ],
    },
    "San Bernardino": {
        "court_name": "Superior Court of California, County of San Bernardino",
        "courthouses": [
            {"name": "San Bernardino District Courthouse", "address": "247 West Third Street", "city": "San Bernardino", "zip": "92415"},
            {"name": "Rancho Cucamonga Courthouse", "address": "8303 Haven Avenue", "city": "Rancho Cucamonga", "zip": "91730"},
            {"name": "Fontana Courthouse", "address": "17780 Arrow Boulevard", "city": "Fontana", "zip": "92335"},
            {"name": "Victorville Courthouse", "address": "14455 Civic Drive, Suite 200", "city": "Victorville", "zip": "92392"},
            {"name": "Barstow Courthouse", "address": "235 East Mountain View Street", "city": "Barstow", "zip": "92311"},
            {"name": "Big Bear Courthouse", "address": "477 Summit Boulevard", "city": "Big Bear Lake", "zip": "92315"},
            {"name": "Joshua Tree Courthouse", "address": "6527 White Feather Road", "city": "Joshua Tree", "zip": "92252"},
            {"name": "Needles Courthouse", "address": "1111 Bailey Avenue", "city": "Needles", "zip": "92363"},
        ],
    },
    "San Diego": {
        "court_name": "Superior Court of California, County of San Diego",
        "courthouses": [
            {"name": "Hall of Justice", "address": "330 West Broadway", "city": "San Diego", "zip": "92101"},
            {"name": "Central Courthouse", "address": "1100 Union Street", "city": "San Diego", "zip": "92101"},
            {"name": "El Cajon Courthouse", "address": "250 East Main Street", "city": "El Cajon", "zip": "92020"},
            {"name": "Vista Courthouse", "address": "325 South Melrose Drive", "city": "Vista", "zip": "92081"},
            {"name": "South County Regional Center (Chula Vista)", "address": "500 Third Avenue", "city": "Chula Vista", "zip": "91910"},
            {"name": "Kearny Mesa Courthouse", "address": "8950 Clairemont Mesa Boulevard", "city": "San Diego", "zip": "92123"},
            {"name": "Ramona Courthouse", "address": "1428 Montecito Road", "city": "Ramona", "zip": "92065"},
        ],
    },
    "San Francisco": {
        "court_name": "Superior Court of California, County of San Francisco",
        "courthouses": [
            {"name": "Civic Center Courthouse", "address": "400 McAllister Street", "city": "San Francisco", "zip": "94102"},
            {"name": "Hall of Justice", "address": "850 Bryant Street", "city": "San Francisco", "zip": "94103"},
        ],
    },
    "San Joaquin": {
        "court_name": "Superior Court of California, County of San Joaquin",
        "courthouses": [
            {"name": "Stockton Courthouse", "address": "222 East Weber Avenue", "city": "Stockton", "zip": "95202"},
            {"name": "Lodi Courthouse", "address": "315 West Elm Street", "city": "Lodi", "zip": "95240"},
            {"name": "Manteca Courthouse", "address": "315 East Center Street", "city": "Manteca", "zip": "95336"},
            {"name": "Tracy Courthouse", "address": "475 East 10th Street", "city": "Tracy", "zip": "95376"},
        ],
    },
    "San Luis Obispo": {
        "court_name": "Superior Court of California, County of San Luis Obispo",
        "courthouses": [
            {"name": "San Luis Obispo Courthouse", "address": "1050 Monterey Street", "city": "San Luis Obispo", "zip": "93408"},
            {"name": "Paso Robles Courthouse", "address": "901 Park Street", "city": "Paso Robles", "zip": "93446"},
            {"name": "Grover Beach Courthouse", "address": "214 South 16th Street", "city": "Grover Beach", "zip": "93433"},
        ],
    },
    "San Mateo": {
        "court_name": "Superior Court of California, County of San Mateo",
        "courthouses": [
            {"name": "Hall of Justice and Records", "address": "400 County Center", "city": "Redwood City", "zip": "94063"},
            {"name": "Southern Branch Courthouse", "address": "1050 Mission Road", "city": "South San Francisco", "zip": "94080"},
        ],
    },
    "Santa Barbara": {
        "court_name": "Superior Court of California, County of Santa Barbara",
        "courthouses": [
            {"name": "Santa Barbara Courthouse", "address": "1100 Anacapa Street", "city": "Santa Barbara", "zip": "93101"},
            {"name": "Santa Maria Courthouse", "address": "312-C East Cook Street", "city": "Santa Maria", "zip": "93454"},
            {"name": "Lompoc Courthouse", "address": "115 Civic Center Plaza", "city": "Lompoc", "zip": "93436"},
            {"name": "Solvang Courthouse", "address": "1745 Mission Drive, Suite C", "city": "Solvang", "zip": "93463"},
        ],
    },
    "Santa Clara": {
        "court_name": "Superior Court of California, County of Santa Clara",
        "courthouses": [
            # Small claims filed at Downtown Superior Court:
            {"name": "Downtown Superior Court", "address": "191 North First Street", "city": "San Jose", "zip": "95113"},
            {"name": "Hall of Justice", "address": "190 West Hedding Street", "city": "San Jose", "zip": "95110"},
            {"name": "Family Justice Center Courthouse", "address": "201 North First Street", "city": "San Jose", "zip": "95113"},
            {"name": "Palo Alto Courthouse", "address": "270 Grant Avenue", "city": "Palo Alto", "zip": "94306"},
            {"name": "South County (Morgan Hill) Courthouse", "address": "301 Diana Avenue", "city": "Morgan Hill", "zip": "95037"},
        ],
    },
    "Santa Cruz": {
        "court_name": "Superior Court of California, County of Santa Cruz",
        "courthouses": [
            {"name": "Santa Cruz Courthouse", "address": "701 Ocean Street", "city": "Santa Cruz", "zip": "95060"},
            {"name": "Watsonville Courthouse", "address": "1430 Freedom Boulevard", "city": "Watsonville", "zip": "95076"},
        ],
    },
    "Shasta": {
        "court_name": "Superior Court of California, County of Shasta",
        "courthouses": [
            {"name": "Redding Courthouse", "address": "1500 Court Street", "city": "Redding", "zip": "96001"},
            {"name": "Burney Courthouse", "address": "20509 Shasta Street", "city": "Burney", "zip": "96013"},
        ],
    },
    "Sierra": {
        "court_name": "Superior Court of California, County of Sierra",
        "courthouses": [
            {"name": "Downieville Courthouse", "address": "100 Courthouse Square", "city": "Downieville", "zip": "95936"},
        ],
    },
    "Siskiyou": {
        "court_name": "Superior Court of California, County of Siskiyou",
        "courthouses": [
            {"name": "Yreka Courthouse", "address": "311 4th Street", "city": "Yreka", "zip": "96097"},
            {"name": "Weed Courthouse", "address": "550 Main Street", "city": "Weed", "zip": "96094"},
            {"name": "Happy Camp Courthouse", "address": "28 Fourth Avenue", "city": "Happy Camp", "zip": "96039"},
        ],
    },
    "Solano": {
        "court_name": "Superior Court of California, County of Solano",
        "courthouses": [
            {"name": "Fairfield Courthouse", "address": "600 Union Avenue", "city": "Fairfield", "zip": "94533"},
            {"name": "Vallejo Courthouse", "address": "321 Tuolumne Street", "city": "Vallejo", "zip": "94590"},
        ],
    },
    "Sonoma": {
        "court_name": "Superior Court of California, County of Sonoma",
        "courthouses": [
            {"name": "Hall of Justice", "address": "600 Administration Drive", "city": "Santa Rosa", "zip": "95403"},
            {"name": "Civil and Family Law Courthouse", "address": "3055 Cleveland Avenue", "city": "Santa Rosa", "zip": "95403"},
        ],
    },
    "Stanislaus": {
        "court_name": "Superior Court of California, County of Stanislaus",
        "courthouses": [
            {"name": "Modesto Courthouse", "address": "800 11th Street", "city": "Modesto", "zip": "95354"},
            {"name": "Ceres Courthouse", "address": "2744 2nd Street", "city": "Ceres", "zip": "95307"},
            {"name": "Turlock Courthouse", "address": "300 Starr Avenue", "city": "Turlock", "zip": "95380"},
        ],
    },
    "Sutter": {
        "court_name": "Superior Court of California, County of Sutter",
        "courthouses": [
            {"name": "Sutter County Courthouse", "address": "463 Second Street", "city": "Yuba City", "zip": "95991"},
        ],
    },
    "Tehama": {
        "court_name": "Superior Court of California, County of Tehama",
        "courthouses": [
            {"name": "Red Bluff Courthouse", "address": "633 Washington Street", "city": "Red Bluff", "zip": "96080"},
            {"name": "Corning Courthouse", "address": "720 Hoag Street", "city": "Corning", "zip": "96021"},
        ],
    },
    "Trinity": {
        "court_name": "Superior Court of California, County of Trinity",
        "courthouses": [
            {"name": "Weaverville Courthouse", "address": "11 Court Street", "city": "Weaverville", "zip": "96093"},
            {"name": "Hayfork Courthouse", "address": "6641B State Highway 3", "city": "Hayfork", "zip": "96041"},
        ],
    },
    "Tulare": {
        "court_name": "Superior Court of California, County of Tulare",
        "courthouses": [
            {"name": "Visalia Courthouse", "address": "221 South Mooney Boulevard", "city": "Visalia", "zip": "93291"},
            {"name": "Porterville Courthouse", "address": "87 East Morton Avenue", "city": "Porterville", "zip": "93257"},
            {"name": "Tulare Courthouse", "address": "425 East Kern Avenue", "city": "Tulare", "zip": "93274"},
            {"name": "Dinuba Courthouse", "address": "640 South Alta Avenue", "city": "Dinuba", "zip": "93618"},
        ],
    },
    "Tuolumne": {
        "court_name": "Superior Court of California, County of Tuolumne",
        "courthouses": [
            {"name": "Sonora Courthouse", "address": "60 North Washington Street", "city": "Sonora", "zip": "95370"},
        ],
    },
    "Ventura": {
        "court_name": "Superior Court of California, County of Ventura",
        # All small claims filings go to the Hall of Justice in Ventura.
        "courthouses": [
            {"name": "Hall of Justice", "address": "800 South Victoria Avenue", "city": "Ventura", "zip": "93009"},
            {"name": "East County Courthouse (Simi Valley)", "address": "3855-F Alamo Street", "city": "Simi Valley", "zip": "93063"},
            {"name": "Juvenile Justice Center (Oxnard)", "address": "4353 East Vineyard Avenue", "city": "Oxnard", "zip": "93036"},
        ],
    },
    "Yolo": {
        "court_name": "Superior Court of California, County of Yolo",
        "courthouses": [
            {"name": "Yolo County Courthouse", "address": "725 Court Street", "city": "Woodland", "zip": "95695"},
        ],
    },
    "Yuba": {
        "court_name": "Superior Court of California, County of Yuba",
        "courthouses": [
            {"name": "Yuba County Courthouse", "address": "215 Fifth Street, Suite 200", "city": "Marysville", "zip": "95901"},
        ],
    },
}


def court_info_string(court: dict) -> str:
    """Build the multi-line court header string from a court dict."""
    county  = court.get("county",  "Alameda")
    address = court.get("address", "1225 Fallon Street")
    city    = court.get("city",    "Oakland")
    zip_    = court.get("zip",     "94612")
    return (
        f"Superior Court of California, County of {county}\n"
        f"{address}\n"
        f"{city}, CA {zip_}"
    )


ALL_COUNTIES: list[str] = sorted(CALIFORNIA_COURTS.keys())


def courthouses_for_county(county: str) -> list[dict]:
    """Return list of courthouse dicts for a county, or [] if unknown."""
    return CALIFORNIA_COURTS.get(county, {}).get("courthouses", [])
