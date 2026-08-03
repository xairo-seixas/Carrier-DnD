"""Country -> major container port(s) reference, for fanning out a
country-level tariff (no named port in the source PDF) across the ports
that actually matter for that country.

STATUS: best-effort, NOT independently verified against the authoritative
UN/LOCODE registry. Cross-checking one candidate offline source against the
single best-known port in the world (Shanghai) produced a code that doesn't
match the one universally used in the industry (it returned CNSGH; the
standard is CNSHA), which means that source's bundled data can't be trusted.
These codes instead come from general shipping-industry knowledge - the
kind of UNLOCODE that appears on essentially every bill of lading for that
port - and should be treated as a strong starting point, not a verified
ground truth. Confirm against BuyCo's own port master data (or the live
UN/LOCODE registry) before this feeds anything used for invoicing or
compliance.

Coverage is deliberately limited to countries with major container gateways
(roughly the top ~40 trading nations by container volume) - not every
country in the world has an entry. Countries without a real seaport
(landlocked) are intentionally omitted; extend PORTS_BY_COUNTRY as needed.
"""

# country name (as it appears after slug.title() in the scraper's discovery
# functions, e.g. "united-states-of-america" -> "United States Of America")
# -> list of (UNLOCODE, port name)
PORTS_BY_COUNTRY: dict[str, list[tuple[str, str]]] = {
    "China": [
        ("CNSHA", "Shanghai"), ("CNNGB", "Ningbo-Zhoushan"), ("CNSZX", "Shenzhen"),
        ("CNTAO", "Qingdao"), ("CNCAN", "Guangzhou"), ("CNTXG", "Tianjin"),
        ("CNXMN", "Xiamen"), ("CNDLC", "Dalian"), ("CNYTN", "Yantian"),
    ],
    "Singapore": [("SGSIN", "Singapore")],
    "South Korea": [("KRPUS", "Busan")],
    "United Arab Emirates": [("AEJEA", "Jebel Ali"), ("AEAUH", "Abu Dhabi"), ("AEKHF", "Khor Fakkan")],
    "Malaysia": [("MYPKG", "Port Klang"), ("MYTPP", "Tanjung Pelepas")],
    "Netherlands": [("NLRTM", "Rotterdam")],
    "Hong Kong": [("HKHKG", "Hong Kong")],
    "Belgium": [("BEANR", "Antwerp")],
    "United States Of America": [
        ("USLAX", "Los Angeles"), ("USLGB", "Long Beach"), ("USNYC", "New York/New Jersey"),
        ("USSAV", "Savannah"), ("USHOU", "Houston"), ("USOAK", "Oakland"),
        ("USCHS", "Charleston"), ("USSEA", "Seattle"), ("USTIW", "Tacoma"), ("USORF", "Norfolk"),
    ],
    "United States": [  # alt slug spelling, kept in sync with the entry above
        ("USLAX", "Los Angeles"), ("USLGB", "Long Beach"), ("USNYC", "New York/New Jersey"),
        ("USSAV", "Savannah"), ("USHOU", "Houston"), ("USOAK", "Oakland"),
        ("USCHS", "Charleston"), ("USSEA", "Seattle"), ("USTIW", "Tacoma"), ("USORF", "Norfolk"),
    ],
    "Morocco": [("MAPTM", "Tanger-Med")],
    "Thailand": [("THLCH", "Laem Chabang")],
    "Taiwan": [("TWKHH", "Kaohsiung"), ("TWKEL", "Keelung")],
    "Vietnam": [("VNSGN", "Ho Chi Minh City"), ("VNHPH", "Haiphong"), ("VNVUT", "Cai Mep / Vung Tau")],
    "India": [("INMUN", "Mundra"), ("INNSA", "Nhava Sheva / Jawaharlal Nehru")],
    "Indonesia": [("IDJKT", "Tanjung Priok / Jakarta"), ("IDSUB", "Tanjung Perak / Surabaya")],
    "Germany": [("DEHAM", "Hamburg"), ("DEBRV", "Bremerhaven")],
    "Sri Lanka": [("LKCMB", "Colombo")],
    "Brazil": [("BRSSZ", "Santos")],
    "Spain": [("ESVLC", "Valencia"), ("ESALG", "Algeciras"), ("ESBCN", "Barcelona")],
    "Panama": [("PACTB", "Cristobal / Colon"), ("PABLB", "Balboa")],
    "Greece": [("GRPIR", "Piraeus")],
    "Japan": [("JPTYO", "Tokyo"), ("JPYOK", "Yokohama"), ("JPUKB", "Kobe"), ("JPNGO", "Nagoya"), ("JPOSA", "Osaka")],
    "Philippines": [("PHMNL", "Manila")],
    "Egypt": [("EGPSD", "Port Said"), ("EGDAM", "Damietta")],
    "Saudi Arabia": [("SAJED", "Jeddah"), ("SADMM", "Dammam")],
    "Oman": [("OMSLL", "Salalah")],
    "Italy": [("ITGIT", "Gioia Tauro"), ("ITGOA", "Genoa")],
    "United Kingdom": [("GBFXT", "Felixstowe"), ("GBSOU", "Southampton")],
    "France": [("FRLEH", "Le Havre"), ("FRMRS", "Marseille")],
    "Australia": [("AUMEL", "Melbourne"), ("AUSYD", "Sydney")],
    "South Africa": [("ZADUR", "Durban")],
    "Canada": [("CAVAN", "Vancouver"), ("CAMTR", "Montreal")],
    "Mexico": [("MXMZT", "Manzanillo"), ("MXVER", "Veracruz")],
    "Colombia": [("COCTG", "Cartagena")],
    "Chile": [("CLVAP", "Valparaiso"), ("CLSAI", "San Antonio")],
    "Peru": [("PECLL", "Callao")],
    "Argentina": [("ARBUE", "Buenos Aires")],
    "Turkiye": [("TRMER", "Mersin"), ("TRAMB", "Ambarli")],
    "Turkey": [("TRMER", "Mersin"), ("TRAMB", "Ambarli")],
    "Poland": [("PLGDN", "Gdansk")],
    "Sweden": [("SEGOT", "Gothenburg")],
    "Portugal": [("PTLEI", "Sines")],
    "Nigeria": [("NGLOS", "Lagos / Apapa")],
    "Kenya": [("KEMBA", "Mombasa")],
    "Pakistan": [("PKKHI", "Karachi"), ("PKBQM", "Port Qasim")],
    "Bangladesh": [("BDCGP", "Chittagong")],
    "New Zealand": [("NZAKL", "Auckland")],
}


def get_ports_for_country(country: str) -> list[tuple[str, str]]:
    """Returns [] if we have no reference entry for this country - callers
    should treat that as 'needs manual port mapping', not 'no ports exist'."""
    return PORTS_BY_COUNTRY.get(country, [])
