"""Transaction categorisation for the Amex and Barclays exports — pure string
rules over a description, no I/O. The most test-friendly part of the pipeline.

Extracted from spending_summary.py on 2026-07-19 — that file had grown to 4,020
lines. Behaviour is unchanged: the code below is the original, moved verbatim.
"""

# ── AMEX categorisation ────────────────────────────────────────────────────────
def categorise_amex(desc: str) -> str | None:
    d = str(desc).upper()
    if "PAYMENT RECEIVED" in d:
        return None
    if "SEATFROG" in d:
        return "Seatfrog"

    holiday_kw = [
        "ADEJE", "TENERIFE", "CRISTIANOS", "ARMENIME", "HIPERDINO",
        "RIU PALACE", "RED LION 005", "WALI HAII", "LIMONELLA",
        "TEMPLE BAR 367", "HARRY S BAR 047", "ROSSO SUL MARE",
        "SUGAR AND SPICY", "WAXY OSHEAS", "RTE. Y BARES", "CANDELARIA",
        "MURPHYS", "COSTA CALETA", "COSTA ADEJE", "MANGO SIAM", "VIPS SIAM",
        "PIZZERIA ALL ITALIANA", "SIAM MALL", "STRADIVARIUS SIAM",
        "ROCA NEGRA", "LM 176", "LM 179", "LM 95", "SSP-TFS",
        "CAFETERIA EL DUQUE", "ORANGE CAFE", "SUPERMERCADO", "MAQUINAS VENDING",
        "TAXI ADEJE", "TAXI-ADEJE", "TAXI LM", "TAXI L.M.", "TAXI VILLA DE ADEJ",
        "HD EX SAN SEBASTIA", "CASA PLAYA", "CHIRINGUITO EL PUE",
        "MINIMARKET TENERIF", "HERTZ UK", "TUI INFLIGHT", "ADMIRAL TRAVEL",
        "NEWCASTLE AIR", "3CPAYMENT*PREMIER INN 4 LONDON",
        "3CPAYMENT*PREMIER INN 4 SOUTHAMP", "3CPAYMENT*NEWCASTLE AIR",
        "IBIS LONDON", "DIVERSE DINING LTD",
    ]
    if any(kw in d for kw in holiday_kw):
        return "Holidays"

    if "EXPLEO" in d or "UK EXCL" in d:
        return "Salary"
    if "LIAM SANGHA" in d:
        return "Liam Sangha"
    if "HMRC" in d:
        return "Tax"

    if any(kw in d for kw in ["OCTOPUS ENERGY", "NORTHUMBRIAN WATER", "TALKTALK", "TV LICENCE", "PARENTPAY"]):
        return "Utilities"
    if any(kw in d for kw in ["TALKMOBILE", "VODAFONE", "APPLE.COM/BILL"]):
        return "Broadband & Phone"
    if any(kw in d for kw in ["NETFLIX", "NORTON", "CLAUDE.AI", "TRADINGVIEW", "THESHIFT.SUBSTACK", "BCS MEMBER"]):
        return "Media & Subscriptions"
    if any(kw in d for kw in ["SAINSBURY'S PETROL", "SAINSBURYS PETROL", "DENTON BURN FILLING",
                                "KWIK FIT", "HALFORDS", "TT2 LTD", "TYNE TUNN", "NYX*EVCHARG"]):
        return "Car Expenses"
    if any(kw in d for kw in ["3CPAYMENT*PREMIER INN 4 NEWCASTLE", "3CPAYMENT*SLALEY HALL",
                                "MALMAISON", "CALEDONIAN HOTEL", "SLALEY HALL", "LONSDALE HOTEL",
                                "6928 - LONSDALE HOTEL", "RAMSIDE HALL", "BEAMISH HALL"]):
        return "Travel & Hotels"
    if any(kw in d for kw in ["LNER CAR PARK", "LNERCARPK", "LONDON NORTH EASTERN", "TFL TRAVEL",
                                "UBER TRIP", "TRAINLINE", "NEXUS TRAVEL", "SUMUP*TAXI",
                                "RINGGO", "PAYBYPHONE", "NEWCASTLE CITY COU", "W H SMITH"]):
        return "Travel & Transport"
    if any(kw in d for kw in ["CINEWORLD", "TYNESIDE BADMINTON", "HOLLYWOOD BOWL", "GO OUTDOORS",
                                "NEWCASTLE THEATRE", "TYNE THEATRE", "SUMUP*TYNE THEATRE",
                                "TICKETMASTER", "TM *TICKETMASTER", "HUSTLER POOL",
                                "UTILITA ARENA", "O2 CITY HALL"]):
        return "Sport & Leisure"
    if any(kw in d for kw in ["BOOTS", "DENPLAN", "WINDMILL ORTHODONT", "WELLNESS EMPOWER",
                                "JG *MACMILLAN", "WWW.DEMENTIAUK", "DACCS HAIR", "L BRITTON",
                                "SP ANCIENT BRAVE", "SP UPCIRCLE", "SP MODIBODI", "SP STEPPRS",
                                "SP ARCHIES", "SP NUTRITION GEEKS", "ROSEDENE", "SEPHORA", "MOLTON BROWN"]):
        return "Health & Wellbeing"
    if any(kw in d for kw in ["BCS MEMBERS", "BCS MEMBERSHIP", "APM", "HM PASSPORT", "STARTFIRST", "PARK ROAD SF"]):
        return "Professional Services"
    if any(kw in d for kw in ["CURRY'S", "CURRYS", "APPLE STORE", "PROCOOK"]):
        return "Electronics"

    food_kw = [
        "SAINSBURY'S SUPERMARKET", "SAINSBURY'S             ", "SAINSBURYS SUPERMARKET",
        "ASDA STORES", "ASDA ", "M&S STORE", "M&S GATESHEAD", "M&S METRO", "M&S NEWCASTLE",
        "MARKS & SPENCER", "MARKS AND SPENCER", "CO-OP WASHINGTON", "TESCO STORE",
        "SANSBRY", "WAITROSE 619", "ELM TREE FARM", "BROCKSBU", "BROOKSIDE NURSERY",
        "RINGTONS", "SWEET SYMPHONY", "SWEETSYMPHONY", "TEAL FARM", "92174 TEAL FARM",
        "DOJO*PITY ME NURSERY",
    ]
    if any(kw in d for kw in food_kw):
        return "Food Shopping"

    eat_kw = [
        "STARBUCKS", "MCDONALDS", "MCDONALD'S", "COSTA ", "FIVE GUYS", "TORTILLA",
        "ZAPATISTA", "BREWDOG", "BABUCHO", "CHAMPS SPORTS BAR", "WASHINGTON ARMS",
        "9012297 WASHINGTON ARMS", "DEEP NORTH", "DOJO*", "FAT FRANKS", "SIMLA",
        "KAMAL", "KYLIN ORIENTAL", "KING NEPTUNE", "JD WETHERSPOON", "WETHERSPOON",
        "DOMINOS", "HARRY'S BAR NEWCASTLE", "LATHAMS", "PRIDE OF SPITALFIELDS",
        "THE SUN WHARF", "LIBERTY BOUNDS", "TRAITORS GATE", "E1 FOODHALL",
        "A & S TAKEAWAY", "ARCHER AND SCOTTS", "STARVED", "SIR WILLIAM DE WES",
        "THREE MILE", "JOB BULLMAN", "BONEMI", "ZETTLE *LAKESIDE", "ZETTLE *OLIVIAS",
        "SHIFT4*EL TORERO", "0033 - GATESHEAD", "0511 - WASHINGTON", "1637 - WASHING",
        "1675 - METRO CENTRE", "5698 - DEARNE VALLEY", "6222 -RAVENSW", "ASTRONOMER",
    ]
    if any(kw in d for kw in eat_kw):
        return "Eating & Drinking Out"

    clothing_kw = [
        "AMAZON", "AMZN", "AMZ*", "NEXT ONLINE", "PRIMARK", "HOLLIST", "LEVI ", "LEVIS",
        "ARGOS", "IKEA", "FENWICK", "OLIVER BONAS", "CARD FACTORY", "HOTEL CHOCOLAT",
        "JELLYCAT", "RADLEY", "WATERSTONES", "SP POM POM", "SP SAINTSOFIA", "SP WUKA",
        "VIVA*FLYING TIGER", "MANGO", "PULL AND BEAR", "STRADIVARIUS", "BECKLE",
        "HOME BARGAINS", "B&Q", "THE RANGE", "LAKELAND", "MORPETH GARDEN",
        "HUSH HOMEWEAR", "WWW.JOHNLEWIS", "JOHN LEWIS", "WHITE STUFF", "SPORTSDIRECT",
        "JG *ISLA COOK", "NATIONAL MERCH", "3CPAYMENT*GEORGE WASHIN",
    ]
    if any(kw in d for kw in clothing_kw):
        return "Clothing & Shopping"

    return "Clothing & Shopping"


# ── Barclays categorisation ────────────────────────────────────────────────────
def categorise_barclays(row) -> str | None:
    memo   = str(row.get("Memo", "") or "").upper().strip()
    subcat = str(row.get("Subcategory", "") or "").strip()
    amount = float(row.get("Amount", 0) or 0)
    spend  = -amount

    if any(kw in memo for kw in ["AMERICAN EXPRESS", "FASL PRIM CLIENT B"]):
        return None

    if amount > 0:
        if "EXPLEO" in memo or "UK EXCL" in memo:
            return "Salary"
        if "FIDELITY FASL PYMT" in memo:
            return "Fidelity"
        return None

    # Susan's BMW car lease — £525.60/month (ends Jul 2027)
    if "BMW FINANCIAL SERV" in memo:
        return "Susan's Car Lease"

    if any(kw in memo for kw in ["NOTEMACHINE", "CARDTRONICS", "LLOYDS", "HALIFAX PLC", "NAT WEST BANK", "HSBC"]):
        return "Cash Withdrawals"
    if subcat == "Cash Withdrawal":
        return "Cash Withdrawals"

    if "WWW.FIDELITY.CO.UK" in memo:
        return "Fidelity"
    if "SEATFROG" in memo:
        return "Seatfrog"
    if "LIAM SANGHA" in memo and "COURSE FEES" in memo:
        return "University Fees"
    if "LIAM SANGHA" in memo:
        return "Liam Sangha"
    if "HMRC" in memo:
        return "Tax"

    if any(kw in memo for kw in ["OCTOPUS ENERGY", "OCTOPUS ENE", "NORTHUMBRIAN WATER",
                                   "TALKTALK", "TV LICENCE", "PARENTPAY", "D&G BOILER"]):
        return "Utilities"
    if any(kw in memo for kw in ["H3G", "TALKMOBILE", "VODAFONE LTD"]):
        return "Broadband & Phone"
    if any(kw in memo for kw in ["NETFLIX", "NORTON", "TRADINGVIEWV", "THESHIFT.SUBSTACK",
                                   "BCS MEMBERS", "BCS MEMBERSHIP"]):
        return "Media & Subscriptions"
    if any(kw in memo for kw in ["KWIK FIT", "HALFORDS", "DENTON BURN", "SAINSBURYS PETROL",
                                   "SAINSBURY'S PETROL", "NYX*EVCHARG", "NYX*GEORGEWASHINGT",
                                   "TT2 LTD", "TYNE TUNN", "MERCEDES BENZ", "MBUK",
                                   "BMW FINANCIAL SERV"]):
        return "Car Expenses"
    if any(kw in memo for kw in ["PREMIER INN", "IBIS", "MALMAISON", "CALEDONIAN HOTEL",
                                   "SLALEY HALL", "LONSDALE HOTEL", "RAMSIDE HALL",
                                   "BEAMISH HALL", "BAY TREE HILGAY"]):
        return "Travel & Hotels"
    if any(kw in memo for kw in ["TUI", "HERTZ", "SPAIN", "EUR", "ADEJE", "TENERIFE",
                                   "CRISTIAN", "AIRPORT"]):
        return "Holidays"
    if any(kw in memo for kw in ["LNER", "TFL TRAVEL", "UBER", "TRAINLINE", "NEXUS TRAVEL",
                                   "RINGGO PARKING", "PAYBYPHONE", "NEWCASTLE CITY COU",
                                   "NEWCASTLE CC PARKI", "PPOINT_*SLATYFORD", "POST OFFICE COUNTE"]):
        return "Travel & Transport"
    if any(kw in memo for kw in ["EVERYONE ACTIVE", "BADMINTON CLUB", "MR THEYAGA SHANDRA",
                                   "HOLLYWOOD BOWL", "CINEWORLD", "NEWCASTLE THEATRE",
                                   "O2 CITY HALL", "UTILITA ARENA", "HUSTLER POOL",
                                   "SUNDERLAND CC", "KESWICK BREWING", "BALLYS NEWCASTLE"]):
        return "Sport & Leisure"
    if any(kw in memo for kw in ["BOOTS UK", "DENPLAN", "WINDMILL ORTHODONT", "WELLNESS EMPOWER",
                                   "L BRITTON", "WWW.HAIRTRADE", "DACCS HAIR",
                                   "GLC LIMITED", "ROSEDENE", "SP WUKA", "SP SAINTSOFIA", "SP POM POM"]):
        return "Health & Wellbeing"
    if any(kw in memo for kw in ["BCS MEMBERS", "BCS MEMBERSHIP", "APM", "HM PASSPORT OFFICE"]):
        return "Professional Services"
    if any(kw in memo for kw in ["CURRY'S", "CURRYS", "APPLE STORE", "AO MANCHESTER", "GEBERIT"]):
        return "Electronics"

    food_kw = [
        "SAINSBURYS S/MKTS", "MARKS&SPENCER PLC", "ASDA", "CO-OP", "TESCO",
        "CLR*KNITSLEY", "ELM TREE FARM", "BRIDGESSTONES", "COSY NEWS", "CLAY'S GARDEN",
        "RINGTONS LTD", "INDIRAN MINIMART", "MORPETH GARDEN", "D AND G NEWS",
        "HOCKWOLD STORES", "HOME BARGAINS", "THE RANGE", "LAKELAND",
        "SNEHAREDDY", "SINGING CANARY", "BOOTS UK ECOMM",
    ]
    if any(kw in memo for kw in food_kw):
        return "Food Shopping"

    eat_kw = [
        "CROSS KEYS", "GREGGS", "KYLIN ORIENTAL", "CHAMPS SPORTS BAR", "THE BOTANIST",
        "TACO BELL", "BABUCHO", "COLMANS CATERERS", "PING ON CHOP SUEY", "BEAMISH MUSEUM",
        "FRIEZ AND BURGZ", "THE GREEN", "SUBWAY", "PERSIAN BITE", "CONCORD TANDOORI",
        "SHIP LANGSTONE", "ZETTLE_*OLIVIAS CO", "6888 - PARK VIEW", "BALTIC FLOUR MILLS",
        "BALTIX FLOUR MILLS", "ELVET ICES", "DOJO*", "NOVELLO'S", "NOVELLO ",
        "O'BRIENS", "O BRIENS", "THE BAKE HOUSE", "ALVINOS",
        "SCREAM SUNDERLAND", "THE BRIDGES SUNDER", "NEWCASTLE HALAL SU",
    ]
    if any(kw in memo for kw in eat_kw):
        return "Eating & Drinking Out"

    clothing_kw = [
        "NEXT ", "PRIMARK", "HOLLISTER", "LEVI", "ARGOS", "IKEA", "FENWICK",
        "OLIVER BONAS", "CARD FACTORY", "HOTEL CHOCOLAT", "JELLYCAT", "RADLEY",
        "W H SMITH", "WATERSTONES", "PULL AND BEAR", "PULLANDBEAR", "STRADIVARIUS",
        "HUSH HOMEWEAR", "WWW.JOHNLEWIS", "WHITE STUFF", "SPORTSDIRECT",
        "JG *ISLA COOK", "MARKS&SPENCER PLC", "AMAZON*", "CLARKS",
    ]
    if any(kw in memo for kw in clothing_kw):
        return "Clothing & Shopping"

    if subcat in ("Funds Transfer", "Standing Order") and spend > 0:
        return "Clothing & Shopping"
    if subcat == "Direct Debit" and spend > 0:
        return "Utilities"

    # Catch-all for any remaining outgoing transactions
    if spend > 0:
        return "Clothing & Shopping"

    return None
