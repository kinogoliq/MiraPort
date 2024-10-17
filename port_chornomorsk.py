# port_chornomorsk.py

FEES_WITHOUT_VAT = {
    "Tonnage dues (In/out)": 0.2784,
    "Canal dues (in/out)": 0.0512,
    "Lighthouse dues": 0.045,
    "Berth dues": 0.028,
    "Sanitary dues": 0.0176,
    "Administrative dues": 0.0176,
    "Port information fee": 0.0065,
}

FEES_WITH_VAT_WITH_MILES = {
    "Inward pilotage in": 0.0139,
    "Inward pilotage out": 0.0139,
    "Outward pilotage in": 0.0014,
    "Outward pilotage out": 0.0014,
}

FEES_WITH_VAT_WITHOUT_MILES = {
    "Services of VTCS": 0.1072799,
    "Tugs in": 0.2720,
    "Tugs out": 0.2720,
    "Mooring in": 0.0136832,
    "Mooring out": 0.0136832,
}

FEES_WITH_INCLUDED_VAT = {
    "Services of VTCS",
    "Tugs in",
    "Tugs out",
    "Mooring in",
    "Mooring out",
}

VAT_RATE = 0.20  # 20%
