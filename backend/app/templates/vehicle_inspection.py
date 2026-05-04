# === FINAL OUTPUT ORDER ===
COLUMNS = [
    "Inspection Date",
    "Registration",
    "Transporter",
    "Model",
    "Origin",
    "Destination",
    "Axleconf",
    "Inspsticker",
    "InsuSticker",
    "Cargo",
    "Permit issue date",
    "Height",
    "Length_",
    "Width_",
    "Abnormal Load Permit",
    "Total tyres",
    "Load Weight",
    "Authorized Weight",
    "Permit No.",
    "Date of Travel",
    "PStartD",
    "PEndD"
]

# === COLUMN MAPPING ===
COLUMN_MAPPING = {
    "Inspection Date": "Inspection Date",
    "registration": "Registration",
    "Transp": "Transporter",
    "Model": "Model",
    "Origin": "Origin",
    "destination": "Destination",
    "Axleconf": "Axleconf",
    "Inspstick": "Inspsticker",
    "InsuaranceStic": "InsuSticker",
    "Cargo": "Cargo",
    "Dpermitissu": "Permit issue date",
    "Height": "Height",
    "Length": "Length_",
    "Width": "Width_",
    "AbnormalLPermit": "Abnormal Load Permit",
    "Totaltyres": "Total tyres",
    "weighofload": "Load Weight",
    "Authweight": "Authorized Weight",
    "Permit No.": "Permit No.",
    "Date of Travel": "Date of Travel",
    "Start Date": "PStartD",
    "End Date": "PEndD"
}

# === FORMATTING RULES ===
DATE_COLUMNS = [
    "Inspection Date",
    "Date of Travel",
    "PStartD",
    "PEndD"
]

COMMA_COLUMNS = [
    "Load Weight",
    "Authorized Weight"
]
