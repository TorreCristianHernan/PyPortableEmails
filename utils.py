import re

# this is working as expected
def extract_sol_num(subject):
    pattern = r'.Nueva\s+asignacion\s+de\s+pedido\s+-\s+(\w+-\d{4}-[A-Z]\d{1,4}-\d+).'
    match = re.search(pattern, subject)
    if match:
        return match.group(1)
    return None

# get the SDATOOL
def extract_sdatool_value(text):
    sdatool_match = re.search(r'SDATOOL:\s*(\w+)', text)
    if sdatool_match:
        return sdatool_match.group(1)
    return None

# get period from sol_num
def extract_period_from_sol_num(subject):
    # find the first occurrence of a year (four consecutive digits) followed by a hyphen and a capital letter
    match = re.search(r'(\d{4})-(Q\d)', subject)
    if match:
        year = match.group(1)
        quarter = match.group(2)
        return f"{year} - {quarter}"
    return None


# get cargo
def extract_creator(text):
    match = re.search(r'El pedido fue cargado por:\s*([A-Z\s,]+)\s*el:', text)
    if match:
        return match.group(1)
    return None


# get MVP from body
def extract_mvp(text):
    match = re.search(r'MVP:\s*(.+)', text)
    if match:
        return match.group(1).replace('\r', '').replace('\n', '')
    return None

# get total hours
def extract_total_hours(text):
    match = re.search(r'Horas totales:\s*([\d,]+)', text)
    if match:
        return float(match.group(1).replace(',', '.'))
    return None

# def extract_total_hours(text):
#     # Updated regular expression to match variations
#     match = re.search(r'\?\s*Horas totales:\s(\d{1,3}(,\d{3})\.\d{1,2}|\d+(,\d{3})(?!\d|,))', text)
#     if match:
#         return match.group(1)
#     return None

def extract_profiles(text):
    # Define the pattern to match each profile
    pattern = r'Perfil\s+(\d+):(.*?)Unidades:(\d+,\d+)'
    
    # Find all matches of the pattern in the text
    matches = re.findall(pattern, text, re.DOTALL)
    
    # Initialize a list to store profile information
    perfiles = []
    
    # Iterate over the matches
    for match in matches:
        # Extract profile number, description, and units
        perfil_number = match[0]
        # perfil_description = match[1].strip()
        perfil_description = match[1].strip().replace('\r', '').replace('\n', '')

        unidades = match[2]
        
        # Store profile information in a dictionary
        perfil_info = {
            "perfil_number": perfil_number,
            "perfil_description": perfil_description,
            "unidades": float(unidades.replace(',', '.')),

        }
        
        # Add profile information to the list
        perfiles.append(perfil_info)
    
    return perfiles


def get_rows_to_append(subject, body):
    # this has profile, units, 
    profiles = extract_profiles(body) 
    sol_num = extract_sol_num(subject)
    period = extract_period_from_sol_num(sol_num)
    sda_tool = extract_sdatool_value(body)
    creator = extract_creator(body)
    mvp = extract_mvp(body)
    total_hours = extract_total_hours(body)
    data_to_insert = []

    for item in profiles:
        row = [
            period,
            sol_num,
            sda_tool,
            creator,
            mvp,
            total_hours,
            item['perfil_description'],
            item['unidades']
        ]
        data_to_insert.append(row)
    return data_to_insert