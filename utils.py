import re


def extract_sol_num(subject):
    pattern = r'.Nueva\s+asignacion\s+de\s+pedido\s+-\s+(\w+-\d{4}-[A-Z]\d{1,4}-\d+).'
    match = re.search(pattern, subject)
    if match:
        return match.group(1)
    return None

# Obtener la SDATOOL
def extract_sdatool_value(text):
    sdatool_match = re.search(r'SDATOOL:\s*(\w+)', text)
    if sdatool_match:
        return sdatool_match.group(1)
    return None


# obtener periodo de sol_num
def extract_period_from_sol_num(subject):
    
# encontrar la primera aparición de un año (cuatro dígitos consecutivos) seguida de un guión y una letra mayúscula
    match = re.search(r'(\d{4})-(Q\d)', subject)
    if match:
        year = match.group(1)
        quarter = match.group(2)
        return f"{year} - {quarter}"
    return None


# obtener cargo
def extract_creator(text):
    match = re.search(r'El pedido fue cargado por:\s*([A-Z\s,]+)\s*el:', text)
    if match:
        return match.group(1)
    return None


# obtener MVP del body
def extract_mvp(text):
    match = re.search(r'MVP:\s*(.+)', text)
    if match:
        return match.group(1).replace('\r', '').replace('\n', '')
    return None

# obtener total horas
def extract_total_hours(text):
    match = re.search(r'Horas totales:\s*([\d,]+)', text)
    if match:
        return float(match.group(1).replace(',', '.'))
    return None


#obtener perfil
def extract_profiles(text):
    # Define  el patrón para que coincida con cada perfil.
    pattern = r'Perfil\s+(\d+):(.*?)Unidades:(\d+,\d+)'
    
    #Encuentra todas las coincidencias del patrón en el texto. 
    matches = re.findall(pattern, text, re.DOTALL)
    
    # Inicializar una lista para almacenar información de perfil
    perfiles = []
    
    # Iterar sobre las coincidencias
    for match in matches:
        # Extraer número de perfil, descripción y unidades.
        perfil_number = match[0]
        perfil_description = match[1].strip().replace('\r', '').replace('\n', '')

        unidades = match[2]
        
        # Almacenar información de perfil en un diccionario
        perfil_info = {
            "perfil_number": perfil_number,
            "perfil_description": perfil_description,
            "unidades": float(unidades.replace(',', '.')),

        }
        
        # Agregar información de perfil a la lista
        perfiles.append(perfil_info)
    
    return perfiles


def get_rows_to_append(subject, body):
    # este tiene perfil, unidades,
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


# this function receives the 'inbox'
def build_folder_dict(folder, path=''):
    folder_dict = {}
    full_path = path + '/' + folder.Name if path else folder.Name
    # folder_dict[full_path] = folder
    folder_dict[folder.Name] = full_path
    subfolders = folder.Folders
    for subfolder in subfolders:
        folder_dict.update(build_folder_dict(subfolder, full_path))
    return folder_dict