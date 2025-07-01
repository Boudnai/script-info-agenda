# IMPORTANT : IL NE FAUT PAS QU'IL Y AI DE CARACTÈRE SPÉCIAUX ENTRE LES BALISES DANS LE FICHIER CONTENANT LES DONNÉES (PAS DE PARENTHÈSES PAR EXEMPLE) 
# SINON CE N'EST PAS PRIS EN COMPTE PAR LE SCRIPT ET LAISSE UNE CASE VIDE
import re
import openpyxl 
from openpyxl import Workbook

def extract_data(file_path, output_excel):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Données" # nom du tableau Excel, à changer si on veut

    sheet.append([
        "Nom Prénom", "Numéro de Tel", "Mail", "Prix", "Moyen de Paiement", "Service"
    ])

    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    name_pattern = re.findall(r'"(.*?)"', content)
    phone_pattern = re.findall(r'#(.*?)#', content)
    mail_pattern = re.findall(r'\+(.*?)\+', content)
    price_pattern = re.findall(r'€(.*?)€', content)
    payment_pattern = re.findall(r'=(.*?)=', content)
    service_pattern = re.findall(r'\*(.*?)\*', content)

    for i in range(len(name_pattern)):
        sheet.append([
            name_pattern[i],
            phone_pattern[i] if i < len(phone_pattern) else "",
            mail_pattern[i] if i < len(mail_pattern) else "",
            price_pattern[i] if i < len(price_pattern) else "",
            payment_pattern[i] if i < len(payment_pattern) else "",
            service_pattern[i] if i < len(service_pattern) else ""
        ])

    workbook.save(output_excel)
    print(f"Fichier Excel généré : {output_excel}")


extract_data("exemple.txt", "donnees.xlsx") # exemple.txt = file which contain data.   # donnees.xlsx = Excel file's name
