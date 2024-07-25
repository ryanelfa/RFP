import argparse
import random
import string
import json
import html
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
from simple_salesforce import Salesforce

def creer_zone_de_texte(service, presentation_id, numero_diapo, x_position_cm, y_position_cm, width_cm, height_cm, texte, couleur_texte, centrer_texte, taille_police, gras=False):
    try:
        if texte:
            text_box_id = ''.join(random.choices(string.ascii_letters + string.digits, k=28))
            cm_to_points = 28.3465
            x_position = x_position_cm * cm_to_points
            y_position = y_position_cm * cm_to_points
            width = width_cm * cm_to_points
            height = height_cm * cm_to_points

            print(f'ID de la boîte de texte : {text_box_id}')
            print(f'Positions : x = {x_position}, y = {y_position}')
            print(f'Dimensions : width = {width}, height = {height}')
            print(f'Texte : {texte}')

            alignment = 'CENTER' if centrer_texte else 'START'

            requests = [{
                'createShape': {
                    'objectId': text_box_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': numero_diapo,
                        'size': {
                            'height': {
                                'magnitude': height,
                                'unit': 'PT'
                            },
                            'width': {
                                'magnitude': width,
                                'unit': 'PT'
                            }
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': x_position,
                            'translateY': y_position,
                            'unit': 'PT'
                        }
                    }
                }
            },
            {
                'insertText': {
                    'objectId': text_box_id,
                    'text': texte
                }
            },
            {
                'updateTextStyle': {
                    'objectId': text_box_id,
                    'style': {
                        'foregroundColor': {
                            'opaqueColor': {
                                'rgbColor': {
                                    'red': couleur_texte[0],
                                    'green': couleur_texte[1],
                                    'blue': couleur_texte[2]
                                }
                            }
                        },
                        'fontSize': {
                            'magnitude': taille_police,
                            'unit': 'PT'
                        },
                        'bold': gras
                    },
                    'fields': 'foregroundColor,fontSize,bold'
                }
            },
            {
                'updateParagraphStyle': {
                    'objectId': text_box_id,
                    'style': {
                        'alignment': alignment
                    },
                    'fields': 'alignment'
                }
            }]

            response = service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()

            print(f'Zone de texte ajoutée avec succès à la diapositive {numero_diapo}.')
        else:
            print("Le texte fourni est vide.")
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la zone de texte : {e}')

def fetch_salesforce_data():
    username = 'ryan@rfp.demo'
    password = 'Salesforce1'
    security_token = ''

    try:
        sf = Salesforce(username=username, password=password, security_token=security_token)
    except Exception as e:
        print(f'Erreur lors de la connexion à Salesforce : {e}')
        return None

    table_name = 'Solutions_Salesforce__c'
    fields = ['Name', 'Solutions_prompt__c', 'Solutions_Json__c']
    query = f"SELECT {', '.join(fields)} FROM {table_name} WHERE Name = 'WINOA'"

    try:
        response = sf.query(query)
        records = response.get('records', [])

        if records:
            record = records[0]
            Solutions_html = record.get('Solutions_Json__c', '{}')
            name = record.get('Name', 'Aucune donnée trouvée')

            def recursive_html_unescape(text):
                unescaped_text = html.unescape(text)
                while unescaped_text != text:
                    text = unescaped_text
                    unescaped_text = html.unescape(text)
                return unescaped_text

            Solutions_json = recursive_html_unescape(Solutions_html)

            def clean_json_string(json_str):
                json_str = json_str.replace('\n', ' ').replace('\r', ' ')
                json_str = json_str.replace("\\'", "'").replace('\\"', '"')
                json_str = re.sub(r'[\x00-\x1F\x7F]', '', json_str)
                return json_str

            enjeux_json = clean_json_string(Solutions_json)

            def parse_json_field(json_str):
                try:
                    data = json.loads(json_str)
                    return data
                except json.JSONDecodeError as e:
                    return f'Format JSON invalide : {e}'

            enjeux_data = parse_json_field(Solutions_json)

            return enjeux_data
        else:
            print("Aucun enregistremen  t trouvé dans Salesforce.")
            return None

    except Exception as e:
        print(f'Erreur lors de la requête Salesforce : {e}')
        return None

def format_json_to_text(json_data):
    if isinstance(json_data, dict):
        text_lines = []
        for key, value in json_data.items():
            text_lines.append(f"{key}:\n{value}\n")
        return text_lines
    elif isinstance(json_data, list):
        return json_data
    else:
        return [str(json_data)]

def main():
    parser = argparse.ArgumentParser(description='Ajouter des zones de texte à une présentation Google Slides.')
    parser.add_argument('presentation_id', type=str, help='ID de la présentation')
    parser.add_argument('numero_diapo', type=str, help='Numéro de la diapositive')

    args = parser.parse_args()

    credentials = service_account.Credentials.from_service_account_file(
        r'C:\Users\ryan el fatihi\Downloads\projetrfp-07b76bd310e7.json',
        scopes=['https://www.googleapis.com/auth/presentations']
    )

    service = build('slides', 'v1', credentials=credentials)

    data = fetch_salesforce_data()
    if data:
        formatted_texts = format_json_to_text(data)
        couleur_texte = (0.0, 0.0, 0.0)
        centrer_texte = False
        taille_police = 7
        gras = False

        x_positions = [15.28, 2.12, 7.9]
        y_positions = [8.52, 9.39,1.19]
        width = 8.82
        height = 3.16

        for i, texte in enumerate(formatted_texts):
            creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, x_positions[i], y_positions[i], width, height, texte, couleur_texte, centrer_texte, taille_police, gras)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo,13.41 , 8.37, 1.75,1.33, 'Sales Cloud', couleur_texte, True, 10, gras)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo,11.43 ,5.25 , 1.75,1.33 , 'Tableau', couleur_texte, True, 10, gras)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo,8.99 , 8.37,2.69 , 1.33, 'Marketing Cloud', couleur_texte, True, 10, gras)

    else:
        print("Aucune donnée récupérée de Salesforce.")

if __name__ == "__main__":
    main()
