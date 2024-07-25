import argparse
import random
import string
import os
import json
import html
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
from simple_salesforce import Salesforce, SalesforceLogin, SFType

def creer_zone_de_texte(service, presentation_id, numero_diapo, x_position_cm, y_position_cm, width_cm, height_cm, texte, couleur_texte, centrer_texte, taille_police, gras=False):
    try:
        if texte:
            # Générer un ID unique pour la nouvelle zone de texte
            text_box_id = ''.join(random.choices(string.ascii_letters + string.digits, k=28))
            
            # Convertir les positions et dimensions de cm en points (1 cm = 28.3465 points)
            cm_to_points = 28.3465
            x_position = x_position_cm * cm_to_points
            y_position = y_position_cm * cm_to_points
            width = width_cm * cm_to_points
            height = height_cm * cm_to_points

            # Debugging print statements
            print(f'ID de la boîte de texte : {text_box_id}')
            print(f'Positions : x = {x_position}, y = {y_position}')
            print(f'Dimensions : width = {width}, height = {height}')
            print(f'Texte : {texte}')
            
            # Définir l'alignement du texte
            alignment = 'CENTER' if centrer_texte else 'START'
            
            # Requête pour créer une zone de texte
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

            # Exécuter la requête pour ajouter la zone de texte
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
    # Chemin vers votre fichier de clé JSON
    credentials = service_account.Credentials.from_service_account_file(
        r'C:\Users\ryan el fatihi\Downloads\projetrfp-07b76bd310e7.json',
        scopes=['https://www.googleapis.com/auth/presentations']
    )

    # Créer un client pour l'API Google Slides
    service = build('slides', 'v1', credentials=credentials)

    # Authentification Salesforce
    username = 'ryan@rfp.demo'
    password = 'Salesforce1'
    security_token = ''

    try:
        sf = Salesforce(username=username, password=password, security_token=security_token)
    except Exception as e:
        return None

    # Nom de votre table Salesforce (sObject)
    table_name = 'Executive_Summary__c'

    # Champ spécifique que vous souhaitez récupérer
    fields = ['Enjeux__c', 'Name', 'Strategie__c', 'Why_Salesforce__c']

    # Requête SOQL pour récupérer tous les champs spécifiés
    query = f"SELECT {', '.join(fields)} FROM {table_name} WHERE Name = 'WINOA'"

    # Effectuer la requête Salesforce
    try:
        response = sf.query(query)
        records = response.get('records', [])

        if records:
            record = records[0]  # Prendre le premier enregistrement trouvé

            # Récupérer et formater les champs
            enjeux_html = record.get('Enjeux__c', '{}')
            exec_name = record.get('Name', 'Aucune donnée trouvée')
            strategie_html = record.get('Strategie__c', '{}')
            why_salesforce_html = record.get('Why_Salesforce__c', '{}')

            # Fonction pour décoder récursivement les entités HTML
            def recursive_html_unescape(text):
                unescaped_text = html.unescape(text)
                while unescaped_text != text:
                    text = unescaped_text
                    unescaped_text = html.unescape(text)
                return unescaped_text

            # Décoder les entités HTML de manière récursive
            enjeux_json = recursive_html_unescape(enjeux_html)
            strategie_json = recursive_html_unescape(strategie_html)
            why_salesforce_json = recursive_html_unescape(why_salesforce_html)

            # Fonction pour nettoyer les caractères de contrôle du JSON
            def clean_json_string(json_str):
                json_str = json_str.replace('\n', ' ').replace('\r', ' ')
                # Échapper correctement les guillemets simples et doubles
                json_str = json_str.replace("\\'", "'").replace('\\"', '"')
                # Supprimer les caractères de contrôle non valides
                json_str = re.sub(r'[\x00-\x1F\x7F]', '', json_str)
                return json_str

            # Nettoyer les chaînes JSON
            enjeux_json = clean_json_string(enjeux_json)
            strategie_json = clean_json_string(strategie_json)
            why_salesforce_json = clean_json_string(why_salesforce_json)

            # Parser les champs JSON
            def parse_json_field(json_str):
                try:
                    data = json.loads(json_str)
                    return data
                except json.JSONDecodeError as e:
                    return f'Format JSON invalide : {e}'

            enjeux_data = parse_json_field(enjeux_json)
            strategie_data = parse_json_field(strategie_json)
            why_salesforce_data = parse_json_field(why_salesforce_json)

            # Préparer les données à retourner
            data = {
                "Name": exec_name,
                "Enjeux__c": enjeux_data,
                "Strategie__c": strategie_data,
                "Why_Salesforce__c": why_salesforce_data
            }

            return data
        else:
            return None

    except Exception as e:
        return None

def format_json_to_text(json_data):
    if isinstance(json_data, dict):
        # Retourner uniquement les valeurs
        text_lines = [str(value) for key, value in json_data.items()]
        return "\n".join(text_lines)
    elif isinstance(json_data, list):
        return "\n".join(json_data)
    else:
        return str(json_data)

def main():
    parser = argparse.ArgumentParser(description='Ajouter des zones de texte à une présentation Google Slides.')
    parser.add_argument('presentation_id', type=str, help='ID de la présentation')
    parser.add_argument('numero_diapo', type=str, help='Numéro de la diapositive')
    
    args = parser.parse_args()
    
    # Chemin vers votre fichier de clé JSON
    credentials = service_account.Credentials.from_service_account_file(
        r'C:\Users\ryan el fatihi\Downloads\projetrfp-07b76bd310e7.json',
        scopes=['https://www.googleapis.com/auth/presentations']
    )

    # Créer un client pour l'API Google Slides
    service = build('slides', 'v1', credentials=credentials)
    
    # Récupérer les données de Salesforce
    data = fetch_salesforce_data()
    if data:
        # Variables de position et de taille pour chaque zone de texte
        positions = [
            ( 4.18,3.89, 8.82,3.16, 'Enjeux__c'),
            (1.14, 2.37, 3.04, 1.75, 'Strategie__c'),
            (1.14, 6.62, 3.04, 7.12, 'Why_Salesforce__c')
        ]
        
        couleur_texte = (0.0, 0.0, 0.0)  # noir
        centrer_texte = False
        taille_police = 7
        gras = False
        
        json_data = data.get('Enjeux__c', 'Enjeux non trouvé')
        texte = format_json_to_text(json_data)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 4.41, 2.77, 8.82, 3.16, texte, couleur_texte, centrer_texte, taille_police, gras)

        json_data = data.get('Strategie__c', 'Strategie non trouvé')
        texte = format_json_to_text(json_data)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 4.36, 6.35, 8.82, 3.16, texte, couleur_texte, centrer_texte, taille_police, gras)
    
        json_data = data.get('Why_Salesforce__c', 'Why Salesforce non trouvé')
        texte = format_json_to_text(json_data)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 4.36, 9.47, 8.82, 3.16, texte, couleur_texte, centrer_texte, taille_police, gras)
    
    else:
        print("Aucune donnée récupérée de Salesforce.")

    # Appel de la fonction pour créer la zone de texte dans une diapositive spécifiqu
    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 1.14, 2.77, 3.04, 3.03,'Vision' , (0.0,0.0,1.0), True, 12, True)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la première zone de texte : {e}')

    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 1.14, 6.36, 3.04, 2.21, 'Vos enjeux', (0.0,0.0,1.0), True, 12, True)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la deuxième zone de texte : {e}')

    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 1.14, 9.12, 3.04, 4.81, 'Notre Engagement', (0.0,0.0,1.0), True, 11, True)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 7.25, 0.38, 10.91, 1.71, 'Synthèse de notre réponse', (0.0,0.0,1.0), True, 20, True)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 13.18, 2.49, 9.54,0.59, 'Nos Solutions', (1.0,1.0,1.0), True, 14, True)

    except Exception as e:
        print(f'Erreur lors de l\'ajout de la troisième zone de texte : {e}')

if __name__ == "__main__":
    main()

