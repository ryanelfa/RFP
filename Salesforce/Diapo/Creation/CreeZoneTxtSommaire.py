import argparse
import random
import string
from google.oauth2 import service_account
from googleapiclient.discovery import build

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
    
    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 2.48, 2.66, 10.54, 1.07, 'Déclarations Conditionnelles',(0.0,0.0,0.0), True, 12, True)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la troisième zone de texte : {e}')
        
    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 2.48, 3.99, 10.54, 1.07, 'Executive Summary',(0.0,0.0,0.0), True, 12, True)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la quatrième zone de texte : {e}')
        
    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 2.48, 5.31, 10.54, 1.07, 'Nos Solutions Salesforce',(0.0,0.0,0.0), True, 12, True)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la cinquième zone de texte : {e}')
    
    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 2.48, 6.64, 10.54, 1.07, 'Customer Success Story',(0.0,0.0,0.0), True, 12, True)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la cinquième zone de texte : {e}')
        
    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 2.48, 7.97, 10.54, 1.07, 'Estimations Tarifaires',(0.0,0.0,0.0), True, 12, True)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la cinquième zone de texte : {e}')
        
    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 2.48, 9.29, 10.54, 1.07, 'Annexes',(0.0,0.0,0.0), True, 12, True)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la cinquième zone de texte : {e}')
    

if __name__ == "__main__":
    main()
