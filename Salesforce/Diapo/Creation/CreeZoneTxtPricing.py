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
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 5.28, 0.49, 18.89, 1.71, 'Proposition licences Salesforce',
(0.0,0.0,1.0), False, 24, True)        
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.28, 2.83, 4.78, 0.9, 'Contexte', (1.0,1.0,1.0), True, 12, True)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 5.39, 2.83, 6.09, 0.9, 'Produits', (1.0,1.0,1.0), True, 12, True)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 11.58, 2.83, 2.83, 0.9, 'Quantité', (1.0,1.0,1.0), True, 12, True)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 14.64, 2.59, 3.31, 0.9, 'PUPM           (Prix Public)', (1.0,1.0,1.0), True, 10, True)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 18.25, 2.59, 3.31, 0.9, 'PUPM   (Remisé)', (1.0,1.0,1.0), True, 10, True)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 21.89, 2.83, 5.31, 0.9, 'Total Annuel', (1.0,1.0,1.0), True, 12, True)
        
        
        
        
        
        ## Idee on écrit une phrase pour chaque solutions et on la place en fonction des images 
        
        ## Sales Cloud
        ##creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.28, 3.76, 5.11, 0.9, 'Suivi des opérations commerciales & service client', (1.0,1.0,1.0), False, 8, True)
        ## Tableau
        ##creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.26, 4.95, 4.78, 0.9, 'Transformer les données en Insight et plans d’actions', (1.0,1.0,1.0), False, 8, True)
        ## Marketing Cloud
        ##creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.24, 6.13, 4.78  , 0.9, 'Communication automatisée et personnalisée', (1.0,1.0,1.0), False, 8, True)
        
        ## A modifier 
        ##creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.28, 3.76, 5.11, 0.9, 'Suivi des opérations commerciales & service client', (1.0,1.0,1.0), False, 8, True)
        ##creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.26, 4.95, 4.78, 0.9, 'Transformer les données en Insight et plans d’actions', (1.0,1.0,1.0), False, 8, True)
        ##creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.24, 6.13, 4.78  , 0.9, 'Communication automatisée et personnalisée', (1.0,1.0,1.0), False, 8, True)
        ##creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.28, 3.76, 5.11, 0.9, 'Suivi des opérations commerciales & service client', (1.0,1.0,1.0), False, 8, True)
        
        
        
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.31, 8.68,21.22 , 0.45, 'Total Licences Annuel', (1.0,1.0,1.0), True, 12, True)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0.31, 8.09,21.22 , 0.45, 'Delta avec le contrat actuel qui sera donc l’investissement en plus', (1.0,1.0,1.0), True, 12, True)
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 0, 13.6,25.4, 0.48, 'Draft- Confidential - Non-binding offer - Only Salesforce Order Forms have a contractual value - All prices are VAT excluded', (1.0,1.0,1.0), True, 12, False)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la zone de texte : {e}')
        
    
    
if __name__ == "__main__":
    main()
