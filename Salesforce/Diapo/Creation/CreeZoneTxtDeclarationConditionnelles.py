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
    
    creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 1.19, 0.14, 21.44, 2.06, '\nDéclarations Conditionnelles', (0.0, 0.0, 1.0), False, 24, True)

    texte2 = """Déclaration conforme à la directive « Safe Harbor » contenue dans la loi américaine intitulée « Private Securities Litigation Reform Act » de 1995 :

Cette présentation est susceptible de comporter des déclarations conditionnelles, qui impliquent nécessairement une certaine prise de risque, des incertitudes et des hypothèses. Si l'une de ces incertitudes se concrétise ou si certaines hypothèses se révèlent incorrectes, les résultats de Salesforce, Inc. pourraient être sensiblement différents de ceux explicitement ou implicitement avancés par nos déclarations conditionnelles. Toutes les déclarations ne portant pas sur des faits historiques peuvent être considérées comme conditionnelles, y compris les projections de disponibilité des produits ou des services, d'augmentation du nombre d'abonnés, de bénéfices, de chiffre d'affaires ou autre valeur financière, toute déclaration concernant les stratégies ou les plans de gestion des opérations à venir, toute opinion personnelle, toute déclaration concernant les services ou les développements technologiques nouveaux, planifiés ou mis à niveau, ainsi que les contrats clients et l'utilisation de nos services.
Les incertitudes et les risques susmentionnés concernent, sans s'y limiter, les risques associés au développement et à la fourniture de nouvelles fonctionnalités pour notre service, aux nouveaux produits et services, à notre nouveau modèle commercial, nos pertes d'exploitation antérieures, les éventuelles fluctuations de nos résultats d'exploitation et de notre taux de croissance, les interruptions ou les retards de notre système d'hébergement, les failles des mesures de sécurité, l'issue des litiges, les risques associés aux fusions et acquisitions réelles et éventuelles, la jeunesse du marché dans lequel nous évoluons, notre historique relativement limité, notre capacité à développer, fidéliser et motiver notre personnel et à gérer notre croissance, les nouvelles éditions de notre service, ainsi que le déploiement réussi chez les clients, notre expérience limitée en matière de revente de produits tiers, et l'utilisation et les ventes à de grands comptes. Vous trouverez plus d'informations sur les facteurs pouvant influencer les résultats financiers de Salesforce, Inc. dans notre rapport annuel (formulaire 10-K) pour l'exercice fiscal le plus récent et dans notre rapport trimestriel (formulaire 10-Q) pour le trimestre fiscal le plus récent. Ce rapport et d'autres documents contenant d'importantes informations sont accessibles sur notre site web dans la partie Informations Investisseurs, section Documents pour la Commission des opérations de bourse (SEC).
Certains services ou fonctions qui ne sont pas encore commercialisés et sont mentionnés ici ou dans d'autres présentations, communiqués de presse ou déclarations publiques, ne sont pas encore disponibles et ne seront peut-être pas livrés à temps, voire pas livrés du tout. Les clients qui achètent nos services doivent prendre leur décision sur la base des fonctions actuellement disponibles. Salesforce, Inc. n'est pas tenu et n'a pas l'intention de mettre à jour ces déclarations conditionnelles.
"""

    try:
        creer_zone_de_texte(service, args.presentation_id, args.numero_diapo, 1.19, 2.5, 22.36, 8.8, texte2 ,(0.0,0.0,0.0), False, 9, False)
    except Exception as e:
        print(f'Erreur lors de l\'ajout de la deuxième zone de texte : {e}')

if __name__ == "__main__":
    main()
