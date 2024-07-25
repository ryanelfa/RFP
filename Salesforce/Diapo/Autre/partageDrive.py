from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Chemin vers votre fichier de clé JSON
credentials = service_account.Credentials.from_service_account_file(
    r'C:\Users\ryan el fatihi\Downloads\projetrfp-07b76bd310e7.json',
    scopes=['https://www.googleapis.com/auth/drive']
)

# Créer un client pour l'API Google Drive
drive_service = build('drive', 'v3', credentials=credentials)

# Demander à l'utilisateur l'ID de la présentation
presentation_id = input("Entrez l'ID de la présentation que vous souhaitez partager : ")

# Demander à l'utilisateur l'adresse e-mail du compte personnel
email_address = input("Entrez l'adresse e-mail du compte personnel : ")

# Définir les nouvelles permissions pour partager avec le compte personnel
new_permission = {
    'type': 'user',
    'role': 'writer',  # Peut être 'reader', 'writer', ou 'commenter'
    'emailAddress': email_address  # Adresse e-mail du compte personnel
}

try:
    # Vérifier que le fichier existe
    file = drive_service.files().get(fileId=presentation_id).execute()
    print(f"Le fichier avec l'ID {presentation_id} a été trouvé : {file['name']}")

    # Ajouter les nouvelles permissions à la présentation
    drive_service.permissions().create(fileId=presentation_id, body=new_permission).execute()
    print('Présentation partagée avec succès avec votre compte personnel.')
except HttpError as error:
    print(f"Une erreur s'est produite : {error}")
