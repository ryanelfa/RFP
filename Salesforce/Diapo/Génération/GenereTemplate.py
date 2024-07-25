import random
import string
import sys
from googleapiclient.discovery import build
from google.oauth2 import service_account

def copy_shape(service, presentation_id, new_slide_id, element):
    element_id = ''.join(random.choices(string.ascii_letters + string.digits, k=28))
    shape_properties = element['shape']['shapeProperties']
    
    create_shape_request = {
        'createShape': {
            'objectId': element_id,
            'shapeType': element['shape']['shapeType'],
            'elementProperties': {
                'pageObjectId': new_slide_id,
                'size': element['size'],
                'transform': element['transform']
            }
        }
    }
    requests = [create_shape_request]
    
    if 'shapeBackgroundFill' in shape_properties:
        shape_background_fill = shape_properties['shapeBackgroundFill']
        if shape_background_fill.get('propertyState') != 'INHERIT':
            requests.append({
                'updateShapeProperties': {
                    'objectId': element_id,
                    'shapeProperties': {
                        'shapeBackgroundFill': shape_background_fill
                    },
                    'fields': 'shapeBackgroundFill'
                }
            })
    if 'outline' in shape_properties:
        outline = shape_properties['outline']
        if outline.get('propertyState') != 'INHERIT':
            requests.append({
                'updateShapeProperties': {
                    'objectId': element_id,
                    'shapeProperties': {
                        'outline': outline
                    },
                    'fields': 'outline'
                }
            })
    if 'shadow' in shape_properties:
        requests.append({
            'updateShapeProperties': {
                'objectId': element_id,
                'shapeProperties': {
                    'shadow': shape_properties['shadow']
                },
                'fields': 'shadow'
            }
        })
    
    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={'requests': requests}
    ).execute()

    if 'text' in element['shape']:
        text_elements = element['shape']['text']['textElements']
        if text_elements:
            for text_element in text_elements:
                text_run = text_element.get('textRun')
                if text_run:
                    insert_text_request = {
                        'insertText': {
                            'objectId': element_id,
                            'text': text_run['content']
                        }
                    }
                    requests = [insert_text_request]
                    service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': requests}
                    ).execute()

def copy_image(service, presentation_id, new_slide_id, element):
    element_id = ''.join(random.choices(string.ascii_letters + string.digits, k=28))
    image_properties = element['image']['imageProperties']
    
    create_image_request = {
        'createImage': {
            'objectId': element_id,
            'url': element['image']['contentUrl'],
            'elementProperties': {
                'pageObjectId': new_slide_id,
                'size': element['size'],
                'transform': element['transform']
            }
        }
    }
    requests = [create_image_request]
    
    if 'outline' in image_properties:
        requests.append({
            'updateImageProperties': {
                'objectId': element_id,
                'imageProperties': {
                    'outline': image_properties['outline']
                },
                'fields': 'outline'
            }
        })
    if 'shadow' in image_properties:
        requests.append({
            'updateImageProperties': {
                'objectId': element_id,
                'imageProperties': {
                    'shadow': image_properties['shadow']
                },
                'fields': 'shadow'
            }
        })
    
    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={'requests': requests}
    ).execute()

def copy_table(service, presentation_id, new_slide_id, element):
    element_id = ''.join(random.choices(string.ascii_letters + string.digits, k=28))
    create_table_request = {
        'createTable': {
            'objectId': element_id,
            'rows': element['table']['rows'],
            'columns': element['table']['columns'],
            'elementProperties': {
                'pageObjectId': new_slide_id,
                'size': element['size'],
                'transform': element['transform']
            }
        }
    }
    requests = [create_table_request]
    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={'requests': requests}
    ).execute()

    for r, row in enumerate(element['table']['tableRows']):
        for c, cell in enumerate(row['tableCells']):
            text_elements = cell['text']['textElements']
            text = ''
            if text_elements:
                for text_element in text_elements:
                    text_run = text_element.get('textRun')
                    if text_run:
                        text += text_run['content']
            insert_text_request = {
                'insertText': {
                    'objectId': element_id,
                    'cellLocation': {'rowIndex': r, 'columnIndex': c},
                    'text': text
                }
            }
            requests = [insert_text_request]
            service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()

def copy_page_element(service, presentation_id, new_slide_id, element):
    if 'shape' in element:
        copy_shape(service, presentation_id, new_slide_id, element)
    elif 'image' in element:
        copy_image(service, presentation_id, new_slide_id, element)
    elif 'table' in element:
        copy_table(service, presentation_id, new_slide_id, element)
    elif 'elementGroup' in element:
        for child in element['elementGroup']['children']:
            copy_page_element(service, presentation_id, new_slide_id, child)
    else:
        print(f"Type d'élément non supporté: {element}")

def main():
    if len(sys.argv) != 4:
        print("Usage: python3 GenereTemplate.py <presentation_id_source> <presentation_id_destination> <slide_id_source>")
        sys.exit(1)
    
    presentation_id_source = sys.argv[1]
    presentation_id_destination = sys.argv[2]
    slide_id_source = sys.argv[3]
    
    new_slide_id = ''.join(random.choices(string.ascii_letters + string.digits, k=28))

    credentials = service_account.Credentials.from_service_account_file(
        r'C:\Users\ryan el fatihi\Downloads\projetrfp-07b76bd310e7.json',
        scopes=['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive']
    )
    service = build('slides', 'v1', credentials=credentials)

    try:
        presentation_source = service.presentations().get(presentationId=presentation_id_source).execute()
        slide_exists = any(page['objectId'] == slide_id_source for page in presentation_source['slides'])
        if not slide_exists:
            print(f"Erreur : La diapositive avec l'ID {slide_id_source} n'a pas été trouvée dans la présentation source.")
            sys.exit(1)
        
        slide_layout_id = next((slide['slideProperties']['layoutObjectId'] for slide in presentation_source['slides'] if slide['objectId'] == slide_id_source), None)
        if slide_layout_id is None:
            raise Exception(f"Impossible de trouver la mise en page de la diapositive source avec l'ID {slide_id_source}.")

        create_slide_request = {
            'requests': [
                {
                    'createSlide': {
                        'objectId': new_slide_id,
                        'insertionIndex': 0,
                        'slideLayoutReference': {
                            'predefinedLayout': 'BLANK'
                        }
                    }
                }
            ]
        }
        service.presentations().batchUpdate(
            presentationId=presentation_id_destination,
            body=create_slide_request
        ).execute()

        for page in presentation_source['slides']:
            if page['objectId'] == slide_id_source:
                for element in page['pageElements']:
                    copy_page_element(service, presentation_id_destination, new_slide_id, element)

        print(new_slide_id)

    except Exception as e:
        print(f"Erreur lors de l'accès à la présentation ou de l'exécution de la requête : {e}")

if __name__ == "__main__":
    main()
