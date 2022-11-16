import json
import os

import requests
from dotenv import load_dotenv

def create_page(co):
    load_dotenv()

    tokenNotion = os.getenv('TOKEN_NOTION')
    database = os.getenv('DATABASE')


    url = "https://api.notion.com/v1/pages/"

    payload = json.dumps({
    "parent": {
        "database_id": database
    },
    "properties": {
        "Pedido": {
            "title": [
                {
                    "text": {
                        "content": co
                    }
                }
            ]
        }
    }
})
    headers = {
    'Authorization': tokenNotion,
    'Content-Type': 'application/json',
    'Notion-Version': '2021-05-13'
}

    response = requests.request("POST", url, headers=headers, data=payload)

    jsonData = json.loads(response.text)
    
    idPage = jsonData["id"]
    
    print(idPage)
    
    create_database(id=idPage, co=co)
    
    
    

def create_database(id, co):
    
    load_dotenv()

    tokenNotion = os.getenv('TOKEN_NOTION')

    url = "https://api.notion.com/v1/databases/"

    payload = json.dumps({
    "parent": {
        "type": "page_id",
        "page_id": id
    },
    "title": [
        {
            "type": "text",
            "text": {
                "content": co,
                "link": None
            }
        }
    ],
    "is_inline": True,
    "properties": {
        "Nombre": {
            "id": ">^X_",
            "name": "Nombre",
            "type": "rich_text",
            "rich_text": {}
        },
        
        "Desc": {
            "id": "KKu[",
            "name": "Desc",
            "type": "rich_text",
            "rich_text": {}
        },
        "MO": {
            "id": "X~v\\",
            "name": "MO",
            "type": "rich_text",
            "rich_text": {}
        },
        "Qty": {
            "id": "ZByb",
            "name": "Qty",
            "type": "number",
            "number": {
                "format": "number"
            }
        },
        "Cortado": {
            "id": "@VRz",
            "name": "Cortado",
            "type": "checkbox",
            "checkbox": {}
        },
        "Item": {
            "id": "title",
            "name": "Item",
            "type": "title",
            "title": {}
        }
    }
})
    headers = {
    'Content-Type': 'application/json',
    'Notion-Version': '2021-05-13',
    'Authorization': tokenNotion
}

    response = requests.request("POST", url, headers=headers, data=payload)
    
    jsonData = json.loads(response.text)

    idDatabase = jsonData["id"]
    
    print(idDatabase)
    
    return (idDatabase)
    
def create_item(database, name, desc, mo, item):
    load_dotenv()

    tokenNotion = os.getenv('TOKEN_NOTION')

    url = "https://api.notion.com/v1/pages/"

    payload = json.dumps({
        "parent": {
            "database_id": database
        },
        "properties": {
            "Item": {
                "title": [
                    {
                        "text": {
                            "content": item
                        }
                    }
                ]
            }
        }
    })
    headers = {
        'Authorization': tokenNotion,
        'Content-Type': 'application/json',
        'Notion-Version': '2021-05-13'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    
    print(response.text)
    
    
