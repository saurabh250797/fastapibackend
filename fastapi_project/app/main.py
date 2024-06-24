from fastapi import FastAPI, HTTPException, Query
from typing import List, Dict, Any
import pandas as pd
import requests
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import os
from dotenv import load_dotenv

load_dotenv()

app = FastAPI()

# Load environment variables
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
SHAREPOINT_CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
SHAREPOINT_DOC_LIBRARY = os.getenv("SHAREPOINT_DOC_LIBRARY")
SANDWAI_API_KEY = os.getenv("SANDWAI_API_KEY")
SANDWAI_API_URL = os.getenv("SANDWAI_API_URL")

# Initialize SharePoint context
ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(
    ClientCredential(SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET))

# In-memory data storage for demonstration purposes
data_store: List[Dict[str, Any]] = []

@app.get("/fetch-data")
def fetch_data_from_sandwai():
    headers = {
        'Authorization': f'Bearer {SANDWAI_API_KEY}'
    }
    response = requests.get(SANDWAI_API_URL, headers=headers)
    if response.status_code != 200:
        raise HTTPException(status_code=response.status_code, detail=response.json())
    
    global data_store
    data = response.json()
    data_store = data  # Update the in-memory data store
    df = pd.DataFrame(data)
    
    csv_path = 'data.csv'
    xlsx_path = 'data.xlsx'
    
    df.to_csv(csv_path, index=False)
    df.to_excel(xlsx_path, index=False)
    
    return {"message": "Data fetched and saved locally"}

@app.post("/upload-file")
def upload_file_to_sharepoint(file_format: str):
    if file_format not in ['csv', 'xlsx']:
        raise HTTPException(status_code=400, detail="Invalid file format")
    
    file_path = f'data.{file_format}'
    
    with open(file_path, 'rb') as file_content:
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/Shared Documents/{SHAREPOINT_DOC_LIBRARY}'
        ctx.web.get_folder_by_server_relative_url(target_folder_url).upload_file(file_path, file_content.read()).execute_query()
    
    return {"message": f"{file_format.upper()} file uploaded to SharePoint"}

@app.get("/data", response_model=List[Dict[str, Any]])
def get_all_data():
    return data_store

@app.get("/data/{item_id}", response_model=Dict[str, Any])
def get_data_item(item_id: int):
    item = next((item for item in data_store if item["id"] == item_id), None)
    if item is None:
        raise HTTPException(status_code=404, detail="Item not found")
    return item

@app.post("/data", response_model=Dict[str, Any])
def create_data_item(item: Dict[str, Any]):
    if "id" not in item:
        raise HTTPException(status_code=400, detail="Item must have an 'id' field")
    if any(existing_item["id"] == item["id"] for existing_item in data_store):
        raise HTTPException(status_code=400, detail="Item with this ID already exists")
    data_store.append(item)
    return item

@app.put("/data/{item_id}", response_model=Dict[str, Any])
def update_data_item(item_id: int, updated_item: Dict[str, Any]):
    for index, existing_item in enumerate(data_store):
        if existing_item["id"] == item_id:
            data_store[index] = updated_item
            return updated_item
    raise HTTPException(status_code=404, detail="Item not found")

@app.delete("/data/{item_id}", response_model=Dict[str, Any])
def delete_data_item(item_id: int):
    global data_store
    item = next((item for item in data_store if item["id"] == item_id), None)
    if item is None:
        raise HTTPException(status_code=404, detail="Item not found")
    data_store = [item for item in data_store if item["id"] != item_id]
    return item
