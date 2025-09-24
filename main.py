from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from googleapiclient.discovery import build
from google.oauth2 import service_account
import os
import json

# ----------------- CONFIG -----------------
SCOPES = ['https://www.googleapis.com/auth/presentations']

# Read service account credentials from environment variable
creds_json = os.environ.get("GOOGLE_CREDS_JSON")
if not creds_json:
    raise ValueError("Missing GOOGLE_CREDS_JSON environment variable!")
creds_info = json.loads(creds_json)
creds = service_account.Credentials.from_service_account_info(creds_info, scopes=SCOPES)

# Read presentation ID from environment variable
PRESENTATION_ID = os.environ.get("PRESENTATION_ID")
if not PRESENTATION_ID:
    raise ValueError("Missing PRESENTATION_ID environment variable!")

# ----------------- SETUP -----------------
slides_service = build('slides', 'v1', credentials=creds)
app = FastAPI()

origins = ["https://json-gen.onrender.com"]

# Allow CORS for frontend domain (adjust to your deployed frontend URL)
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,  # Change "*" to your frontend URL in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ----------------- HELPER -----------------
def build_replacement_map(data: dict):
    # Weekly keys
    weekly_keys = [
        'week_number', 'week_suffix', 'BN_offering',
        'MN_offering', 'PN_offering', 'BN_SundayS', 'MN_SundayS'
    ]
    replacement_map = {k: str(data.get(k, '')) for k in weekly_keys}

    # Songs
    songs = data.get('songs', {})
    for song_key, lyrics in songs.items():
        if isinstance(lyrics, dict):
            replacement_map[song_key] = lyrics.get('main', '')
            replacement_map[f"{song_key}_eng"] = lyrics.get('eng', '')

    return replacement_map

# ----------------- ENDPOINTS -----------------
@app.post("/update-slides")
async def update_slides(request: Request):
    try:
        data = await request.json()
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid JSON: {e}")

    replacement_map = build_replacement_map(data)

    # Fetch current presentation
    presentation = slides_service.presentations().get(
        presentationId=PRESENTATION_ID
    ).execute()

    requests_list = []

    # Iterate slides and elements
    for slide in presentation.get('slides', []):
        for element in slide.get('pageElements', []):
            shape = element.get('shape')
            if not shape:
                continue
            alt_desc = element.get('description', '')
            if not alt_desc:
                continue
            if alt_desc in replacement_map:
                text = replacement_map[alt_desc]
                object_id = element['objectId']
                # Delete existing text
                requests_list.append({
                    'deleteText': {'objectId': object_id, 'textRange': {'type': 'ALL'}}
                })
                # Insert new text
                requests_list.append({
                    'insertText': {'objectId': object_id, 'insertionIndex': 0, 'text': text}
                })

    if requests_list:
        slides_service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={'requests': requests_list}
        ).execute()
        return {"status": "success", "message": "Slides updated successfully!"}
    else:
        return {"status": "warning", "message": "No matching alt text found in slides."}


@app.post("/reset-slides")
async def reset_slides():
    # Weekly keys
    weekly_keys = [
        'week_number', 'week_suffix', 'BN_offering',
        'MN_offering', 'PN_offering', 'BN_SundayS', 'MN_SundayS'
    ]

    # Song keys (adjust max number if needed)
    song_keys = [f"song_{i}" for i in range(1, 11)]

    requests_list = []

    # Reset weekly info
    for key in weekly_keys:
        requests_list.append({
            'replaceAllText': {
                'containsText': {'text': key, 'matchCase': True},
                'replaceText': key
            }
        })

    # Reset songs and English
    for song_key in song_keys:
        requests_list.append({
            'replaceAllText': {
                'containsText': {'text': song_key, 'matchCase': True},
                'replaceText': song_key
            }
        })
        eng_key = f"{song_key}_eng"
        requests_list.append({
            'replaceAllText': {
                'containsText': {'text': eng_key, 'matchCase': True},
                'replaceText': eng_key
            }
        })

    try:
        slides_service.presentations().batchUpdate(
            presentationId=PRESENTATION_ID,
            body={'requests': requests_list}
        ).execute()
        return {"status": "success", "message": "Template placeholders reset successfully!"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to reset template: {e}")
