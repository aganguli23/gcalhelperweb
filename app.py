import os
import json
import re
import warnings
import io
import contextlib
import datetime
import uuid
import webbrowser
from datetime import datetime, timezone, timedelta

from flask import Flask, render_template, request, redirect, url_for, flash, session
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

# Load environment variables from .env (for local development)
load_dotenv()

# Silence deprecation warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# External dependencies
from openai import OpenAI
from pvrecorder import PvRecorder
from playsound import playsound
from IPython.display import Image, display

# Google OAuth and Calendar imports
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# OCR and file conversion imports
from PIL import Image as PILImage
import pytesseract
from pdf2image import convert_from_path
from docx2pdf import convert

# Time zone imports
from tzlocal import get_localzone
from zoneinfo import ZoneInfo

# Flask app setup
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev_secret_key")
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'pdf', 'docx'}
SCOPES = ['https://www.googleapis.com/auth/calendar']

# --- Utility Functions ---

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Create credentials.json file programmatically if not present.
# (For local development only; on Heroku you will upload your downloaded credentials.json)
def create_credentials_file():
    if not os.path.exists('credentials.json'):
        google_client_id = os.environ.get("GOOGLE_CLIENT_ID")
        google_client_secret = os.environ.get("GOOGLE_CLIENT_SECRET")
        google_project_id = os.environ.get("GOOGLE_PROJECT_ID", "your_project_id")
        if not google_client_id or not google_client_secret:
            raise Exception("Google client ID and/or client secret not set in environment variables!")
        credentials_data = {
            "web": {
                "client_id": google_client_id,
                "project_id": google_project_id,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": google_client_secret,
                "redirect_uris": [
                    url_for('oauth2callback', _external=True)
                ]
            }
        }
        with open('credentials.json', 'w') as f:
            json.dump(credentials_data, f, indent=4)
        print("credentials.json created successfully!")
    else:
        print("credentials.json already exists.")

def credentials_to_dict(creds):
    return {
        'token': creds.token,
        'refresh_token': creds.refresh_token,
        'token_uri': creds.token_uri,
        'client_id': creds.client_id,
        'client_secret': creds.client_secret,
        'scopes': creds.scopes
    }

# --- OAuth2 Web Flow Routes ---

@app.route('/authorize')
def authorize():
    flow = Flow.from_client_secrets_file(
        'credentials.json',
        scopes=SCOPES,
        redirect_uri=url_for('oauth2callback', _external=True)
    )
    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true'
    )
    session['state'] = state
    return redirect(authorization_url)

@app.route('/oauth2callback')
def oauth2callback():
    state = session.get('state')
    flow = Flow.from_client_secrets_file(
        'credentials.json',
        scopes=SCOPES,
        state=state,
        flow.redirect_uri = url_for('authorize', _external=True)
    )
    flow.fetch_token(authorization_response=request.url)
    creds = flow.credentials
    session['credentials'] = credentials_to_dict(creds)
    # Optionally, write to token.json (note: Heroku’s filesystem is ephemeral)
    with open('token.json', 'w') as token_file:
        token_file.write(creds.to_json())
    flash("Google Calendar authenticated successfully.")
    return redirect(url_for('index'))

def get_credentials():
    if 'credentials' in session:
        creds = Credentials(**session['credentials'])
    elif os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        session['credentials'] = credentials_to_dict(creds)
    else:
        creds = None
    if not creds or not creds.valid:
        return None
    return creds

# --- OCR and Input Functions ---

def ocr_image(file_path_or_object, selected_pages=None):
    try:
        if isinstance(file_path_or_object, str):
            if file_path_or_object.lower().endswith(".pdf"):
                pdf_images = convert_from_path(file_path_or_object, dpi=300)
                if selected_pages:
                    filtered_images = []
                    for page_num in selected_pages:
                        if 1 <= page_num <= len(pdf_images):
                            filtered_images.append(pdf_images[page_num - 1])
                    pdf_images = filtered_images
                extracted_text = ""
                for page_number, image in enumerate(pdf_images, start=1):
                    print(f"Processing page {page_number} from PDF...")
                    page_text = pytesseract.image_to_string(image)
                    extracted_text += f"--- Page {page_number} ---\n{page_text}\n"
                return extracted_text
            elif file_path_or_object.lower().endswith(".docx"):
                temp_pdf = file_path_or_object.rsplit('.', 1)[0] + '_temp.pdf'
                try:
                    convert(file_path_or_object, temp_pdf)
                except Exception as e:
                    print(f"Error during DOCX conversion: {e}")
                    return ""
                pdf_images = convert_from_path(temp_pdf, dpi=300)
                if selected_pages:
                    filtered_images = []
                    for page_num in selected_pages:
                        if 1 <= page_num <= len(pdf_images):
                            filtered_images.append(pdf_images[page_num - 1])
                    pdf_images = filtered_images
                extracted_text = ""
                for page_number, image in enumerate(pdf_images, start=1):
                    page_text = pytesseract.image_to_string(image)
                    extracted_text += f"--- Page {page_number} ---\n{page_text}\n"
                os.remove(temp_pdf)
                return extracted_text
            else:
                image = PILImage.open(file_path_or_object)
                text = pytesseract.image_to_string(image)
                return text
        else:
            image = PILImage.open(file_path_or_object)
            text = pytesseract.image_to_string(image)
            return text
    except Exception as e:
        print("Error during OCR:", e)
        return ""

def combine_inputs(user_input, ocr_text):
    combined_parts = []
    if user_input:
        combined_parts.append(user_input)
    if ocr_text:
        combined_parts.append(ocr_text)
    return " ".join(combined_parts)

# --- GPT-4o Chatbot Class ---

class GPT4o:
    def __init__(self, client, json_file='gpt4oContext1.json'):
        self.client = client
        self.context = []
        self.json_file = json_file

    def chat(self, message, save=False):
        message = (message or "") + " "
        if not self.context:
            self.context.append({"role": "system", "content": "You are a helpful assistant."})
        self.context.append({"role": "user", "content": message})
        response = self.client.chat.completions.create(
            model="gpt-4o",
            messages=self.context
        )
        response_content = response.choices[0].message.content
        self.context.append({"role": "assistant", "content": response_content})
        self.save_to_json(message, response_content, save)
        if not save:
            json_files_to_clear = ['gpt4oContext1.json', 'gpt4oMiniContext1.json', 'gpt3pt5TurboContext1.json']
            self.clear_json_files(json_files_to_clear)
        self.print_response(response_content)
        return response_content

    def clear_json_files(self, json_files):
        for json_file in json_files:
            with open(json_file, 'w') as file:
                json.dump({}, file)
            print(f"The contents of {json_file} have been cleared.")

    def save_to_json(self, input_text, output_text, save):
        if save:
            try:
                with open(self.json_file, 'r') as file:
                    data = json.load(file)
            except FileNotFoundError:
                data = {}
            data[input_text] = output_text
            with open(self.json_file, 'w') as file:
                json.dump(data, file, indent=4)
            print(f"Data successfully saved to {self.json_file}.")
        else:
            print("Save flag is False. No data was saved.")

    def print_response(self, response_content):
        print(f'BOT: {response_content}')

    def print_full_chat(self):
        for message in self.context:
            if message["role"] == "user":
                print(f'USER: {message["content"]}')
            elif message["role"] == "assistant":
                print(f'BOT: {message["content"]}')
        if self.context:
            print("\nFINAL OUTPUT")
            print(f'BOT: {self.context[-1]["content"]}')

openai_api_key = os.environ.get("OPENAI_API_KEY")
if not openai_api_key:
    raise Exception("OPENAI_API_KEY not set in environment variables!")
client = OpenAI(api_key=openai_api_key)
gpt4o = GPT4o(client)

def get_gpt4o_response(input_text):
    try:
        print("Sending to GPT-4o:", input_text)
        response_content = gpt4o.chat(f"{input_text}")
        if not response_content:
            print("GPT-4o returned an empty response.")
            return None
        return response_content.strip()
    except Exception as e:
        print(f"Error during GPT-4o call: {e}")
        return None

def extract_code(response_text):
    code_match = re.search(r"```python\s*(.*?)\s*```", response_text, re.DOTALL)
    if code_match:
        return code_match.group(1).strip()
    return ""

# --- Flask Routes ---

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():
    user_input = request.form.get("text_input", "")
    ocr_text = ""
    selected_pages = []
    selected_pages_str = request.form.get("selected_pages", "")
    if selected_pages_str:
        try:
            selected_pages = [int(num.strip()) for num in selected_pages_str.split(",") if num.strip().isdigit()]
            if len(selected_pages) > 2:
                flash("Please select a maximum of 2 pages.")
                return redirect(url_for('index'))
        except ValueError:
            flash("Invalid page numbers entered.")
            return redirect(url_for('index'))
    file = request.files.get("file_upload")
    if file and file.filename != "":
        if not allowed_file(file.filename):
            flash("File type not allowed! Please upload an image, PDF, or DOCX file.")
            return redirect(url_for('index'))
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        ext = filename.rsplit('.', 1)[1].lower()
        if ext in ['pdf', 'docx']:
            ocr_text = ocr_image(file_path, selected_pages=selected_pages)
        else:
            with open(file_path, "rb") as f:
                ocr_text = ocr_image(f)
        os.remove(file_path)
    combined_input = combine_inputs(user_input, ocr_text)
    print("Combined Input:", combined_input)
    prompt = f"""
Generate a Python script to add all the calendar event(s) (that you identify in this user input text)
to Google Calendar: {combined_input}.
Carefully meet all the criteria and follow all the directions below:
All API setup has been completed and authentication is managed via OAuth2 using the "web" client credentials defined in credentials.json.
Ensure that the script utilizes the Google OAuth2 web flow for user authentication and stores tokens appropriately.
Do not use service account credentials, as those require fields (such as client_email and token_uri) which are not present in credentials.json.
Use the Google Calendar API and include proper timezone handling by
    1. Setting the Time in Local Timezone: The start_time is now set in the local timezone (local_tz) instead of UTC: start_time = datetime.combine(next_wednesday, datetime.min.time(), tzinfo=local_tz)
    2. Avoiding Unnecessary UTC Conversion: By setting the time directly in the local timezone, you avoid the need to convert from UTC to local time later.
    3. Ensuring Correct Time in Google Calendar: The create_calendar_event function already sends the local time and timezone to Google Calendar:
    'dateTime': start_local.isoformat(),
    'timeZone': str(local_tz)
    This ensures that the event is created at the correct time in your local timezone.

Ensure the year and month are correct. If not provided, extract from:
    from datetime import datetime, timezone
    now = datetime.now(timezone.utc).isoformat()
and convert to local time.
Ensure that event titles are human yet professional--short, concise, and descriptive.
Unless otherwise specified, include reminders at 10 minutes, 1 hour, and 1 day before as notifications.
If a Google Meet link is explicitly required and certain, include conferenceData with a createRequest (using a unique requestId and conferenceSolutionKey set as 'hangoutsMeet'), and when calling events.insert or events.update include conferenceDataVersion=1.
After event creation, use Python's webbrowser module to open the event link in the default browser.
At the end, include a summary of how many events were created along with additional details.

IMPORTANT: Only use the following external dependencies when generating the code. Do not include any libraries or modules outside this list (aside from Python's standard library):

Flask>=2.0.0  
gunicorn  
google-auth-oauthlib>=0.4.6  
google-api-python-client>=2.70.0  
google-auth>=2.3.3  
Pillow>=9.0.0  
pytesseract>=0.3.10  
openai  
pvrecorder  
playsound==1.2.2  
IPython  
pytz  
tzlocal  
pdf2image  
docx2pdf  
python-dotenv  
requests>=2.25.0  
httplib2>=0.20.0  
uritemplate>=3.0.1  
oauthlib>=3.1.0  
six>=1.15.0  
Jinja2>=3.0.0  
MarkupSafe>=2.0.0  
itsdangerous>=2.0.0  
click>=8.0.0

SYSTEM DEPENDENCIES (use Homebrew on macOS):
- Tesseract OCR (for pytesseract) → `brew install tesseract`
- Poppler (for pdf2image) → `brew install poppler`
- LibreOffice (for docx2pdf, if Microsoft Word is not available) → `brew install --cask libreoffice`

Do not include code that requires dependencies outside of these.
"""
    # Clear context files if they exist
    for file_name in ['gpt4oContext1.json', 'gpt4oMiniContext1.json']:
        if os.path.exists(file_name):
            with open(file_name, 'w') as f:
                json.dump({}, f)
    response_text = get_gpt4o_response(prompt)
    generated_code = extract_code(response_text)
    execution_output = ""
    if generated_code:
        try:
            f = io.StringIO()
            with contextlib.redirect_stdout(f):
                exec(generated_code, globals())
            execution_output = f.getvalue()
        except Exception as e:
            execution_output = f"Execution Error: {e}"
    return render_template("result.html",
                           combined_input=combined_input,
                           generated_code=generated_code,
                           execution_output=execution_output)

# --- Run the App ---
if __name__ == "__main__":
    create_credentials_file()
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))