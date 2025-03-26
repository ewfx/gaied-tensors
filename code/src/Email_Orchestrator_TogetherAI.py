import os
import json
import email
from email import policy
from email.parser import BytesParser
import fitz  # PyMuPDF
import docx
import pytesseract
from PIL import Image
from together import Together
import extract_msg
from flask import Flask, request, jsonify
import email
import base64
import uuid

app = Flask(__name__)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "pdf", "txt", "docx", "eml", "msg"}
source_directory = "./temp/"

# Load configuration
config_path = './gaied-tensors/code/src/config.json'
if os.path.getsize(config_path) > 0:
    with open(config_path, 'r') as f:
        config = json.load(f)
else:
    raise ValueError("The configuration file is empty or not found.")

# Initialize the Together client
client = Together()

# Function to read and parse emails
def read_eml(file_path):
    with open(file_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    return msg

# Function to read .msg files
def read_msg(file_path):
    msg = extract_msg.Message(file_path)
    # msg.load()
    return msg

# Function to read PDF files
def read_pdf(file_path):
    doc = fitz.open(file_path)
    text = ""
    for page in doc:
        text += page.get_text()
        # If no text is found, use OCR
        if not text.strip():
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            text += pytesseract.image_to_string(img)
    return text

# Function to read DOC files
def read_doc(file_path):
    doc = docx.Document(file_path)
    text = ""
    for para in doc.paragraphs:
        text += para.text
    # If no text is found, use OCR on images in the document
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img = Image.open(rel.target_part.blob)
            text += pytesseract.image_to_string(img)
    return text

# Function to read image files
def read_image(file_path):
    img = Image.open(file_path)
    text = pytesseract.image_to_string(img)
    return text

# Function to read text files
def read_txt(file_path):
    with open(file_path, 'r') as f:
        text = f.read()
    return text



def extract_attachments(eml_file_path, output_folder):
    # Ensure the output directory exists
    os.makedirs(output_folder, exist_ok=True)

    # Read the .eml file
    with open(eml_file_path, 'rb') as f:
        msg = email.message_from_bytes(f.read())

    # Iterate through email parts
    for part in msg.walk():
        # Check if it's an attachment
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            if filename:
                file_path = os.path.join(output_folder, filename)
                
                # Write the attachment to file
                with open(file_path, 'wb') as attachment_file:
                    attachment_file.write(part.get_payload(decode=True))
                
                print(f"Attachment saved: {file_path}")




# Function to classify and extract information using Together API
def classify_and_extract(content, config):
    request_types = "\n".join([f"- {req_type}: {', '.join(sub_types)}" for req_type, sub_types in config["request_types"].items()])
    fields_to_extract = "\n".join([f"- {field}" for field in config["fields_to_extract"].values()])
    
    prompt = f"""
    You are an AI assistant. The following is an email/document content. Based on the provided configuration, classify the content into request types and sub-request types, and extract relevant information.

    Configuration:
    Request Types and Sub Request Types:
    {request_types}

    Fields to Extract:
    {fields_to_extract}

    Content: {content}

    Please provide the classification and extracted information in the following format without using comments or explaination or double slashes:
    {{
        "request_type": "Request Type",
        "sub_request_type": "Sub Request Type",
        "confidence_score": "Confidence Score",
        "reasoning": "Reasoning",
        "extracted_information": {{
            "field1": "value1",
            "field2": "value2",
            ...
        }}
    }}
    """
    
    # Execute the Request
    try:
        response = client.chat.completions.create(
            model="meta-llama/Llama-Vision-Free",
            messages=[{"role": "user", "content": prompt}],
        )
        print("API Response:", response)
        
        # Handle the response
        result = response.choices[0].message.content.strip()
        return json.loads(result)
    except Exception as error:
        print("Error:", error)
        return {}
    except (KeyError, IndexError, TypeError) as e:
        print("Error processing response:", e)
        return {}

# Function to detect duplicates
def detect_duplicates(email_list):
    # Placeholder for duplicate detection logic
    duplicates = []
    return duplicates

# Main function to process files
def process_files(file_path, config):
    filename = os.path.basename(file_path)
    processed_files = []
    # for filename in os.listdir(source_dir):
    # file_path = os.path.join(source_dir, filename)
    if file_path.endswith(".eml"):
        email_content = read_eml(file_path)
        email_body = email_content.get_body(preferencelist=('plain', 'html')).get_content()
        new_guid = source_directory+uuid.uuid4().hex
        extract_attachments(file_path, new_guid)
        attachments_content = ""
        folder_path = new_guid
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(file_path):  # Check if it's a file
                if file_path.endswith(".pdf"):
                    attachments_content += read_pdf(file_path)
                elif file_path.endswith((".docx",".doc")):
                    attachments_content += read_doc(file_path)
                elif file_path.endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
                    attachments_content += read_image(file_path)
                elif file_path.endswith(".txt"):
                    attachments_content += read_txt(file_path)
            # Add more conditions for other attachment types if needed
        full_content = email_body + "\n" + attachments_content
        result = classify_and_extract(full_content, config)
        return {
            "filename": filename,
            "result": result
        }
    if file_path.endswith(".msg"):
        email_content = read_msg(file_path)
        email_body = email_content.body
        new_guid = source_directory+uuid.uuid4().hex
        for attachment in email_content.attachments:
            os.makedirs(new_guid, exist_ok=True)
            attachment.save(customPath=new_guid)  # Save attachments to 'attachments/' folder
            print(f"Attachment saved: {attachment.longFilename}")
        attachments_content = ""
        folder_path = new_guid
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(file_path):  # Check if it's a file
                if file_path.endswith(".pdf"):
                    attachments_content += read_pdf(file_path)
                elif file_path.endswith((".docx",".doc")):
                    attachments_content += read_doc(file_path)
                elif file_path.endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
                    attachments_content += read_image(file_path)
                elif file_path.endswith(".txt"):
                    attachments_content += read_txt(file_path)
            # Add more conditions for other attachment types if needed
        full_content = email_body + "\n" + attachments_content
        result = classify_and_extract(full_content, config)
        return {
            "filename": filename,
            "result": result
        }
    
    if file_path.endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
        email_body = read_image(file_path)
        full_content = email_body + "\n" + attachments_content
        result = classify_and_extract(full_content, config)
        return {
            "filename": filename,
            "result": result
        }
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/home")
def home():
    return jsonify({"message": "Welcome to Flask API!"})

# POST method for file upload
@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    if allowed_file(file.filename):
        file_path = os.path.join(source_directory, file.filename)
        file.save(file_path)
        output= process_files(file_path, config)
        return jsonify(output), 200
    else:
        return jsonify({"error": "File type not allowed"}), 400
    
if __name__ == "__main__":
    app.run(debug=True)

