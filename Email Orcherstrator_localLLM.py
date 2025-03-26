import requests
import json
import os
import email
from email import policy
from email.parser import BytesParser
import fitz  # PyMuPDF
import docx
import pytesseract
from PIL import Image
import extract_msg

# Load configuration
config_path = 'C:/Users/subba/Hackathon2025/gaied-tensors/code/src/config.json'
if os.path.getsize(config_path) > 0:
    with open(config_path, 'r') as f:
        config = json.load(f)
else:
    raise ValueError("The configuration file is empty or not found.")

# Function to read and parse emails
def read_eml(file_path):
    with open(file_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    return msg

# Function to read .msg files
def read_msg(file_path):
    msg = extract_msg.Message(file_path)
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

# Function to classify and extract information using the LLM API
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

    Please provide the classification and extracted information in the following format without using comments or double slashes:
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
    
    # Set up the base URL for the local Ollama API
    url = "http://localhost:11434/api/chat"

    # Define the payload (your input prompt)
    payload = {
        "model": "phi3",  # Replace with the model name you're using
        "messages": [{"role": "user", "content": prompt}]
    }

    # Send the HTTP POST request with streaming enabled
    response = requests.post(url, json=payload, stream=True)

    # Check the response status
    if response.status_code == 200:
        result = ""
        for line in response.iter_lines(decode_unicode=True):
            if line:  # Ignore empty lines
                try:
                    # Parse each line as a JSON object
                    json_data = json.loads(line)
                    # Extract and print the assistant's message content
                    if "message" in json_data and "content" in json_data["message"]:
                        result += json_data["message"]["content"]
                except json.JSONDecodeError:
                    print(f"\nFailed to parse line: {line}")
        return json.loads(result)
    else:
        print(f"Error: {response.status_code}")
        print(response.text)
        return {}

# Function to detect duplicates
def detect_duplicates(email_list):
    # Placeholder for duplicate detection logic
    duplicates = []
    return duplicates

# Main function to process files
def process_files(source_dir, config):
    processed_files = []
    for filename in os.listdir(source_dir):
        file_path = os.path.join(source_dir, filename)
        if filename.endswith(".eml"):
            email_content = read_eml(file_path)
            email_body = email_content.get_body(preferencelist=('plain', 'html')).get_content()
            attachments_content = ""
            for part in email_content.iter_attachments():
                content_type = part.get_content_type()
                if content_type == 'application/pdf':
                    attachments_content += read_pdf(part.get_payload(decode=True))
                elif content_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                    attachments_content += read_doc(part.get_payload(decode=True))
                elif content_type.startswith('image/'):
                    attachments_content += read_image(part.get_payload(decode=True))
                elif content_type == 'text/plain':
                    attachments_content += read_txt(part.get_payload(decode=True))
                # Add more conditions for other attachment types if needed
            full_content = email_body + "\n" + attachments_content
            result = classify_and_extract(full_content, config)
            processed_files.append({
                "filename": filename,
                "result": result
            })
        elif filename.endswith(".msg"):
            msg_content = read_msg(file_path)
            email_body = msg_content.body
            attachments_content = ""
            for attachment in msg_content.attachments:
                if attachment.longFilename.endswith(".pdf"):
                    attachments_content += read_pdf(attachment.data)
                elif attachment.longFilename.endswith(".docx"):
                    attachments_content += read_doc(attachment.data)
                elif attachment.longFilename.endswith(".txt"):
                    attachments_content += read_txt(attachment.data)
                elif attachment.longFilename.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
                    attachments_content += read_image(attachment.data)
                # Add more conditions for other attachment types if needed
            full_content = email_body + "\n" + attachments_content
            result = classify_and_extract(full_content, config)
            processed_files.append({
                "filename": filename,
                "result": result
            })
        elif filename.endswith(".pdf"):
            pdf_content = read_pdf(file_path)
            result = classify_and_extract(pdf_content, config)
            processed_files.append({
                "filename": filename,
                "result": result
            })
        elif filename.endswith(".docx"):
            doc_content = read_doc(file_path)
            result = classify_and_extract(doc_content, config)
            processed_files.append({
                "filename": filename,
                "result": result
            })
        elif filename.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
            image_content = read_image(file_path)
            result = classify_and_extract(image_content, config)
            processed_files.append({
                "filename": filename,
                "result": result
            })
        elif filename.endswith(".txt"):
            txt_content = read_txt(file_path)
            result = classify_and_extract(txt_content, config)
            processed_files.append({
                "filename": filename,
                "result": result
            })
    duplicates = detect_duplicates(processed_files)
    return processed_files, duplicates

if __name__ == "__main__":
    source_directory = "C:/Users/subba/Downloads/Hackathon Data"
    processed_files, duplicate_files = process_files(source_directory, config)
    
    # Output the results in JSON format
    output = []
    for file in processed_files:
        output.append({
            "filename": file["filename"],
            "request_type": file["result"].get("request_type", ""),
            "sub_request_type": file["result"].get("sub_request_type", ""),
            "confidence_score": file["result"].get("confidence_score", ""),
            "reasoning": file["result"].get("reasoning", ""),
            "extracted_information": file["result"].get("extracted_information", {})
        })
    
    with open('output.json', 'w') as f:
        json.dump(output, f, indent=4)
    
    print("Processed Files:", processed_files)
    print("Duplicate Files:", duplicate_files)