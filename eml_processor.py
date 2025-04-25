from common_utils import extract_emailer_name, rem_inv_chars
import email
import email.policy
from bs4 import BeautifulSoup
import base64
import pdfkit
import os
from io import BytesIO
from email.generator import BytesGenerator

# Functions specific to .eml processing

# Function to create attachment list from EmailMessage Object
def create_attachment_list(msg):
    attachment_list = []
    for attachment in msg.iter_attachments():
        attachment_list.append(attachment.get_filename())
    
    if len(attachment_list) >= 1:
        return ", ".join(attachment_list)
    else:
        return "None"
    
# Function to insert header info into html for entire PDF
def insert_header_info(soup, attachments_html):
    # Find div containing email body
    html_body_div = soup.find("div", class_ = "WordSection1")
    # Insert before body
    html_body_div.insert_before(attachments_html)

# Function to recursively find the first text/html part
def find_html_part(part):
    if part.get_content_type() == "text/html":
        return part.get_content()
    if part.is_multipart():
        for subpart in part.iter_parts():
            result = find_html_part(subpart)
            if result:
                return result
    return None

# Function to extract html body from EmailMessage object
def extract_html_from_email(msg):
    html = find_html_part(msg)
    if html is None:
        raise ValueError("No text/html part found anywhere in message.")
    return html

# Function to retrieve datetime from .eml
def eml_get_date(filepath):
    with open(filepath, "r") as f:
        msg = email.message_from_file(f, policy=email.policy.default)
    date = msg.get("Date").datetime.astimezone()
    return date

# Main function called to process eml
def eml_to_pdf(filepath, output_folder, index, wkhtmltopdf_path, log_callback=None):
    # Open .eml file and set EmailMessage object as variable
    with open(filepath, "r") as f:
        msg = email.message_from_file(f, policy=email.policy.default)
    
    # Get HTML content
    html_content = extract_html_from_email(msg)
    soup = BeautifulSoup(html_content, "html.parser")

    # Replace cid: images with data URIs
    for part in msg.walk():
        content_id = part.get("Content-ID")
        if part.get_content_maintype() == "image" and content_id:
            cid = content_id.strip("<>")
            img_data = part.get_payload(decode=True)
            content_type = part.get_content_type()
            b64_data = base64.b64encode(img_data).decode("utf-8")
            data_uri = f"data:{content_type};base64,{b64_data}"

            for img_tag in soup.find_all("img", src=f"cid:{cid}"):
                img_tag["src"] = data_uri

    # Extract info from EmailMessage object to create PDF name
    sender_info = msg.get("From")
    receiver_info = msg.get("To")
    subject = msg.get("Subject")
    cc_info = msg.get("CC")
    date = msg.get("Date").datetime.astimezone()
    time = date.strftime("%I.%M %p")

    sender = extract_emailer_name(sender_info)
    receiver = extract_emailer_name(receiver_info)
    
    # Insert header info
    raw_html_header = f"""
        <div>
        <p class="MsoNormal">
        <b>
        From:
        </b>
        {sender_info}
        <br/>
        <b>
        Sent:
        </b>
        {date}
        <br/>
        <b>
        To:
        </b>
        {receiver_info}
        <br/>
        <b>
        Cc:
        </b>
        {cc_info}
        <br/>
        <b>
        Subject:
        </b>
        {subject}
        <br/>
        <b>
        Attachments:
        </b>
        {create_attachment_list(msg)}
        <br/><br/> 
        </p>
        <div>
    """

    parsed_html_header = BeautifulSoup(raw_html_header, "html.parser")

    insert_header_info(soup, parsed_html_header)

    # Configure wkhtmltopdf
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

    # Convert HTML to PDF
    pdfname = f'{index}. {sender} to {receiver}-{date.day}.{date.month}.{date.year}-{subject}-{time}.pdf'
    pdfname = rem_inv_chars(pdfname)
    pdfpath = os.path.join(output_folder, pdfname)
    pdfkit.from_string(str(soup), pdfpath, configuration=config, options={ 'page-size': 'Letter'})
    log_callback(f'Saving PDF: {pdfname}')

# Function to extract attachments from .eml file
def eml_extract_attachments(filepath, output_folder, index, log_callback=None):
    # Open .eml file and set EmailMessage object as variable
    with open(filepath, "r") as f:
        msg = email.message_from_file(f, policy=email.policy.default)
    
    # Loop through attachments, creating name and saving each one
    for attachment_index, attachment in enumerate(msg.iter_attachments(), start=1):
        attachment_prefix = f'{index}{chr(96 + attachment_index)}'
        # If attachment is a .msg file
        if attachment.get_content_type() == "message/rfc822":
            log_callback("!!! Problematic Attachment - Manual Resave Required !!!")
            subject = attachment.get_payload()[0].get("Subject", failobj = "Untitled Email")
            filename = f'{attachment_prefix}. {subject} (RESAVE IN OUTLOOK).eml' 

            # Prepare output buffer and use BytesGenerator to serialize it
            buffer = BytesIO()
            gen = BytesGenerator(buffer, policy = email.policy.compat32)
            gen.flatten(attachment.get_payload()[0])  # Write the full message to the buffer
            raw_bytes = buffer.getvalue()
            
            attachmentpath = os.path.join(output_folder, filename)
            with open(attachmentpath, "wb") as f:
                f.write(raw_bytes)
                log_callback(f'Saved {filename}')
                log_callback(f'Remember to open {filepath} in Outlook and resave attachment {subject}.msg')  

        else:
            filename = f'{attachment_prefix}. {attachment.get_filename()}'
            payload = attachment.get_payload(decode=True)
            attachmentpath = os.path.join(output_folder, filename)
            with open(attachmentpath, "wb") as f:
                f.write(payload)
                log_callback(f'Saved {filename}')