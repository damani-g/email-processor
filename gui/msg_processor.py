from common_utils import extract_emailer_name, rem_inv_chars
import os
import extract_msg
import mimetypes
import re
import sys
import pdfkit

from bs4 import BeautifulSoup

# Function to create list of attachments with no content ID (so that output excludes signature images)
def no_cid_attachment_list(msg_attachments):
    attachment_list = []
    for attachment in msg_attachments:
        if attachment.cid == None:
            attachment_list.append(attachment.longFilename)
    return attachment_list

# Function to print attachment list for use in attachments_html
def make_attachment_list(attachment_list):
    if len(attachment_list) >= 1:
        return ", ".join(attachment_list)
    else:
        return "None"

# Function to create html block for insertion of attachment info into PDF
def create_attachments_html(attachment_list):
    raw_html = f"""<br/>
    <b>
    Attachments:
    </b>
    {make_attachment_list(attachment_list)}
    <br/><br/> 
    """
    parsed_html = BeautifulSoup(raw_html, "html.parser")

    return parsed_html

# Function to insert html block for attachment info into html for entire PDF
def insert_attachments_html(soup, attachments_html):
    # Find <p> tag containing email info
    email_info_p = soup.find("p", class_ = "MsoNormal")
    # Insert as last element before <o:p></o:p>
    email_info_p.contents[-1].insert_before(attachments_html)

# Function to retrieve datetime from .msg
def msg_get_date(filepath):
    msg = extract_msg.Message(filepath)
    return msg.date

# Main function called to process .msg
def msg_to_pdf(filepath, output_folder, index, wkhtmltopdf_path, log_callback=None):
    # Store Message object in variable for processing 
    msg = extract_msg.Message(filepath)

    # Extract email as html and convert to BeautifulSoup object
    msg_body = msg.getSaveHtmlBody(preparedHtml=True)
    soup = BeautifulSoup(msg_body, "html.parser")

    # Create attachment list and insert as html into BeautifulSoup object
    attachment_list = no_cid_attachment_list(msg.attachments)
    attachments_html = create_attachments_html(attachment_list)
    insert_attachments_html(soup,attachments_html)

    # Extract information to create PDF name
    subject = msg.subject
    sendername = extract_emailer_name(msg.sender)
    receivername = extract_emailer_name(msg.to)
    date = msg.date
    time = date.strftime("%I.%M %p")

    pdf_name = f'{index}. {sendername} to {receivername}-{date.day}.{date.month}.{date.year}-{subject}-{time}'
    pdf_name = rem_inv_chars(pdf_name)

    # Convert BeautifulSoup object into PDF and save to Output folder
    pdf_path = os.path.join(output_folder, f'{pdf_name}.pdf')
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
    pdfkit.from_string(str(soup), pdf_path,configuration=config, options={ 'page-size': 'Letter'})
    log_callback(f'Saving PDF: {pdf_name}')

# Function to extract attachments from .msg file
def msg_extract_attachments(filepath, output_folder, index, log_callback=None):
    # Store Message object in variable for processing 
    msg = extract_msg.Message(filepath)

    # Loop through attachments, creating name and saving each one
    for attachment_index, attachment in enumerate(msg.attachments, start=1):
        # If attachment is an email
        if attachment.cid == None:
            attachment_filename = attachment.getFilename() 
            attachment_filename = rem_inv_chars(attachment_filename)
            attachment_name = f"{index}{chr(96 + attachment_index)}. {attachment_filename}"
            log_callback(f'Saving attachment: {attachment_name}')
            attachment.save(customPath=output_folder, extractEmbedded=True, customFilename=attachment_name)
