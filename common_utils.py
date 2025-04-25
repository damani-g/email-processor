import re

# Utility functions for both .eml and .msg processing

# Function to remove invalid characters from filenames
def rem_inv_chars(filename):
    return re.sub(r'[<>:"/\\|?*]', '', filename)

# Function to extract sender/receiver name for PDF naming
def extract_emailer_name(email_string):
    if not email_string:
        return "No Name"
    
    first_emailer = email_string.split(';', 1)[0].strip()

    # Case 1: Name and email, e.g., "John Doe <john@company.com>"
    match_with_email = re.match(r"(.+?)\s*<[^<>]+>", first_emailer)
    if match_with_email:
        name_str = match_with_email.group(1).strip()
    # Case 2: Only email in angle brackets, e.g., "<john@company.com>"
    elif re.match(r"<[^<>]+>", first_emailer):
        non_inv_name = rem_inv_chars(first_emailer)
        return re.sub(r'@', '[at]', non_inv_name)
    # Case 3: Just a name, e.g., "John Doe"
    else:
        name_str = first_emailer.strip()

    parts = name_str.split()
    if len(parts) == 1:
        return parts[0]
    else:
        return f"{parts[0][0]} {parts[-1]}"
    
