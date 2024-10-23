import requests
from bs4 import BeautifulSoup
from docx import Document
from io import BytesIO
import streamlit as st
import time
from urllib.parse import urlparse
import re

# Function to scrape and save content to DOCX
def scrape_and_save_to_docx(url):
    # Add a delay between requests (e.g., 5 seconds)
    time.sleep(5)

    # Proxies (you can replace the IP with your actual proxy)
    proxies = {
        "http": "http://122.160.30.99:80",
        # "https": "https://your_proxy_ip:port",
    }

    # Headers to mimic a browser request
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:102.0) Gecko/20100101 Firefox/102.0"
    }

    # Fetch the webpage with the provided proxy and headers
    response = requests.get(url, headers=headers, proxies=proxies)

    # Check for response success
    if response.status_code != 200:
        st.error(f"Failed to retrieve the webpage. Status code: {response.status_code}")
        return None
    
    # Parse the webpage content using BeautifulSoup
    soup = BeautifulSoup(response.content, 'html.parser')
    doc = Document()

    # Extract the domain name or main part of the URL for the heading
    parsed_url = urlparse(url)
    domain = parsed_url.netloc if parsed_url.netloc else url

    # Clean up the domain name to create a valid filename (removes special characters)
    valid_filename = re.sub(r'[^a-zA-Z0-9_-]', '_', domain)

    # Add a custom heading to the DOCX document with the domain
    doc.add_heading(f'Content from {domain}', level=1)

    # Function to add different content types (headings, paragraphs, lists, etc.) to the DOCX document
    def add_content_to_docx(tag, doc):
        if tag.name in ['h1', 'h2', 'h3', 'h4']:
            level = int(tag.name[-1])
            doc.add_heading(tag.get_text(strip=True), level=level)
        elif tag.name == 'p':
            doc.add_paragraph(tag.get_text(strip=True))
        elif tag.name == 'ul':
            for li in tag.find_all('li'):
                doc.add_paragraph(f"• {li.get_text(strip=True)}", style='List Bullet')
        elif tag.name == 'ol':
            for i, li in enumerate(tag.find_all('li'), start=1):
                doc.add_paragraph(f"{i}. {li.get_text(strip=True)}", style='List Number')
        elif tag.name == 'img':
            doc.add_paragraph("[Image Placeholder: " + tag.get('src', 'Image URL') + "]")
        elif tag.name == 'audio':
            doc.add_paragraph("[Audio Placeholder: " + tag.get('src', 'Audio URL') + "]")
        elif tag.name == 'video':
            doc.add_paragraph("[Video Placeholder: " + tag.get('src', 'Video URL') + "]")

    # Scraping the relevant sections (header, body, footer) and extracting the desired tags
    sections = ['header', 'body', 'footer']
    for section in sections:
        section_tag = soup.find(section)
        if section_tag:
            for tag in section_tag.find_all(['h1', 'h2', 'h3', 'h4', 'p', 'ul', 'ol', 'img', 'audio', 'video']):
                add_content_to_docx(tag, doc)
    
    # Save the DOCX to a BytesIO object (avoids saving to a physical file)
    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)

    # Save the DOCX file with the domain in its name
    file_name = f"{valid_filename}.docx"
    
    return doc_io, file_name

# Function to display DOCX content on the screen
def display_docx_content(docx_io):
    document = Document(docx_io)
    content = ""
    
    for para in document.paragraphs:
        content += para.text + "\n"
    
    return content

# Streamlit app
st.title("Web Scraper to DOCX")

# Input for URL
url = st.text_input("Enter the URL of the webpage to scrape")

# Button to scrape and generate DOCX file
if st.button("Generate and View DOCX"):
    if url:
        with st.spinner("Scraping content and generating DOCX..."):
            docx_io, file_name = scrape_and_save_to_docx(url)
            
            if docx_io:
                # Display success message
                st.success(f"DOCX file '{file_name}' generated successfully!")

                # Display the DOCX content on the screen
                docx_content = display_docx_content(docx_io)
                st.text_area("DOCX Content", docx_content, height=300)
                
                # Allow user to download DOCX
                st.download_button(
                    label="Download DOCX",
                    data=docx_io,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    else:
        st.error("Please enter a valid URL.")