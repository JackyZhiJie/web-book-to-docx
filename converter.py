import win32com.client
import os
import time
import argparse
import urllib.request
from urllib.parse import urljoin, urlparse
from html.parser import HTMLParser

class LinkParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links = []
        
    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            for attr, val in attrs:
                if attr == 'href':
                    self.links.append(val)

def extract_chapter_links(url):
    print(f"Fetching main page to find chapters: {url}")
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    html = urllib.request.urlopen(req).read().decode('utf-8')
    
    parser = LinkParser()
    parser.feed(html)
    
    domain = urlparse(url).netloc
    chapter_urls = []
    
    for link in parser.links:
        full_url = urljoin(url, link)
        # Filter for links that are on the same domain and end with .html
        if urlparse(full_url).netloc == domain and full_url.endswith('.html'):
            if full_url not in chapter_urls:
                chapter_urls.append(full_url)
                
    return chapter_urls

def url_to_docx_via_word(url, output_filename):
    # First, get all the chapter links
    chapter_urls = extract_chapter_links(url)
    
    if not chapter_urls:
        print("No chapter sublinks found. Converting the single page instead.")
        chapter_urls = [url]
    else:
        print(f"Found {len(chapter_urls)} chapters to merge.")
    
    output_path = os.path.abspath(output_filename)
    word = None
    
    try:
        # Launch Word in the background
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False 
        
        # Create a master document to hold everything
        master_doc = word.Documents.Add()
        
        for i, chapter_url in enumerate(chapter_urls):
            print(f"[{i+1}/{len(chapter_urls)}] Downloading and merging: {chapter_url}")
            
            # Insert the chapter's HTML directly at the end of the master document
            rng = master_doc.Range(master_doc.Content.End - 1, master_doc.Content.End - 1)
            rng.InsertFile(chapter_url)
            
            # Insert a page break if it's not the last chapter
            if i < len(chapter_urls) - 1:
                rng_break = master_doc.Range(master_doc.Content.End - 1, master_doc.Content.End - 1)
                rng_break.InsertBreak(Type=7) # 7 is wdPageBreak
                
        # Save the final merged document
        print(f"Saving merged document to {output_path}...")
        master_doc.SaveAs(output_path, FileFormat=16) # 16 is .docx
        master_doc.Close(False)
        
        print(f"Success! Conversion complete. Saved to: {output_path}")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        
    finally:
        # Make sure we quit Word
        if word:
            word.Quit()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert a book URL (with chapters) to a merged DOCX file using Microsoft Word")
    parser.add_argument("url", help="The URL of the webpage to convert")
    parser.add_argument("-o", "--output", default="output.docx", help="The output DOCX filename (default: output.docx)")
    
    args = parser.parse_args()
    
    url_to_docx_via_word(args.url, args.output)
