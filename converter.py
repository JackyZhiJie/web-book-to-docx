import win32com.client  # Used for COM automation to interact with Microsoft Word
import os  # Used for file path manipulations
import time  # Used for potential sleep/delay operations
import argparse  # Used for parsing command-line arguments
import urllib.request  # Used for making HTTP requests to fetch web pages
from urllib.parse import urljoin, urlparse  # Used for parsing and constructing URLs
from html.parser import HTMLParser  # Used for parsing HTML content to find links

class LinkParser(HTMLParser):
    """
    A custom HTML parser that extracts all 'href' links from '<a>' tags in an HTML document.
    """
    def __init__(self):
        super().__init__()
        self.links = []  # List to store the extracted URLs
        
    def handle_starttag(self, tag, attrs):
        """
        Called automatically by HTMLParser whenever an opening tag (like <a>) is encountered.
        """
        if tag == 'a':
            for attr, val in attrs:
                if attr == 'href':
                    self.links.append(val)  # Save the href link

def extract_chapter_links(url):
    """
    Fetches the main webpage at the given URL and extracts all internal links 
    that likely represent chapters of the book.
    """
    print(f"Fetching main page to find chapters: {url}")
    
    # Send an HTTP request masquerading as a web browser (Mozilla/5.0) to avoid being blocked
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    # Read and decode the HTML response
    html = urllib.request.urlopen(req).read().decode('utf-8')
    
    # Initialize our custom parser and feed it the HTML content
    parser = LinkParser()
    parser.feed(html)
    
    # Extract the domain name from the main URL to ensure we stay on the same site
    domain = urlparse(url).netloc
    chapter_urls = []
    
    # Process each link found by the parser
    for link in parser.links:
        # Convert relative links (e.g., '/chapter1.html') into full URLs
        full_url = urljoin(url, link)
        
        # Filter for links that are on the same domain and end with .html
        if urlparse(full_url).netloc == domain and full_url.endswith('.html'):
            # Avoid duplicate links
            if full_url not in chapter_urls:
                chapter_urls.append(full_url)
                
    return chapter_urls

def url_to_docx_via_word(url, output_filename):
    """
    Downloads the chapters of a web book and merges them into a single Microsoft Word document.
    """
    # First, get all the chapter links from the main index page
    chapter_urls = extract_chapter_links(url)
    
    # If no valid chapter links were found, just convert the provided URL itself
    if not chapter_urls:
        print("No chapter sublinks found. Converting the single page instead.")
        chapter_urls = [url]
    else:
        print(f"Found {len(chapter_urls)} chapters to merge.")
    
    # Get the absolute (full) file path for saving the document later
    output_path = os.path.abspath(output_filename)
    word = None
    
    try:
        # Launch Microsoft Word in the background via COM automation
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Keep the Word window hidden from the user
        
        # Create a new, blank master document to hold everything
        master_doc = word.Documents.Add()
        
        # Iterate over all chapter URLs
        for i, chapter_url in enumerate(chapter_urls):
            print(f"[{i+1}/{len(chapter_urls)}] Downloading and merging: {chapter_url}")
            
            # Create a Range object pointing to the very end of the master document
            rng = master_doc.Range(master_doc.Content.End - 1, master_doc.Content.End - 1)
            
            # Instruct Word to download and render the HTML directly into the document.
            # This preserves styling, images, and layout natively.
            rng.InsertFile(chapter_url)
            
            # Insert a page break if it's not the final chapter, so each chapter starts on a new page
            if i < len(chapter_urls) - 1:
                rng_break = master_doc.Range(master_doc.Content.End - 1, master_doc.Content.End - 1)
                rng_break.InsertBreak(Type=7) # 7 represents wdPageBreak in Word COM
                
        # Save the final merged document
        print(f"Saving merged document to {output_path}...")
        # Save as format 16, which corresponds to the standard .docx format
        master_doc.SaveAs(output_path, FileFormat=16) 
        master_doc.Close(False)  # Close the document without prompting to save changes
        
        print(f"Success! Conversion complete. Saved to: {output_path}")
        
    except Exception as e:
        # Catch and display any errors (like network failures or Word crashes)
        print(f"An error occurred: {e}")
        
    finally:
        # A finally block ensures Word is closed even if an error occurs,
        # preventing hidden Word processes from staying open forever in the background.
        if word:
            word.Quit()

# This block ensures the following code only runs if the script is executed directly 
# from the command line, and not if it's imported as a module in another script.
if __name__ == "__main__":
    # Set up the command-line argument parser
    parser = argparse.ArgumentParser(description="Convert a book URL (with chapters) to a merged DOCX file using Microsoft Word")
    
    # Define a required 'url' positional argument
    parser.add_argument("url", help="The URL of the webpage to convert")
    
    # Define an optional '-o' or '--output' argument for the destination filename
    parser.add_argument("-o", "--output", default="output.docx", help="The output DOCX filename (default: output.docx)")
    
    # Parse the arguments provided by the user in the terminal
    args = parser.parse_args()
    
    # Start the conversion process using the provided arguments
    url_to_docx_via_word(args.url, args.output)
