
import docx
import os
import sys

# Ensure stdout uses utf-8
sys.stdout.reconfigure(encoding='utf-8')

def extract_hyperlinks_and_text(file_path):
    doc = docx.Document(file_path)
    content = []
    
    # Iterate through paragraphs
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        links = []
        para_links = []
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        
        # We need to look into the paragraph xml
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        hyperlink_tags = para._element.findall('.//w:hyperlink', ns)
        for tag in hyperlink_tags:
            r_id = tag.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if r_id and r_id in doc.part.rels:
                rel = doc.part.rels[r_id]
                if rel.is_external:
                    para_links.append(rel.target_ref)
        
        # Also check for direct text that looks like a URL if para_links is empty?
        # The user provided text sometimes has "Read More" as the link anchor.
        
        content.append({
            "text": text,
            "links": para_links
        })

    return content

file_path = r"c:\Users\ASUS\Desktop\UIU Newsletter\January Newsletter\January Newsletter Content.docx"
output_path = r"c:\Users\ASUS\Desktop\UIU Newsletter\extracted_links_utf8.txt"

with open(output_path, 'w', encoding='utf-8') as f:
    try:
        data = extract_hyperlinks_and_text(file_path)
        for item in data:
            f.write(f"Text: {item['text'][:50]}...\n")
            if item['links']:
                f.write(f"Links: {item['links']}\n")
            f.write("-" * 20 + "\n")
    except Exception as e:
        f.write(f"Error: {e}\n")
