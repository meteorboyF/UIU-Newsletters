
import docx
import os

def extract_hyperlinks_and_text(file_path):
    doc = docx.Document(file_path)
    content = []
    
    # Iterate through paragraphs
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
        links = []
        # Access the xml to find hyperlinks since python-docx doesn't directly expose them easily in all versions
        # But let's try a simpler approach first or parse the relationships.
        
        # Actually, python-docx doesn't make extracting hyperlinks super specific to the text easy without xml parsing.
        # Let's try to just dump relationships and see if we can map them, or use a specific function for xml.
        
        # A workaround to get hyperlinks from a paragraph
        rels = doc.part.rels
        for rel in rels.values():
            if "hyperlink" in rel.reltype:
                 links.append(rel.target_ref)
                 
        # This gives ALL links in the doc. We need them associated with text.
        # Let's try to parse the XML of the paragraph.
        
        para_links = []
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        
        # We need to look into the paragraph xml
        # namespace map
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        hyperlink_tags = para._element.findall('.//w:hyperlink', ns)
        for tag in hyperlink_tags:
            r_id = tag.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if r_id and r_id in doc.part.rels:
                rel = doc.part.rels[r_id]
                if rel.is_external:
                    para_links.append(rel.target_ref)
        
        content.append({
            "text": text,
            "links": para_links
        })

    return content

file_path = r"c:\Users\ASUS\Desktop\UIU Newsletter\January Newsletter\January Newsletter Content.docx"
try:
    data = extract_hyperlinks_and_text(file_path)
    for item in data:
        print(f"Text: {item['text'][:50]}...")
        if item['links']:
            print(f"Links: {item['links']}")
        print("-" * 20)
except Exception as e:
    print(f"Error: {e}")
