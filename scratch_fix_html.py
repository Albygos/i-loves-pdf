import os
import glob

def fix_html_files():
    print("Fixing all HTML files in base and seo_pages directories...")
    
    html_files = glob.glob("*.html") + glob.glob("seo_pages/*.html")
    
    replacements = {
        # Fix back button syntax error
        'onclick="window.location.href="/""': 'onclick="window.location.href=\'/\'"',
        
        # Fix old tool links to clean Flask routes
        'href="/merge_pdf"': 'href="/merge"',
        'href="/split_pdf"': 'href="/split"',
        'href="/compress_pdf"': 'href="/compress"',
        'href="/rotate_pdf"': 'href="/rotate"',
        'href="/unlock_pdf"': 'href="/unlock"',
        'href="/pdf_add_watermark"': 'href="/watermark"',
        'href="/add_pdf_page_number"': 'href="/page-numbers"',
        'href="/remove-pages"': 'href="/delete-pages"',
        'href="/organize-pdf"': 'href="/organize"',
        'href="/protect-pdf"': 'href="/protect"',
        'href="/word_to_pdf"': 'href="/word-to-pdf"',
        'href="/pdf_to_word"': 'href="/pdf-to-word"',
        'href="/jpg_to_pdf"': 'href="/jpg-to-pdf"',
        'href="/pdf_to_jpg"': 'href="/pdf-to-jpg"'
    }
    
    count = 0
    for filepath in html_files:
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
                
            original_content = content
            for old, new in replacements.items():
                content = content.replace(old, new)
                
            if content != original_content:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(content)
                count += 1
                if count % 1000 == 0:
                    print(f"Fixed {count} files...")
        except Exception as e:
            print(f"Error processing {filepath}: {e}")
            
    print(f"Successfully fixed {count} files out of {len(html_files)}.")

if __name__ == '__main__':
    fix_html_files()
