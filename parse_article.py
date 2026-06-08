import re
from html.parser import HTMLParser

class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.recording = False
        self.text_data = []
        self.in_content = False

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        # Main content container in Bangkok Biz News is usually content-detail or article body
        if tag == 'div' and ('content-detail' in attrs_dict.get('class', '') or 'article-wrapper' in attrs_dict.get('class', '')):
            self.in_content = True
        
        # We want to record text inside paragraphs, headings, and lists
        if tag in ['p', 'h1', 'h2', 'h3', 'h4', 'li', 'strong']:
            self.recording = True

    def handle_endtag(self, tag):
        if tag in ['p', 'h1', 'h2', 'h3', 'h4', 'li', 'strong']:
            self.recording = False
            self.text_data.append('\n')

    def handle_data(self, data):
        if self.recording:
            text = data.strip()
            if text:
                self.text_data.append(text)

def main():
    html_path = r"C:\Users\poons\AppData\Local\Temp\cortex_url_content.html"
    # Fallback to the saved md/html file in steps if Temp doesn't exist
    fallback_path = r"C:\Users\poons\.gemini\antigravity\brain\432dffda-e635-4620-ac22-df18237a5de6\.system_generated\steps\132\content.md"
    
    import os
    path_to_read = fallback_path
    if not os.path.exists(path_to_read):
        print("Path not found:", path_to_read)
        return

    with open(path_to_read, 'r', encoding='utf-8') as f:
        html_content = f.read()

    parser = TextExtractor()
    parser.feed(html_content)
    
    clean_text = " ".join(parser.text_data)
    # Basic clean up of whitespace
    clean_text = re.sub(r'\s+', ' ', clean_text)
    clean_text = re.sub(r' \n ', '\n', clean_text)
    
    output_lines = []
    for line in clean_text.split('\n'):
        line = line.strip()
        if len(line) > 10:
            output_lines.append(line)

    with open('article_text.txt', 'w', encoding='utf-8') as f_out:
        f_out.write('\n'.join(output_lines))
    print("[SUCCESS] Extracted text saved to article_text.txt")

if __name__ == '__main__':
    main()
