import re
import markdown
from bs4 import BeautifulSoup
from typing import Tuple, Optional


class ContentConverter:
    MARKDOWN_TABLE_PATTERN = re.compile(
        r'\|(?:.+\|)+\n\|(?:[-:]+)+\|'
        r'(?:\n\|(?:.+\|)+)+',
        re.MULTILINE
    )
    
    HTML_PATTERN = re.compile(r'<[^>]+>')
    
    def __init__(self, html_formatting: bool = True):
        self.html_formatting = html_formatting
    
    def detect_content_type(self, content: str) -> str:
        if self._is_markdown_table(content):
            return 'markdown_table'
        if self._is_html(content):
            return 'html'
        return 'plain_text'
    
    def _is_markdown_table(self, content: str) -> bool:
        return bool(self.MARKDOWN_TABLE_PATTERN.search(content))
    
    def _is_html(self, content: str) -> bool:
        return bool(self.HTML_PATTERN.search(content))
    
    def format_markdown_delimiters(self, content: str) -> str:
        if not self.html_formatting:
            return content
        
        content = re.sub(r'~~([^~]+)~~', r'<del>\1</del>', content)
        content = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', content)
        content = re.sub(r'\*([^*]+)\*', r'<em>\1</em>', content)
        content = re.sub(r'`([^`]+)`', r'<code>\1</code>', content)
        
        return content
    
    def markdown_to_html(self, content: str) -> str:
        content = self.format_markdown_delimiters(content)
        html = markdown.markdown(content, extensions=['tables'])
        return html
    
    def html_to_richtext(self, html: str) -> Tuple[str, str]:
        soup = BeautifulSoup(html, 'lxml')
        
        plain_text = soup.get_text(separator='\n')
        
        return html, plain_text
    
    def markdown_table_to_csv(self, table_content: str) -> str:
        lines = table_content.strip().split('\n')
        
        if len(lines) < 2:
            return table_content
        
        header_line = self._parse_table_line(lines[0])
        separator_line = lines[1] if len(lines) > 1 else ''
        
        data_lines = []
        for line in lines[2:]:
            if line.strip():
                data_lines.append(self._parse_table_line(line))
        
        csv_lines = [header_line] + data_lines
        
        return '\n'.join(csv_lines)
    
    def _parse_table_line(self, line: str) -> str:
        parts = line.strip().strip('|').split('|')
        cells = [self._escape_csv_cell(p.strip()) for p in parts]
        return ','.join(cells)
    
    def _escape_csv_cell(self, cell: str) -> str:
        if ',' in cell or '"' in cell or '\n' in cell:
            cell = cell.replace('"', '""')
            return f'"{cell}"'
        return cell
    
    def convert(self, content: str) -> dict:
        content_type = self.detect_content_type(content)
        
        result = {
            'original': content,
            'type': content_type,
            'html': None,
            'plain_text': None,
            'csv': None
        }
        
        if content_type == 'markdown_table':
            result['csv'] = self.markdown_table_to_csv(content)
            result['html'] = self.markdown_to_html(content)
            result['plain_text'] = content
        
        elif content_type == 'html':
            result['html'], result['plain_text'] = self.html_to_richtext(content)
        
        else:
            result['html'] = self.markdown_to_html(content)
            result['plain_text'] = content
        
        return result
