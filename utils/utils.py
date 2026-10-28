from datetime import datetime


def parse_date_safely(date_str, format_str='%Y-%m-%d'):
    try:
        return datetime.strptime(date_str, format_str)
    except (ValueError, TypeError):
        return None


def safe_float_conversion(value):
    if isinstance(value, (int, float)):
        return float(value)
    elif isinstance(value, str):
        try:
            return float(value)
        except ValueError:
            return None
    else:
        return None


def extract_name_from_content(content):
    import re
    name = None
    cleaned_content = content
    
    # 括弧内の文字列を抽出
    name_match = re.search(r'\((.*?)\)', content)
    if name_match:
        name = name_match.group(1)
        cleaned_content = re.sub(r'\(.*?\)', '', content).strip()
        
        # 最初の単語だけ取得
        if cleaned_content:
            cleaned_content = cleaned_content.split()[0]
            
    return name, cleaned_content
