import sys
import os
import io
import re
import pdfplumber
import json

# 中文处理方式说明：
# 1. 设置标准输出和标准错误的编码为utf-8，解决控制台输出中文乱码问题
# 2. 使用 pdfplumber 的 extract_text() 方法提取文本，该方法会自动处理中文编码
# 3. 确保输出文本为 utf-8 编码，避免编码转换问题
# 4. 使用 errors='ignore' 参数处理无法解码的字符，确保程序不会崩溃
# 5. 最终输出时使用 encode('utf-8').decode('utf-8') 确保文本格式正确

# 设置标准输出和标准错误的编码为utf-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def get_heading_level(font_size, max_font_size):
    """根据字体大小确定标题级别"""
    if font_size >= max_font_size * 0.9:
        return 1
    elif font_size >= max_font_size * 0.8:
        return 2
    elif font_size >= max_font_size * 0.7:
        return 3
    elif font_size >= max_font_size * 0.6:
        return 4
    elif font_size >= max_font_size * 0.5:
        return 5
    else:
        return 6

def is_new_paragraph(word, prev_word, page_width, font_size):
    """判断是否是新段落的开始"""
    if not prev_word:
        return False
    
    # 如果当前单词在上一行之后很远，可能是新段落
    if word['top'] - prev_word['top'] > font_size * 1.5:
        return True
    
    # 如果当前单词在页面左侧，可能是新段落
    if word['x0'] < page_width * 0.1:  # 假设页面左侧10%的位置是段落缩进
        return True
    
    return False

def is_list_item(text):
    """判断是否是列表项"""
    # 有序列表
    if re.match(r'^\d+[\.\)]\s*', text):
        return True
    # 无序列表
    if re.match(r'^[-•*]\s*', text):
        return True
    return False

def get_format_info(chars, text, x0, y0):
    """获取文本的格式信息"""
    # 找到与文本位置最接近的字符
    closest_char = None
    min_distance = float('inf')
    
    for char in chars:
        # 计算字符与文本的距离
        distance = abs(char['x0'] - x0) + abs(char['top'] - y0)
        if distance < min_distance:
            min_distance = distance
            closest_char = char
    
    if closest_char:
        return {
            'size': closest_char.get('size', 12),
            'fontname': closest_char.get('fontname', '').lower()
        }
    return {'size': 12, 'fontname': ''}

def process_text_with_formatting(page, output_format='text'):
    """处理页面文本，根据输出格式决定是否保留格式信息"""
    if output_format == 'text':
        # 对于纯文本模式，直接提取文本
        text = page.extract_text()
        if text:
            # 清理多余的空行
            text = re.sub(r'\n\s*\n\s*\n', '\n\n', text)
            return text.strip()
        return ""
    else:
        # 对于markdown模式，先提取文本保持顺序，再添加格式
        # 1. 先提取纯文本保持顺序
        text = page.extract_text()
        if not text:
            return ""
        
        # 2. 获取所有字符的格式信息
        chars = page.chars
        if not chars:
            return text
        
        # 3. 获取页面宽度和最大字体大小
        page_width = page.width
        max_font_size = max((char.get('size', 0) for char in chars), default=12)
        
        # 4. 将文本分割成行
        lines = text.split('\n')
        formatted_lines = []
        
        for line in lines:
            if not line.strip():
                formatted_lines.append('')
                continue
                
            # 5. 获取这一行的格式信息
            # 找到这一行对应的字符
            line_chars = []
            for char in chars:
                # 使用y坐标判断是否在同一行
                if abs(char['top'] - (chars[0]['top'] if line_chars else char['top'])) < 5:
                    line_chars.append(char)
            
            if not line_chars:
                formatted_lines.append(line)
                continue
            
            # 6. 处理这一行的格式
            # 获取这一行的字体大小和字体名称
            font_size = max((char.get('size', 12) for char in line_chars), default=12)
            font_name = line_chars[0].get('fontname', '').lower()
            
            # 7. 处理标题
            if font_size >= max_font_size * 0.5:
                level = get_heading_level(font_size, max_font_size)
                line = f"{'#' * level} {line}"
            
            # 8. 处理粗体和斜体
            if 'bold' in font_name:
                line = f"**{line}**"
            if 'italic' in font_name:
                line = f"*{line}*"
            
            # 9. 处理列表项
            if is_list_item(line):
                # 根据缩进确定列表级别
                indent_level = int(line_chars[0]['x0'] / (page_width * 0.05))
                list_level = max(0, indent_level - 1)
                # 添加适当的缩进
                line = '  ' * list_level + line
            
            formatted_lines.append(line)
        
        return '\n'.join(formatted_lines)

def convert_pdf_to_text(pdf_path, output_format='markdown'):
    """将PDF文件转换为文本，支持markdown和纯文本格式"""
    try:
        # 确保输出编码为UTF-8
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
        
        # 检查文件是否存在
        if not os.path.exists(pdf_path):
            error_msg = f"文件不存在: {pdf_path}"
            print(f'data: {json.dumps({"type": "error", "content": error_msg}, ensure_ascii=False)}\n\n', flush=True)
            return None
            
        # 检查文件大小
        file_size = os.path.getsize(pdf_path)
        print(f'data: {json.dumps({"type": "debug", "content": f"文件大小: {file_size} 字节"}, ensure_ascii=False)}\n\n', flush=True)
        
        if file_size == 0:
            error_msg = "文件为空"
            print(f'data: {json.dumps({"type": "error", "content": error_msg}, ensure_ascii=False)}\n\n', flush=True)
            return None
            
        # 打开PDF文件
        print(f'data: {json.dumps({"type": "debug", "content": "正在打开PDF文件..."}, ensure_ascii=False)}\n\n', flush=True)
        
        with pdfplumber.open(pdf_path) as pdf:
            # 存储所有页面的文本
            all_text = []
            
            # 遍历所有页面
            print(f'data: {json.dumps({"type": "debug", "content": f"PDF总页数: {len(pdf.pages)}"}, ensure_ascii=False)}\n\n', flush=True)
            
            for i, page in enumerate(pdf.pages):
                print(f'data: {json.dumps({"type": "debug", "content": f"正在处理第 {i+1} 页..."}, ensure_ascii=False)}\n\n', flush=True)
                
                try:
                    # 尝试提取文本，忽略CropBox警告
                    text = page.extract_text()
                    if text:
                        # 确保文本是utf-8编码
                        if not isinstance(text, str):
                            text = text.decode('utf-8', errors='ignore')
                        all_text.append(text)
                except Exception as e:
                    print(f'data: {json.dumps({"type": "debug", "content": f"处理第 {i+1} 页时出现警告: {str(e)}"}, ensure_ascii=False)}\n\n', flush=True)
                    continue
            
            # 合并所有页面的文本
            if all_text:
                final_text = "\n\n".join(all_text)
                # 清理多余的空行
                final_text = re.sub(r'\n\s*\n\s*\n', '\n\n', final_text)
                
                print(f'data: {json.dumps({"type": "debug", "content": f"提取文本长度: {len(final_text)} 字符"}, ensure_ascii=False)}\n\n', flush=True)
                
                # 输出结果
                result = {
                    "type": "complete",
                    "content": final_text.strip()
                }
                print(f'data: {json.dumps(result, ensure_ascii=False)}\n\n', flush=True)
                return final_text.strip()
            else:
                error_msg = "警告：没有提取到任何文本"
                print(f'data: {json.dumps({"type": "error", "content": error_msg}, ensure_ascii=False)}\n\n', flush=True)
                return None

    except Exception as e:
        error_msg = f"转换过程中出错: {str(e)}"
        print(f'data: {json.dumps({"type": "error", "content": error_msg}, ensure_ascii=False)}\n\n', flush=True)
        
        import traceback
        error_detail = traceback.format_exc()
        print(f'data: {json.dumps({"type": "error", "content": f"错误详情: {error_detail}"}, ensure_ascii=False)}\n\n', flush=True)
        return None

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pdf_to_md.py <pdf_file_path> [format]", file=sys.stderr)
        print("format: 'markdown' (default) or 'text'", file=sys.stderr)
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    output_format = sys.argv[2] if len(sys.argv) > 2 else 'text'
    
    if output_format not in ['markdown', 'text']:
        print("Error: format must be either 'markdown' or 'text'", file=sys.stderr)
        sys.exit(1)
    
    result = convert_pdf_to_text(pdf_path, output_format)
    if result:
        # 使用utf-8编码输出结果，确保中文正确显示
        if isinstance(result, str):
            print(result.encode('utf-8').decode('utf-8'))
        else:
            print(result.decode('utf-8'))
    else:
        print("\n转换失败", file=sys.stderr)
        sys.exit(1) 