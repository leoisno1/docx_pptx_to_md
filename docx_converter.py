import io
import os

from docx import Document
from PIL import Image


def process_json_content(text):
    """处理JSON内容，将{和}移到最左边"""
    lines = text.split('\n')
    processed_lines = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('{') or stripped.startswith('}') or stripped.startswith('[') or stripped.startswith(']'):
            processed_lines.append(stripped)
        else:
            processed_lines.append(line)
    return '\n'.join(processed_lines)


def docx_to_markdown(docx_path, output_path):
    try:
        docx_basename = os.path.splitext(os.path.basename(docx_path))[0]
        output_dir = os.path.dirname(output_path)

        img_dir = os.path.join(output_dir, 'img')
        if not os.path.exists(img_dir):
            os.makedirs(img_dir)

        doc = Document(docx_path)
        markdown_content = []

        # 按文档顺序处理段落和表格
        for elem in doc.element.body:
            # 判断元素类型：w:p 是段落，w:tbl 是表格
            if elem.tag.endswith('}p'):
                # 段落处理
                from docx.text.paragraph import Paragraph
                para = Paragraph(elem, doc)
                style_name = para.style.name if para.style else 'Normal'
                text = para.text

                if style_name.startswith('Heading'):
                    try:
                        level = int(style_name.replace('Heading', ''))
                        heading = '#' * level + ' ' + text
                        markdown_content.append(heading)
                        continue
                    except:
                        pass

                is_list = False
                if para.style and para.style.name.startswith('List'):
                    is_list = True
                elif elem.xml.find('w:numPr') != -1:
                    is_list = True

                run_text = ''
                has_content = False

                for run in para.runs:
                    run_content = run.text

                    if run_content.strip() or hasattr(run._element, 'drawing_lst') and run._element.drawing_lst:
                        has_content = True

                    if run.bold:
                        run_content = f'**{run_content}**'

                    if run.italic:
                        run_content = f'*{run_content}*'

                    if run.underline:
                        run_content = f'__{run_content}__'

                    if hasattr(run._element, 'drawing_lst') and run._element.drawing_lst:
                        for drawing in run._element.drawing_lst:
                            try:
                                rid = drawing.xpath('.//a:blip/@r:embed')[0]
                                image = doc.part.related_parts[rid]
                                # 文件名中的空格替换为下划线，避免Markdown引用失败
                                safe_basename = docx_basename.replace(' ', '_')
                                image_filename = f'{safe_basename}_image_{len(os.listdir(img_dir))}.png'
                                image_path = os.path.join(img_dir, image_filename)
                                # 处理透明背景图片，转换为白底
                                img = Image.open(io.BytesIO(image.blob))
                                if img.mode in ('RGBA', 'LA', 'P'):
                                    background = Image.new('RGB', img.size, (255, 255, 255))
                                    if img.mode == 'P':
                                        img = img.convert('RGBA')
                                    if img.mode in ('RGBA', 'LA'):
                                        background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else img.split()[0])
                                    else:
                                        background.paste(img)
                                    background.save(image_path, 'PNG')
                                else:
                                    with open(image_path, 'wb') as f:
                                        f.write(image.blob)
                                run_content += f'\n![image](img/{image_filename})\n'
                            except:
                                pass

                    run_text += run_content

                if run_text.strip() or has_content:
                    if is_list:
                        markdown_content.append('* ' + run_text.strip())
                    else:
                        # 处理JSON内容
                        processed_text = process_json_content(run_text)
                        markdown_content.append(processed_text)

            elif elem.tag.endswith('}tbl'):
                # 表格处理
                from docx.table import Table
                table = Table(elem, doc)
                header_cells = table.rows[0].cells
                header_row = '| ' + ' | '.join(' '.join(cell.text.split()) for cell in header_cells) + ' |'
                markdown_content.append(header_row)
                separator_row = '| ' + ' | '.join('---' for _ in header_cells) + ' |'
                markdown_content.append(separator_row)
                for row in table.rows[1:]:
                    row_cells = row.cells
                    # 检测是否是只有第一列有内容的标题行（如 "Cdr参数说明："）
                    first_cell_text = ' '.join(row_cells[0].text.split())
                    other_cells_text = [' '.join(cell.text.split()) for cell in row_cells[1:]]
                    # 只有当文本以冒号结尾时才视为标题行（如 "参数说明："）
                    is_title_row = (
                        first_cell_text and 
                        all(t == '' for t in other_cells_text) and
                        not any(c in first_cell_text for c in ['|', '\n', '\t']) and
                        len(first_cell_text) < 50 and
                        first_cell_text.endswith('：')
                    )
                    if is_title_row:
                        # 作为独立段落处理，而不是表格行
                        markdown_content.append(first_cell_text)
                    else:
                        data_row = '| ' + ' | '.join(' '.join(cell.text.split()) for cell in row_cells) + ' |'
                        markdown_content.append(data_row)
                # 表格结束后添加空行
                markdown_content.append('')

        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                in_table = False
                for item in markdown_content:
                    # 将tab替换为4个空格
                    item = item.replace('\t', '    ')
                    # 去掉每行最前面的空格
                    lines = item.split('\n')
                    lines = [line.lstrip() for line in lines]
                    item = '\n'.join(lines)
                    f.write(item)
                    # 检测是否进入/离开表格块
                    if item.startswith('|'):
                        in_table = True
                        f.write('\n')
                    elif in_table:
                        # 之前在表格块中，现在是非表格行，添加空行分隔
                        in_table = False
                        f.write('\n\n')
                    else:
                        f.write('\n\n')
        except:
            return False

        return True
    except:
        return False
