import tkinter as tk
from tkinter import filedialog
from docx import Document

def docx_to_html_with_formatting(docx_path, html_path):
    doc = Document(docx_path)

    html_content = '''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {
                font-family: "Times New Roman", serif;
                margin: 20px auto;
                width: 21cm;
                height: 29.7cm;
                padding: 1cm;
                background: white;
                box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            }
            p {
                margin: 0;
                white-space: pre-wrap;
            }
            table {
                width: 100%;
                border-collapse: collapse;
            }
            th, td {
                padding: 5px;
                text-align: left;
                vertical-align: top;
            }
        </style>
    </head>
    <body>
    '''

    def get_run_style(run):
        """Extract style from a run object."""
        styles = []
        if run.bold:
            styles.append("font-weight: bold;")
        if run.italic:
            styles.append("font-style: italic;")
        if run.font.size:
            styles.append(f"font-size: {run.font.size.pt}pt;")
        if run.font.color and run.font.color.rgb:
            styles.append(f"color: #{run.font.color.rgb};")
        return "".join(styles)

    def get_paragraph_style(paragraph):
        """Extract spacing, indent, and alignment from a paragraph."""
        styles = []
        if paragraph.alignment == 0:
            styles.append("text-align: left;")
        elif paragraph.alignment == 1:
            styles.append("text-align: center;")
        elif paragraph.alignment == 2:
            styles.append("text-align: right;")
        elif paragraph.alignment == 3:
            styles.append("text-align: justify;")

        if paragraph.paragraph_format.space_before:
            styles.append(f"margin-top: {paragraph.paragraph_format.space_before.pt}pt;")
        if paragraph.paragraph_format.space_after:
            styles.append(f"margin-bottom: {paragraph.paragraph_format.space_after.pt}pt;")
        if paragraph.paragraph_format.line_spacing:
            styles.append(f"line-height: {paragraph.paragraph_format.line_spacing};")
        if paragraph.paragraph_format.first_line_indent:
            styles.append(f"text-indent: {paragraph.paragraph_format.first_line_indent.pt}pt;")

        return "".join(styles)

    def get_cell_style(cell):
        """Extract alignment, background color, and border for table cells."""
        cell_style = ["text-align: left;", "vertical-align: top;"]
        cell_fill = cell._element.xpath(".//w:shd[@w:fill]")
        if cell_fill:
            color = cell_fill[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill")
            if color and color != "auto":
                cell_style.append(f"background-color: #{color};")
        return "".join(cell_style)

    for block in doc.element.body.iterchildren():
        if block.tag.endswith('p'):  # Paragraphs
            for paragraph in doc.paragraphs:
                if paragraph._element == block:
                    paragraph_style = get_paragraph_style(paragraph)
                    html_content += f'<p style="{paragraph_style}">'
                    for run in paragraph.runs:
                        run_style = get_run_style(run)
                        html_content += f'<span style="{run_style}">{run.text}</span>'
                    html_content += '</p>'
                    break

        elif block.tag.endswith('tbl'):  # Tables
            for table in doc.tables:
                if table._element == block:
                    html_content += '<table style="border: none;">'
                    for row in table.rows:
                        html_content += '<tr>'
                        for cell in row.cells:
                            cell_style = get_cell_style(cell)
                            cell_content = ""
                            for paragraph in cell.paragraphs:
                                paragraph_style = get_paragraph_style(paragraph)
                                cell_content += f'<p style="{paragraph_style}">'
                                for run in paragraph.runs:
                                    run_style = get_run_style(run)
                                    cell_content += f'<span style="{run_style}">{run.text}</span>'
                                cell_content += '</p>'
                            html_content += f'<td style="{cell_style}">{cell_content}</td>'
                        html_content += '</tr>'
                    html_content += '</table>'
                    break

    html_content += '</body></html>'

    with open(html_path, 'w', encoding='utf-8') as html_file:
        html_file.write(html_content)

def main():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Chọn file DOCX",
        filetypes=[("Word Documents", "*.docx")]
    )

    if not file_path:
        print("Không có file nào được chọn.")
        return

    output_path = filedialog.asksaveasfilename(
        title="Lưu file HTML",
        defaultextension=".html",
        filetypes=[("HTML files", "*.html")]
    )

    if not output_path:
        print("Không có nơi lưu file được chọn.")
        return

    print(f"Chuyển đổi file: {file_path} → {output_path}")
    docx_to_html_with_formatting(file_path, output_path)
    print(f"Hoàn tất! File HTML đã được lưu tại: {output_path}")

if __name__ == "__main__":
    main()
