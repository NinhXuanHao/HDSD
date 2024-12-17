# import tkinter as tk
# from tkinter import filedialog
# from docx import Document

# def docx_to_html_with_formatting(docx_path, html_path):
#     doc = Document(docx_path)

#     html_content = '''
#     <!DOCTYPE html>
#     <html>
#     <head>
#         <meta charset="UTF-8">
#         <style>
#             body {
#                 font-family: "Times New Roman", serif;
#                 margin: 20px auto;
#                 width: 21cm;
#                 height: 29.7cm;
#                 padding: 1cm;
#                 background: white;
#                 box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
#             }
#             p {
#                 margin: 0;
#                 white-space: pre-wrap;
#             }
#             table {
#                 width: 100%;
#                 border-collapse: collapse;
#             }
#             th, td {
#                 padding: 5px;
#                 text-align: left;
#                 vertical-align: top;
#             }
#         </style>
#     </head>
#     <body>
#     '''

#     def get_run_style(run):
#         """Extract style from a run object."""
#         styles = []
#         if run.bold:
#             styles.append("font-weight: bold;")
#         if run.italic:
#             styles.append("font-style: italic;")
#         if run.font.size:
#             styles.append(f"font-size: {run.font.size.pt}pt;")
#         if run.font.color and run.font.color.rgb:
#             styles.append(f"color: #{run.font.color.rgb};")
#         return "".join(styles)

#     def get_paragraph_style(paragraph):
#         """Extract spacing, indent, and alignment from a paragraph."""
#         styles = []
#         if paragraph.alignment == 0:
#             styles.append("text-align: left;")
#         elif paragraph.alignment == 1:
#             styles.append("text-align: center;")
#         elif paragraph.alignment == 2:
#             styles.append("text-align: right;")
#         elif paragraph.alignment == 3:
#             styles.append("text-align: justify;")

#         if paragraph.paragraph_format.space_before:
#             styles.append(f"margin-top: {paragraph.paragraph_format.space_before.pt}pt;")
#         if paragraph.paragraph_format.space_after:
#             styles.append(f"margin-bottom: {paragraph.paragraph_format.space_after.pt}pt;")
#         if paragraph.paragraph_format.line_spacing:
#             styles.append(f"line-height: {paragraph.paragraph_format.line_spacing};")
#         if paragraph.paragraph_format.first_line_indent:
#             styles.append(f"text-indent: {paragraph.paragraph_format.first_line_indent.pt}pt;")

#         return "".join(styles)

#     def get_cell_style(cell):
#         """Extract alignment, background color, and border for table cells."""
#         cell_style = ["text-align: left;", "vertical-align: top;"]
#         cell_fill = cell._element.xpath(".//w:shd[@w:fill]")
#         if cell_fill:
#             color = cell_fill[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill")
#             if color and color != "auto":
#                 cell_style.append(f"background-color: #{color};")
#         return "".join(cell_style)

#     for block in doc.element.body.iterchildren():
#         if block.tag.endswith('p'):  # Paragraphs
#             for paragraph in doc.paragraphs:
#                 if paragraph._element == block:
#                     paragraph_style = get_paragraph_style(paragraph)
#                     html_content += f'<p style="{paragraph_style}">'
#                     for run in paragraph.runs:
#                         run_style = get_run_style(run)
#                         html_content += f'<span style="{run_style}">{run.text}</span>'
#                     html_content += '</p>'
#                     break

#         elif block.tag.endswith('tbl'):  # Tables
#             for table in doc.tables:
#                 if table._element == block:
#                     html_content += '<table style="border: none;">'
#                     for row in table.rows:
#                         html_content += '<tr>'
#                         for cell in row.cells:
#                             cell_style = get_cell_style(cell)
#                             cell_content = ""
#                             for paragraph in cell.paragraphs:
#                                 paragraph_style = get_paragraph_style(paragraph)
#                                 cell_content += f'<p style="{paragraph_style}">'
#                                 for run in paragraph.runs:
#                                     run_style = get_run_style(run)
#                                     cell_content += f'<span style="{run_style}">{run.text}</span>'
#                                 cell_content += '</p>'
#                             html_content += f'<td style="{cell_style}">{cell_content}</td>'
#                         html_content += '</tr>'
#                     html_content += '</table>'
#                     break

#     html_content += '</body></html>'

#     with open(html_path, 'w', encoding='utf-8') as html_file:
#         html_file.write(html_content)

# def main():
#     root = tk.Tk()
#     root.withdraw()

#     file_path = filedialog.askopenfilename(
#         title="Chọn file DOCX",
#         filetypes=[("Word Documents", "*.docx")]
#     )

#     if not file_path:
#         print("Không có file nào được chọn.")
#         return

#     output_path = filedialog.asksaveasfilename(
#         title="Lưu file HTML",
#         defaultextension=".html",
#         filetypes=[("HTML files", "*.html")]
#     )

#     if not output_path:
#         print("Không có nơi lưu file được chọn.")
#         return

#     print(f"Chuyển đổi file: {file_path} → {output_path}")
#     docx_to_html_with_formatting(file_path, output_path)
#     print(f"Hoàn tất! File HTML đã được lưu tại: {output_path}")

# if __name__ == "__main__":
#     main()

import tkinter as tk
from tkinter import filedialog
import subprocess
import win32com.client as win32  # Chỉ dùng cho Windows
from docx import Document
import os
def docx_to_html_with_full_content(docx_path, html_path):
    
    doc = Document(docx_path)
    script =r'''
    <script>
        function getUrlParameter(name) {
            name = name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]');
            var regex = new RegExp('[\\?&]' + name + '=([^&#]*)');
            var results = regex.exec(location.search);
            return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
        }

        // Function to set text content if element exists
        function setTextContentById(id, value) {
            var element = document.getElementById(id);
            if (element) {
                element.textContent = value;
            }
        }

        // Get parameters from URL and set content
        var parameters = ['a', 'b']; //, 'c', 'aj', 'tinh','ad', 'y', 'ss', 'ae', 'l', 'ak', 'al', 'am', 'an', 'f', 'n'];
        parameters.forEach(function (param) {
            var value = getUrlParameter(param);
            setTextContentById(param, value);
        });
        var y = getUrlParameter('y');
        document.getElementById('y1').textContent = y;
        var l = getUrlParameter('l1');
        document.getElementById('l1').textContent = l;
        document.getElementById('l2').textContent = l;
        document.getElementById('l3').textContent = l;

    </script>
    <script>
        function exportHTML() {
            var header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' " +
                "xmlns:w='urn:schemas-microsoft-com:office:word' " +
                "xmlns='http://www.w3.org/TR/REC-html40'>" +
                "<head><meta charset='utf-8'><title>Export HTML to Word Document with JavaScript</title>" +
                "<style>" +
                "@page {" +
                "    size: A4;" +
                "    margin: 20mm;" +
                "}" +
                "body {" +
                "    font-family: 'Times New Roman', Times, serif;" +
                "    width: 210mm;" +
                "    margin: 0 auto;" +
                "}" +
                ".center {" +
                "    text-align: center;" +
                "}" +
                ".right {" +
                "    text-align: right;" +
                "}" +
                "td p {" +
                "        margin: 0; /* Đảm bảo văn bản trong ô không có spacing */" +
                "}" +
                "p {" +
                "        text-align: justify;" +
                "        line-height: 1;" +
                "        margin: 0;" +
                "}" +
                ".bold {" +
                "    font-weight: bold;" +
                "}" +
                ".underline {" +
                "    text-decoration: underline;" +
                "}" +
                ".signature {" +
                "    margin-top: 50px;" +
                "}" +
                ".content {" +
                "    margin-left: 40px;" +
                "}" +
                ".red {" +
                "    color: red;" +
                "}" +
                ".buttons {" +
                "    text-align: center;" +
                "    margin-bottom: 20px;" +
                "}" +
                "</style>" +
                "</head><body>";

            var sourceHTML = header + document.getElementById("content").innerHTML;

            var source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
            var fileDownload = document.createElement("a");
            document.body.appendChild(fileDownload);
            fileDownload.href = source;
            fileDownload.download = 'document.doc';
            fileDownload.click();
            document.body.removeChild(fileDownload);
        }
    </script>
    '''
    html_content = r'''
    <!DOCTYPE html>
<html lang="vi">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Phối hợp Kiểm tra gửi NHNN</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet"
        integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">

    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.11.8/dist/umd/popper.min.js"
        integrity="sha384-I7E8VVD/ismYTF4hNIPjVp/Zjvgyol6VFvRkX/vR+Vc4jQkC+hVqc2pM8ODewa9r"
        crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.min.js"
        integrity="sha384-0pUGZvbkm6XF6gxjEnlmuGrJXVbNuzT9qBBavbLwCsOGabYfZo0T0to5eqruptLy"
        crossorigin="anonymous"></script>
    <style>
        @page {
            size: A4;
            margin: 20mm;
        }

        body {
            font-family: "Times New Roman", Times, serif;
            width: 210mm;
            /* Width of A4 */
            margin: 0 auto;
        }

        # p {
        #     text-align: justify;
        # }
        p, h1, h2, h3, div:not(table *) {
            text-align: justify; line-height: 1; /* Single spacing */
        }
        table td, table th {
            line-height: 1; /* Single spacing */
        }
        td p {
        margin: 0; /* Đảm bảo văn bản trong ô không có spacing */
        }
        .center {
            text-align: center;
        }

        .right {
            text-align: right;
        }

        .bold {
            font-weight: bold;
        }

        .underline {
            text-decoration: underline;
        }

        .signature {
            margin-top: 50px;
        }

        .content {
            margin-left: 40px;
        }

        .red {
            color: red;
        }

        .buttons {
            text-align: center;
            margin-bottom: 20px;
        }

        @media print {
            #footer {
                position: fixed;
                bottom: 0;
                width: 100%;
                text-align: right;
                /* Đảm bảo rằng nền footer là trắng để tránh chồng lấp */
                padding: 10px 0;
                /* Đường kẻ ngăn cách footer với nội dung */
            }

            /* Ẩn các phần tử khác nếu cần */
            #in * {
                visibility: hidden;
            }

            #footer,
            #footer,
            #content * {
                visibility: visible;
            }

            /* Nếu bạn muốn hiện các phần khác của nội dung, hãy tùy chỉnh các selector */
        }
    </style>

</head>

<body>
    <div id="in"><button type="button" class="btn btn-outline-primary" onclick="window.print()">In hoặc xuất
        PDF</button>
        <button id="exportBtn" type="button" class="btn btn-outline-primary" onclick="exportHTML()">Xuất thành file DOC</button></div>

    <div id="content">
    '''

    def get_run_style(run):
        """ Lấy định dạng chữ như in đậm, in nghiêng, size. """
        styles = []
        if run.bold:
            styles.append("font-weight: bold;")
        if run.italic:
            styles.append("font-style: italic;")
        if run.font.size:
            styles.append(f"font-size: {run.font.size.pt}pt;")
        return "".join(styles)

    def get_paragraph_style(paragraph, is_in_table):
        """ Lấy căn lề, thụt dòng đầu tiên và spacing. """
        styles = []

        # Căn lề
        if paragraph.alignment == 0:
            styles.append("text-align: left;")
        elif paragraph.alignment == 1:
            styles.append("text-align: center;")
        elif paragraph.alignment == 2:
            styles.append("text-align: right;")

        # Thụt đầu dòng
        if paragraph.paragraph_format.first_line_indent:
            styles.append(f"text-indent: {paragraph.paragraph_format.first_line_indent.pt}pt;")

        # Spacing Before/After: Kiểm tra spacing cho bảng và ngoài bảng
        # if is_in_table:
        #     styles.append("margin:0;margin-top: 0;margin-bottom:0")  # Trong bảng, loại bỏ spacing
        # else:
        #     spacing_before = paragraph.paragraph_format.space_before
        #     spacing_after = paragraph.paragraph_format.space_after
        #     if spacing_before:
        #         styles.append(f"margin-top: {spacing_before.pt}pt;")
        #     if spacing_after:
        #         styles.append(f"margin-bottom: {spacing_after.pt}pt;")

        spacing_before = paragraph.paragraph_format.space_before
        spacing_after = paragraph.paragraph_format.space_after
        if spacing_before:
            styles.append(f"margin-top: {spacing_before.pt}pt;")
        if spacing_after:
            styles.append(f"margin-bottom: {spacing_after.pt}pt;")
        return "".join(styles)

    def process_paragraph(paragraph, is_in_table=False):
        """ Xử lý văn bản cho từng đoạn văn bản. """
        paragraph_style = get_paragraph_style(paragraph, is_in_table)
        content = f'<p style="{paragraph_style}">'
        for run in paragraph.runs:
            run_style = get_run_style(run)
            content += f'<span style="{run_style}">{run.text}</span>'
        content += '</p>'
        return content

    def get_cell_width(cell):
        """ Lấy chiều rộng ô. """
        cell_width = cell._element.xpath(".//w:tcW[@w:w]")
        if cell_width:
            width = cell_width[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w")
            return f"width: {int(width) / 20}pt;"
        return ""

    def get_table_column_widths(table):
        """ Lấy chiều rộng các cột trong bảng. """
        grid_cols = table._element.xpath(".//w:tblGrid/w:gridCol")
        column_widths = [int(col.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w")) / 20
                         for col in grid_cols]
        return column_widths

    def get_cell_border_color(cell):
        """ Lấy màu border của ô nếu có, trả về None nếu không có border. """
        try:
            borders = cell._element.xpath('.//w:tcBorders')
            if not borders:
                return None  # Không có border
            for border in borders[0]:  # Duyệt qua từng đường viền (top, bottom, left, right)
                color = border.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                if color and color != 'auto':
                    return f"#{color}"  # Trả về màu border
            return None
        except Exception as e:
            return None
    
    def process_table(table):
        """ Xử lý bảng để giữ đúng kích thước và định dạng. """
        column_widths = get_table_column_widths(table)
        table_html = '<table>'
        for row in table.rows:
            table_html += '<tr>'
            for i, cell in enumerate(row.cells):
                cell_style = get_cell_width(cell)
                if i < len(column_widths):
                    cell_style += f" width: {column_widths[i]}pt;"

                # Kiểm tra border trong ô
                border_color = get_cell_border_color(cell)
                if not border_color:
                    cell_style += " border: none;"  # Không có border
                else:
                    cell_style += f" border: 1px solid {border_color};"
                
                cell_content = ""
                for paragraph in cell.paragraphs:
                    cell_content += process_paragraph(paragraph)
                table_html += f'<td style="{cell_style}">{cell_content}</td>'
            table_html += '</tr>'
        table_html += '</table>'
        return table_html

    # Xử lý từng phần từ trên xuống dưới (văn bản hoặc bảng)
    for element in doc.element.body:
        if element.tag.endswith('tbl'):
            table = next((t for t in doc.tables if t._element == element), None)
            if table:
                html_content += process_table(table)
        elif element.tag.endswith('p'):
            paragraph = next((p for p in doc.paragraphs if p._element == element), None)
            if paragraph:
                html_content += process_paragraph(paragraph, is_in_table=False)

    html_content += script 
    html_content +='</body></html>'

    with open(html_path, 'w', encoding='utf-8') as html_file:
        html_file.write(html_content)

def convert_doc_to_docx(input_path):
    """ Chuyển file .doc sang .docx bằng Microsoft Word (win32com). """
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # Đảm bảo Word không hiển thị giao diện
    output_path = os.path.splitext(input_path)[0] + ".docx"
    
    try:
        print("Đang chuyển đổi file DOC sang DOCX...")
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
        doc.Close()
        print(f"Chuyển đổi thành công: {output_path}")
        return output_path
    except Exception as e:
        print(f"Lỗi trong quá trình chuyển đổi: {e}")
        return None
    finally:
        word.Quit()

def main():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Chọn file DOCX",
        filetypes=[("Word files", "*.doc;*.docx")]
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
    """ Chuyển file DOCX thành HTML và giữ nguyên định dạng. """
    # Kiểm tra định dạng file và chuyển đổi nếu cần
    if file_path.endswith(".doc"):
        file_path = convert_doc_to_docx(file_path)
        if not file_path:
            print("Không thể chuyển đổi file DOC sang DOCX.")
            return

    docx_to_html_with_full_content(file_path, output_path)
    print(f"Hoàn tất! File HTML đã được lưu tại: {output_path}")

if __name__ == "__main__":
    main()
