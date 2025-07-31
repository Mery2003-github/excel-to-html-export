import zipfile
import base64
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image
import io

EMU_PER_PIXEL = 9525
PIXELS_PER_CHAR = 7  
LINE_HEIGHT = 19     

MAX_IMAGE_WIDTH = 300 
QUALITY = 50           

def get_image_data(zipf, media_path):
    with zipf.open(media_path) as f:
        data = f.read()

    mime = "image/webp"

    try:
        img = Image.open(io.BytesIO(data)).convert("RGB")

        # Redimensionner si trop large
        if img.width > MAX_IMAGE_WIDTH:
            ratio = MAX_IMAGE_WIDTH / img.width
            new_height = int(img.height * ratio)
            img = img.resize((MAX_IMAGE_WIDTH, new_height), Image.LANCZOS)

        output = io.BytesIO()
        img.save(output, format="WEBP", quality=QUALITY, method=6)
        optimized_data = output.getvalue()

        return f"data:{mime};base64," + base64.b64encode(optimized_data).decode()

    except Exception as e:
        print(f"⚠ Erreur lors de l’image {media_path}: {e}")
        # Fallback : base64 original (non compressé)
        ext = media_path.split('.')[-1].lower()
        mime_fallback = "image/png" if ext == "png" else "image/jpeg"
        return f"data:{mime_fallback};base64," + base64.b64encode(data).decode()

def get_sheet_data(wb, sheet_name=None):
    ws = wb[sheet_name] if sheet_name else wb.active
    data = []
    for row in ws.iter_rows():
        row_data = []
        for cell in row:
            value = cell.value if cell.value is not None else ""
            font = cell.font
            alignment = cell.alignment
            
            color = None
            if font.color and font.color.type == 'rgb':
                color = font.color.rgb[2:]  # Supprime 'FF' initial
            elif font.color and font.color.type == 'theme':
                color = '000000'
            
            style = {
                'bold': font.bold,
                'italic': font.italic,
                'underline': font.underline,
                'size': font.size,
                'color': color,
                'wrap': alignment.wrapText if alignment else False,
                'align': alignment.horizontal if alignment else 'general'
            }
            row_data.append({'value': str(value), 'style': style})
        data.append(row_data)
    return data

def get_column_widths(wb, sheet_name=None):
    ws = wb[sheet_name] if sheet_name else wb.active
    col_widths = {}
    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        col_dim = ws.column_dimensions.get(col_letter)
        col_widths[col] = col_dim.width if col_dim and col_dim.width else 8.43
    return col_widths

def get_row_heights(wb, sheet_name=None):
    ws = wb[sheet_name] if sheet_name else wb.active
    row_heights = {}
    for row in range(1, ws.max_row + 1):
        row_dim = ws.row_dimensions.get(row)
        row_heights[row] = row_dim.height if row_dim and row_dim.height else 15
    return row_heights

def calculate_position(start_idx, offset_emu, dimensions, is_column=True):
    if is_column:
        total = sum(dimensions.get(i, 8.43) * PIXELS_PER_CHAR for i in range(1, start_idx))
    else:
        total = sum(dimensions.get(i, 15) * (LINE_HEIGHT/15) for i in range(1, start_idx))
    return total + (offset_emu / EMU_PER_PIXEL)

def parse_drawing(zipf, drawing_path, col_widths, row_heights):
    try:
        rels_path = drawing_path.replace('drawings/', 'drawings/_rels/') + '.rels'
        with zipf.open(rels_path) as f:
            rels_root = ET.parse(f).getroot()
        
        rels = {rel.attrib['Id']: rel.attrib['Target'].replace('../', 'xl/')
                for rel in rels_root.findall('{*}Relationship')}

        ns = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        with zipf.open(drawing_path) as f:
            root = ET.parse(f).getroot()

        images = []
        for anchor in root.findall('xdr:oneCellAnchor', ns) + root.findall('xdr:twoCellAnchor', ns):
            try:
                from_elem = anchor.find('xdr:from', ns)
                if from_elem is None:
                    continue
                    
                col = int(from_elem.find('xdr:col', ns).text) + 1
                row = int(from_elem.find('xdr:row', ns).text) + 1
                colOff = int(from_elem.find('xdr:colOff', ns).text)
                rowOff = int(from_elem.find('xdr:rowOff', ns).text)

                ext = anchor.find('xdr:ext', ns)
                if ext is None:
                    continue
                    
                cx, cy = int(ext.attrib['cx']), int(ext.attrib['cy'])

                blip = anchor.find('.//a:blip', ns)
                if blip is None:
                    continue
                    
                embed = blip.attrib.get(f"{{{ns['r']}}}embed")
                if embed not in rels:
                    continue

                left = calculate_position(col, colOff, col_widths, True)
                top = calculate_position(row, rowOff, row_heights, False)
                width = cx / EMU_PER_PIXEL
                height = cy / EMU_PER_PIXEL

                images.append({
                    'row': row,
                    'col': col,
                    'left': left,
                    'top': top,
                    'width': width,
                    'height': height,
                    'data_uri': get_image_data(zipf, rels[embed])
                })
            except Exception as e:
                print(f"⚠ Erreur image (nested): {e}")
                continue

        return images
    except Exception as e:
        print(f"⚠ Erreur lors de l'analyse du dessin: {e}")
        return []

def generate_html(sheet_data, images, col_widths, row_heights, output_file):
    max_width = max((img['left'] + img['width'] for img in images), default=800)
    max_height = max((img['top'] + img['height'] for img in images), default=600)

    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Export Excel fidèle</title>
    <style>
        body {{ font-family: Arial; margin: 20px; }}
        .excel-container {{
            position: relative;
            width: {max_width}px;
            height: {max_height}px;
            border: 1px solid #ddd;
        }}
        .cell {{
            position: absolute;
            white-space: nowrap;
            overflow: hidden;
            font-size: 11px;
            padding: 2px;
        }}
        .image-container {{
            position: absolute;
            border: 1px solid rgba(0,0,0,0.2);
        }}
        .image-container img {{
            width: 100%;
            height: 100%;
            object-fit: contain;
        }}
    </style>
</head>
<body>
<h1>Export Excel avec positionnement exact</h1>
<div class="excel-container">
"""

    for row_idx, row in enumerate(sheet_data, 1):
        for col_idx, cell in enumerate(row, 1):
            if not cell['value']:
                continue
                
            left = calculate_position(col_idx, 0, col_widths, True)
            top = calculate_position(row_idx, 0, row_heights, False)
            width = col_widths.get(col_idx, 8.43) * PIXELS_PER_CHAR
            height = row_heights.get(row_idx, 15) * (LINE_HEIGHT/15)
            
            style = ""
            if cell['style']['bold']:
                style += "font-weight:bold;"
            if cell['style']['italic']:
                style += "font-style:italic;"
            if cell['style']['underline']:
                style += "text-decoration:underline;"
            if cell['style']['color']:
                style += f"color:#{cell['style']['color']};"
            
            html += f"""
<div class="cell" 
     style="left:{left}px; top:{top}px; 
            width:{width}px; height:{height}px;
            {style}">
    {cell['value']}
</div>
"""

    for img in images:
        html += f"""
<div class="image-container" 
     style="left:{img['left']}px; top:{img['top']}px;
            width:{img['width']}px; height:{img['height']}px;">
    <img src="{img['data_uri']}" alt="Image">
</div>
"""

    html += """
</div>
</body>
</html>
"""

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"✅ Fichier HTML généré: {output_file}")

def main():
    input_file = "Etiquette CLEMENTINE.xlsx"
    output_file = "export_fidele.html"

    print("📂 Chargement du fichier Excel...")
    wb = load_workbook(input_file)
    sheet_name = wb.sheetnames[0]
    
    print("📝 Extraction du texte et styles...")
    sheet_data = get_sheet_data(wb, sheet_name)
    col_widths = get_column_widths(wb, sheet_name)
    row_heights = get_row_heights(wb, sheet_name)

    print("🖼 Extraction des images...")
    all_images = []
    with zipfile.ZipFile(input_file) as zipf:
        drawings = [f for f in zipf.namelist() if f.startswith('xl/drawings/drawing')]
        
        for drawing_path in drawings:
            images = parse_drawing(zipf, drawing_path, col_widths, row_heights)
            all_images.extend(images)

    print(f"✅ {len(all_images)} images trouvées")
    print("🧱 Génération du HTML fidèle...")
    generate_html(sheet_data, all_images, col_widths, row_heights, output_file)

if __name__ == "__main__":
    main()
