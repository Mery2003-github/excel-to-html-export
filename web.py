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

def argb_to_hex(argb):
    if argb is None:
        return None
    argb = argb.upper()
    if len(argb) == 8:  # ARGB
        return argb[2:]
    elif len(argb) == 6:  # déjà RGB
        return argb
    else:
        return argb

def extract_styles_from_xml(zipf):
    styles = {}
    try:
        with zipf.open('xl/styles.xml') as f:
            styles_xml = ET.parse(f).getroot()
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

            fonts = {}
            for i, font in enumerate(styles_xml.findall('main:fonts/main:font', ns)):
                fonts[i] = {
                    'bold': font.find('main:b', ns) is not None,
                    'italic': font.find('main:i', ns) is not None,
                    'underline': font.find('main:u', ns) is not None,
                    'size': float(font.find('main:sz', ns).get('val')) if font.find('main:sz', ns) is not None else 11,
                    'color': argb_to_hex(font.find('main:color', ns).get('rgb')) if font.find('main:color', ns) is not None else None
                }

            fills = {}
            for i, fill in enumerate(styles_xml.findall('main:fills/main:fill', ns)):
                pattern = fill.find('main:patternFill', ns)
                if pattern is not None:
                    fgColor = pattern.find('main:fgColor', ns)
                    fills[i] = {
                        'bg_color': argb_to_hex(fgColor.get('rgb')) if fgColor is not None else None
                    }

            borders = {}
            for i, border in enumerate(styles_xml.findall('main:borders/main:border', ns)):
                borders[i] = {}
                for side in ['left', 'right', 'top', 'bottom']:
                    side_elem = border.find(f'main:{side}', ns)
                    if side_elem is not None:
                        color_elem = side_elem.find('main:color', ns)
                        borders[i][side] = {
                            'style': side_elem.get('style'),
                            'color': argb_to_hex(color_elem.get('rgb')) if color_elem is not None else '000000'
                        }

            alignments = {}
            for i, xf in enumerate(styles_xml.findall('main:cellXfs/main:xf', ns)):
                align = xf.find('main:alignment', ns)
                alignments[i] = {
                    'fontId': int(xf.get('fontId', 0)),
                    'fillId': int(xf.get('fillId', 0)),
                    'borderId': int(xf.get('borderId', 0)),
                    'wrap': align.get('wrapText') == '1' if align is not None else False,
                    'horizontal': align.get('horizontal') if align is not None else 'general',
                    'vertical': align.get('vertical') if align is not None else 'bottom'
                }

            styles = {
                'fonts': fonts,
                'fills': fills,
                'borders': borders,
                'alignments': alignments
            }

    except Exception as e:
        print(f"⚠ Erreur lors de la lecture du fichier styles.xml: {e}")

    return styles

def get_cell_style(cell, styles):
    if not styles:
        return {
            'bold': False,
            'italic': False,
            'underline': False,
            'size': 11,
            'color': None,
            'bg_color': None,
            'border': {},
            'wrap': False,
            'align': 'general'
        }

    style_index = getattr(cell, 'style_id', 0)
    alignment = styles['alignments'].get(style_index, {})

    font = styles['fonts'].get(alignment.get('fontId', 0), {})
    fill = styles['fills'].get(alignment.get('fillId', 0), {})
    border = styles['borders'].get(alignment.get('borderId', 0), {})

    return {
        'bold': font.get('bold', False),
        'italic': font.get('italic', False),
        'underline': font.get('underline', False),
        'size': font.get('size', 11),
        'color': font.get('color'),
        'bg_color': fill.get('bg_color'),
        'border': border,
        'wrap': alignment.get('wrap', False),
        'align': alignment.get('horizontal', 'general')
    }

def get_image_data(zipf, media_path):
    with zipf.open(media_path) as f:
        data = f.read()

    mime = "image/webp"
    try:
        img = Image.open(io.BytesIO(data)).convert("RGB")

        if img.width > MAX_IMAGE_WIDTH:
            ratio = MAX_IMAGE_WIDTH / img.width
            new_height = int(img.height * ratio)
            img = img.resize((MAX_IMAGE_WIDTH, new_height), Image.LANCZOS)

        output = io.BytesIO()
        img.save(output, format="WEBP", quality=QUALITY, method=6)
        optimized_data = output.getvalue()
        return f"data:{mime};base64," + base64.b64encode(optimized_data).decode()

    except Exception as e:
        print(f"⚠ Erreur lors de l'image {media_path}: {e}")
        ext = media_path.split('.')[-1].lower()
        mime_fallback = "image/png" if ext == "png" else "image/jpeg"
        return f"data:{mime_fallback};base64," + base64.b64encode(data).decode()

def get_sheet_data(wb, zipf, sheet_name=None):
    styles = extract_styles_from_xml(zipf)
    ws = wb[sheet_name] if sheet_name else wb.active
    data = []

    for row in ws.iter_rows():
        row_data = []
        for cell in row:
            value = cell.value if cell.value is not None else ""
            style = get_cell_style(cell, styles)
            row_data.append({
                'value': str(value),
                'style': style,
                'row': cell.row,
                'col': cell.column
            })
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

        return images
    except Exception as e:
        print(f"⚠ Erreur lors de l'analyse du dessin: {e}")
        return []

def generate_html(sheet_data, images, col_widths, row_heights, output_file):
    total_width = sum(col_widths.values()) * PIXELS_PER_CHAR
    total_height = sum(h * (LINE_HEIGHT/15) for h in row_heights.values())

    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Export Excel fidèle</title>
</head>
<body>
<div style="position:relative;width:{total_width}px;height:{total_height}px;">
"""

    # Affichage du texte / cellules
    for row in sheet_data:
        for cell in row:
            left = calculate_position(cell['col'], 0, col_widths, True)
            top = calculate_position(cell['row'], 0, row_heights, False)
            width = col_widths.get(cell['col'], 8.43) * PIXELS_PER_CHAR
            height = row_heights.get(cell['row'], 15) * (LINE_HEIGHT/15)

            style = ""
            if cell['style']['bold']:
                style += "font-weight:bold;"
            if cell['style']['italic']:
                style += "font-style:italic;"
            if cell['style']['underline']:
                style += "text-decoration:underline;"
            if cell['style']['color']:
                style += f"color:#{cell['style']['color']};"
            if cell['style']['bg_color']:
                style += f"background-color:#{cell['style']['bg_color']};"

            for side, border in cell['style']['border'].items():
                if border and border.get('style') and border.get('color'):
                    style += f"border-{side}: 1px solid #{border['color']};"

            align_map = {
                'left': 'left',
                'right': 'right',
                'center': 'center',
                'justify': 'justify',
                'general': 'left',
                'distributed': 'justify'
            }
            align = align_map.get(cell['style']['align'], 'left')
            style += f"text-align:{align};"

            style += f"font-size:{cell['style']['size']}px;"

            if cell['style']['wrap']:
                style += "white-space:normal; overflow:visible;"

            html += f"""
<div style="position:absolute;left:{left}px; top:{top}px; width:{width}px; height:{height}px; {style} overflow:hidden;">
    {cell['value']}
</div>
"""

    # Affichage des images (avec position:absolute)
    for img in images:
     html += f"""
<div style="position:absolute; left:{img['left']}px; top:{img['top']}px; width:{img['width']}px; height:{img['height']}px; box-sizing:border-box;">
    <img src="{img['data_uri']}" alt="Image" style="width:100%; height:100%; object-fit:contain;">
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
    input_file = "mmmm (3).xlsx"
    output_file = "export_fidele.html"

    print("📂 Chargement du fichier Excel...")
    wb = load_workbook(input_file)
    sheet_name = wb.sheetnames[0]

    with zipfile.ZipFile(input_file) as zipf:
        print("📝 Extraction du texte et styles...")
        sheet_data = get_sheet_data(wb, zipf, sheet_name)
        col_widths = get_column_widths(wb, sheet_name)
        row_heights = get_row_heights(wb, sheet_name)

        print("🖼 Extraction des images...")
        all_images = []
        drawings = [f for f in zipf.namelist() if f.startswith('xl/drawings/drawing')]

        for drawing_path in drawings:
            images = parse_drawing(zipf, drawing_path, col_widths, row_heights)
            all_images.extend(images)

    print(f"✅ {len(all_images)} images trouvées")
    print("🧱 Génération du HTML fidèle...")
    generate_html(sheet_data, all_images, col_widths, row_heights, output_file)

if __name__ == "__main__":
    main()
