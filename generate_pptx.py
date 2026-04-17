"""
PPTX Generator for Select by G Group Supplier Presentations
Uses the template and fills in supplier data to generate a new presentation.
"""

import os
import re
import copy
import shutil
import zipfile
import tempfile
from pathlib import Path
from lxml import etree

try:
    from PIL import Image as PILImage
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

# XML Namespaces
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'


def find_template():
    """Find template file using config."""
    try:
        from config import find_template as _find_template
        return _find_template()
    except Exception:
        pass
    for root, dirs, files in os.walk('/sessions'):
        for f in files:
            if 'Template' in f and f.endswith('.pptx') and 'Fournisseurs' in f:
                return Path(os.path.join(root, f))
    raise FileNotFoundError("Template PPTX not found")


def find_shape_by_name(root_elem, shape_name):
    """Find a <p:sp> element by its cNvPr name attribute."""
    for sp in root_elem.findall('.//{%s}sp' % NS_P):
        for elem in sp.iter():
            if '}cNvPr' in elem.tag:
                if elem.get('name') == shape_name:
                    return sp
    return None


def get_txbody(sp_elem):
    """Get the txBody element from a shape (handles p: namespace)."""
    body = sp_elem.find('{%s}txBody' % NS_P)
    if body is None:
        body = sp_elem.find('.//{%s}txBody' % NS_P)
    return body


def set_textbox_lines(sp_elem, lines, use_breaks=True):
    """
    Replace text content in a shape's txBody with the given list of lines.
    Preserves existing paragraph/run formatting.
    """
    body = get_txbody(sp_elem)
    if body is None:
        return False

    existing_paras = body.findall('{%s}p' % NS_A)
    if not existing_paras:
        return False

    first_para = existing_paras[0]
    first_ppr = first_para.find('{%s}pPr' % NS_A)

    first_rpr = None
    for para in existing_paras:
        for run in para.findall('{%s}r' % NS_A):
            rpr = run.find('{%s}rPr' % NS_A)
            if rpr is not None:
                first_rpr = copy.deepcopy(rpr)
                break
        if first_rpr is not None:
            break

    new_para = etree.Element('{%s}p' % NS_A)
    if first_ppr is not None:
        new_para.append(copy.deepcopy(first_ppr))

    non_empty_lines = [l for l in lines if l.strip()]
    for i, line in enumerate(non_empty_lines):
        run = etree.SubElement(new_para, '{%s}r' % NS_A)
        if first_rpr is not None:
            run.append(copy.deepcopy(first_rpr))
        t = etree.SubElement(run, '{%s}t' % NS_A)
        t.text = line
        if i < len(non_empty_lines) - 1:
            br = etree.SubElement(new_para, '{%s}br' % NS_A)
            if first_rpr is not None:
                br.append(copy.deepcopy(first_rpr))

    for p in existing_paras:
        body.remove(p)
    body.append(new_para)
    return True


def set_textbox_single(sp_elem, text):
    """Set a single text string in a shape, replacing all existing text."""
    return set_textbox_lines(sp_elem, [text], use_breaks=False)


def update_slide1(slide_xml_bytes, data):
    """Update slide 1 (supplier info slide) with the given data."""
    root = etree.fromstring(slide_xml_bytes)

    title_text = data.get('company_name', '')
    if data.get('category'):
        title_text = f"{title_text} - {data['category']}"
    sp = find_shape_by_name(root, 'Sous-titre 2')
    if sp is not None:
        set_textbox_single(sp, title_text)

    history = data.get('history', '')
    history_lines = [l.strip() for l in history.split('\n') if l.strip()]
    sp = find_shape_by_name(root, 'TextBox 13')
    if sp is not None:
        set_textbox_lines(sp, history_lines)

    projects = data.get('projects', '')
    projects_lines = [l.strip() for l in projects.split('\n') if l.strip()]
    sp = find_shape_by_name(root, 'TextBox 15')
    if sp is not None:
        set_textbox_lines(sp, projects_lines)

    identity = data.get('identity', '')
    identity_lines = [l.strip() for l in identity.split('\n') if l.strip()]
    sp = find_shape_by_name(root, 'TextBox 18')
    if sp is not None:
        set_textbox_lines(sp, identity_lines)

    added_values = [v for v in data.get('added_values', []) if v.strip()]
    sp = find_shape_by_name(root, 'TextBox 20')
    if sp is not None:
        set_textbox_lines(sp, added_values)

    references = [r for r in data.get('references', []) if r.strip()]
    sp = find_shape_by_name(root, 'TextBox 23')
    if sp is not None:
        set_textbox_lines(sp, references)

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def update_hotel_name_on_slide(root, hotel_name):
    """Update the hotel name text box on a photo slide."""
    possible_names = ['Image 1', 'ZoneTexte 48', 'Maison Heler Hotel\u2013 Metz ****']
    for name in possible_names:
        sp = find_shape_by_name(root, name)
        if sp is not None:
            body = get_txbody(sp)
            if body is not None:
                set_textbox_single(sp, hotel_name)
                return True
    return False


def get_image_rids_from_slide(root, rels_root):
    """Get list of (rid, pic_shape_name) for all image relationships in a slide."""
    image_rids = []
    seen_rids = set()
    for pic in root.findall('.//{%s}pic' % NS_P):
        name = None
        for elem in pic.iter():
            if '}cNvPr' in elem.tag:
                name = elem.get('name')
                break
        blip = pic.find('.//{%s}blip' % NS_A)
        if blip is not None:
            rid = blip.get('{%s}embed' % NS_R)
            if rid and rid not in seen_rids:
                image_rids.append((rid, name))
                seen_rids.add(rid)
    return image_rids


def replace_images_in_slide(root, rels_root, photos, slide_num, tmpdir):
    """Replace images in a slide with uploaded photos."""
    image_rids = get_image_rids_from_slide(root, rels_root)

    for i, (photo_path) in enumerate(photos[:len(image_rids)]):
        if i >= len(image_rids):
            break
        rid, pic_name = image_rids[i]
        if not Path(photo_path).exists():
            continue
        ext = Path(photo_path).suffix.lower()
        for rel in rels_root.findall('{%s}Relationship' % NS_PKG):
            if rel.get('Id') == rid:
                new_media_name = f'hotel_s{slide_num}_p{i+1}{ext}'
                media_dest = Path(tmpdir) / 'ppt' / 'media' / new_media_name
                shutil.copy2(photo_path, str(media_dest))
                rel.set('Target', f'../media/{new_media_name}')
                break

    ct_path = Path(tmpdir) / '[Content_Types].xml'
    ct_root = etree.fromstring(ct_path.read_bytes())
    existing_exts = {d.get('Extension') for d in ct_root.findall('{%s}Default' % NS_CT)}
    mime_map = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'png': 'image/png',
                'gif': 'image/gif', 'webp': 'image/webp'}
    for photo_path in photos:
        ext = Path(photo_path).suffix.lower().lstrip('.')
        if ext not in existing_exts:
            new_default = etree.SubElement(ct_root, '{%s}Default' % NS_CT)
            new_default.set('Extension', ext)
            new_default.set('ContentType', mime_map.get(ext, 'image/jpeg'))
            existing_exts.add(ext)
    ct_path.write_bytes(etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True))


def add_slide_to_presentation(tmpdir, new_slide_num, template_slide_xml, template_rels_xml,
                               template_slide_num=2):
    """Add a new slide to the presentation based on a template slide."""
    slides_dir = Path(tmpdir) / 'ppt' / 'slides'
    rels_dir = Path(tmpdir) / 'ppt' / 'slides' / '_rels'

    new_slide_path = slides_dir / f'slide{new_slide_num}.xml'
    new_rels_path = rels_dir / f'slide{new_slide_num}.xml.rels'
    new_slide_path.write_bytes(template_slide_xml)
    new_rels_path.write_bytes(template_rels_xml)

    prs_path = Path(tmpdir) / 'ppt' / 'presentation.xml'
    prs_root = etree.fromstring(prs_path.read_bytes())
    sld_id_lst = prs_root.find('{%s}sldIdLst' % NS_P)
    existing = sld_id_lst.findall('{%s}sldId' % NS_P)
    next_id = max(int(s.get('id')) for s in existing) + 1

    new_rid = f'rId_slide{new_slide_num}'
    new_sld_id = etree.SubElement(sld_id_lst, '{%s}sldId' % NS_P)
    new_sld_id.set('id', str(next_id))
    new_sld_id.set('{%s}id' % NS_R, new_rid)
    prs_path.write_bytes(etree.tostring(prs_root, xml_declaration=True, encoding='UTF-8', standalone=True))

    prs_rels_path = Path(tmpdir) / 'ppt' / '_rels' / 'presentation.xml.rels'
    prs_rels_root = etree.fromstring(prs_rels_path.read_bytes())
    new_rel = etree.SubElement(prs_rels_root, '{%s}Relationship' % NS_PKG)
    new_rel.set('Id', new_rid)
    new_rel.set('Type', f'{NS_R}/slide')
    new_rel.set('Target', f'slides/slide{new_slide_num}.xml')
    prs_rels_path.write_bytes(etree.tostring(prs_rels_root, xml_declaration=True,
                                               encoding='UTF-8', standalone=True))

    ct_path = Path(tmpdir) / '[Content_Types].xml'
    ct_root = etree.fromstring(ct_path.read_bytes())
    existing_parts = {o.get('PartName') for o in ct_root.findall('{%s}Override' % NS_CT)}
    part_name = f'/ppt/slides/slide{new_slide_num}.xml'
    if part_name not in existing_parts:
        new_override = etree.SubElement(ct_root, '{%s}Override' % NS_CT)
        new_override.set('PartName', part_name)
        new_override.set('ContentType',
            'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
    ct_path.write_bytes(etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True))


def remove_slides_from_presentation(tmpdir, slide_nums_to_remove):
    """Remove specified slides from the presentation."""
    prs_path = Path(tmpdir) / 'ppt' / 'presentation.xml'
    prs_rels_path = Path(tmpdir) / 'ppt' / '_rels' / 'presentation.xml.rels'
    prs_root = etree.fromstring(prs_path.read_bytes())
    prs_rels_root = etree.fromstring(prs_rels_path.read_bytes())
    sld_id_lst = prs_root.find('{%s}sldIdLst' % NS_P)
    existing = sld_id_lst.findall('{%s}sldId' % NS_P)

    for slide_num in sorted(slide_nums_to_remove, reverse=True):
        idx = slide_num - 1
        if idx < len(existing):
            sld_elem = existing[idx]
            rid = sld_elem.get('{%s}id' % NS_R)
            sld_id_lst.remove(sld_elem)
            for rel in prs_rels_root.findall('{%s}Relationship' % NS_PKG):
                if rel.get('Id') == rid:
                    prs_rels_root.remove(rel)
                    break
            slide_file = Path(tmpdir) / 'ppt' / 'slides' / f'slide{slide_num}.xml'
            rels_file = Path(tmpdir) / 'ppt' / 'slides' / '_rels' / f'slide{slide_num}.xml.rels'
            for f in [slide_file, rels_file]:
                if f.exists():
                    f.unlink()

    prs_path.write_bytes(etree.tostring(prs_root, xml_declaration=True, encoding='UTF-8', standalone=True))
    prs_rels_path.write_bytes(etree.tostring(prs_rels_root, xml_declaration=True,
                                               encoding='UTF-8', standalone=True))


def _get_image_dimensions(path):
    """Return (width, height) in pixels using PIL, or (None, None) if unavailable."""
    if not _PIL_AVAILABLE:
        return None, None
    try:
        with PILImage.open(str(path)) as img:
            return img.width, img.height
    except Exception:
        return None, None


def _fit_pic_to_image(pic_elem, img_w, img_h):
    """
    Resize a <p:pic> shape so the image fits inside its current bounding box
    without stretching, keeping the shape centred on its original position.
    """
    if not img_w or not img_h:
        return
    xfrm = pic_elem.find('.//{%s}xfrm' % NS_A)
    if xfrm is None:
        return
    off = xfrm.find('{%s}off' % NS_A)
    ext = xfrm.find('{%s}ext' % NS_A)
    if off is None or ext is None:
        return

    cx = int(ext.get('cx'))
    cy = int(ext.get('cy'))
    x  = int(off.get('x'))
    y  = int(off.get('y'))

    scale   = min(cx / img_w, cy / img_h)
    new_cx  = int(img_w * scale)
    new_cy  = int(img_h * scale)
    new_x   = x + (cx - new_cx) // 2
    new_y   = y + (cy - new_cy) // 2

    off.set('x',  str(new_x))
    off.set('y',  str(new_y))
    ext.set('cx', str(new_cx))
    ext.set('cy', str(new_cy))


def replace_supplier_logo(logo_path, tmpdir):
    """
    Replace the supplier logo image on all slides.
    Slide 1: shape 'Image 15' — Slides 2+: shape 'Image 46'
    Falls back to smallest image if shape name not found.
    """
    if not logo_path or not Path(logo_path).exists():
        return

    slides_dir = Path(tmpdir) / 'ppt' / 'slides'
    rels_dir   = slides_dir / '_rels'
    media_dir  = Path(tmpdir) / 'ppt' / 'media'

    logo_ext   = Path(logo_path).suffix.lower()
    logo_media = f'supplier_logo{logo_ext}'
    logo_dest  = media_dir / logo_media
    shutil.copy2(logo_path, str(logo_dest))

    ct_path = Path(tmpdir) / '[Content_Types].xml'
    ct_root = etree.fromstring(ct_path.read_bytes())
    existing_exts = {d.get('Extension') for d in ct_root.findall('{%s}Default' % NS_CT)}
    mime_map = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
                'png': 'image/png',  'gif':  'image/gif', 'webp': 'image/webp'}
    ext_key = logo_ext.lstrip('.')
    if ext_key not in existing_exts:
        new_default = etree.SubElement(ct_root, '{%s}Default' % NS_CT)
        new_default.set('Extension', ext_key)
        new_default.set('ContentType', mime_map.get(ext_key, 'image/png'))
    ct_path.write_bytes(etree.tostring(ct_root, xml_declaration=True,
                                        encoding='UTF-8', standalone=True))

    LOGO_SHAPE_NAMES = {'slide1': 'Image 15', 'default': 'Image 46'}

    for slide_file in sorted(slides_dir.glob('slide*.xml')):
        rels_file = rels_dir / (slide_file.name + '.rels')
        if not rels_file.exists():
            continue

        slide_root = etree.fromstring(slide_file.read_bytes())
        rels_root  = etree.fromstring(rels_file.read_bytes())

        slide_key  = slide_file.stem
        logo_shape = LOGO_SHAPE_NAMES.get(slide_key, LOGO_SHAPE_NAMES['default'])

        target_rid = None
        for pic in slide_root.findall('.//{%s}pic' % NS_P):
            for elem in pic.iter():
                if '}cNvPr' in elem.tag and elem.get('name') == logo_shape:
                    blip = pic.find('.//{%s}blip' % NS_A)
                    if blip is not None:
                        target_rid = blip.get('{%s}embed' % NS_R)
                    break
            if target_rid:
                break

        if not target_rid:
            smallest_size = None
            for pic in slide_root.findall('.//{%s}pic' % NS_P):
                blip = pic.find('.//{%s}blip' % NS_A)
                if blip is None:
                    continue
                rid = blip.get('{%s}embed' % NS_R)
                for rel in rels_root.findall('{%s}Relationship' % NS_PKG):
                    if rel.get('Id') == rid:
                        img_name = rel.get('Target', '').split('/')[-1]
                        img_path = media_dir / img_name
                        if img_path.exists():
                            sz = img_path.stat().st_size
                            if smallest_size is None or sz < smallest_size[0]:
                                smallest_size = (sz, rid)
            if smallest_size:
                target_rid = smallest_size[1]

        if not target_rid:
            continue

        for rel in rels_root.findall('{%s}Relationship' % NS_PKG):
            if rel.get('Id') == target_rid:
                rel.set('Target', f'../media/{logo_media}')
                break

        img_w, img_h = _get_image_dimensions(logo_path)
        if img_w and img_h:
            for pic in slide_root.findall('.//{%s}pic' % NS_P):
                blip = pic.find('.//{%s}blip' % NS_A)
                if blip is not None and blip.get('{%s}embed' % NS_R) == target_rid:
                    _fit_pic_to_image(pic, img_w, img_h)
                    break

        slide_file.write_bytes(etree.tostring(slide_root, xml_declaration=True,
                                               encoding='UTF-8', standalone=True))
        rels_file.write_bytes(etree.tostring(rels_root, xml_declaration=True,
                                              encoding='UTF-8', standalone=True))


def replace_slide1_photos(slide_root, rels_root, supplier_photos, tmpdir):
    """
    Replace the two content photo slots on slide 1 with uploaded supplier photos.
    Shape names: 'Picture 17' and 'Picture 25'
    """
    SLIDE1_PHOTO_SHAPES = ['Picture 17', 'Picture 25']

    media_dir = Path(tmpdir) / 'ppt' / 'media'
    ct_path   = Path(tmpdir) / '[Content_Types].xml'
    ct_root   = etree.fromstring(ct_path.read_bytes())
    existing_exts = {d.get('Extension') for d in ct_root.findall('{%s}Default' % NS_CT)}
    mime_map  = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'png': 'image/png',
                 'gif': 'image/gif',  'webp': 'image/webp'}
    ct_changed = False

    for idx, (shape_name, photo_path) in enumerate(zip(SLIDE1_PHOTO_SHAPES, supplier_photos)):
        if not photo_path or not Path(photo_path).exists():
            continue

        target_rid = None
        for pic in slide_root.findall('.//{%s}pic' % NS_P):
            for elem in pic.iter():
                if '}cNvPr' in elem.tag and elem.get('name') == shape_name:
                    blip = pic.find('.//{%s}blip' % NS_A)
                    if blip is not None:
                        target_rid = blip.get('{%s}embed' % NS_R)
                    break
            if target_rid:
                break

        if not target_rid:
            print(f"  WARNING: shape '{shape_name}' not found on slide 1")
            continue

        ext = Path(photo_path).suffix.lower()
        new_media_name = f'slide1_supplier_{idx + 1}{ext}'
        shutil.copy2(str(photo_path), str(media_dir / new_media_name))

        for rel in rels_root.findall('{%s}Relationship' % NS_PKG):
            if rel.get('Id') == target_rid:
                rel.set('Target', f'../media/{new_media_name}')
                break

        ext_key = ext.lstrip('.')
        if ext_key not in existing_exts:
            new_default = etree.SubElement(ct_root, '{%s}Default' % NS_CT)
            new_default.set('Extension', ext_key)
            new_default.set('ContentType', mime_map.get(ext_key, 'image/jpeg'))
            existing_exts.add(ext_key)
            ct_changed = True

    if ct_changed:
        ct_path.write_bytes(etree.tostring(ct_root, xml_declaration=True,
                                            encoding='UTF-8', standalone=True))


def pack_pptx(source_dir, output_path):
    """Pack a directory back into a PPTX file."""
    source_dir = Path(source_dir)
    with zipfile.ZipFile(str(output_path), 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(source_dir.rglob('*')):
            if file_path.is_file():
                arcname = str(file_path.relative_to(source_dir))
                zf.write(str(file_path), arcname)


def generate_presentation(data, photo_groups, supplier_logo_path=None, supplier_photos=None):
    """
    Generate a PPTX presentation from supplier data.

    data: dict with company_name, category, history, identity, projects,
          added_values (list), references (list)
    photo_groups: list of hotel dicts {hotel_name, photos} — each gets its own slide
    supplier_photos: list of up to 2 paths for the resume slide (slide 1)

    Returns: Path to the generated file
    """
    template_path = find_template()
    try:
        from config import GDRIVE_FOLDER
        output_dir = GDRIVE_FOLDER
    except Exception:
        output_dir = Path(__file__).parent.parent

    company_slug = re.sub(r'[^\w\s-]', '', data.get('company_name', 'Supplier')).strip().replace(' ', '_')[:30]
    output_filename = f"{company_slug}_SelectByG.pptx"
    output_path = output_dir / output_filename

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(str(template_path), 'r') as zf:
            zf.extractall(tmpdir)

        slides_dir = Path(tmpdir) / 'ppt' / 'slides'
        rels_dir = Path(tmpdir) / 'ppt' / 'slides' / '_rels'

        # Update slide 1 text
        slide1_path      = slides_dir / 'slide1.xml'
        slide1_rels_path = rels_dir   / 'slide1.xml.rels'
        slide1_bytes     = slide1_path.read_bytes()
        updated_slide1   = update_slide1(slide1_bytes, data)
        slide1_path.write_bytes(updated_slide1)

        # Place supplier photos on slide 1
        if supplier_photos:
            slide1_root      = etree.fromstring(slide1_path.read_bytes())
            slide1_rels_root = etree.fromstring(slide1_rels_path.read_bytes())
            replace_slide1_photos(slide1_root, slide1_rels_root, supplier_photos, tmpdir)
            slide1_path.write_bytes(
                etree.tostring(slide1_root, xml_declaration=True, encoding='UTF-8', standalone=True))
            slide1_rels_path.write_bytes(
                etree.tostring(slide1_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True))

        # Hotel photo slides
        slide2_template_bytes = (slides_dir / 'slide2.xml').read_bytes()
        slide2_rels_template_bytes = (rels_dir / 'slide2.xml.rels').read_bytes()
        remove_slides_from_presentation(tmpdir, [3, 4, 5])

        if photo_groups:
            first_group = photo_groups[0]
            slide2_root = etree.fromstring(slide2_template_bytes)
            slide2_rels_root = etree.fromstring(slide2_rels_template_bytes)
            update_hotel_name_on_slide(slide2_root, first_group.get('hotel_name', ''))
            if first_group.get('photos'):
                replace_images_in_slide(slide2_root, slide2_rels_root,
                                        first_group['photos'], 2, tmpdir)
            (slides_dir / 'slide2.xml').write_bytes(
                etree.tostring(slide2_root, xml_declaration=True, encoding='UTF-8', standalone=True))
            (rels_dir / 'slide2.xml.rels').write_bytes(
                etree.tostring(slide2_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True))

            for i, group in enumerate(photo_groups[1:], start=3):
                new_slide_root = etree.fromstring(slide2_template_bytes)
                new_rels_root = etree.fromstring(slide2_rels_template_bytes)
                update_hotel_name_on_slide(new_slide_root, group.get('hotel_name', ''))
                if group.get('photos'):
                    replace_images_in_slide(new_slide_root, new_rels_root,
                                            group['photos'], i, tmpdir)
                new_slide_bytes = etree.tostring(new_slide_root, xml_declaration=True,
                                                  encoding='UTF-8', standalone=True)
                new_rels_bytes = etree.tostring(new_rels_root, xml_declaration=True,
                                                 encoding='UTF-8', standalone=True)
                add_slide_to_presentation(tmpdir, i, new_slide_bytes, new_rels_bytes)
        else:
            remove_slides_from_presentation(tmpdir, [2])

        # Replace supplier logo on all slides
        if supplier_logo_path:
            replace_supplier_logo(supplier_logo_path, tmpdir)

        pack_pptx(tmpdir, output_path)

    print(f"Generated: {output_path}")
    return output_path


if __name__ == '__main__':
    data = {
        'company_name': 'TEST SUPPLIER',
        'category': 'FIXED & LOOSE FURNITURE',
        'history': '30 years of history\nA family business\n1 factory in Europe with 200 employees',
        'identity': 'We create bespoke furniture solutions for the luxury hospitality sector across Europe.',
        'projects': 'Each project is carefully planned from concept to delivery.\nDedicated project managers oversee every phase.',
        'added_values': ['Custom design for luxury hotels', 'European manufacturing',
                         'On-site installation', 'Sustainable materials'],
        'references': ['Hotel Ritz Paris ***** - Rooms', 'Four Seasons Geneva ***** - Suites'],
    }
    result = generate_presentation(data, [])
    print(f"Success: {result}")
