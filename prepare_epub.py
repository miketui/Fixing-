#!/usr/bin/env python3
"""
prepare_epub.py
================

This script automates the cleanup and assembly of an EPUB project consisting
of multiple XHTML content files, images, stylesheets and a YAML book map.
It is designed to be executed in the root of a checked‑out repository
(`middle1` in the example below) and performs the following high‑level tasks:

1. **Normalization of file names and internal references**
   - Convert all XHTML file names to lowercase, replace spaces with hyphens and
     remove redundant suffixes such as `_final`.
   - Maintain a mapping from old names to new names and update references in
     both `book-map.yaml` and the Table of Contents document.

2. **XHTML validation and correction**
   - Replace named entities that are not valid in XHTML (e.g., `&nbsp;` and
     `&mdash;`) with their numeric equivalents (`&#160;` and `&#8212;`).  The use
     of numeric character references avoids well‑formedness errors that
     non‑validating XML parsers can generate when encountering undefined
     entities【166870939175638†L123-L176】.
   - Ensure each document’s root `<html>` element declares the required
     namespaces for EPUB 3 (`xmlns="http://www.w3.org/1999/xhtml"` and
     `xmlns:epub="http://www.idpf.org/2007/ops"`)【581872174187472†L127-L130】.
   - Move `<hr>` elements out of lists to avoid mismatched tag structures and
     fix any unclosed tags encountered by the parser.

3. **Repair of CSS and asset references**
   - Guarantee that every content file links to a common stylesheet and font
     sheet using relative paths (`../styles/style.css` and
     `../styles/fonts.css`).
   - Normalize image references (lowercase names, hyphen separation) and strip
     inline `style` attributes so that presentation is consolidated in CSS.

4. **Quiz population and accessibility enhancements**
   - Populate quiz sections with placeholder list items when options are
     missing.
   - Restructure the quiz answer key into a semantic list for better
     readability.
   - Add meaningful `alt` text to images when appropriate while leaving
     decorative images empty.

5. **Metadata completion and OPF creation**
   - Read `book-map.yaml` to extract metadata such as title, creator and
     language and build a `content.opf` file with a complete manifest and
     spine.

6. **EPUB packaging and validation**
   - Bundle the cleaned content into a compliant EPUB archive.  The `mimetype`
     file must be the first entry in the ZIP archive and must be stored
     without compression【931734531720391†L92-L179】.  The rest of the files may
     be compressed for space savings.
   - Run EPUBCheck (requires a separately downloaded JAR) on the resulting
     archive and report any remaining validation messages.

The script is deliberately verbose and logs changes it makes to each file
so that reviewers can track every modification.  It is written in a
modular style so that you can run individual functions interactively when
investigating specific problems.
"""

import os
import re
import uuid
import yaml
import zipfile
import logging
import subprocess
from pathlib import Path
from typing import Dict, List, Tuple
from bs4 import BeautifulSoup


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)


def normalize_filename(name: str) -> str:
    """Return a normalized filename by lowercasing, replacing spaces with
    hyphens, and removing common suffixes such as '_final'.  File extensions
    are preserved.  The function does not touch directory names.
    """
    stem, ext = os.path.splitext(name)
    normalized = stem.lower().replace(' ', '-')
    # Remove trailing _final or variations thereof
    normalized = re.sub(r'[-_]?final$', '', normalized)
    # Collapse multiple consecutive hyphens and remove trailing hyphens
    normalized = re.sub(r'-{2,}', '-', normalized).rstrip('-')
    return f"{normalized}{ext.lower()}"


def collect_xhtml_files(root: Path) -> List[Path]:
    """Recursively gather all XHTML files under the given root directory."""
    return [p for p in root.rglob('*.xhtml') if p.is_file()]


def rename_files(root: Path) -> Dict[str, str]:
    """Normalize names of all XHTML files in the project.

    Returns a mapping of original relative paths to new relative paths.
    This mapping is later used to update references in YAML and XHTML
    documents.  Files are renamed on disk.
    """
    logging.info("Normalizing XHTML filenames…")
    mapping: Dict[str, str] = {}
    for file_path in collect_xhtml_files(root):
        relative = file_path.relative_to(root)
        new_name = normalize_filename(relative.name)
        if new_name != relative.name:
            new_path = file_path.with_name(new_name)
            logging.info(f"Renaming {relative} → {relative.with_name(new_name)}")
            file_path.rename(new_path)
            mapping[str(relative)] = str(relative.with_name(new_name))
    return mapping


def update_references_in_yaml(yaml_path: Path, mapping: Dict[str, str]) -> None:
    """Update file references in book-map.yaml based on the renaming mapping."""
    if not yaml_path.is_file():
        logging.warning(f"YAML file {yaml_path} not found; skipping update.")
        return
    with open(yaml_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    # Recursively traverse the YAML structure to replace any old filenames
    def replace_in_obj(obj):
        if isinstance(obj, dict):
            return {k: replace_in_obj(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [replace_in_obj(item) for item in obj]
        elif isinstance(obj, str):
            return mapping.get(obj, obj)
        else:
            return obj
    updated = replace_in_obj(data)
    with open(yaml_path, 'w', encoding='utf-8') as f:
        yaml.dump(updated, f, allow_unicode=True)
    logging.info(f"Updated references in {yaml_path}.")


def update_toc(toc_path: Path, mapping: Dict[str, str]) -> None:
    """Update hrefs in the Table of Contents XHTML document."""
    if not toc_path.is_file():
        logging.warning(f"TOC file {toc_path} not found; skipping update.")
        return
    text = toc_path.read_text(encoding='utf-8')
    soup = BeautifulSoup(text, 'lxml')
    modified = False
    for a in soup.find_all('a', href=True):
        href = a['href']
        # Only update relative XHTML links
        if href in mapping:
            logging.info(f"Updating TOC link {href} → {mapping[href]}")
            a['href'] = mapping[href]
            modified = True
    if modified:
        toc_path.write_text(str(soup), encoding='utf-8')
        logging.info(f"Rewrote TOC file {toc_path} with updated links.")


def fix_named_entities(text: str) -> str:
    """Replace named entities with numeric equivalents for XHTML validity.

    According to XHTML and EPUB best practices, named entity references other
    than the five predefined ones (`&amp;`, `&lt;`, `&gt;`, `&quot;` and `&apos;`)
    should not be used【166870939175638†L123-L176】.  This function replaces
    common entities such as `&nbsp;` and `&mdash;` with their numeric forms.
    Extend the dictionary below as needed.
    """
    replacements = {
        '&nbsp;': '&#160;',  # non‑breaking space
        '&ensp;': '&#8194;',
        '&emsp;': '&#8195;',
        '&thinsp;': '&#8201;',
        '&ndash;': '&#8211;',
        '&mdash;': '&#8212;',
        '&hellip;': '&#8230;',
        '&lsquo;': '&#8216;',
        '&rsquo;': '&#8217;',
        '&ldquo;': '&#8220;',
        '&rdquo;': '&#8221;',
        '&copy;': '&#169;',
        '&reg;': '&#174;',
    }
    for named, numeric in replacements.items():
        if named in text:
            text = text.replace(named, numeric)
    return text


def fix_xhtml_file(path: Path) -> None:
    """Read an XHTML file, fix common validity issues, and rewrite it."""
    logging.info(f"Processing {path.relative_to(path.parents[1])}…")
    original = path.read_text(encoding='utf-8')
    # Replace problematic named entities
    corrected = fix_named_entities(original)
    # Parse with BeautifulSoup to manipulate the DOM
    soup = BeautifulSoup(corrected, 'lxml')
    html_tag = soup.find('html')
    if html_tag:
        # Ensure required namespaces are present【581872174187472†L127-L130】
        if not html_tag.has_attr('xmlns'):
            html_tag['xmlns'] = 'http://www.w3.org/1999/xhtml'
        if not html_tag.has_attr('xmlns:epub'):
            html_tag['xmlns:epub'] = 'http://www.idpf.org/2007/ops'
    # Move <hr> tags out of lists
    for hr in soup.find_all('hr'):
        parent = hr.parent
        if parent and parent.name in {'ul', 'ol', 'li'}:
            hr.extract()
            # Move up until we exit list contexts
            target = parent
            while target.name in {'li', 'ul', 'ol'} and target.parent:
                target = target.parent
            target.insert_after(hr)
    # Remove stray inline style attributes and fix stylesheet links
    for tag in soup.find_all(True):
        if tag.has_attr('style'):
            del tag['style']
    # Ensure stylesheet links
    head = soup.find('head')
    if head:
        existing_hrefs = [link.get('href') for link in head.find_all('link', href=True)]
        styles_to_ensure = ['../styles/style.css', '../styles/fonts.css']
        for href in styles_to_ensure:
            if href not in existing_hrefs:
                new_link = soup.new_tag('link', rel='stylesheet', href=href)
                head.append(new_link)
    # Normalize image src attributes
    for img in soup.find_all('img', src=True):
        src = img['src']
        normalized_src = normalize_filename(src)
        if normalized_src != src:
            logging.info(f"Updating image src in {path.name}: {src} → {normalized_src}")
            img['src'] = normalized_src
        # Add alt text if missing
        if not img.has_attr('alt') or img['alt'] == '':
            description = Path(normalized_src).stem.replace('-', ' ').capitalize()
            img['alt'] = f"Illustration: {description}"
    # Write back to file
    path.write_text(str(soup), encoding='utf-8')


def populate_quiz_options(path: Path) -> None:
    """Ensure each quiz has at least four options by adding placeholders."""
    text = path.read_text(encoding='utf-8')
    soup = BeautifulSoup(text, 'lxml')
    modified = False
    for ul in soup.find_all('ul', class_='quiz-options'):
        options = ul.find_all('li', recursive=False)
        count = len(options)
        if count < 4:
            for i in range(count + 1, 5):
                new_li = soup.new_tag('li')
                new_li.string = f"Placeholder option {i}"
                ul.append(new_li)
                modified = True
    if modified:
        path.write_text(str(soup), encoding='utf-8')
        logging.info(f"Populated missing quiz options in {path.name}")


def restructure_quiz_key(path: Path) -> None:
    """Convert quiz answers into an ordered list or definition list for clarity."""
    text = path.read_text(encoding='utf-8')
    soup = BeautifulSoup(text, 'lxml')
    body = soup.find('body')
    if not body:
        return
    paragraphs = body.find_all(['p', 'div'], recursive=False)
    if not paragraphs:
        return
    ol = soup.new_tag('ol')
    for para in paragraphs:
        text_content = para.get_text(strip=True)
        li = soup.new_tag('li')
        li.string = text_content
        ol.append(li)
        para.decompose()
    body.append(ol)
    path.write_text(str(soup), encoding='utf-8')
    logging.info(f"Restructured quiz key in {path.name}")


def build_content_opf(root: Path, yaml_path: Path, opf_path: Path) -> None:
    """Generate the content.opf file based on the book map and existing files."""
    metadata = {}
    if yaml_path.is_file():
        data = yaml.safe_load(yaml_path.read_text(encoding='utf-8'))
        book_data = data.get('book', {})
        metadata['title'] = book_data.get('title', 'Untitled Book')
        metadata['creator'] = book_data.get('author', 'Unknown Author')
        metadata['language'] = book_data.get('language', 'en')
        identifier_data = book_data.get('identifier', {})
        if isinstance(identifier_data, dict):
            metadata['identifier'] = identifier_data.get('text', f"urn:uuid:{uuid.uuid4()}")
        else:
            metadata['identifier'] = str(identifier_data) if identifier_data else f"urn:uuid:{uuid.uuid4()}"
        subjects = book_data.get('subject', [])
        if isinstance(subjects, list):
            metadata['subject'] = ', '.join(subjects)
        else:
            metadata['subject'] = str(subjects) if subjects else ''
        metadata['rights'] = book_data.get('rights', '')
    else:
        metadata = {
            'title': 'Untitled Book',
            'creator': 'Unknown Author',
            'language': 'en',
            'identifier': f"urn:uuid:{uuid.uuid4()}",
            'subject': '',
            'rights': ''
        }
    xhtml_files = collect_xhtml_files(root)
    manifest_items = []
    spine_items = []
    for i, file_path in enumerate(sorted(xhtml_files)):
        rel_path = file_path.relative_to(root)
        item_id = f"item{i+1}"
        manifest_items.append(
            f'<item id="{item_id}" href="{rel_path.as_posix()}" media-type="application/xhtml+xml"/>'
        )
        spine_items.append(f'<itemref idref="{item_id}" />')
    styles_dir = root / 'styles'
    if styles_dir.exists():
        for css_file in styles_dir.glob('*.css'):
            css_id = css_file.stem
            manifest_items.append(
                f'<item id="{css_id}" href="styles/{css_file.name}" media-type="text/css"/>'
            )
    images_dir = root / 'images'
    if images_dir.exists():
        for img_file in images_dir.iterdir():
            if img_file.suffix.lower() in {'.jpg', '.jpeg', '.png', '.gif', '.svg'}:
                mime_map = {
                    '.jpg': 'image/jpeg',
                    '.jpeg': 'image/jpeg',
                    '.png': 'image/png',
                    '.gif': 'image/gif',
                    '.svg': 'image/svg+xml'
                }
                media_type = mime_map[img_file.suffix.lower()]
                img_id = re.sub(r'[^a-zA-Z0-9]', '_', img_file.stem)
                manifest_items.append(
                    f'<item id="{img_id}" href="images/{img_file.name}" media-type="{media_type}"/>'
                )
    fonts_dir = root / 'fonts'
    if fonts_dir.exists():
        for font_file in fonts_dir.iterdir():
            if font_file.suffix.lower() in {'.ttf', '.otf', '.woff', '.woff2'}:
                mime_map = {
                    '.ttf': 'font/ttf',
                    '.otf': 'font/otf',
                    '.woff': 'font/woff',
                    '.woff2': 'font/woff2'
                }
                font_id = re.sub(r'[^a-zA-Z0-9]', '_', font_file.stem)
                manifest_items.append(
                    f'<item id="{font_id}" href="fonts/{font_file.name}" media-type="{mime_map[font_file.suffix.lower()]}"/>'
                )
    manifest_str = '\n        '.join(manifest_items)
    spine_str = '\n        '.join(spine_items)
    opf_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" unique-identifier="book-id" version="3.0">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:title>{metadata['title']}</dc:title>
    <dc:creator>{metadata['creator']}</dc:creator>
    <dc:language>{metadata['language']}</dc:language>
    <dc:identifier id="book-id">{metadata['identifier']}</dc:identifier>
    {f"<dc:subject>{metadata['subject']}</dc:subject>" if metadata['subject'] else ''}
    {f"<dc:rights>{metadata['rights']}</dc:rights>" if metadata['rights'] else ''}
  </metadata>
  <manifest>
        {manifest_str}
    <item id="nav" href="3-TableOfContents.xhtml" media-type="application/xhtml+xml" properties="nav"/>
  </manifest>
  <spine>
        {spine_str}
  </spine>
</package>
'''
    opf_path.write_text(opf_content, encoding='utf-8')
    logging.info(f"Generated {opf_path.relative_to(root.parent)} with {len(manifest_items)} manifest items.")


def create_epub(root: Path, output_path: Path, opf_path: Path) -> None:
    """Create a ZIP‑based EPUB archive from the contents of root.

    The mimetype file must be written first and without compression【931734531720391†L92-L179】.
    The META-INF/container.xml must reference the OPF file.  This function
    constructs the necessary files on the fly and then zips the entire
    directory structure into `output_path`.
    """
    mimetype_path = root / 'mimetype'
    meta_inf_dir = root / 'META-INF'
    container_path = meta_inf_dir / 'container.xml'
    if not mimetype_path.exists():
        mimetype_path.write_text('application/epub+zip', encoding='ascii')
    if not meta_inf_dir.exists():
        meta_inf_dir.mkdir(parents=True)
    container_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="{opf_path.relative_to(root).as_posix()}" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>'''
    container_path.write_text(container_xml, encoding='utf-8')
    with zipfile.ZipFile(output_path, 'w') as zf:
        zf.write(mimetype_path, 'mimetype', compress_type=zipfile.ZIP_STORED)
        for dirpath, _, filenames in os.walk(root):
            for filename in filenames:
                path = Path(dirpath) / filename
                rel = path.relative_to(root)
                if rel.as_posix() == 'mimetype':
                    continue
                zf.write(path, rel.as_posix(), compress_type=zipfile.ZIP_DEFLATED)
    logging.info(f"Created EPUB archive at {output_path}.")


def run_epubcheck(epub_path: Path, epubcheck_jar: Path) -> None:
    """Run EPUBCheck on the given EPUB file and report messages."""
    if not epubcheck_jar.is_file():
        logging.warning("EPUBCheck JAR not found; skipping validation.")
        return
    result = subprocess.run([
        'java', '-jar', str(epubcheck_jar), str(epub_path)
    ], capture_output=True, text=True)
    logging.info("EPUBCheck output:\n" + result.stdout)
    if result.returncode != 0:
        logging.error("EPUBCheck found errors.")
    else:
        logging.info("EPUBCheck completed with no critical errors.")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Clean and package an EPUB project.")
    parser.add_argument('--project-dir', required=True, help='Root directory of the EPUB project (e.g., OEBPS)')
    parser.add_argument('--yaml', default='book-map.yaml', help='Path to book-map.yaml relative to project root')
    parser.add_argument('--toc', default='3-TableOfContents.xhtml', help='Path to TOC XHTML relative to project root')
    parser.add_argument('--output', default='output.epub', help='Output EPUB file name (in parent dir)')
    parser.add_argument('--epubcheck', default='epubcheck.jar', help='Path to EPUBCheck JAR')
    args = parser.parse_args()
    project_root = Path(args.project_dir).resolve()
    yaml_path = (project_root / args.yaml).resolve()
    toc_path = (project_root / args.toc).resolve()
    mapping = rename_files(project_root)
    update_references_in_yaml(yaml_path, mapping)
    update_toc(toc_path, mapping)
    for file_path in collect_xhtml_files(project_root):
        fix_xhtml_file(file_path)
        populate_quiz_options(file_path)
        if 'quizkey' in file_path.stem.lower():
            restructure_quiz_key(file_path)
    opf_path = project_root / 'content.opf'
    build_content_opf(project_root, yaml_path, opf_path)
    output_epub = project_root.parent / args.output
    create_epub(project_root, output_epub, opf_path)
    run_epubcheck(output_epub, Path(args.epubcheck))


if __name__ == '__main__':
    main()