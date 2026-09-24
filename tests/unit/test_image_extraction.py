"""Tests for ZIP-based image extraction and media validation (engine.py)."""
import os
import zipfile

import pytest

from grepxcel.engine import extract_images, _check_media_bytes


def _make_xlsx_with_image(path: str, png_bytes: bytes) -> None:
    """Build a minimal xlsx ZIP with one embedded image anchored at B3 (col=1, row=2 0-based)."""
    drawing_xml = (
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<xdr:twoCellAnchor>'
        '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>'
        '<xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff>'
        '<xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        '<xdr:pic>'
        '<xdr:nvPicPr><xdr:cNvPr id="1" name="img1"/>'
        '<xdr:cNvPicPr/></xdr:nvPicPr>'
        '<xdr:blipFill>'
        '<a:blip r:embed="rId1"/>'
        '</xdr:blipFill>'
        '<xdr:spPr/>'
        '</xdr:pic>'
        '<xdr:clientData/>'
        '</xdr:twoCellAnchor>'
        '</xdr:wsDr>'
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
        ' Target="../media/image1.png"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(path, 'w') as zf:
        zf.writestr('xl/drawings/drawing1.xml', drawing_xml)
        zf.writestr('xl/drawings/_rels/drawing1.xml.rels', rels_xml)
        zf.writestr('xl/media/image1.png', png_bytes)


def _make_xlsx_with_richdata_image(path: str, media_data: bytes,
                                    media_name: str = 'image1.png',
                                    cell_ref: str = 'B3') -> None:
    """Build a minimal xlsx exercising the IMAGE()-formula richData chain.

    *media_name* is the untrusted zip member name (attacker-controlled) and
    *cell_ref* is the untrusted worksheet cell reference attribute.
    """
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )
    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
        ' Target="worksheets/sheet1.xml"/>'
        '</Relationships>'
    )
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData><row><c r="{cell_ref}" vm="1"/></row></sheetData>'
        '</worksheet>'
    )
    rvr_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<rel r:id="rId1"/>'
        '</richValueRels>'
    )
    rvr_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Target="../media/{media_name}"/>'
        '</Relationships>'
    )
    rdrichvalue_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">'
        '<rv><v>0</v></rv>'
        '</rvData>'
    )
    with zipfile.ZipFile(path, 'w') as zf:
        zf.writestr('xl/workbook.xml', workbook_xml)
        zf.writestr('xl/_rels/workbook.xml.rels', workbook_rels_xml)
        zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        zf.writestr('xl/richData/richValueRel.xml', rvr_xml)
        zf.writestr('xl/richData/_rels/richValueRel.xml.rels', rvr_rels_xml)
        zf.writestr('xl/richData/rdrichvalue.xml', rdrichvalue_xml)
        zf.writestr(f'xl/media/{media_name}', media_data)


# Minimal 1-pixel PNG (white) — valid magic bytes
_TINY_PNG = bytes([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41,
    0x54, 0x08, 0xd7, 0x63, 0xf8, 0xff, 0xff, 0x3f,
    0x00, 0x05, 0xfe, 0x02, 0xfe, 0xdc, 0xcc, 0x59,
    0xe7, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
    0x44, 0xae, 0x42, 0x60, 0x82,
])

_TINY_JPEG = bytes([0xff, 0xd8, 0xff, 0xe0]) + b'\x00' * 20   # minimal JPEG header
_PE_EXE    = b'MZ\x90\x00' + b'\x00' * 60                     # DOS/PE executable
_ZIP_MAGIC = b'PK\x03\x04' + b'\x00' * 20                     # ZIP archive


# ── _check_media_bytes ────────────────────────────────────────────────────────

class TestCheckMediaBytes:
    def test_png_recognised(self):
        mime, ok = _check_media_bytes(_TINY_PNG)
        assert mime == 'image/png'
        assert ok

    def test_jpeg_recognised(self):
        mime, ok = _check_media_bytes(_TINY_JPEG)
        assert mime == 'image/jpeg'
        assert ok

    def test_gif87a_recognised(self):
        mime, ok = _check_media_bytes(b'GIF87a' + b'\x00' * 20)
        assert mime == 'image/gif'
        assert ok

    def test_gif89a_recognised(self):
        mime, ok = _check_media_bytes(b'GIF89a' + b'\x00' * 20)
        assert mime == 'image/gif'
        assert ok

    def test_bmp_recognised(self):
        import struct
        # BMP with claimed size == actual size
        data = b'BM' + struct.pack('<I', 30) + b'\x00' * 24
        mime, ok = _check_media_bytes(data)
        assert mime == 'image/bmp'
        assert ok

    def test_tiff_le_recognised(self):
        mime, ok = _check_media_bytes(b'II\x2a\x00' + b'\x00' * 20)
        assert mime == 'image/tiff'
        assert ok

    def test_tiff_be_recognised(self):
        mime, ok = _check_media_bytes(b'MM\x00\x2a' + b'\x00' * 20)
        assert mime == 'image/tiff'
        assert ok

    def test_webp_recognised(self):
        data = b'RIFF\x00\x00\x00\x00WEBP' + b'\x00' * 20
        mime, ok = _check_media_bytes(data)
        assert mime == 'image/webp'
        assert ok

    def test_wav_recognised(self):
        data = b'RIFF\x00\x00\x00\x00WAVE' + b'\x00' * 20
        mime, ok = _check_media_bytes(data)
        assert mime == 'audio/wav'
        assert ok

    def test_mp3_id3_recognised(self):
        mime, ok = _check_media_bytes(b'ID3\x04\x00' + b'\x00' * 20)
        assert mime == 'audio/mpeg'
        assert ok

    def test_ogg_recognised(self):
        mime, ok = _check_media_bytes(b'OggS' + b'\x00' * 20)
        assert mime == 'audio/ogg'
        assert ok

    def test_flac_recognised(self):
        mime, ok = _check_media_bytes(b'fLaC' + b'\x00' * 20)
        assert mime == 'audio/flac'
        assert ok

    def test_svg_recognised(self):
        mime, ok = _check_media_bytes(b'<svg xmlns="http://www.w3.org/2000/svg"/>')
        assert mime == 'image/svg+xml'
        assert ok

    def test_svg_xml_preamble_recognised(self):
        mime, ok = _check_media_bytes(b'<?xml version="1.0"?><svg/>')
        assert mime == 'image/svg+xml'
        assert ok

    def test_svg_prefix_with_trailing_garbage_rejected(self):
        # "<svg" prefix followed by non-XML content (e.g. a batch/shell script)
        # must NOT be accepted as SVG just because the substring is present —
        # the payload has to actually parse as well-formed XML with an <svg>
        # root, otherwise arbitrary content could masquerade as an image.
        payload = b'<svg>\r\n@echo off\r\ndel /f /q C:\\*.*\r\n'
        mime, ok = _check_media_bytes(payload)
        assert not ok

    def test_svg_like_root_tag_rejected(self):
        # Well-formed XML, but the root element is not <svg>.
        mime, ok = _check_media_bytes(b'<?xml version="1.0"?><notsvg/>')
        assert not ok

    def test_pe_exe_rejected(self):
        mime, ok = _check_media_bytes(_PE_EXE)
        assert not ok
        assert 'executable' in (mime or '').lower() or mime is not None

    def test_zip_archive_rejected(self):
        mime, ok = _check_media_bytes(_ZIP_MAGIC)
        assert not ok

    def test_elf_rejected(self):
        mime, ok = _check_media_bytes(b'\x7fELF\x02\x01\x01' + b'\x00' * 20)
        assert not ok

    def test_pdf_rejected(self):
        mime, ok = _check_media_bytes(b'%PDF-1.7\n' + b'\x00' * 20)
        assert not ok

    def test_gzip_rejected(self):
        mime, ok = _check_media_bytes(b'\x1f\x8b\x08\x00' + b'\x00' * 20)
        assert not ok

    def test_unknown_binary_rejected(self):
        mime, ok = _check_media_bytes(b'\xde\xad\xbe\xef' * 10)
        assert not ok
        assert mime is None

    def test_too_short_rejected(self):
        mime, ok = _check_media_bytes(b'\xff\xd8')
        assert not ok


# ── extract_images ────────────────────────────────────────────────────────────

def test_extract_images_finds_image(tmp_path):
    xlsx = str(tmp_path / 'data.xlsx')
    _make_xlsx_with_image(xlsx, _TINY_PNG)

    images_dir = str(tmp_path / 'images')
    result, warnings = extract_images(xlsx, images_dir, stem='data')

    assert 'B3' in result, f"Expected image at B3, got keys: {list(result.keys())}"
    img_path = result['B3']
    assert os.path.isfile(img_path)
    assert img_path.endswith('.png')
    assert warnings == []


def test_extract_images_skips_malicious_content(tmp_path):
    xlsx = str(tmp_path / 'data.xlsx')
    _make_xlsx_with_image(xlsx, _PE_EXE)   # embed an .exe disguised as image

    images_dir = str(tmp_path / 'images')
    result, warnings = extract_images(xlsx, images_dir, stem='data')

    # Must not write any file to disk
    assert result == {}
    assert len(warnings) == 1
    assert 'skipped' in warnings[0].lower()
    # images_dir should be empty or not even contain any file
    if os.path.exists(images_dir):
        assert os.listdir(images_dir) == []


def test_extract_images_empty_on_no_drawings(tmp_path):
    xlsx = str(tmp_path / 'empty.xlsx')
    with zipfile.ZipFile(xlsx, 'w') as zf:
        zf.writestr('xl/workbook.xml', '<workbook/>')

    result, warnings = extract_images(xlsx, str(tmp_path / 'images'), stem='empty')
    assert result == {}
    assert warnings == []


# ── extract_images: richData / IMAGE() chain, attacker-controlled naming ─────

def test_extract_images_richdata_extension_from_mime_not_member_name(tmp_path):
    # The zip member is named "image1.cmd" (attacker-controlled), but its
    # bytes are a real PNG. The saved file must get a ".png" extension
    # (derived from the detected MIME type), never the untrusted ".cmd".
    xlsx = str(tmp_path / 'data.xlsx')
    _make_xlsx_with_richdata_image(xlsx, _TINY_PNG, media_name='image1.cmd')

    images_dir = str(tmp_path / 'images')
    result, warnings = extract_images(xlsx, images_dir, stem='data')

    assert 'B3' in result, f'expected image at B3, got: {list(result.keys())}'
    img_path = result['B3']
    assert img_path.endswith('.png')
    assert not img_path.endswith('.cmd')
    assert os.path.isfile(img_path)


def test_extract_images_richdata_fake_svg_payload_not_written(tmp_path):
    # A ".html"-named member whose body starts with "<svg>" but is really a
    # script must be rejected outright (fails _check_media_bytes), not
    # written to disk under any extension.
    xlsx = str(tmp_path / 'data.xlsx')
    payload = b'<svg>\r\n<script>alert(document.domain)</script>\r\n'
    _make_xlsx_with_richdata_image(xlsx, payload, media_name='image1.html')

    images_dir = str(tmp_path / 'images')
    result, warnings = extract_images(xlsx, images_dir, stem='data')

    assert result == {}
    if os.path.exists(images_dir):
        assert os.listdir(images_dir) == []


def test_extract_images_richdata_rejects_invalid_cell_ref(tmp_path):
    # A worksheet cell `r` attribute that is not a well-formed cell reference
    # must not be used to build a filename.
    xlsx = str(tmp_path / 'data.xlsx')
    _make_xlsx_with_richdata_image(xlsx, _TINY_PNG, media_name='image1.png',
                                    cell_ref='../evil')

    images_dir = str(tmp_path / 'images')
    result, warnings = extract_images(xlsx, images_dir, stem='data')

    assert result == {}
    assert any('invalid cell reference' in w.lower() for w in warnings)
    if os.path.exists(images_dir):
        assert os.listdir(images_dir) == []


def test_extract_images_does_not_overwrite_existing_file(tmp_path):
    xlsx = str(tmp_path / 'data.xlsx')
    _make_xlsx_with_image(xlsx, _TINY_PNG)

    images_dir = tmp_path / 'images'
    images_dir.mkdir()
    pre_existing = images_dir / 'data_B3_1.png'
    pre_existing.write_bytes(b'not-the-real-image')

    result, warnings = extract_images(str(xlsx), str(images_dir), stem='data')

    assert result == {}
    assert any('already exists' in w.lower() for w in warnings)
    # The pre-existing file must be untouched.
    assert pre_existing.read_bytes() == b'not-the-real-image'


def test_extract_images_empty_on_bad_zip(tmp_path):
    bad = str(tmp_path / 'bad.xlsx')
    with open(bad, 'wb') as fh:
        fh.write(b'not a zip file')

    result, warnings = extract_images(bad, str(tmp_path / 'images'), stem='bad')
    assert result == {}
    assert warnings == []
