"""Tests for ZIP-based image extraction (extract_images) in engine.py."""
import os
import sys
import zipfile
import tempfile
import pytest

from grepxcel.engine import extract_images


def _make_xlsx_with_image(path: str, png_bytes: bytes) -> None:
    """Build a minimal xlsx ZIP with one embedded image anchored at B3 (col=1, row=2 0-based)."""
    import xml.etree.ElementTree as ET

    # Minimal OOXML parts needed for extract_images to find the image.
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


# Minimal 1-pixel PNG (white).
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


def test_extract_images_finds_image(tmp_path):
    xlsx = str(tmp_path / 'data.xlsx')
    _make_xlsx_with_image(xlsx, _TINY_PNG)

    images_dir = str(tmp_path / 'images')
    result = extract_images(xlsx, images_dir, stem='data')

    assert 'B3' in result, f"Expected image at B3, got keys: {list(result.keys())}"
    img_path = result['B3']
    assert os.path.isfile(img_path)
    assert img_path.endswith('.png')


def test_extract_images_empty_on_no_drawings(tmp_path):
    xlsx = str(tmp_path / 'empty.xlsx')
    # Create a ZIP with no drawings at all.
    with zipfile.ZipFile(xlsx, 'w') as zf:
        zf.writestr('xl/workbook.xml', '<workbook/>')

    result = extract_images(xlsx, str(tmp_path / 'images'), stem='empty')
    assert result == {}


def test_extract_images_empty_on_bad_zip(tmp_path):
    bad = str(tmp_path / 'bad.xlsx')
    with open(bad, 'wb') as fh:
        fh.write(b'not a zip file')

    result = extract_images(bad, str(tmp_path / 'images'), stem='bad')
    assert result == {}
