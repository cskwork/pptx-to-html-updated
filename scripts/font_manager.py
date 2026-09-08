#!/usr/bin/env python3
"""
Embedded Font Extraction and Conversion
생성일: 2025-10-21
설명: PPTX에 포함된 폰트를 추출하여 WOFF 포맷으로 변환하고 @font-face CSS를 생성
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from fontTools.ttLib import TTFont


class FontManager:
    """PPTX 임베디드 폰트 추출 및 CSS 정의 생성"""

    def __init__(self, zip_ref, output_dir: Path, logger=None):
        """
        Args:
            zip_ref: PPTX ZipFile 객체
            output_dir: 출력 디렉토리
            logger: ConversionLogger 인스턴스
        """
        self.zip_ref = zip_ref
        self.output_dir = Path(output_dir)
        self.logger = logger

        self.ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        self.fonts_dir = self.output_dir / "fonts"

    def extract_embedded_fonts(self) -> List[str]:
        """
        PPTX에서 임베디드 폰트를 추출하고 WOFF로 변환

        Returns:
            @font-face CSS 정의 리스트
        """
        font_faces: List[str] = []

        try:
            font_table_xml = ET.fromstring(self.zip_ref.read('ppt/fontTable.xml'))
            rels_xml = ET.fromstring(self.zip_ref.read('ppt/_rels/fontTable.xml.rels'))
        except KeyError:
            if self.logger:
                self.logger.info("No embedded fonts found in presentation")
            return font_faces

        rel_map = self._build_relationship_map(rels_xml)

        for font in font_table_xml.findall('.//a:font', self.ns):
            typeface = font.get('typeface')
            if not typeface:
                continue

            embed_variants = [
                ('Regular', font.find('a:embedRegular', self.ns), 'normal', 'normal'),
                ('Bold', font.find('a:embedBold', self.ns), 'bold', 'normal'),
                ('Italic', font.find('a:embedItalic', self.ns), 'normal', 'italic'),
                ('BoldItalic', font.find('a:embedBoldItalic', self.ns), 'bold', 'italic')
            ]

            for variant_name, embed_elem, font_weight, font_style in embed_variants:
                if embed_elem is None:
                    continue

                rel_id = embed_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                key = embed_elem.get('id')
                if not rel_id or not key:
                    continue

                target = rel_map.get(rel_id)
                if not target:
                    if self.logger:
                        self.logger.warning(f"Font relationship not resolved for {typeface} ({variant_name})")
                    continue

                font_bytes = self._load_font_data(target, key)
                if not font_bytes:
                    continue

                css_def = self._convert_to_woff(
                    font_bytes,
                    typeface,
                    variant_name.lower(),
                    font_weight,
                    font_style
                )

                if css_def:
                    font_faces.append(css_def)

        return font_faces

    def _build_relationship_map(self, rels_xml: ET.Element) -> Dict[str, str]:
        """Relationship XML을 기반으로 ID → Target 매핑 생성"""
        rel_map = {}
        for rel in rels_xml.findall('.//r:Relationship', self.ns):
            rel_id = rel.get('Id')
            target = rel.get('Target')
            if rel_id and target:
                normalized = f"ppt/{target.replace('..', '').lstrip('/')}"
                rel_map[rel_id] = normalized
        return rel_map

    def _load_font_data(self, target: str, guid_key: str) -> Optional[bytes]:
        """
        임베디드 폰트 데이터를 로드하고 ODTTF일 경우 디옵퓨스케이트 처리

        Args:
            target: ZIP 내부 경로
            guid_key: GUID 문자열

        Returns:
            디옵퓨스케이트된 폰트 데이터 (bytes)
        """
        try:
            raw_data = bytearray(self.zip_ref.read(target))
        except KeyError:
            if self.logger:
                self.logger.warning(f"Embedded font file missing: {target}")
            return None

        suffix = Path(target).suffix.lower()
        if suffix == '.odttf':
            key_bytes = self._guid_to_bytes(guid_key)
            if not key_bytes:
                if self.logger:
                    self.logger.warning(f"Invalid GUID for embedded font: {guid_key}")
                return None

            for idx in range(min(32, len(raw_data))):
                raw_data[idx] ^= key_bytes[idx % len(key_bytes)]

        return bytes(raw_data)

    def _convert_to_woff(self, font_bytes: bytes, typeface: str, variant: str,
                          font_weight: str, font_style: str) -> Optional[str]:
        """
        TTF/OTF 데이터를 WOFF로 변환하고 CSS 정의 반환
        """
        try:
            font_stream = io.BytesIO(font_bytes)
            tt_font = TTFont(font_stream)

            family_name = self._extract_font_family(tt_font) or typeface

            safe_name = self._slugify(f"{family_name}-{variant}")
            self.fonts_dir.mkdir(parents=True, exist_ok=True)
            output_path = self.fonts_dir / f"{safe_name}.woff"

            tt_font.flavor = 'woff'
            tt_font.save(output_path)

            css = (
                "@font-face {\n"
                f"    font-family: '{family_name}';\n"
                f"    src: url('fonts/{output_path.name}') format('woff');\n"
                f"    font-weight: {font_weight};\n"
                f"    font-style: {font_style};\n"
                "    font-display: swap;\n"
                "}"
            )

            if self.logger:
                self.logger.info(f"Embedded font exported: {family_name} ({variant})")

            return css

        except Exception as exc:
            if self.logger:
                self.logger.error(f"Failed to convert embedded font '{typeface} ({variant})'", exception=exc)
            return None

    @staticmethod
    def _extract_font_family(tt_font: TTFont) -> Optional[str]:
        """TTFont에서 family name 추출 (NameID 1 우선)"""
        name_table = tt_font['name']

        for record in name_table.names:
            if record.nameID == 1:
                try:
                    return record.toUnicode()
                except UnicodeDecodeError:
                    continue

        return None

    @staticmethod
    def _guid_to_bytes(guid_str: str) -> Optional[bytearray]:
        """GUID 문자열을 바이트 배열로 변환"""
        cleaned = guid_str.strip('{}').replace('-', '')
        if len(cleaned) != 32:
            return None
        try:
            return bytearray.fromhex(cleaned)
        except ValueError:
            return None

    @staticmethod
    def _slugify(value: str) -> str:
        """파일명 안전 문자열 생성"""
        allowed = []
        for ch in value.lower():
            if ch.isalnum():
                allowed.append(ch)
            elif ch in (' ', '-', '_'):
                allowed.append('-')
        slug = ''.join(allowed).strip('-')
        return slug or 'font'
