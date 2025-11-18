#!/usr/bin/env python3
"""
Enhanced PowerPoint to HTML Converter - Phase 2
생성일: 2025-01-21
설명: Chart.js, 커스텀 도형, SmartArt, 애니메이션을 지원하는 고급 변환기

Phase 2 Features:
- Chart rendering with Chart.js
- Custom shape geometries (SVG)
- SmartArt text extraction
- Animation preservation (CSS/JavaScript)
- Shadow and reflection effects
- 150 DPI image export
- Comprehensive logging and error handling
"""

import html
import json
import math
import mimetypes
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from urllib.parse import parse_qs, urlparse
import xml.etree.ElementTree as ET

# 모듈 임포트
from logger import ConversionLogger, get_logger, LogLevel
from chart_extractor import ChartExtractor
from shape_geometry import ShapeGeometryConverter
from smartart_parser import SmartArtParser
from animation_handler import AnimationHandler
from font_manager import FontManager
from transition_handler import TransitionHandler


class EnhancedPPTXToHTMLV2:
    """Phase 2 기능을 포함한 향상된 PPTX to HTML 변환기"""

    def __init__(self, pptx_path, output_dir=None, dpi=150, log_file=None):
        """
        Args:
            pptx_path: PPTX 파일 경로
            output_dir: 출력 디렉토리 (기본값: PPTX 파일과 같은 위치)
            dpi: 이미지 DPI (기본값: 150, 기존 72에서 향상)
            log_file: 로그 파일 경로 (옵션)
        """
        self.pptx_path = Path(pptx_path)
        self.output_dir = Path(output_dir) if output_dir else self.pptx_path.parent
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.image_dpi = dpi
        self.layout_dpi = 96  # 브라우저 렌더링을 위한 고정 DPI

        # 로거 초기화
        self.logger = get_logger(log_file, LogLevel.INFO)
        self.logger.info(f"Initializing Enhanced PPTX Converter (layout DPI: {self.layout_dpi}, image DPI: {dpi})")

        self.slides_data = []
        self.slide_size = {'width': 10080000, 'height': 7560000}  # 기본 16:9 EMU

        # XML 네임스페이스
        self.ns = {
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
            'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        }

        # 관계 파일 캐시
        self.rels_cache = {}

        # 차트 카운터
        self.chart_counter = 0
        self.z_index_counter = 1
        self.theme_colors = self._default_theme_colors()
        self.theme_background_fills: List[Dict] = []
        self.layout_background_cache: Dict[str, Optional[Dict]] = {}
        self.master_background_cache: Dict[str, Optional[Dict]] = {}
        self.layout_placeholder_cache: Dict[str, Dict[Tuple[str, str, str, str], Dict]] = {}
        self.master_placeholder_cache: Dict[str, Dict[Tuple[str, str, str, str], Dict]] = {}
        self.transition_handler = TransitionHandler(self.ns, self.logger)

    # === 유틸리티 메서드 ===

    @staticmethod
    def emu_to_px(emu: int, dpi: int = 96) -> float:
        """EMU를 픽셀로 변환 (DPI 반영)"""
        return (emu / 914400) * dpi

    def emu_to_layout_px(self, emu: int) -> float:
        """EMU를 픽셀로 변환 (레이아웃 DPI 사용)"""
        return self.emu_to_px(emu, self.layout_dpi)

    def _next_z_index(self) -> int:
        """다음 z-index 값 반환"""
        value = self.z_index_counter
        self.z_index_counter += 1
        return value

    @staticmethod
    def _default_theme_colors() -> Dict[str, str]:
        """테마 컬러 기본값 초기화"""
        return {
            'tx1': '#000000',
            'tx2': '#000000',
            'bg1': '#FFFFFF',
            'bg2': '#FFFFFF',
            'accent1': '#4472C4',
            'accent2': '#ED7D31',
            'accent3': '#A5A5A5',
            'accent4': '#FFC000',
            'accent5': '#5B9BD5',
            'accent6': '#70AD47',
            'hlink': '#0563C1',
            'folHlink': '#954F72'
        }

    def _load_theme(self, zip_ref):
        """프레젠테이션 테마 컬러 및 배경 스타일 로드"""
        try:
            pres_rels = ET.fromstring(zip_ref.read('ppt/_rels/presentation.xml.rels'))
        except KeyError:
            return

        theme_path = None
        for rel in pres_rels.findall('.//rel:Relationship', self.ns):
            rel_type = rel.get('Type', '')
            if rel_type.endswith('/theme'):
                target = rel.get('Target')
                if target:
                    theme_path = f"ppt/{target.replace('..', '').lstrip('/')}"
                    break

        if not theme_path:
            return

        try:
            theme_xml = ET.fromstring(zip_ref.read(theme_path))
        except KeyError:
            return

        clr_scheme = theme_xml.find('.//a:themeElements/a:clrScheme', self.ns)
        if clr_scheme is not None:
            for color_child in clr_scheme:
                name = self._strip_namespace(color_child.tag)
                resolved = self._resolve_color(color_child, self.theme_colors)
                if resolved:
                    self.theme_colors[name] = resolved

        bg_fill_lst = theme_xml.find('.//a:themeElements/a:fmtScheme/a:bgFillStyleLst', self.ns)
        background_fills: List[Dict] = []
        if bg_fill_lst is not None:
            for fill in bg_fill_lst:
                tag = self._strip_namespace(fill.tag)
                if tag == 'solidFill':
                    background_fills.append({
                        'type': 'solid',
                        'color': self.color_to_hex(fill)
                    })
                elif tag == 'gradFill':
                    stops = []
                    for gs in fill.findall('.//a:gs', self.ns):
                        pos = float(gs.get('pos', 0)) / 100000
                        color = self.color_to_hex(gs)
                        stops.append((pos, color))
                    background_fills.append({
                        'type': 'gradient',
                        'stops': stops
                    })
        if background_fills:
            self.theme_background_fills = background_fills

    def color_to_hex(self, color_elem) -> str:
        """색상 엘리먼트를 hex 코드로 변환"""
        if color_elem is None:
            return '#000000'

        resolved = self._resolve_color(color_elem, self.theme_colors)
        return resolved or '#000000'

    @staticmethod
    def _clamp_color(value: float) -> int:
        """0-255 범위로 색상 값 클램프"""
        return max(0, min(255, int(round(value))))

    def _apply_color_modifiers(self, base_hex: str, modifier_elem: ET.Element) -> str:
        """색상 수정 요소(lumMod, tint 등)를 적용"""
        base_hex = base_hex.lstrip('#') or '000000'
        r = int(base_hex[0:2], 16)
        g = int(base_hex[2:4], 16)
        b = int(base_hex[4:6], 16)

        lum_mod = 1.0
        lum_off = 0.0

        for child in list(modifier_elem):
            tag = self._strip_namespace(child.tag)
            val_attr = child.get('val')
            if val_attr is None:
                continue
            try:
                raw_val = int(val_attr)
            except ValueError:
                continue

            if tag == 'lumMod':
                lum_mod *= raw_val / 100000
            elif tag == 'lumOff':
                lum_off += raw_val / 100000
            elif tag == 'tint':
                ratio = raw_val / 100000
                r = self._clamp_color(r + (255 - r) * ratio)
                g = self._clamp_color(g + (255 - g) * ratio)
                b = self._clamp_color(b + (255 - b) * ratio)
            elif tag == 'shade':
                ratio = 1 - (raw_val / 100000)
                r = self._clamp_color(r * ratio)
                g = self._clamp_color(g * ratio)
                b = self._clamp_color(b * ratio)

        if lum_mod != 1.0 or lum_off != 0.0:
            r = self._clamp_color(r * lum_mod + 255 * lum_off)
            g = self._clamp_color(g * lum_mod + 255 * lum_off)
            b = self._clamp_color(b * lum_mod + 255 * lum_off)

        return f"#{r:02X}{g:02X}{b:02X}"

    def _resolve_color(self, color_elem: ET.Element, theme_map: Dict[str, str]) -> Optional[str]:
        """색상 요소를 실제 HEX 값으로 해석"""
        # 직접 SRGB 값
        srgb = color_elem.find('.//a:srgbClr', self.ns)
        if srgb is not None and srgb.get('val'):
            base = srgb.get('val')
            return self._apply_color_modifiers(base, srgb)

        # 테마 색상
        scheme = color_elem.find('.//a:schemeClr', self.ns)
        if scheme is not None:
            scheme_name = scheme.get('val', '')
            base_hex = theme_map.get(scheme_name, theme_map.get(scheme_name.lower()))
            if base_hex:
                return self._apply_color_modifiers(base_hex.lstrip('#'), scheme)

        # 시스템 색상
        sys_clr = color_elem.find('.//a:sysClr', self.ns)
        if sys_clr is not None:
            last = sys_clr.get('lastClr', '000000')
            return self._apply_color_modifiers(last, sys_clr)

        # 미리 설정된 색상
        prst = color_elem.find('.//a:prstClr', self.ns)
        if prst is not None and prst.get('val'):
            base_hex = theme_map.get(prst.get('val'), '#000000')
            return self._apply_color_modifiers(base_hex.lstrip('#'), prst)

        return None

    # === 기존 메서드들 (PREVIOUS CODE에서 복사) ===

    def _placeholder_key(self, ph_elem: ET.Element) -> Tuple[str, str, str, str]:
        """플레이스홀더 식별 키 생성"""
        ph_type = ph_elem.get('type', 'body')
        ph_idx = ph_elem.get('idx', '')
        ph_orient = ph_elem.get('orient', '')
        ph_sz = ph_elem.get('sz', '')
        return (ph_type or 'body', ph_idx or '', ph_orient or '', ph_sz or '')

    def _get_xfrm_element(self, elem: Optional[ET.Element]) -> Optional[ET.Element]:
        """요소에서 xfrm 엘리먼트 추출"""
        if elem is None:
            return None
        tag = getattr(elem, 'tag', '')
        if tag.endswith('}xfrm'):
            return elem
        return elem.find('a:xfrm', self.ns)

    def _extract_transform_components(self, elem: Optional[ET.Element]) -> Dict[str, Optional[float]]:
        """xfrm 요소에서 좌표/크기/회전 값을 추출"""
        components: Dict[str, Optional[float]] = {
            'x': None,
            'y': None,
            'width': None,
            'height': None,
            'rotation': None
        }

        xfrm = self._get_xfrm_element(elem)
        if xfrm is None:
            return components

        off = xfrm.find('a:off', self.ns)
        if off is not None:
            x_val = off.get('x')
            y_val = off.get('y')
            try:
                if x_val is not None:
                    components['x'] = self.emu_to_layout_px(int(x_val))
                if y_val is not None:
                    components['y'] = self.emu_to_layout_px(int(y_val))
            except ValueError:
                pass

        ext = xfrm.find('a:ext', self.ns)
        if ext is not None:
            width_val = ext.get('cx')
            height_val = ext.get('cy')
            try:
                if width_val is not None:
                    components['width'] = self.emu_to_layout_px(int(width_val))
                if height_val is not None:
                    components['height'] = self.emu_to_layout_px(int(height_val))
            except ValueError:
                pass

        rot = xfrm.get('rot')
        if rot is not None:
            try:
                components['rotation'] = int(rot) / 60000
            except ValueError:
                components['rotation'] = None

        return components

    def _get_transform_source_for_element(self, element: ET.Element) -> Optional[ET.Element]:
        """요소 유형에 따른 변환 소스 반환"""
        tag = self._strip_namespace(element.tag)
        if tag == 'graphicFrame':
            xfrm = element.find('p:xfrm', self.ns)
            if xfrm is not None:
                return xfrm

        sp_pr = element.find('p:spPr', self.ns)
        if sp_pr is None:
            sp_pr = element.find('pic:spPr', self.ns)
        if sp_pr is not None:
            xfrm = sp_pr.find('a:xfrm', self.ns)
            return xfrm if xfrm is not None else sp_pr

        xfrm = element.find('a:xfrm', self.ns)
        if xfrm is not None:
            return xfrm

        return None

    def _collect_placeholder_info(self, container: Optional[ET.Element]) -> Dict[Tuple[str, str, str, str], Dict]:
        """주어진 spTree에서 플레이스홀더 정보를 수집"""
        placeholders: Dict[Tuple[str, str, str, str], Dict] = {}
        if container is None:
            return placeholders

        for child in container:
            tag = self._strip_namespace(child.tag)
            if tag in {'nvGrpSpPr', 'grpSpPr'}:
                continue

            ph_elem = child.find('.//p:nvPr/p:ph', self.ns)
            if ph_elem is not None:
                key = self._placeholder_key(ph_elem)
                if key not in placeholders:
                    transform_elem = self._get_transform_source_for_element(child)
                    if transform_elem is not None:
                        placeholders[key] = {
                            'position': self.extract_shape_position(transform_elem),
                            'tag': tag
                        }

            if tag == 'grpSp':
                nested = self._collect_placeholder_info(child)
                for nested_key, value in nested.items():
                    placeholders.setdefault(nested_key, value)

        return placeholders

    def _get_layout_master_path(self, zip_ref, layout_path: str) -> Optional[str]:
        """레이아웃에 연결된 마스터 경로 반환"""
        layout_rels_path = layout_path.replace('slideLayouts/', 'slideLayouts/_rels/').replace('.xml', '.xml.rels')
        rels_tree = self.get_relationships(zip_ref, layout_rels_path)
        if rels_tree is None:
            return None

        for rel in rels_tree.findall('.//rel:Relationship', self.ns):
            rel_type = rel.get('Type', '')
            if rel_type.endswith('/slideMaster'):
                target = rel.get('Target')
                if target:
                    return f"ppt/{target.replace('..', '').lstrip('/')}"

        return None

    def _build_placeholder_context(self, zip_ref, slide_rels_path: str,
                                   layout_path: Optional[str] = None) -> Dict[str, Dict]:
        """슬라이드에 적용되는 레이아웃/마스터 플레이스홀더 맵 구성"""
        context = {'layout': {}, 'master': {}}

        if layout_path is None:
            layout_path = self._get_related_slide_layout(zip_ref, slide_rels_path)
        if layout_path:
            if layout_path not in self.layout_placeholder_cache:
                try:
                    layout_xml = ET.fromstring(zip_ref.read(layout_path))
                    sp_tree = layout_xml.find('.//p:cSld/p:spTree', self.ns)
                    self.layout_placeholder_cache[layout_path] = self._collect_placeholder_info(sp_tree)
                except Exception:
                    self.layout_placeholder_cache[layout_path] = {}
            context['layout'] = self.layout_placeholder_cache.get(layout_path, {})

            master_path = self._get_layout_master_path(zip_ref, layout_path)
            if master_path:
                if master_path not in self.master_placeholder_cache:
                    try:
                        master_xml = ET.fromstring(zip_ref.read(master_path))
                        sp_tree = master_xml.find('.//p:cSld/p:spTree', self.ns)
                        self.master_placeholder_cache[master_path] = self._collect_placeholder_info(sp_tree)
                    except Exception:
                        self.master_placeholder_cache[master_path] = {}
                context['master'] = self.master_placeholder_cache.get(master_path, {})

        return context

    def _apply_placeholder_inheritance(self, element: ET.Element, position: Dict,
                                       placeholder_context: Dict[str, Dict]) -> Dict:
        """슬라이드 요소에 레이아웃/마스터 플레이스홀더 좌표를 적용"""
        if not placeholder_context:
            return position

        ph_elem = element.find('.//p:nvPr/p:ph', self.ns)
        if ph_elem is None:
            return position

        key = self._placeholder_key(ph_elem)
        base_info = (placeholder_context.get('layout') or {}).get(key)
        if base_info is None:
            base_info = (placeholder_context.get('master') or {}).get(key)

        if base_info is None:
            return position

        merged = dict(position)
        base_position = base_info.get('position', {})
        for field in ('x', 'y', 'width', 'height', 'rotation'):
            value = base_position.get(field)
            if value is not None:
                merged[field] = value

        transform_elem = self._get_transform_source_for_element(element)
        overrides = self._extract_transform_components(transform_elem)
        for field, value in overrides.items():
            if value is not None:
                merged[field] = value

        if merged.get('rotation') is None:
            merged['rotation'] = 0.0

        return merged

    def _ensure_position_defaults(self, position: Optional[Dict]) -> Dict:
        """위치 정보에 기본 피벗/회전 값을 보장"""
        if position is None:
            return {'x': 0.0, 'y': 0.0, 'width': 0.0, 'height': 0.0, 'rotation': 0.0,
                    'pivot_x': 0.0, 'pivot_y': 0.0}

        position.setdefault('x', 0.0)
        position.setdefault('y', 0.0)
        position.setdefault('width', 0.0)
        position.setdefault('height', 0.0)
        position.setdefault('rotation', 0.0)

        width = self._safe_dimension(position.get('width'))
        height = self._safe_dimension(position.get('height'))

        if position.get('pivot_x') is None:
            position['pivot_x'] = position['x'] + width / 2.0
        if position.get('pivot_y') is None:
            position['pivot_y'] = position['y'] + height / 2.0

        return position

    def _align_position_to_pivot(self, position: Dict) -> Dict:
        """피벗 기준으로 좌상단 좌표를 재조정"""
        position = self._ensure_position_defaults(position)
        width = self._safe_dimension(position.get('width'))
        height = self._safe_dimension(position.get('height'))
        pivot_x = position.get('pivot_x', 0.0)
        pivot_y = position.get('pivot_y', 0.0)
        position['x'] = pivot_x - width / 2.0
        position['y'] = pivot_y - height / 2.0
        return position

    def get_relationships(self, zip_ref, rels_path):
        """관계 파일 파싱 및 캐싱"""
        if rels_path in self.rels_cache:
            return self.rels_cache[rels_path]

        try:
            rels_content = zip_ref.read(rels_path)
            rels_tree = ET.fromstring(rels_content)
            self.rels_cache[rels_path] = rels_tree
            return rels_tree
        except Exception as e:
            self.logger.error(f"Failed to load relationships: {rels_path}", exception=e)
            return None

    def resolve_relationship(self, zip_ref, rels_path, rel_id):
        """관계 ID를 타겟 경로로 해결"""
        rels_tree = self.get_relationships(zip_ref, rels_path)
        if rels_tree is None:
            return None, None, None

        for rel in rels_tree.findall('.//rel:Relationship', self.ns):
            if rel.get('Id') == rel_id:
                return rel.get('Target'), rel.get('Type'), rel.get('TargetMode')

        return None, None, None

    # === 도형 추출 (기존 로직 + Phase 2 개선사항) ===

    def extract_shape_position(self, sp_pr) -> Dict:
        """도형 위치 및 크기 추출"""
        position = {'x': 0, 'y': 0, 'width': 0, 'height': 0, 'rotation': 0}

        if sp_pr is None:
            return position

        xfrm = None
        # sp_pr may already be an a:xfrm element (e.g., tables within graphicFrame)
        tag = sp_pr.tag if hasattr(sp_pr, 'tag') else None
        if tag and tag.endswith('}xfrm'):
            xfrm = sp_pr
        else:
            xfrm = sp_pr.find('.//a:xfrm', self.ns)

        if xfrm is not None:
            off = xfrm.find('a:off', self.ns)
            if off is not None:
                position['x'] = self.emu_to_layout_px(int(off.get('x', 0)))
                position['y'] = self.emu_to_layout_px(int(off.get('y', 0)))

            ext = xfrm.find('a:ext', self.ns)
            if ext is not None:
                position['width'] = self.emu_to_layout_px(int(ext.get('cx', 0)))
                position['height'] = self.emu_to_layout_px(int(ext.get('cy', 0)))

            rot = xfrm.get('rot')
            if rot:
                position['rotation'] = int(rot) / 60000

            if xfrm.get('flipH') == '1' or xfrm.get('flipH') == 'true':
                position['flip_h'] = True
            if xfrm.get('flipV') == '1' or xfrm.get('flipV') == 'true':
                position['flip_v'] = True

        if position['width'] and position['height']:
            position['pivot_x'] = position['x'] + position['width'] / 2.0
            position['pivot_y'] = position['y'] + position['height'] / 2.0
        else:
            position['pivot_x'] = position['x']
            position['pivot_y'] = position['y']

        return position

    def extract_shape_fill(self, sp_pr) -> Dict:
        """도형 채우기 추출"""
        fill = {
            'type': 'none',
            'color': '#FFFFFF',
            'gradient': None,
            'gradient_angle': 0,
            'gradient_type': 'linear'
        }

        if sp_pr is None:
            return fill

        if sp_pr.find('a:noFill', self.ns) is not None:
            return fill

        solid = sp_pr.find('a:solidFill', self.ns)
        if solid is not None:
            fill['type'] = 'solid'
            fill['color'] = self.color_to_hex(solid)
            return fill

        blip_fill = sp_pr.find('a:blipFill', self.ns)
        if blip_fill is not None:
            blip = blip_fill.find('a:blip', self.ns)
            if blip is not None:
                rel_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                src_rect = blip_fill.find('a:srcRect', self.ns)
                crop = {}
                if src_rect is not None:
                    for attr in ('l', 'r', 't', 'b'):
                        if src_rect.get(attr) is not None:
                            try:
                                crop[attr] = int(src_rect.get(attr)) / 100000
                            except ValueError:
                                continue
                stretch = blip_fill.find('a:stretch', self.ns) is not None
                return {
                    'type': 'image',
                    'rel_id': rel_id,
                    'crop': crop,
                    'stretch': stretch
                }

        grad = sp_pr.find('a:gradFill', self.ns)
        if grad is not None:
            fill['type'] = 'gradient'
            stops = []
            for gs in grad.findall('.//a:gs', self.ns):
                pos = float(gs.get('pos', 0)) / 100000
                color = self.color_to_hex(gs)
                stops.append((pos, color))
            lin = grad.find('.//a:lin', self.ns)
            if lin is not None and lin.get('ang'):
                try:
                    fill['gradient_angle'] = int(lin.get('ang', 0)) / 60000
                except ValueError:
                    fill['gradient_angle'] = 0

            path = grad.find('.//a:path', self.ns)
            if path is not None and path.get('path'):
                fill['gradient_type'] = path.get('path')

            fill['gradient'] = stops
            return fill

        no_fill = sp_pr.find('.//a:noFill', self.ns)
        if no_fill is not None:
            fill['type'] = 'none'

        return fill

    def extract_shape_border(self, sp_pr) -> Dict:
        """도형 테두리 추출"""
        border = {'width': 0, 'color': '#000000', 'style': 'solid'}

        if sp_pr is None:
            return border

        ln = sp_pr.find('.//a:ln', self.ns)
        if ln is not None:
            w = ln.get('w')
            if w:
                border['width'] = self.emu_to_layout_px(int(w))

            solid = ln.find('.//a:solidFill', self.ns)
            if solid is not None:
                border['color'] = self.color_to_hex(solid)

            dash = ln.find('.//a:prstDash', self.ns)
            if dash is not None:
                dash_val = dash.get('val', 'solid')
                border['style'] = 'dashed' if 'dash' in dash_val else 'solid'

        return border

    # === 텍스트 추출 (하이퍼링크 포함) ===

    @staticmethod
    def _strip_namespace(tag: str) -> str:
        """네임스페이스 제거"""
        return tag.split('}')[-1] if '}' in tag else tag

    def _hpt_to_px(self, value: int) -> float:
        """1/100 포인트 값을 픽셀로 변환"""
        points = value / 100.0
        return points * (self.layout_dpi / 72.0)

    def _pt_to_px(self, value: float) -> float:
        """포인트 값을 픽셀로 변환"""
        return float(value) * (self.layout_dpi / 72.0)

    def _default_indent_px(self, level: int) -> float:
        """레벨 기반 기본 들여쓰기"""
        base_emu = 457200  # 0.5 inch
        return self.emu_to_layout_px(base_emu * max(level, 0))

    def _parse_line_spacing(self, ln_spc_elem) -> Optional[Dict]:
        """줄 간격 파싱"""
        if ln_spc_elem is None:
            return None

        spc_pct = ln_spc_elem.find('a:spcPct', self.ns)
        if spc_pct is not None and spc_pct.get('val'):
            try:
                value = int(spc_pct.get('val')) / 100000
                return {'value': value, 'unit': 'multiple'}
            except ValueError:
                return None

        spc_pts = ln_spc_elem.find('a:spcPts', self.ns)
        if spc_pts is not None and spc_pts.get('val'):
            try:
                value_px = self._hpt_to_px(int(spc_pts.get('val')))
                return {'value': value_px, 'unit': 'px'}
            except ValueError:
                return None

        return None

    def _parse_spacing(self, spc_elem) -> float:
        """문단 앞/뒤 간격 파싱 (픽셀 반환)"""
        if spc_elem is None:
            return 0.0

        spc_pts = spc_elem.find('a:spcPts', self.ns)
        if spc_pts is not None and spc_pts.get('val'):
            try:
                return self._hpt_to_px(int(spc_pts.get('val')))
            except ValueError:
                return 0.0

        spc_pct = spc_elem.find('a:spcPct', self.ns)
        if spc_pct is not None and spc_pct.get('val'):
            try:
                percent = int(spc_pct.get('val')) / 1000  # 예: 120000 -> 120%
                default_line_height = self._hpt_to_px(1200)  # 약 12pt
                return default_line_height * (percent / 100)
            except ValueError:
                return 0.0

        return 0.0

    def _parse_paragraph_properties(self, p_pr, level_hint: int) -> Dict:
        """문단 속성 파싱"""
        alignment = 'left'
        level = level_hint
        list_type = None
        bullet_char = None
        bullet_font = None
        numbering = None

        margin_left = None
        margin_right = None
        text_indent = 0.0

        space_before = 0.0
        space_after = 0.0
        line_spacing = None

        if p_pr is not None:
            align = p_pr.get('algn')
            if align == 'ctr':
                alignment = 'center'
            elif align == 'r':
                alignment = 'right'
            elif align == 'just':
                alignment = 'justify'

            if p_pr.get('lvl') is not None:
                try:
                    level = int(p_pr.get('lvl'))
                except ValueError:
                    level = level_hint

            bu_none = p_pr.find('a:buNone', self.ns)
            if bu_none is None:
                bu_auto = p_pr.find('a:buAutoNum', self.ns)
                bu_char = p_pr.find('a:buChar', self.ns)

                if bu_auto is not None:
                    list_type = 'number'
                    numbering = {
                        'type': bu_auto.get('type', 'arabicPeriod'),
                        'start_at': int(bu_auto.get('startAt', 1)),
                        'level': level
                    }
                elif bu_char is not None:
                    list_type = 'bullet'
                    bullet_char = bu_char.get('char', '•')
                    bu_font = p_pr.find('a:buFont', self.ns)
                    if bu_font is not None:
                        bullet_font = bu_font.get('typeface')

            mar_l = p_pr.get('marL')
            if mar_l is not None:
                try:
                    margin_left = self.emu_to_layout_px(int(mar_l))
                except ValueError:
                    margin_left = None

            mar_r = p_pr.get('marR')
            if mar_r is not None:
                try:
                    margin_right = self.emu_to_layout_px(int(mar_r))
                except ValueError:
                    margin_right = None

            indent = p_pr.get('indent')
            if indent is not None:
                try:
                    text_indent = self.emu_to_layout_px(int(indent))
                except ValueError:
                    text_indent = 0.0

            space_before = self._parse_spacing(p_pr.find('a:spcBef', self.ns))
            space_after = self._parse_spacing(p_pr.find('a:spcAft', self.ns))
            line_spacing = self._parse_line_spacing(p_pr.find('a:lnSpc', self.ns))

        if margin_left is None:
            margin_left = self._default_indent_px(level)

        return {
            'alignment': alignment,
            'level': level,
            'list_type': list_type,
            'bullet_char': bullet_char,
            'bullet_font': bullet_font,
            'numbering': numbering,
            'margin_left': margin_left,
            'margin_right': margin_right or 0.0,
            'text_indent': text_indent,
            'space_before': space_before,
            'space_after': space_after,
            'line_spacing': line_spacing
        }

    def _parse_text_run(self, run_elem, zip_ref=None, slide_rels_path=None) -> Optional[Dict]:
        """텍스트 런 파싱"""
        t = run_elem.find('a:t', self.ns)
        if t is None or t.text is None:
            # 빈 런이라도 강제 개행 처리 가능
            if run_elem.findall('.//a:br', self.ns):
                return {'text': '<br/>', 'is_break': True}
            return None

        text = t.text
        r_pr = run_elem.find('a:rPr', self.ns)

        default_text_color = self.theme_colors.get('tx1', '#000000')

        formatting = {
            'font_family': 'Arial',
            'font_size': 18,
            'color': default_text_color,
            'bold': False,
            'italic': False,
            'underline': False
        }

        hyperlink = None

        if r_pr is not None:
            latin = r_pr.find('.//a:latin', self.ns)
            if latin is not None:
                formatting['font_family'] = latin.get('typeface', 'Arial')

            ea = r_pr.find('.//a:ea', self.ns)
            if ea is not None and ea.get('typeface'):
                formatting['font_family'] = ea.get('typeface')

            sz = r_pr.get('sz')
            if sz:
                try:
                    formatting['font_size'] = int(sz) / 100
                except ValueError:
                    formatting['font_size'] = 18

            solid = r_pr.find('.//a:solidFill', self.ns)
            if solid is not None:
                formatting['color'] = self.color_to_hex(solid)

            formatting['bold'] = r_pr.get('b') == '1'
            formatting['italic'] = r_pr.get('i') == '1'
            formatting['underline'] = r_pr.get('u') != 'none' if r_pr.get('u') else False

            if zip_ref and slide_rels_path:
                hlink = r_pr.find('.//a:hlinkClick', self.ns)
                if hlink is not None:
                    rel_id = hlink.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if rel_id:
                        target, _, _ = self.resolve_relationship(zip_ref, slide_rels_path, rel_id)
                        if target:
                            hyperlink = target

        return {
            'text': text,
            'formatting': formatting,
            'hyperlink': hyperlink,
            'is_break': False
        }

    def _build_gradient_css(self, fill: Dict) -> str:
        """그라디언트 CSS 생성"""
        stops = fill.get('gradient') or []
        if not stops:
            return 'transparent'

        stop_css = ', '.join([f"{color} {pos * 100:.2f}%" for pos, color in stops])

        gradient_type = fill.get('gradient_type', 'linear')
        if gradient_type in ('shape', 'rect', 'path'):
            return f"radial-gradient(circle, {stop_css})"

        angle = fill.get('gradient_angle', 0)
        return f"linear-gradient({angle:.2f}deg, {stop_css})"

    def _build_background_style(self, bg: Dict) -> str:
        """슬라이드 배경 CSS 문자열 생성"""
        if bg.get('type') == 'gradient' and bg.get('gradient'):
            return f"background: {self._build_gradient_css(bg)};"
        return f"background-color: {bg.get('color', '#FFFFFF')};"

    def _to_roman(self, number: int, uppercase: bool = True) -> str:
        """정수를 로마 숫자로 변환"""
        numeral_map = (
            (1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
            (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
            (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')
        )
        result = []
        n = max(1, number)
        for value, numeral in numeral_map:
            while n >= value:
                result.append(numeral)
                n -= value
        roman = ''.join(result)
        return roman if uppercase else roman.lower()

    def _to_alpha(self, number: int, uppercase: bool = True) -> str:
        """정수를 알파벳 시퀀스로 변환 (1 → A)"""
        n = max(1, number)
        chars = []
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            chars.append(chr(ord('A') + remainder))
        alpha = ''.join(reversed(chars))
        return alpha if uppercase else alpha.lower()

    def _format_number_bullet(self, number: int, num_type: str) -> str:
        """자동 번호 스타일을 텍스트로 변환"""
        num_type = num_type or 'arabicPeriod'
        mapping = {
            'arabicPeriod': f"{number}.",
            'arabicParenR': f"{number})",
            'romanUpperPeriod': f"{self._to_roman(number)}.",
            'romanLowerPeriod': f"{self._to_roman(number, uppercase=False)}.",
            'alphaUpperPeriod': f"{self._to_alpha(number)}.",
            'alphaLowerPeriod': f"{self._to_alpha(number, uppercase=False)}.",
            'alphaLowerParenR': f"{self._to_alpha(number, uppercase=False)})",
            'alphaUpperParenR': f"{self._to_alpha(number)})"
        }

        return mapping.get(num_type, f"{number}.")

    @staticmethod
    def _safe_dimension(value: Optional[float]) -> float:
        """치수 값 안전 반환"""
        if value is None:
            return 0.0
        return float(value)

    def _apply_single_transform(self, position: Dict, transform: Dict) -> Dict:
        """단일 그룹 변환을 위치에 적용"""
        pos = {
            'x': self._safe_dimension(position.get('x')),
            'y': self._safe_dimension(position.get('y')),
            'width': self._safe_dimension(position.get('width')),
            'height': self._safe_dimension(position.get('height')),
            'rotation': position.get('rotation', 0.0),
            'pivot_x': position.get('pivot_x'),
            'pivot_y': position.get('pivot_y')
        }

        offset_x = transform.get('offset_x', 0.0)
        offset_y = transform.get('offset_y', 0.0)
        origin_x = transform.get('origin_x', 0.0)
        origin_y = transform.get('origin_y', 0.0)
        scale_x = transform.get('scale_x', 1.0)
        scale_y = transform.get('scale_y', 1.0)
        rotation = transform.get('rotation', 0.0)
        transform_pivot_x = transform.get('pivot_x', offset_x + pos['width'] / 2.0)
        transform_pivot_y = transform.get('pivot_y', offset_y + pos['height'] / 2.0)

        pos['x'] = offset_x + (pos['x'] - origin_x) * scale_x
        pos['y'] = offset_y + (pos['y'] - origin_y) * scale_y
        pos['width'] *= scale_x
        pos['height'] *= scale_y

        pivot_x = position.get('pivot_x')
        pivot_y = position.get('pivot_y')
        if pivot_x is None:
            pivot_x = position.get('x', 0.0) + position.get('width', 0.0) / 2.0
        if pivot_y is None:
            pivot_y = position.get('y', 0.0) + position.get('height', 0.0) / 2.0

        pivot_x = offset_x + (pivot_x - origin_x) * scale_x
        pivot_y = offset_y + (pivot_y - origin_y) * scale_y

        if rotation:
            theta = math.radians(rotation)
            dx = pos['x'] - transform_pivot_x
            dy = pos['y'] - transform_pivot_y
            rotated_x = transform_pivot_x + dx * math.cos(theta) - dy * math.sin(theta)
            rotated_y = transform_pivot_y + dx * math.sin(theta) + dy * math.cos(theta)
            pos['x'] = rotated_x
            pos['y'] = rotated_y

            dx_p = pivot_x - transform_pivot_x
            dy_p = pivot_y - transform_pivot_y
            pivot_x = transform_pivot_x + dx_p * math.cos(theta) - dy_p * math.sin(theta)
            pivot_y = transform_pivot_y + dx_p * math.sin(theta) + dy_p * math.cos(theta)

        pos['rotation'] = pos.get('rotation', 0.0) + rotation
        pos['pivot_x'] = pivot_x
        pos['pivot_y'] = pivot_y

        return pos

    def _apply_transform_chain(self, position: Dict, transform_chain: List[Dict]) -> Dict:
        """여러 그룹 변환을 순차적으로 적용"""
        adjusted = self._ensure_position_defaults(dict(position))
        for transform in reversed(transform_chain):
            adjusted = self._apply_single_transform(adjusted, transform)
        adjusted = self._ensure_position_defaults(adjusted)
        adjusted = self._align_position_to_pivot(adjusted)
        return adjusted

    def _render_paragraphs(self, paragraphs: List[Dict], wrap_text: bool = True) -> str:
        """문단 리스트를 HTML 문자열로 변환"""
        html_parts: List[str] = []
        number_counters: Dict[int, int] = {}

        for para in paragraphs:
            runs = para.get('runs', [])
            if not runs:
                continue

            bullet_text = None
            bullet_font = para.get('bullet_font')

            if para.get('list_type') == 'number' and para.get('numbering'):
                numbering = para['numbering']
                level = numbering.get('level', 0)
                for deeper in list(number_counters.keys()):
                    if deeper > level:
                        number_counters.pop(deeper, None)

                start_at = numbering.get('start_at', 1)
                current = number_counters.get(level, start_at - 1) + 1
                number_counters[level] = current
                bullet_text = self._format_number_bullet(current, numbering.get('type'))
            elif para.get('list_type') == 'bullet':
                bullet_text = para.get('bullet_char') or '•'

            para_styles = [
                f"text-align: {para.get('alignment', 'left')}",
                f"margin-top: {para.get('space_before', 0.0):.2f}px",
                f"margin-bottom: {para.get('space_after', 0.0):.2f}px"
            ]

            margin_left = para.get('margin_left', 0.0)
            margin_right = para.get('margin_right', 0.0)
            text_indent = para.get('text_indent', 0.0)

            para_styles.append(f"margin-left: {margin_left:.2f}px")
            para_styles.append(f"margin-right: {margin_right:.2f}px")

            if text_indent:
                para_styles.append(f"text-indent: {text_indent:.2f}px")

            line_spacing = para.get('line_spacing')
            if line_spacing:
                if line_spacing.get('unit') == 'multiple':
                    para_styles.append(f"line-height: {line_spacing['value']:.3f}")
                else:
                    para_styles.append(f"line-height: {line_spacing['value']:.2f}px")

            if bullet_text is not None:
                para_styles.extend([
                    "display: flex",
                    "gap: 8px",
                    "align-items: flex-start"
                ])
                if not wrap_text:
                    para_styles.append("white-space: nowrap")
            else:
                para_styles.append("white-space: pre-wrap" if wrap_text else "white-space: nowrap")

            run_fragments: List[str] = []
            for run in runs:
                if run.get('is_break'):
                    run_fragments.append('<br>')
                    continue

                fmt = run.get('formatting', {})
                # Convert font size to px for consistent rendering
                font_size_pt = float(fmt.get('font_size', 18))
                font_size_px = self._pt_to_px(font_size_pt)
                
                span_styles = [
                    f"font-family: '{fmt.get('font_family', 'Arial')}'",
                    f"font-size: {font_size_px:.1f}px",
                    f"color: {fmt.get('color', '#000000')}"
                ]
                if fmt.get('bold'):
                    span_styles.append("font-weight: bold")
                if fmt.get('italic'):
                    span_styles.append("font-style: italic")
                if fmt.get('underline'):
                    span_styles.append("text-decoration: underline")

                sanitized = self._escape_html(run.get('text', ''))
                sanitized = sanitized.replace('\n', '<br>')

                content_html = f'<span style="{"; ".join(span_styles)}">{sanitized}</span>'
                if run.get('hyperlink'):
                    content_html = (
                        f'<a href="{run["hyperlink"]}" target="_blank" '
                        f'style="text-decoration: underline; color: {fmt.get("color", "#1a0dab")}">'
                        f'{content_html}</a>'
                    )

                run_fragments.append(content_html)

            if not run_fragments and bullet_text is None:
                continue

            if bullet_text is not None:
                bullet_styles = ["flex: 0 0 auto", "min-width: 24px"]
                if bullet_font:
                    bullet_styles.append(f"font-family: '{bullet_font}'")
                bullet_html = f'<span class="ppt-bullet" style="{"; ".join(bullet_styles)}">{self._escape_html(bullet_text)}</span>'
                text_container = (
                    f'<span class="ppt-run-container" '
                    f'style="flex: 1 1 auto; white-space: {"pre-wrap" if wrap_text else "nowrap"}; '
                    f'overflow: visible; min-width: 0;">'
                    f'{"".join(run_fragments)}</span>'
                )
                html_parts.append(
                    f'<p class="ppt-paragraph" style="{"; ".join(para_styles)}">'
                    f'{bullet_html}{text_container}</p>'
                )
            else:
                html_parts.append(
                    f'<p class="ppt-paragraph" style="{"; ".join(para_styles)}">'
                    f'{"".join(run_fragments)}</p>'
                )

        return ''.join(html_parts)

    def _build_chart_element(self, element: Dict) -> str:
        """Chart.js 요소 HTML 생성"""
        pos = element.get('position', {})
        styles = [
            "position: absolute",
            f"left: {pos.get('x', 0.0):.2f}px",
            f"top: {pos.get('y', 0.0):.2f}px",
            f"width: {pos.get('width', 0.0):.2f}px",
            f"height: {pos.get('height', 0.0):.2f}px",
            f"z-index: {element.get('z_index', 1)}"
        ]
        config_json = json.dumps(element.get('chart_config', {}), ensure_ascii=False)
        config_attr = html.escape(config_json, quote=True)
        return (
            f'<div class="ppt-element chart-element" style="{"; ".join(styles)}">'
            f'<canvas id="{element.get("chart_id", "")}" data-chart-config="{config_attr}"></canvas>'
            f'</div>'
        )

    def extract_text_with_formatting(self, shape, zip_ref=None, slide_rels_path=None) -> Tuple[List[Dict], str, Dict[str, float], Dict[str, object]]:
        """텍스트 및 서식 추출 (하이퍼링크 포함)"""
        paragraphs = []

        body_pr = shape.find('.//a:bodyPr', self.ns)
        anchor = 't'
        padding = {'left': 0.0, 'right': 0.0, 'top': 0.0, 'bottom': 0.0}
        wrap_text = True
        if body_pr is not None:
            if body_pr.get('anchor'):
                anchor = body_pr.get('anchor')
            for attr, key in [('lIns', 'left'), ('rIns', 'right'), ('tIns', 'top'), ('bIns', 'bottom')]:
                val = body_pr.get(attr)
                if val is not None:
                    try:
                        padding[key] = self.emu_to_layout_px(int(val))
                    except ValueError:
                        padding[key] = 0.0
            wrap_attr = body_pr.get('wrap')
            if wrap_attr and wrap_attr.lower() == 'none':
                wrap_text = False

        for p in shape.findall('.//a:p', self.ns):
            p_pr = p.find('a:pPr', self.ns)
            paragraph_info = self._parse_paragraph_properties(p_pr, level_hint=0)
            runs = []

            for child in p:
                tag = self._strip_namespace(child.tag)
                if tag == 'r':
                    run_data = self._parse_text_run(child, zip_ref, slide_rels_path)
                    if run_data:
                        runs.append(run_data)
                elif tag == 'br':
                    runs.append({'text': '<br/>', 'is_break': True})
                elif tag == 'fld':
                    for fld_run in child.findall('.//a:r', self.ns):
                        run_data = self._parse_text_run(fld_run, zip_ref, slide_rels_path)
                        if run_data:
                            runs.append(run_data)

            if runs:
                paragraph_info['runs'] = runs
                paragraphs.append(paragraph_info)

        max_font_px = 0.0
        total_lines = 0

        for para in paragraphs:
            runs = para.get('runs', [])
            if not runs:
                continue

            para_max_pt = 0.0
            explicit_breaks = 0
            for run in runs:
                if run.get('is_break'):
                    explicit_breaks += 1
                    continue
                fmt = run.get('formatting', {})
                try:
                    para_max_pt = max(para_max_pt, float(fmt.get('font_size', 18)))
                except (TypeError, ValueError):
                    para_max_pt = max(para_max_pt, 18.0)

            if para_max_pt <= 0.0:
                para_max_pt = 18.0

            para_max_px = self._pt_to_px(para_max_pt)
            max_font_px = max(max_font_px, para_max_px)

            line_count = explicit_breaks + 1
            line_spacing = para.get('line_spacing')
            if line_spacing:
                if line_spacing.get('unit') == 'multiple':
                    line_height_px = para_max_px * line_spacing['value']
                else:
                    line_height_px = float(line_spacing['value'])
            else:
                line_height_px = para_max_px * 1.15

            para_height = line_height_px * line_count
            para_height += para.get('space_before', 0.0) + para.get('space_after', 0.0)

            total_lines += line_count

        text_props = {
            'wrap_text': wrap_text,
            'max_font_px': max_font_px,
            'line_count': total_lines,
            'paragraph_count': len(paragraphs)
        }

        return paragraphs, anchor, padding, text_props

    # === 미디어 추출 (이미지, 비디오, 오디오) - DPI 향상 ===

    def extract_media(self, zip_ref, slide_rels_path, rel_id, slide_num, media_type='image'):
        """미디어 파일 추출 (내장 및 외부 링크 지원)"""
        try:
            target, rel_type, target_mode = self.resolve_relationship(zip_ref, slide_rels_path, rel_id)
            if not target:
                return None

            rel_type_lower = (rel_type or '').lower()
            mode_lower = (target_mode or '').lower()
            media_type_hint = (media_type or '').lower()

            is_audio = media_type_hint == 'audio' or 'audio' in rel_type_lower
            is_video = media_type_hint == 'video' or ('video' in rel_type_lower) or ('media' in rel_type_lower and not is_audio)
            is_external = mode_lower == 'external'

            normalized_target = target
            if not is_external and normalized_target.lower().startswith(('http://', 'https://')):
                is_external = True

            media_kind = 'video' if is_video else ('audio' if is_audio else 'image')
            default_mime = 'video/mp4' if is_video else ('audio/mpeg' if is_audio else 'image/png')

            if is_external:
                mime_type = self._guess_mime_type(normalized_target, rel_type, default_mime)
                if is_video:
                    self.logger.increment_video()
                elif is_audio:
                    self.logger.increment_audio()
                else:
                    self.logger.increment_image()
                return {
                    'path': normalized_target,
                    'type': media_kind,
                    'external': True,
                    'mime_type': mime_type
                }

            media_path = f"ppt/{normalized_target.replace('..', '').lstrip('/')}"

            try:
                media_data = zip_ref.read(media_path)
            except KeyError:
                self.logger.warning(f"Media file not found: {media_path}", slide_num=slide_num)
                return None

            mime_type = self._guess_mime_type(media_path, rel_type, default_mime)
            ext = Path(media_path).suffix
            if not ext and mime_type:
                guessed_ext = mimetypes.guess_extension(mime_type.split(';')[0])
                if guessed_ext:
                    ext = guessed_ext
            if not ext:
                ext = '.mp4' if is_video else ('.mp3' if is_audio else '.png')

            assets_dir = self.output_dir / "assets"
            assets_dir.mkdir(exist_ok=True)

            if is_video:
                prefix = 'video'
                self.logger.increment_video()
            elif is_audio:
                prefix = 'audio'
                self.logger.increment_audio()
            else:
                prefix = 'img'
                self.logger.increment_image()

            media_filename = f"slide{slide_num}_{prefix}_{rel_id}{ext}"
            media_file_path = assets_dir / media_filename
            media_file_path.write_bytes(media_data)

            return {
                'path': f"assets/{media_filename}",
                'type': media_kind,
                'external': False,
                'mime_type': mime_type
            }
        except Exception as e:
            self.logger.error(f"Failed to extract media", exception=e, slide_num=slide_num)
            return None

    def _guess_mime_type(self, path_or_url: str, rel_type: Optional[str], default: str) -> str:
        """파일 경로나 URL 기반 MIME 타입 추정"""
        mime, _ = mimetypes.guess_type(path_or_url)
        if mime:
            return mime

        rel_lower = (rel_type or '').lower()
        if 'video' in rel_lower:
            return 'video/mp4'
        if 'audio' in rel_lower:
            return 'audio/mpeg'
        if 'image' in rel_lower:
            return 'image/png'

        return default

    def _extract_video_payload(self, pic_elem, zip_ref, slide_rels_path, slide_num: int,
                               poster_path: Optional[str]) -> Optional[Dict]:
        """비디오 도형에서 영상 메타데이터 추출"""
        nv_pr = pic_elem.find('.//p:nvPicPr/p:nvPr', self.ns)
        if nv_pr is None:
            return None

        rel_attr = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
        embed_ids = set()
        link_ids = set()

        video_file = nv_pr.find('a:videoFile', self.ns)
        if video_file is not None:
            link_rel = video_file.get(f'{rel_attr}link')
            if link_rel:
                link_ids.add(link_rel)

        for media_elem in nv_pr.findall('.//p14:media', self.ns):
            embed_rel = media_elem.get(f'{rel_attr}embed')
            if embed_rel:
                embed_ids.add(embed_rel)
            link_rel = media_elem.get(f'{rel_attr}link')
            if link_rel:
                link_ids.add(link_rel)

        if not embed_ids and not link_ids:
            return None

        sources: List[Dict[str, Optional[str]]] = []
        iframe_src: Optional[str] = None
        fallback_link: Optional[str] = None
        seen_paths: set = set()

        for rel_id in embed_ids:
            media = self.extract_media(zip_ref, slide_rels_path, rel_id, slide_num, media_type='video')
            if not media:
                continue
            src = media.get('path')
            if not src or src in seen_paths:
                continue
            seen_paths.add(src)
            sources.append({
                'src': src,
                'mime': media.get('mime_type') or 'video/mp4',
                'external': media.get('external', False)
            })

        for rel_id in link_ids:
            media = self.extract_media(zip_ref, slide_rels_path, rel_id, slide_num, media_type='video')
            if not media:
                continue
            link_src = media.get('path')
            if not link_src:
                continue
            if media.get('external'):
                youtube_embed = self._build_youtube_embed_url(link_src)
                if youtube_embed:
                    iframe_src = youtube_embed
                    fallback_link = link_src
                    continue
            if link_src in seen_paths:
                continue
            seen_paths.add(link_src)
            sources.append({
                'src': link_src,
                'mime': media.get('mime_type') or 'video/mp4',
                'external': media.get('external', False)
            })
            if media.get('external'):
                fallback_link = link_src

        if not sources and not iframe_src:
            return None

        return {
            'sources': sources,
            'poster': poster_path,
            'iframe': iframe_src,
            'fallback': fallback_link
        }

    def _build_video_payload_from_link(self, url: str, poster: Optional[str]) -> Optional[Dict]:
        """하이퍼링크 URL 기반 비디오 페이로드 생성"""
        youtube_embed = self._build_youtube_embed_url(url)
        if youtube_embed:
            return {
                'sources': [],
                'poster': poster,
                'iframe': youtube_embed,
                'fallback': url
            }

        mime_type = self._guess_mime_type(url, None, 'video/mp4')
        if mime_type.startswith('video'):
            return {
                'sources': [{
                    'src': url,
                    'mime': mime_type,
                    'external': True
                }],
                'poster': poster,
                'iframe': None,
                'fallback': url
            }

        return None

    def _build_youtube_embed_url(self, url: str) -> Optional[str]:
        """YouTube URL을 iframe용 embed 주소로 변환"""
        try:
            parsed = urlparse(url)
        except Exception:
            return None

        host = parsed.netloc.lower()
        path = parsed.path or ''

        video_id: Optional[str] = None

        if 'youtube.com' in host:
            if path.startswith('/watch'):
                params = parse_qs(parsed.query)
                video_id = params.get('v', [None])[0]
            elif path.startswith('/embed/'):
                parts = [p for p in path.split('/') if p]
                if len(parts) >= 2:
                    video_id = parts[1]
            elif path.startswith('/shorts/'):
                parts = [p for p in path.split('/') if p]
                if len(parts) >= 2:
                    video_id = parts[1]
        elif 'youtu.be' in host:
            video_id = path.lstrip('/')

        if not video_id:
            return None

        return f"https://www.youtube.com/embed/{video_id}"

    # === 테이블 추출 (기존 코드 유지) ===

    def extract_table(self, graphic_frame, zip_ref, slide_rels_path) -> Optional[Dict]:
        """테이블 구조 및 스타일 추출"""
        tbl = graphic_frame.find('.//a:tbl', self.ns)
        if tbl is None:
            return None

        table_data = {
            'type': 'table',
            'rows': [],
            'col_widths': []
        }

        tbl_grid = tbl.find('.//a:tblGrid', self.ns)
        if tbl_grid is not None:
            for grid_col in tbl_grid.findall('.//a:gridCol', self.ns):
                width = int(grid_col.get('w', 0))
                table_data['col_widths'].append(self.emu_to_layout_px(width))

        for tr in tbl.findall('.//a:tr', self.ns):
            row_data = {
                'height': self.emu_to_layout_px(int(tr.get('h', 0))),
                'cells': []
            }

            for tc in tr.findall('.//a:tc', self.ns):
                cell_text, cell_anchor, cell_padding, cell_props = self.extract_text_with_formatting(tc, zip_ref, slide_rels_path)
                cell_data = {
                    'text': cell_text,
                    'text_anchor': cell_anchor,
                    'text_padding': cell_padding,
                    'text_props': cell_props,
                    'fill': self.extract_cell_fill(tc),
                    'borders': self.extract_cell_borders(tc),
                    'colspan': int(tc.get('gridSpan', 1)),
                    'rowspan': int(tc.get('rowSpan', 1)),
                    'vertical_align': tc.get('anchor', 't')
                }
                row_data['cells'].append(cell_data)

            table_data['rows'].append(row_data)

        self.logger.increment_table()
        return table_data

    def extract_cell_fill(self, tc) -> Dict:
        """테이블 셀 채우기 색상 추출"""
        tc_pr = tc.find('.//a:tcPr', self.ns)
        if tc_pr is None:
            return {'type': 'none', 'color': '#FFFFFF'}

        solid = tc_pr.find('.//a:solidFill', self.ns)
        if solid is not None:
            return {
                'type': 'solid',
                'color': self.color_to_hex(solid)
            }

        return {'type': 'none', 'color': '#FFFFFF'}

    def extract_cell_borders(self, tc) -> Dict:
        """테이블 셀 테두리 추출"""
        tc_pr = tc.find('.//a:tcPr', self.ns)
        borders = {}

        if tc_pr is not None:
            border_sides = {
                'left': 'a:lnL',
                'right': 'a:lnR',
                'top': 'a:lnT',
                'bottom': 'a:lnB'
            }

            for side_name, side_tag in border_sides.items():
                ln = tc_pr.find(f'.//{side_tag}', self.ns)
                if ln is not None:
                    width = int(ln.get('w', 12700))
                    borders[side_name] = {
                        'width': max(1, self.emu_to_layout_px(width)),
                        'color': self.color_to_hex(ln.find('.//a:solidFill', self.ns)),
                        'style': 'solid'
                    }
                else:
                    borders[side_name] = {
                        'width': 1,
                        'color': '#000000',
                        'style': 'solid'
                    }

        return borders

    # === 슬라이드 배경 처리 ===

    def _parse_background(self, bg_elem: Optional[ET.Element]) -> Optional[Dict]:
        """배경 요소를 파싱하여 색상 정보 반환"""
        if bg_elem is None:
            return None

        bg_pr = bg_elem.find('.//p:bgPr', self.ns)
        if bg_pr is not None:
            solid = bg_pr.find('.//a:solidFill', self.ns)
            if solid is not None:
                return {
                    'type': 'solid',
                    'color': self.color_to_hex(solid),
                    'gradient': None,
                    'gradient_angle': 0,
                    'gradient_type': 'linear'
                }

            grad = bg_pr.find('.//a:gradFill', self.ns)
            if grad is not None:
                stops = []
                for gs in grad.findall('.//a:gs', self.ns):
                    pos = float(gs.get('pos', 0)) / 100000
                    color = self.color_to_hex(gs)
                    stops.append((pos, color))

                angle = 0
                grad_type = 'linear'
                lin = grad.find('.//a:lin', self.ns)
                if lin is not None and lin.get('ang'):
                    try:
                        angle = int(lin.get('ang', 0)) / 60000
                    except ValueError:
                        angle = 0

                path = grad.find('.//a:path', self.ns)
                if path is not None and path.get('path'):
                    grad_type = path.get('path')

                return {
                    'type': 'gradient',
                    'color': self.theme_colors.get('bg1', '#FFFFFF'),
                    'gradient': stops,
                    'gradient_angle': angle,
                    'gradient_type': grad_type
                }

        bg_ref = bg_elem.find('.//p:bgRef', self.ns)
        if bg_ref is not None:
            color = self.color_to_hex(bg_ref)
            return {
                'type': 'solid',
                'color': color,
                'gradient': None,
                'gradient_angle': 0,
                'gradient_type': 'linear'
            }

        return None

    def _get_related_slide_layout(self, zip_ref, slide_rels_path: str) -> Optional[str]:
        """슬라이드와 연결된 레이아웃 경로 반환"""
        rels_tree = self.get_relationships(zip_ref, slide_rels_path)
        if rels_tree is None:
            return None

        for rel in rels_tree.findall('.//rel:Relationship', self.ns):
            rel_type = rel.get('Type', '')
            if rel_type.endswith('/slideLayout'):
                target = rel.get('Target')
                if target:
                    return f"ppt/{target.replace('..', '').lstrip('/')}"
        return None

    def _get_master_background(self, zip_ref, master_path: str) -> Optional[Dict]:
        """슬라이드 마스터 배경 조회"""
        if master_path in self.master_background_cache:
            return self.master_background_cache[master_path]

        try:
            master_xml = ET.fromstring(zip_ref.read(master_path))
        except KeyError:
            self.master_background_cache[master_path] = None
            return None

        bg = self._parse_background(master_xml.find('.//p:bg', self.ns))
        if bg is None:
            bg = {
                'type': 'solid',
                'color': self.theme_colors.get('bg1', '#FFFFFF'),
                'gradient': None,
                'gradient_angle': 0,
                'gradient_type': 'linear'
            }
        self.master_background_cache[master_path] = bg
        return bg

    def _get_layout_background(self, zip_ref, layout_path: str) -> Optional[Dict]:
        """슬라이드 레이아웃 배경 조회"""
        if layout_path in self.layout_background_cache:
            return self.layout_background_cache[layout_path]

        try:
            layout_xml = ET.fromstring(zip_ref.read(layout_path))
        except KeyError:
            self.layout_background_cache[layout_path] = None
            return None

        bg = self._parse_background(layout_xml.find('.//p:bg', self.ns))
        if bg:
            self.layout_background_cache[layout_path] = bg
            return bg

        layout_rels_path = layout_path.replace('slideLayouts/', 'slideLayouts/_rels/').replace('.xml', '.xml.rels')
        rels_tree = self.get_relationships(zip_ref, layout_rels_path)
        master_path = None
        if rels_tree is not None:
            for rel in rels_tree.findall('.//rel:Relationship', self.ns):
                rel_type = rel.get('Type', '')
                if rel_type.endswith('/slideMaster'):
                    target = rel.get('Target')
                    if target:
                        master_path = f"ppt/{target.replace('..', '').lstrip('/')}"
                    break

        bg = self._get_master_background(zip_ref, master_path) if master_path else None
        self.layout_background_cache[layout_path] = bg
        return bg

    def _resolve_background(self, zip_ref, slide_xml: ET.Element, slide_rels_path: str) -> Dict:
        """슬라이드 배경을 레이아웃/마스터까지 고려하여 결정"""
        slide_bg = self._parse_background(slide_xml.find('.//p:bg', self.ns))
        if slide_bg:
            return slide_bg

        layout_path = self._get_related_slide_layout(zip_ref, slide_rels_path)
        if layout_path:
            layout_bg = self._get_layout_background(zip_ref, layout_path)
            if layout_bg:
                return layout_bg

        master_bg = {
            'type': 'solid',
            'color': self.theme_colors.get('bg1', '#FFFFFF'),
            'gradient': None,
            'gradient_angle': 0,
            'gradient_type': 'linear'
        }
        return master_bg

    # === 도형 처리 (Phase 2 통합) ===

    def _extract_template_elements(self, zip_ref, template_path: str, rels_path: str, idx: int,
                                   chart_extractor, shape_converter, smartart_parser,
                                   animation_handler) -> List[Dict]:
        """레이아웃/마스터 템플릿 요소 추출"""
        elements: List[Dict] = []
        try:
            template_xml = ET.fromstring(zip_ref.read(template_path))
        except KeyError:
            return elements

        sp_tree = template_xml.find('.//p:cSld/p:spTree', self.ns)
        if sp_tree is None:
            return elements

        self._process_sp_tree(
            sp_tree,
            zip_ref,
            rels_path,
            idx,
            {'layout': {}, 'master': {}},
            chart_extractor,
            shape_converter,
            smartart_parser,
            animation_handler,
            elements,
            transform_chain=[],
            skip_placeholders=True
        )
        return elements

    def process_shape(self, shape, zip_ref, slide_rels_path, slide_num,
                      shape_converter, animation_handler,
                      placeholder_context: Dict[str, Dict],
                      transform_chain: Optional[List[Dict]] = None) -> Dict:
        """도형 처리 (커스텀 지오메트리, 그림자, 반사 효과 포함)"""
        transform_chain = transform_chain or []
        sp_pr = shape.find('.//p:spPr', self.ns)

        shape_id = None
        c_nv_pr = None
        nv_sp_pr = shape.find('.//p:nvSpPr', self.ns)
        if nv_sp_pr is not None:
            c_nv_pr = nv_sp_pr.find('.//p:cNvPr', self.ns)
            if c_nv_pr is not None:
                shape_id = c_nv_pr.get('id')

        paragraphs, text_anchor, text_padding, text_props = self.extract_text_with_formatting(shape, zip_ref, slide_rels_path)

        shape_data = {
            'type': 'shape',
            'position': self.extract_shape_position(sp_pr),
            'fill': self.extract_shape_fill(sp_pr),
            'border': self.extract_shape_border(sp_pr),
            'paragraphs': paragraphs,
            'text_anchor': text_anchor,
            'text_padding': text_padding,
            'image': None,
            'video': None,
            'audio': None,
            'hyperlink': None,
            'custom_geometry': None,
            'shadow': None,
            'reflection': None,
            'shape_id': shape_id,
            'z_index': self._next_z_index(),
            'text_props': text_props
        }

        shape_data['position'] = self._apply_placeholder_inheritance(shape, shape_data['position'], placeholder_context)
        if shape_data['position'].get('rotation') is None:
            shape_data['position']['rotation'] = 0.0
        shape_data['position'] = self._align_position_to_pivot(shape_data['position'])

        # Phase 2: 커스텀 지오메트리 추출
        if sp_pr is not None:
            custom_geom = shape_converter.extract_custom_geometry(sp_pr)
            if custom_geom:
                shape_data['custom_geometry'] = custom_geom

            # Phase 2: 그림자 효과
            shadow_css = animation_handler.apply_shadow_effects(sp_pr)
            if shadow_css:
                shape_data['shadow'] = shadow_css

            # Phase 2: 반사 효과
            reflection_css = animation_handler.apply_reflection_effects(sp_pr)
            if reflection_css:
                shape_data['reflection'] = reflection_css

        # 도형 레벨 하이퍼링크
        if nv_sp_pr is not None and c_nv_pr is not None:
            hlink = c_nv_pr.find('.//a:hlinkClick', self.ns)
            if hlink is not None:
                rel_id = hlink.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rel_id:
                    target, _, _ = self.resolve_relationship(zip_ref, slide_rels_path, rel_id)
                    if target:
                            shape_data['hyperlink'] = target

        # 이미지/비디오/오디오 확인
        pic = shape.find('.//pic:pic', self.ns)
        if pic is not None:
            poster_path: Optional[str] = None
            blip = pic.find('.//a:blip', self.ns)
            if blip is not None:
                embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed:
                    media = self.extract_media(zip_ref, slide_rels_path, embed, slide_num)
                    if media:
                        if media['type'] == 'video':
                            shape_data['video'] = {
                                'sources': [{
                                    'src': media['path'],
                                    'mime': media.get('mime_type') or 'video/mp4',
                                    'external': media.get('external', False)
                                }],
                                'poster': None,
                                'iframe': None,
                                'fallback': media['path'] if media.get('external') else None
                            }
                        elif media['type'] == 'audio':
                            shape_data['audio'] = media['path']
                        else:
                            shape_data['image'] = media['path']
                            poster_path = media['path']

            if poster_path is None:
                poster_path = shape_data.get('image')

            video_payload = self._extract_video_payload(pic, zip_ref, slide_rels_path, slide_num, poster_path)
            if video_payload:
                shape_data['video'] = video_payload
            elif shape_data.get('hyperlink'):
                link_video = self._build_video_payload_from_link(shape_data['hyperlink'], poster_path)
                if link_video:
                    shape_data['video'] = link_video
                    shape_data['hyperlink'] = None

        if transform_chain:
            shape_data['position'] = self._apply_transform_chain(shape_data['position'], transform_chain)
        shape_data['position'] = self._ensure_position_defaults(shape_data['position'])

        # 이미지 채우기 처리
        if shape_data['fill'].get('type') == 'image':
            rel_id = shape_data['fill'].get('rel_id')
            if rel_id:
                media = self.extract_media(zip_ref, slide_rels_path, rel_id, slide_num)
                if media and media.get('type') == 'image':
                    shape_data['image'] = media['path']
                    shape_data['image_crop'] = shape_data['fill'].get('crop')
                    shape_data['image_stretch'] = shape_data['fill'].get('stretch')
            shape_data['fill'] = {'type': 'none'}

        if shape_data.get('hyperlink') and not shape_data.get('video'):
            poster = shape_data.get('image')
            link_video = self._build_video_payload_from_link(shape_data['hyperlink'], poster)
            if link_video:
                shape_data['video'] = link_video
                shape_data['hyperlink'] = None

        self.logger.increment_shape()
        return shape_data

    # === 메인 변환 프로세스 ===

    def convert(self):
        """메인 변환 프로세스 (Phase 2 통합)"""
        self.logger.info(f"Starting conversion: {self.pptx_path.name}")

        try:
            self.theme_colors = self._default_theme_colors()
            self.theme_background_fills = []
            self.layout_background_cache.clear()
            self.master_background_cache.clear()
            self.layout_placeholder_cache.clear()
            self.master_placeholder_cache.clear()
            font_faces: List[str] = []
            with zipfile.ZipFile(self.pptx_path, 'r') as zip_ref:
                # 모듈 초기화
                chart_extractor = ChartExtractor(zip_ref, self.logger)
                shape_converter = ShapeGeometryConverter(self.ns, self.logger)
                smartart_parser = SmartArtParser(zip_ref, self.ns, self.logger)
                animation_handler = AnimationHandler(self.ns, self.logger)
                transition_handler = self.transition_handler
                font_manager = FontManager(zip_ref, self.output_dir, self.logger)
                self._load_theme(zip_ref)

                # 슬라이드 크기 가져오기
                try:
                    pres_xml = ET.fromstring(zip_ref.read('ppt/presentation.xml'))
                    sld_sz = pres_xml.find('.//p:sldSz', self.ns)
                    if sld_sz is not None:
                        self.slide_size['width'] = int(sld_sz.get('cx', self.slide_size['width']))
                        self.slide_size['height'] = int(sld_sz.get('cy', self.slide_size['height']))

                    slide_ids = pres_xml.findall('.//p:sldId', self.ns)
                    self.logger.info(f"Found {len(slide_ids)} slide(s)")

                    # 각 슬라이드 처리
                    for idx, slide_id in enumerate(slide_ids, 1):
                        self.logger.info(f"Processing slide {idx}...")
                        try:
                            self.process_slide(zip_ref, slide_id, idx, chart_extractor,
                                               shape_converter, smartart_parser, animation_handler,
                                               transition_handler)
                            self.logger.increment_slide()
                        except Exception as e:
                            self.logger.error(f"Failed to process slide {idx}", exception=e, slide_num=idx)
                            # 계속 진행

                except Exception as e:
                    self.logger.critical("Failed to read presentation structure", exception=e)
                    return None

                # 폰트 추출
                try:
                    font_faces = font_manager.extract_embedded_fonts()
                except Exception as font_exc:
                    self.logger.warning(f"Embedded font extraction failed: {font_exc}")

            # HTML 생성
            html_output, css_output, js_output = self.generate_bundle(animation_handler, font_faces)
            html_path = self.output_dir / f"{self.pptx_path.stem}.html"
            css_path = self.output_dir / f"{self.pptx_path.stem}.css"
            js_path = self.output_dir / f"{self.pptx_path.stem}.js"

            html_path.write_text(html_output, encoding='utf-8')
            css_path.write_text(css_output, encoding='utf-8')
            js_path.write_text(js_output, encoding='utf-8')

            self.logger.info(f"Conversion complete: {html_path}")
            self.logger.print_summary()

            # 상세 리포트 저장
            report_path = self.output_dir / f"{self.pptx_path.stem}_report.md"
            self.logger.save_detailed_report(report_path)

            return html_path

        except Exception as e:
            self.logger.critical("Conversion failed", exception=e)
            return None

    def _extract_group_transform(self, grp_sp) -> Dict:
        """그룹 도형 변환 정보 추출"""
        xfrm = grp_sp.find('.//p:grpSpPr/a:xfrm', self.ns)
        if xfrm is None:
            return {
                'offset_x': 0.0,
                'offset_y': 0.0,
                'origin_x': 0.0,
                'origin_y': 0.0,
                'scale_x': 1.0,
                'scale_y': 1.0,
                'rotation': 0.0
            }

        def get_point(tag):
            elem = xfrm.find(f'a:{tag}', self.ns)
            if elem is None:
                return {'x': 0.0, 'y': 0.0}
            return {
                'x': self.emu_to_layout_px(int(elem.get('x', 0))),
                'y': self.emu_to_layout_px(int(elem.get('y', 0)))
            }

        def get_size(tag):
            elem = xfrm.find(f'a:{tag}', self.ns)
            if elem is None:
                return {'width': 0.0, 'height': 0.0}
            return {
                'width': self.emu_to_layout_px(int(elem.get('cx', 0))),
                'height': self.emu_to_layout_px(int(elem.get('cy', 0)))
            }

        off = get_point('off')
        ext = get_size('ext')
        ch_off = get_point('chOff')
        ch_ext = get_size('chExt')

        scale_x = ext['width'] / ch_ext['width'] if ch_ext['width'] else 1.0
        scale_y = ext['height'] / ch_ext['height'] if ch_ext['height'] else 1.0

        rotation = 0.0
        if xfrm.get('rot'):
            try:
                rotation = int(xfrm.get('rot')) / 60000
            except ValueError:
                rotation = 0.0

        pivot_x = off['x'] + (ext['width'] / 2.0)
        pivot_y = off['y'] + (ext['height'] / 2.0)

        return {
            'offset_x': off['x'],
            'offset_y': off['y'],
            'origin_x': ch_off['x'],
            'origin_y': ch_off['y'],
            'scale_x': scale_x if scale_x != 0 else 1.0,
            'scale_y': scale_y if scale_y != 0 else 1.0,
            'rotation': rotation,
            'pivot_x': pivot_x,
            'pivot_y': pivot_y,
            'ext_width': ext['width'],
            'ext_height': ext['height']
        }

    def _process_picture(self, pic, zip_ref, slide_rels_path, idx, placeholder_context: Dict[str, Dict],
                         transform_chain: Optional[List[Dict]], elements: List[Dict]):
        """그림 도형 처리"""
        transform_chain = transform_chain or []
        sp_pr = pic.find('.//p:spPr', self.ns)

        element = {
            'type': 'image',
            'position': self.extract_shape_position(sp_pr),
            'fill': {'type': 'none'},
            'border': {'width': 0},
            'paragraphs': [],
            'image': None,
            'video': None,
            'audio': None,
            'hyperlink': None,
            'z_index': self._next_z_index()
        }

        element['position'] = self._apply_placeholder_inheritance(pic, element['position'], placeholder_context)
        if element['position'].get('rotation') is None:
            element['position']['rotation'] = 0.0
        element['position'] = self._ensure_position_defaults(element['position'])

        hyperlink_target = None
        c_nv_pr = pic.find('.//p:nvPicPr/p:cNvPr', self.ns)
        if c_nv_pr is not None:
            hlink = c_nv_pr.find('.//a:hlinkClick', self.ns)
            if hlink is not None:
                rel_id = hlink.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rel_id:
                    target, _, _ = self.resolve_relationship(zip_ref, slide_rels_path, rel_id)
                    if target:
                        hyperlink_target = target

        blip = pic.find('.//a:blip', self.ns)
        if blip is not None:
            embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if embed:
                media = self.extract_media(zip_ref, slide_rels_path, embed, idx)
                if media:
                    if media['type'] == 'video':
                        element['type'] = 'video'
                        element['video'] = {
                            'sources': [{
                                'src': media['path'],
                                'mime': media.get('mime_type') or 'video/mp4',
                                'external': media.get('external', False)
                            }],
                            'poster': None,
                            'iframe': None,
                            'fallback': media['path'] if media.get('external') else None
                        }
                    elif media['type'] == 'audio':
                        element['type'] = 'audio'
                        element['audio'] = media['path']
                    else:
                        element['image'] = media['path']

        poster_path = element.get('image')
        video_payload = self._extract_video_payload(pic, zip_ref, slide_rels_path, idx, poster_path)
        if video_payload:
            element['type'] = 'video'
            element['video'] = video_payload
            element['hyperlink'] = None
        elif hyperlink_target:
            link_video = self._build_video_payload_from_link(hyperlink_target, poster_path)
            if link_video:
                element['type'] = 'video'
                element['video'] = link_video
                element['hyperlink'] = None
            else:
                element['hyperlink'] = hyperlink_target

        if transform_chain:
            element['position'] = self._apply_transform_chain(element['position'], transform_chain)
        else:
            element['position'] = self._align_position_to_pivot(element['position'])

        elements.append(element)

    def _process_graphic_frame(self, graphic_frame, zip_ref, slide_rels_path, idx,
                               placeholder_context: Dict[str, Dict],
                               chart_extractor, smartart_parser,
                               transform_chain: Optional[List[Dict]], elements: List[Dict]):
        """그래픽 프레임 처리 (차트, 테이블, SmartArt)"""
        transform_chain = transform_chain or []

        # 차트 처리
        chart_config = chart_extractor.extract_chart_from_graphic_frame(graphic_frame, slide_rels_path)
        if chart_config:
            self.chart_counter += 1
            position = self.extract_shape_position(graphic_frame.find('.//p:xfrm', self.ns))
            position = self._apply_placeholder_inheritance(graphic_frame, position, placeholder_context)
            if position.get('rotation') is None:
                position['rotation'] = 0.0
            position = self._ensure_position_defaults(position)
            if transform_chain:
                position = self._apply_transform_chain(position, transform_chain)
            else:
                position = self._align_position_to_pivot(position)

            elements.append({
                'type': 'chart',
                'position': position,
                'chart_config': chart_config,
                'chart_id': f'chart_{idx}_{self.chart_counter}',
                'z_index': self._next_z_index()
            })
            return

        # 테이블 처리
        table_data = self.extract_table(graphic_frame, zip_ref, slide_rels_path)
        if table_data:
            xfrm = graphic_frame.find('.//p:xfrm', self.ns)
            position = self.extract_shape_position(xfrm)
            position = self._apply_placeholder_inheritance(graphic_frame, position, placeholder_context)
            if position.get('rotation') is None:
                position['rotation'] = 0.0
            position = self._ensure_position_defaults(position)
            if transform_chain:
                position = self._apply_transform_chain(position, transform_chain)
            else:
                position = self._align_position_to_pivot(position)
            table_data['position'] = position
            table_data['z_index'] = self._next_z_index()
            elements.append(table_data)
            return

        # SmartArt 처리
        smartart_data = smartart_parser.extract_smartart_text(graphic_frame, slide_rels_path)
        if smartart_data:
            xfrm = graphic_frame.find('.//p:xfrm', self.ns)
            position = self.extract_shape_position(xfrm)
            position = self._apply_placeholder_inheritance(graphic_frame, position, placeholder_context)
            if position.get('rotation') is None:
                position['rotation'] = 0.0
            position = self._ensure_position_defaults(position)
            if transform_chain:
                position = self._apply_transform_chain(position, transform_chain)
            else:
                position = self._align_position_to_pivot(position)
            smartart_data['position'] = position
            smartart_data['z_index'] = self._next_z_index()
            elements.append(smartart_data)

    def _process_group(self, grp_sp, zip_ref, slide_rels_path, idx, placeholder_context: Dict[str, Dict],
                       chart_extractor, shape_converter, smartart_parser,
                       animation_handler, elements: List[Dict], transform_chain: Optional[List[Dict]],
                       skip_placeholders: bool = False):
        """그룹 도형 처리 (Nested approach)"""
        # Note: We ignore the incoming transform_chain because we are building a nested structure.
        # The parent group (if any) handles the transform for this group.
        
        if skip_placeholders:
            ph = grp_sp.find('.//p:nvGrpSpPr/p:nvPr/p:ph', self.ns)
            if ph is not None:
                return

        # Extract group transform info
        group_transform = self._extract_group_transform(grp_sp)
        
        # Create container for children
        group_children: List[Dict] = []
        
        # Process children (pass empty transform_chain to avoid flattening)
        for child in grp_sp:
            tag = self._strip_namespace(child.tag)
            if tag in {'nvGrpSpPr', 'grpSpPr'}:
                continue
            if skip_placeholders:
                ph_elem = child.find('.//p:nvPr/p:ph', self.ns)
                if ph_elem is not None:
                    continue

            if tag in {'sp', 'cxnSp'}:
                # Process shape and append explicitly
                child_element = self.process_shape(child, zip_ref, slide_rels_path, idx,
                                                   shape_converter, animation_handler,
                                                   placeholder_context, transform_chain=[])
                group_children.append(child_element)
                
            elif tag == 'pic':
                self._process_picture(child, zip_ref, slide_rels_path, idx,
                                      placeholder_context, [], group_children)
                                      
            elif tag == 'graphicFrame':
                self._process_graphic_frame(child, zip_ref, slide_rels_path, idx,
                                            placeholder_context, chart_extractor,
                                            smartart_parser, [], group_children)
                                            
            elif tag == 'grpSp':
                self._process_group(child, zip_ref, slide_rels_path, idx,
                                    placeholder_context, chart_extractor, shape_converter, smartart_parser,
                                    animation_handler, group_children, [], skip_placeholders=skip_placeholders)

        # Normalize children positions relative to this group
        # Target coordinate system: 0..ext_width, 0..ext_height
        # Source coordinate system: chOff.x..chOff.x+chExt.cx, ...
        
        ox = group_transform['origin_x']
        oy = group_transform['origin_y']
        sx = group_transform['scale_x']
        sy = group_transform['scale_y']
        
        for child in group_children:
            pos = child['position']
            # Apply scaling and translation to map child coords to group visual coords
            pos['x'] = (pos['x'] - ox) * sx
            pos['y'] = (pos['y'] - oy) * sy
            pos['width'] *= sx
            pos['height'] *= sy
            
            # Adjust pivot if present
            if 'pivot_x' in pos:
                pos['pivot_x'] = (pos['pivot_x'] - ox) * sx
            if 'pivot_y' in pos:
                pos['pivot_y'] = (pos['pivot_y'] - oy) * sy

        # Create Group Element
        group_element = {
            'type': 'group',
            'position': {
                'x': group_transform['offset_x'],
                'y': group_transform['offset_y'],
                'width': group_transform['ext_width'],
                'height': group_transform['ext_height'],
                'rotation': group_transform['rotation']
            },
            'children': group_children,
            'z_index': self._next_z_index()
        }
        
        elements.append(group_element)

    def _process_sp_tree(self, container, zip_ref, slide_rels_path, idx, placeholder_context: Dict[str, Dict],
                         chart_extractor, shape_converter, smartart_parser,
                         animation_handler, elements: List[Dict], transform_chain: Optional[List[Dict]] = None,
                         skip_placeholders: bool = False):
        """슬라이드 spTree 재귀 처리"""
        transform_chain = transform_chain or []

        for child in container:
            tag = self._strip_namespace(child.tag)
            if tag in {'nvGrpSpPr', 'grpSpPr'}:
                continue
            if skip_placeholders:
                ph_elem = child.find('.//p:nvPr/p:ph', self.ns)
                if ph_elem is not None:
                    continue

            if tag in {'sp', 'cxnSp'}:
                elements.append(
                    self.process_shape(child, zip_ref, slide_rels_path, idx,
                                       shape_converter, animation_handler,
                                       placeholder_context, transform_chain=transform_chain)
                )
            elif tag == 'pic':
                self._process_picture(child, zip_ref, slide_rels_path, idx,
                                      placeholder_context, transform_chain, elements)
            elif tag == 'graphicFrame':
                self._process_graphic_frame(child, zip_ref, slide_rels_path, idx,
                                            placeholder_context, chart_extractor,
                                            smartart_parser, transform_chain, elements)
            elif tag == 'grpSp':
                self._process_group(child, zip_ref, slide_rels_path, idx,
                                    placeholder_context, chart_extractor, shape_converter, smartart_parser,
                                    animation_handler, elements, transform_chain, skip_placeholders=skip_placeholders)

    def process_slide(self, zip_ref, slide_id, idx, chart_extractor,
                      shape_converter, smartart_parser, animation_handler,
                      transition_handler):
        """단일 슬라이드 처리"""
        slide_rel_id = slide_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

        # 슬라이드 경로 찾기
        pres_rels = ET.fromstring(zip_ref.read('ppt/_rels/presentation.xml.rels'))
        slide_path = None

        for rel in pres_rels.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            if rel.get('Id') == slide_rel_id:
                slide_path = f"ppt/{rel.get('Target')}"
                break

        if not slide_path:
            raise Exception(f"Slide path not found for ID: {slide_rel_id}")

        # 슬라이드 파싱
        slide_rels_path = slide_path.replace('slides/', 'slides/_rels/').replace('.xml', '.xml.rels')

        slide_xml = ET.fromstring(zip_ref.read(slide_path))

        layout_path = self._get_related_slide_layout(zip_ref, slide_rels_path)
        master_path = self._get_layout_master_path(zip_ref, layout_path) if layout_path else None

        # 배경 추출
        background = self._resolve_background(zip_ref, slide_xml, slide_rels_path)

        # Phase 2: 애니메이션 추출
        animations = animation_handler.extract_slide_animations(slide_xml)
        slide_transition = transition_handler.extract_transition(slide_xml)

        placeholder_context = self._build_placeholder_context(zip_ref, slide_rels_path, layout_path=layout_path)

        # 모든 요소 처리
        elements = []

        # z-index 초기화
        self.z_index_counter = 1

        # 템플릿 요소(마스터 → 레이아웃 순) 추가
        if master_path:
            master_rels_path = master_path.replace('slideMasters/', 'slideMasters/_rels/').replace('.xml', '.xml.rels')
            elements.extend(
                self._extract_template_elements(
                    zip_ref,
                    master_path,
                    master_rels_path,
                    idx,
                    chart_extractor,
                    shape_converter,
                    smartart_parser,
                    animation_handler
                )
            )

        if layout_path:
            layout_rels_path = layout_path.replace('slideLayouts/', 'slideLayouts/_rels/').replace('.xml', '.xml.rels')
            elements.extend(
                self._extract_template_elements(
                    zip_ref,
                    layout_path,
                    layout_rels_path,
                    idx,
                    chart_extractor,
                    shape_converter,
                    smartart_parser,
                    animation_handler
                )
            )

        sp_tree = slide_xml.find('.//p:cSld/p:spTree', self.ns)
        if sp_tree is not None:
            self._process_sp_tree(sp_tree, zip_ref, slide_rels_path, idx, placeholder_context,
                                  chart_extractor, shape_converter, smartart_parser,
                                  animation_handler, elements, transform_chain=[])

        self.slides_data.append({
            'number': idx,
            'background': background,
            'elements': elements,
            'animations': animations,
            'transition': slide_transition
        })

        self.logger.debug(f"Slide {idx}: {len(elements)} element(s) extracted")

    # === HTML 생성 (Phase 2 통합) ===

    def _generate_html_legacy(self, animation_handler):
        """[Deprecated] Phase 2 HTML 생성 로직"""
        slide_width = self.emu_to_layout_px(self.slide_size['width'])
        slide_height = self.emu_to_layout_px(self.slide_size['height'])

        # Phase 2: 각 슬라이드 HTML 생성
        slides_html = []

        for slide in self.slides_data:
            # 배경 스타일
            bg = slide['background']
            bg_style = f"background-color: {bg['color']}"
            if bg['type'] == 'gradient' and bg['gradient']:
                gradient_stops = ', '.join([f"{color} {pos*100}%" for pos, color in bg['gradient']])
                bg_style = f"background: linear-gradient(135deg, {gradient_stops})"

            # 요소 HTML 생성
            elements_html = []
            for element in slide['elements']:
                if element.get('type') == 'chart':
                    # Phase 2: Chart.js 차트
                    chart_extractor = ChartExtractor(None, self.logger)
                    chart_html = chart_extractor.generate_chartjs_html(
                        element['chart_config'],
                        element['chart_id'],
                        element['position'],
                        slide_width,
                        slide_height
                    )
                    elements_html.append(chart_html)

                elif element.get('type') == 'smartart':
                    # Phase 2: SmartArt
                    smartart_parser = SmartArtParser(None, self.ns, self.logger)
                    smartart_html = smartart_parser.generate_smartart_html(
                        element,
                        element['position'],
                        slide_width,
                        slide_height
                    )
                    elements_html.append(smartart_html)

                elif element.get('custom_geometry'):
                    # Phase 2: 커스텀 도형 (SVG)
                    shape_converter = ShapeGeometryConverter(self.ns, self.logger)
                    svg_html = shape_converter.generate_svg_html(
                        element['custom_geometry'],
                        element['position'],
                        element.get('fill', {}),
                        element.get('border', {}),
                        element.get('z_index', 1),
                        element.get('shadow'),
                        element.get('reflection')
                    )
                    elements_html.append(svg_html)

                else:
                    # 기존 도형/테이블 렌더링
                    shape_html = self.generate_element_html(element, slide_width, slide_height)
                    elements_html.append(shape_html)

            slides_html.append(f'''
            <div class="slide" data-slide="{slide['number']}" style="{bg_style}">
                <div class="slide-container">
                    {''.join(elements_html)}
                </div>
                <div class="slide-number">Slide {slide['number']}</div>
            </div>
            ''')

        # Phase 2: CSS 애니메이션
        animation_css = animation_handler.generate_css_animations()

        # 완전한 HTML
        html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{self.pptx_path.stem}</title>
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: Arial, sans-serif;
            background: #1a1a1a;
            overflow: hidden;
        }}

        #presentation {{
            width: 100vw;
            height: 100vh;
            position: relative;
        }}

        .slide {{
            width: 100%;
            height: 100%;
            display: none;
            position: absolute;
            top: 0;
            left: 0;
        }}

        .slide.active {{
            display: block;
            animation: slideIn 0.5s ease-out;
        }}

        @keyframes slideIn {{
            from {{
                opacity: 0;
                transform: translateX(50px);
            }}
            to {{
                opacity: 1;
                transform: translateX(0);
            }}
        }}

        .slide-container {{
            width: {slide_width}px;
            height: {slide_height}px;
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            max-width: 90vw;
            max-height: 90vh;
        }}

        @media (max-aspect-ratio: {int(slide_width)}/{int(slide_height)}) {{
            .slide-container {{
                width: 90vw;
                height: auto;
            }}
        }}

        @media (min-aspect-ratio: {int(slide_width)}/{int(slide_height)}) {{
            .slide-container {{
                width: auto;
                height: 90vh;
            }}
        }}

        .slide-number {{
            position: absolute;
            bottom: 20px;
            right: 30px;
            font-size: 1rem;
            color: rgba(255,255,255,0.7);
            z-index: 1000;
        }}

        .controls {{
            position: fixed;
            bottom: 40px;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            gap: 20px;
            z-index: 2000;
        }}

        .controls button {{
            background: rgba(255, 255, 255, 0.2);
            border: 2px solid rgba(255, 255, 255, 0.3);
            color: white;
            padding: 12px 24px;
            font-size: 1rem;
            cursor: pointer;
            border-radius: 8px;
            transition: all 0.3s ease;
            backdrop-filter: blur(10px);
        }}

        .controls button:hover {{
            background: rgba(255, 255, 255, 0.3);
            transform: scale(1.05);
        }}

        .controls button:disabled {{
            opacity: 0.3;
            cursor: not-allowed;
        }}

        .progress-bar {{
            position: fixed;
            top: 0;
            left: 0;
            height: 4px;
            background: #4CAF50;
            width: 0%;
            transition: width 0.3s ease;
            z-index: 2000;
        }}

        /* Phase 2: Chart container */
        .chart-container {{
            position: relative;
        }}

        {animation_css}
    </style>
</head>
<body>
    <div class="progress-bar" id="progress"></div>

    <div id="presentation">
        {''.join(slides_html)}
    </div>

    <div class="controls">
        <button id="prev" onclick="prevSlide()">← Previous</button>
        <button id="next" onclick="nextSlide()">Next →</button>
    </div>

    <script>
        let currentSlide = 0;
        const slides = document.querySelectorAll('.slide');
        const totalSlides = slides.length;

        function showSlide(n) {{
            slides.forEach(slide => slide.classList.remove('active'));

            if (n >= totalSlides) {{
                currentSlide = totalSlides - 1;
            }} else if (n < 0) {{
                currentSlide = 0;
            }} else {{
                currentSlide = n;
            }}

            slides[currentSlide].classList.add('active');

            const progress = ((currentSlide + 1) / totalSlides) * 100;
            document.getElementById('progress').style.width = progress + '%';

            document.getElementById('prev').disabled = currentSlide === 0;
            document.getElementById('next').disabled = currentSlide === totalSlides - 1;

            // 슬라이드 변경 이벤트 발생
            const event = new Event('slideChanged');
            document.dispatchEvent(event);
        }}

        function nextSlide() {{
            showSlide(currentSlide + 1);
        }}

        function prevSlide() {{
            showSlide(currentSlide - 1);
        }}

        document.addEventListener('keydown', (e) => {{
            if (e.key === 'ArrowRight' || e.key === ' ') {{
                nextSlide();
            }} else if (e.key === 'ArrowLeft') {{
                prevSlide();
            }}
        }});

        showSlide(0);
    </script>
</body>
</html>'''

        return html

    def generate_bundle(self, animation_handler, font_faces: List[str]) -> Tuple[str, str, str]:
        """HTML/CSS/JS 번들 생성"""
        slide_width = self.emu_to_layout_px(self.slide_size['width'])
        slide_height = self.emu_to_layout_px(self.slide_size['height'])

        slides_markup: List[str] = []
        total_slides = len(self.slides_data)

        smartart_parser = SmartArtParser(None, self.ns, self.logger)

        for index, slide in enumerate(self.slides_data, start=1):
            bg_style = self._build_background_style(slide['background'])
            elements_html: List[str] = []

            for element in slide['elements']:
                if element.get('type') == 'chart':
                    elements_html.append(self._build_chart_element(element))
                elif element.get('type') == 'smartart':
                    elements_html.append(
                        smartart_parser.generate_smartart_html(
                            element,
                            element['position'],
                            slide_width,
                            slide_height,
                            absolute=True,
                            z_index=element.get('z_index', 1)
                        )
                    )
                else:
                    elements_html.append(self.generate_element_html(element, slide_width, slide_height))

            transition_meta = slide.get('transition')
            data_attrs = ''
            if transition_meta:
                effect_attr = html.escape(transition_meta.effect or '', quote=True)
                direction_attr = html.escape(transition_meta.direction or '', quote=True)
                data_attrs = (
                    f' data-transition-effect="{effect_attr}"'
                    f' data-transition-direction="{direction_attr}"'
                )

            slide_markup = (
                f'<div class="slide{" active" if index == 1 else ""}" '
                f'data-slide="{index}"{data_attrs} style="{bg_style}">'
                f'{"".join(elements_html)}'
                f'<div class="slide-number">{index} / {total_slides}</div>'
                '</div>'
            )
            slides_markup.append(slide_markup)

        animation_css = animation_handler.generate_css_animations()
        transition_css = self.transition_handler.generate_transition_css()
        transition_data = self.transition_handler.serialize_transitions(self.slides_data)
        transition_json = json.dumps(transition_data, ensure_ascii=False)

        css_template = """:root {
    --slide-width: __SLIDE_WIDTH__px;
    --slide-height: __SLIDE_HEIGHT__px;
}

* {
    box-sizing: border-box;
    margin: 0;
    padding: 0;
}

body {
    font-family: 'Arial', sans-serif;
    background-color: #1a1a1a;
    color: #111;
    width: 100vw;
    height: 100vh;
    overflow: hidden;
}

#presentation {
    position: relative;
    width: 100vw;
    height: 100vh;
    display: flex;
    flex-direction: column;
}

.stage-wrapper {
    flex: 1 1 auto;
    position: relative;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 24px;
    overflow: hidden;
    min-width: 0;
    min-height: 0;
}

.slide-stage {
    width: var(--slide-width);
    height: var(--slide-height);
    position: relative;
    transform-origin: center center;
    overflow: hidden;
    background-color: #000;
}

.slide {
    position: absolute;
    inset: 0;
    width: var(--slide-width);
    height: var(--slide-height);
    overflow: hidden;
    background-position: center;
    background-repeat: no-repeat;
    background-size: cover;
}

.slide-number {
    position: absolute;
    right: 24px;
    bottom: 16px;
    background: rgba(17, 17, 17, 0.35);
    color: #ffffff;
    padding: 6px 12px;
    border-radius: 12px;
    font-size: 13px;
    backdrop-filter: blur(6px);
    z-index: 25;
}

.ppt-element {
    position: absolute;
    overflow: visible;
}

.ppt-link {
    text-decoration: none;
    color: inherit;
    display: block;
}

.ppt-table {
    background-color: transparent;
}

.ppt-table td, .ppt-table th {
    white-space: pre-wrap;
}

.chart-element {
    pointer-events: auto;
}

.chart-element canvas {
    width: 100%;
    height: 100%;
}

.ppt-paragraph {
    color: inherit;
}

.ppt-bullet {
    display: inline-block;
    line-height: 1;
}

.controls {
    display: flex;
    gap: 16px;
    justify-content: center;
    padding: 12px 0 24px;
    z-index: 30;
}

.controls button {
    background: rgba(255, 255, 255, 0.18);
    border: 1px solid rgba(255, 255, 255, 0.3);
    color: #ffffff;
    padding: 10px 22px;
    font-size: 15px;
    cursor: pointer;
    border-radius: 10px;
    transition: all 0.25s ease;
    backdrop-filter: blur(8px);
}

.controls button:hover {
    background: rgba(255, 255, 255, 0.28);
}

.controls button:disabled {
    opacity: 0.35;
    cursor: not-allowed;
}

.progress-bar {
    position: absolute;
    top: 0;
    left: 0;
    height: 4px;
    background: #4CAF50;
    width: 0;
    transition: width 0.3s ease;
    z-index: 40;
}
""".strip()

        base_css = (css_template
                    .replace('__SLIDE_WIDTH__', f"{slide_width:.2f}")
                    .replace('__SLIDE_HEIGHT__', f"{slide_height:.2f}"))

        css_sections: List[str] = []
        if font_faces:
            css_sections.append('\n'.join(font_faces))
        if transition_css:
            css_sections.append(transition_css)
        if base_css:
            css_sections.append(base_css)
        if animation_css:
            css_sections.append(animation_css)
        css_content = '\n\n'.join(css_sections)

        js_template = """(function() {
    const slideWidth = __SLIDE_WIDTH__;
    const slideHeight = __SLIDE_HEIGHT__;
    const slides = Array.from(document.querySelectorAll('.slide'));
    const stage = document.querySelector('.slide-stage');
    const progress = document.getElementById('progress');
    const totalSlides = slides.length;
    const prevBtn = document.getElementById('prev');
    const nextBtn = document.getElementById('next');
    const blackout = document.getElementById('transitionBlackout');
    const slideTransitions = __TRANSITION_DATA__;
    const defaultTransition = {
        effect: 'cut',
        duration: 0,
        delay: 0,
        direction: null,
        orientation: null,
        throughBlack: false,
        spokes: null,
        shape: null,
        speed: null,
        advanceOnClick: true,
        advanceAfter: null
    };
    const cssVarKeys = [
        '--ppt-transition-duration',
        '--ppt-transition-delay',
        '--ppt-translate-in-x',
        '--ppt-translate-in-y',
        '--ppt-translate-out-x',
        '--ppt-translate-out-y',
        '--ppt-scale-in',
        '--ppt-scale-out',
        '--ppt-rotate-in',
        '--ppt-rotate-out',
        '--ppt-mask-start',
        '--ppt-mask-angle',
        '--ppt-mask-shape',
        '--ppt-clip-inset-top',
        '--ppt-clip-inset-right',
        '--ppt-clip-inset-bottom',
        '--ppt-clip-inset-left',
        '--ppt-wheel-spokes',
        '--ppt-clip-shape'
    ];
    let currentSlide = 0;
    let isTransitioning = false;
    let autoAdvanceTimer = null;

    function applyScale() {
        if (!stage || !stage.parentElement) {
            return;
        }
        const wrapper = stage.parentElement;
        const scaleX = wrapper.clientWidth / slideWidth;
        const scaleY = wrapper.clientHeight / slideHeight;
        const scale = Math.min(scaleX, scaleY);
        stage.style.transform = `scale(${scale})`;
    }

    function mergeTransition(index) {
        const raw = slideTransitions[index] || {};
        const config = Object.assign({}, defaultTransition);
        Object.keys(raw).forEach(key => {
            if (raw[key] !== undefined && raw[key] !== null) {
                config[key] = raw[key];
            }
        });
        if (!config.spokes && config.effect === 'wheel') {
            config.spokes = 6;
        }
        return config;
    }

    function invertDirection(direction) {
        switch (direction) {
            case 'l': return 'r';
            case 'r': return 'l';
            case 'u': return 'd';
            case 'd': return 'u';
            case 'tl': return 'br';
            case 'tr': return 'bl';
            case 'bl': return 'tr';
            case 'br': return 'tl';
            default: return direction;
        }
    }

    function directionVector(direction) {
        switch (direction) {
            case 'l':
                return { inX: '-100%', inY: '0%', outX: '100%', outY: '0%', maskAngle: '90deg', maskStart: '-125%', clip: { top: '0%', right: '0%', bottom: '0%', left: '100%' }, rotate: '90deg' };
            case 'r':
                return { inX: '100%', inY: '0%', outX: '-100%', outY: '0%', maskAngle: '270deg', maskStart: '125%', clip: { top: '0%', right: '100%', bottom: '0%', left: '0%' }, rotate: '-90deg' };
            case 'u':
                return { inX: '0%', inY: '-100%', outX: '0%', outY: '100%', maskAngle: '0deg', maskStart: '-125%', clip: { top: '100%', right: '0%', bottom: '0%', left: '0%' }, rotate: '90deg' };
            case 'd':
                return { inX: '0%', inY: '100%', outX: '0%', outY: '-100%', maskAngle: '180deg', maskStart: '125%', clip: { top: '0%', right: '0%', bottom: '100%', left: '0%' }, rotate: '-90deg' };
            case 'tl':
                return { inX: '-100%', inY: '-100%', outX: '100%', outY: '100%', maskAngle: '45deg', maskStart: '-125%', clip: { top: '100%', right: '0%', bottom: '0%', left: '100%' }, rotate: '90deg' };
            case 'tr':
                return { inX: '100%', inY: '-100%', outX: '-100%', outY: '100%', maskAngle: '315deg', maskStart: '125%', clip: { top: '100%', right: '100%', bottom: '0%', left: '0%' }, rotate: '-90deg' };
            case 'bl':
                return { inX: '-100%', inY: '100%', outX: '100%', outY: '-100%', maskAngle: '135deg', maskStart: '-125%', clip: { top: '0%', right: '0%', bottom: '100%', left: '100%' }, rotate: '90deg' };
            case 'br':
                return { inX: '100%', inY: '100%', outX: '-100%', outY: '-100%', maskAngle: '225deg', maskStart: '125%', clip: { top: '0%', right: '100%', bottom: '100%', left: '0%' }, rotate: '-90deg' };
            default:
                return { inX: '0%', inY: '0%', outX: '0%', outY: '0%', maskAngle: '0deg', maskStart: '125%', clip: { top: '0%', right: '0%', bottom: '0%', left: '0%' }, rotate: '90deg' };
        }
    }

    function setClip(target, top, right, bottom, left) {
        if (!target) {
            return;
        }
        target.style.setProperty('--ppt-clip-inset-top', top);
        target.style.setProperty('--ppt-clip-inset-right', right);
        target.style.setProperty('--ppt-clip-inset-bottom', bottom);
        target.style.setProperty('--ppt-clip-inset-left', left);
    }

    function shapeClipFor(shape) {
        switch ((shape || '').toLowerCase()) {
            case 'circle':
                return 'circle(0% at 50% 50%)';
            case 'diamond':
                return 'polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)';
            case 'plus':
                return 'polygon(40% 0%, 60% 0%, 60% 40%, 100% 40%, 100% 60%, 60% 60%, 60% 100%, 40% 100%, 40% 60%, 0% 60%, 0% 40%, 40% 40%)';
            case 'star':
                return 'polygon(50% 0%, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)';
            case 'square':
                return 'inset(0% 0% 0% 0%)';
            default:
                return 'circle(0% at 50% 50%)';
        }
    }

    function clearTransitionClasses(slide) {
        if (!slide) {
            return;
        }
        slide.classList.remove('ppt-transition-enter', 'ppt-transition-exit', 'pre-active');
        Array.from(slide.classList).forEach(cls => {
            if (cls.startsWith('ppt-transition-')) {
                slide.classList.remove(cls);
            }
        });
    }

    function clearTransitionVariables(slide) {
        if (!slide) {
            return;
        }
        cssVarKeys.forEach(key => slide.style.removeProperty(key));
    }

    function getTransitionClasses(effect) {
        switch (effect) {
            case 'cover':
                return { inClass: 'ppt-transition-cover-in', outClass: null };
            case 'uncover':
                return { inClass: null, outClass: 'ppt-transition-uncover-out' };
            case 'flash':
                return { inClass: 'ppt-transition-flash-in', outClass: 'ppt-transition-flash-out' };
            default:
                return { inClass: `ppt-transition-${effect}-in`, outClass: `ppt-transition-${effect}-out` };
        }
    }

    function applyEffectVariables(incoming, outgoing, config, vectors, forward) {
        if (incoming) {
            incoming.style.setProperty('--ppt-transition-duration', `${config.duration}ms`);
            incoming.style.setProperty('--ppt-transition-delay', `${config.delay}ms`);
        }
        if (outgoing) {
            outgoing.style.setProperty('--ppt-transition-duration', `${config.duration}ms`);
        }

        switch (config.effect) {
            case 'push':
                if (incoming) {
                    incoming.style.setProperty('--ppt-translate-in-x', vectors.inX);
                    incoming.style.setProperty('--ppt-translate-in-y', vectors.inY);
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-translate-out-x', vectors.outX);
                    outgoing.style.setProperty('--ppt-translate-out-y', vectors.outY);
                }
                break;
            case 'cover':
                if (incoming) {
                    incoming.style.setProperty('--ppt-translate-in-x', vectors.inX);
                    incoming.style.setProperty('--ppt-translate-in-y', vectors.inY);
                }
                break;
            case 'uncover':
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-translate-out-x', vectors.outX);
                    outgoing.style.setProperty('--ppt-translate-out-y', vectors.outY);
                }
                if (incoming) {
                    incoming.classList.remove('pre-active');
                    incoming.classList.add('active');
                }
                break;
            case 'wipe':
                if (incoming) {
                    setClip(incoming, vectors.clip.top, vectors.clip.right, vectors.clip.bottom, vectors.clip.left);
                }
                if (outgoing) {
                    setClip(outgoing, '0%', '0%', '0%', '0%');
                }
                break;
            case 'split': {
                const orient = (config.orientation || '').toLowerCase();
                if (incoming) {
                    if (orient === 'horz') {
                        setClip(incoming, '50%', '0%', '50%', '0%');
                    } else {
                        setClip(incoming, '0%', '50%', '0%', '50%');
                    }
                }
                if ((config.direction || 'in').toLowerCase() === 'out' && outgoing) {
                    if (orient === 'horz') {
                        setClip(outgoing, '50%', '0%', '50%', '0%');
                    } else {
                        setClip(outgoing, '0%', '50%', '0%', '50%');
                    }
                }
                break;
            }
            case 'bars':
            case 'blinds':
            case 'comb':
                if (incoming) {
                    incoming.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                    incoming.style.setProperty('--ppt-mask-start', vectors.maskStart);
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                    outgoing.style.setProperty('--ppt-mask-start', vectors.maskStart);
                }
                break;
            case 'strips':
                if (incoming) {
                    incoming.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                    incoming.style.setProperty('--ppt-mask-start', vectors.maskStart);
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                    outgoing.style.setProperty('--ppt-mask-start', vectors.maskStart);
                }
                break;
            case 'checker':
                if (incoming) {
                    incoming.style.setProperty('--ppt-mask-start', '220%');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-mask-start', '220%');
                }
                break;
            case 'zoom':
                if (incoming) {
                    incoming.style.setProperty('--ppt-scale-in', forward ? '0.35' : '1.15');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-scale-out', forward ? '1.1' : '0.35');
                }
                break;
            case 'wheel':
                if (incoming) {
                    incoming.style.setProperty('--ppt-wheel-spokes', `${config.spokes || 6}`);
                    incoming.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-wheel-spokes', `${config.spokes || 6}`);
                    outgoing.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                }
                break;
            case 'shape':
                if (incoming) {
                    incoming.style.setProperty('--ppt-clip-shape', shapeClipFor(config.shape));
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-clip-shape', shapeClipFor(config.shape));
                }
                break;
            case 'gallery':
            case 'flip':
            case 'cube':
                if (incoming) {
                    incoming.style.setProperty('--ppt-rotate-in', forward ? '-75deg' : '75deg');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-rotate-in', forward ? '75deg' : '-75deg');
                }
                break;
            case 'switch':
                if (incoming) {
                    incoming.style.setProperty('--ppt-rotate-in', forward ? '-55deg' : '55deg');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-rotate-in', forward ? '55deg' : '-55deg');
                }
                break;
            case 'conveyor':
                if (incoming) {
                    incoming.style.setProperty('--ppt-translate-in-x', vectors.inX);
                    incoming.style.setProperty('--ppt-translate-in-y', vectors.inY);
                    incoming.style.setProperty('--ppt-rotate-in', forward ? '-25deg' : '25deg');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-translate-out-x', vectors.outX);
                    outgoing.style.setProperty('--ppt-translate-out-y', vectors.outY);
                    outgoing.style.setProperty('--ppt-rotate-in', forward ? '25deg' : '-25deg');
                }
                break;
            case 'pan':
                if (incoming) {
                    incoming.style.setProperty('--ppt-translate-in-x', vectors.inX);
                    incoming.style.setProperty('--ppt-translate-in-y', vectors.inY);
                    incoming.style.setProperty('--ppt-scale-in', '1.1');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-scale-out', '0.85');
                }
                break;
            case 'rotate':
                if (incoming) {
                    incoming.style.setProperty('--ppt-rotate-in', forward ? '-120deg' : '120deg');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-rotate-in', forward ? '120deg' : '-120deg');
                }
                break;
            case 'box':
                if (incoming) {
                    setClip(incoming, '50%', '25%', '50%', '25%');
                }
                break;
            default:
                break;
        }
    }

    function activateBlackout(active, duration) {
        if (!blackout) {
            return;
        }
        blackout.style.setProperty('--ppt-transition-duration', `${duration}ms`);
        blackout.classList.toggle('active', active);
    }

    function finalizeTransition(previousIndex, nextIndex, config) {
        const outgoing = slides[previousIndex];
        const incoming = slides[nextIndex];

        clearTransitionClasses(outgoing);
        clearTransitionClasses(incoming);
        clearTransitionVariables(outgoing);
        clearTransitionVariables(incoming);

        if (outgoing) {
            outgoing.classList.remove('active');
            outgoing.classList.remove('pre-active');
        }
        if (incoming) {
            incoming.classList.remove('pre-active');
            incoming.classList.add('active');
        }

        if (config.throughBlack) {
            window.setTimeout(() => activateBlackout(false, Math.max(200, Math.round(config.duration / 2))), 20);
        }

        currentSlide = nextIndex;
        isTransitioning = false;
        updateControls();
        initializeCharts();
        scheduleAutoAdvance(currentSlide);
        if (progress) {
            progress.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
        }
        document.dispatchEvent(new Event('slideChanged'));
    }

    function playTransition(nextIndex, config, meta) {
        isTransitioning = true;
        clearTimeout(autoAdvanceTimer);

        const previousIndex = currentSlide;
        const outgoing = slides[previousIndex];
        const incoming = slides[nextIndex];
        const vectors = directionVector(config.direction);

        if (incoming) {
            clearTransitionClasses(incoming);
            clearTransitionVariables(incoming);
            incoming.classList.add('ppt-transition-enter');
            if (config.effect !== 'uncover') {
                incoming.classList.add('pre-active');
            } else {
                incoming.classList.add('active');
            }
        }

        if (outgoing) {
            clearTransitionClasses(outgoing);
            clearTransitionVariables(outgoing);
            outgoing.classList.add('ppt-transition-exit');
            outgoing.classList.add('active');
        }

        applyEffectVariables(incoming, outgoing, config, vectors, meta.forward);

        const classes = getTransitionClasses(config.effect);
        window.requestAnimationFrame(() => {
            if (incoming && classes.inClass) {
                incoming.classList.add(classes.inClass);
            }
            if (outgoing && classes.outClass) {
                outgoing.classList.add(classes.outClass);
            }
        });

        if (config.throughBlack) {
            activateBlackout(true, Math.max(200, Math.round(config.duration / 2)));
        }

        const totalDuration = config.duration + config.delay;
        window.setTimeout(() => finalizeTransition(previousIndex, nextIndex, config), totalDuration + 48);
    }

    function performImmediate(index) {
        clearTimeout(autoAdvanceTimer);
        const previousIndex = currentSlide;
        const outgoing = slides[previousIndex];
        const incoming = slides[index];

        if (outgoing) {
            clearTransitionClasses(outgoing);
            clearTransitionVariables(outgoing);
            outgoing.classList.remove('active');
            outgoing.classList.remove('pre-active');
        }

        if (incoming) {
            clearTransitionClasses(incoming);
            clearTransitionVariables(incoming);
            incoming.classList.add('active');
            incoming.classList.remove('pre-active');
        }

        currentSlide = index;
        isTransitioning = false;
        updateControls();
        initializeCharts();
        scheduleAutoAdvance(currentSlide);
        if (progress) {
            progress.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
        }
        document.dispatchEvent(new Event('slideChanged'));
    }

    function scheduleAutoAdvance(index) {
        clearTimeout(autoAdvanceTimer);
        const cfg = mergeTransition(index);
        if (cfg.advanceAfter && cfg.advanceAfter > 0) {
            autoAdvanceTimer = window.setTimeout(() => {
                if (index === currentSlide) {
                    const nextIndex = Math.min(totalSlides - 1, index + 1);
                    if (nextIndex !== index) {
                        showSlide(nextIndex, { force: true });
                    }
                }
            }, cfg.advanceAfter);
        }
    }

    function updateControls() {
        if (prevBtn) {
            prevBtn.disabled = currentSlide === 0;
        }
        if (nextBtn) {
            const cfg = mergeTransition(currentSlide);
            nextBtn.disabled = currentSlide === totalSlides - 1 || cfg.advanceOnClick === false;
        }
        if (progress) {
            progress.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
        }
    }

    function initializeCharts() {
        if (typeof Chart === 'undefined') {
            return;
        }
        document.querySelectorAll('canvas[data-chart-config]').forEach(canvas => {
            if (canvas.dataset.initialized === '1') {
                return;
            }
            try {
                const config = JSON.parse(canvas.dataset.chartConfig);
                new Chart(canvas.getContext('2d'), config);
                canvas.dataset.initialized = '1';
            } catch (error) {
                console.error('Chart initialization failed', error);
            }
        });
    }

    function showSlide(index, options = {}) {
        if (index < 0 || index >= totalSlides || index === currentSlide) {
            return;
        }
        if (isTransitioning) {
            return;
        }
        const forward = index > currentSlide;
        const currentConfig = mergeTransition(currentSlide);
        if (forward && !options.force && currentConfig.advanceOnClick === false) {
            return;
        }

        const nextConfig = mergeTransition(index);
        if (!forward && nextConfig.direction) {
            nextConfig.direction = invertDirection(nextConfig.direction);
        }

        if (nextConfig.effect === 'cut') {
            performImmediate(index);
            return;
        }

        if (!nextConfig.duration || nextConfig.duration < 1) {
            nextConfig.duration = 700;
        }

        playTransition(index, nextConfig, { forward });
    }

    if (prevBtn) {
        prevBtn.addEventListener('click', () => showSlide(currentSlide - 1, { force: true }));
    }
    if (nextBtn) {
        nextBtn.addEventListener('click', () => showSlide(currentSlide + 1));
    }

    document.addEventListener('keydown', event => {
        if (event.key === 'ArrowRight' || event.key === 'PageDown' || event.key === ' ') {
            if (mergeTransition(currentSlide).advanceOnClick === false) {
                return;
            }
            event.preventDefault();
            showSlide(Math.min(totalSlides - 1, currentSlide + 1));
        } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
            event.preventDefault();
            showSlide(Math.max(0, currentSlide - 1), { force: true });
        }
    });

    window.addEventListener('resize', () => window.requestAnimationFrame(applyScale));

    window.addEventListener('load', () => {
        slides.forEach((slide, idx) => {
            slide.classList.toggle('active', idx === 0);
            slide.classList.toggle('pre-active', idx !== 0);
        });
        updateControls();
        applyScale();
        initializeCharts();
        scheduleAutoAdvance(0);
        document.dispatchEvent(new Event('slideChanged'));
    });

    window.showSlide = showSlide;
    window.nextSlide = () => showSlide(currentSlide + 1);
    window.prevSlide = () => showSlide(currentSlide - 1, { force: true });
})();
""".strip()

        base_js = (js_template
                   .replace('__SLIDE_WIDTH__', f"{slide_width:.2f}")
                   .replace('__SLIDE_HEIGHT__', f"{slide_height:.2f}")
                   .replace('__TRANSITION_DATA__', transition_json))

        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{self.pptx_path.stem}</title>
    <link rel="stylesheet" href="{self.pptx_path.stem}.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js" defer></script>
    <script src="{self.pptx_path.stem}.js" defer></script>
</head>
<body>
    <div id="presentation" data-slide-count="{total_slides}">
        <div class="progress-bar" id="progress"></div>
        <div class="stage-wrapper">
            <div class="slide-stage">
                {''.join(slides_markup)}
                <div class="transition-blackout" id="transitionBlackout"></div>
            </div>
        </div>
        <div class="controls">
            <button id="prev" type="button">← Previous</button>
            <button id="next" type="button">Next →</button>
        </div>
    </div>
</body>
</html>
""".strip()

        return html_content, css_content, base_js

    def generate_element_html(self, element, slide_width, slide_height) -> str:
        """요소 HTML 생성 (도형, 텍스트, 미디어 등)"""
        if element.get('type') == 'group':
            # Phase 2: Nested Group Rendering
            pos = element['position']
            styles = [
                f"position: absolute",
                f"left: {pos['x']:.2f}px",
                f"top: {pos['y']:.2f}px",
                f"width: {pos['width']:.2f}px",
                f"height: {pos['height']:.2f}px",
                f"z-index: {element.get('z_index', 1)}"
            ]
            if pos.get('rotation'):
                styles.append(f"transform: rotate({pos['rotation']:.3f}deg)")

            # Render children recursively
            children_html = []
            for child in element['children']:
                # Pass parent dimensions (though child positions are relative, function sig needs them)
                children_html.append(self.generate_element_html(child, pos['width'], pos['height']))
            
            return f'<div class="group-shape" style="{"; ".join(styles)}">{"".join(children_html)}</div>'

        if element.get('type') == 'table':
            return self.generate_table_html(element, slide_width, slide_height)

        pos = element['position']
        styles = [
            f"position: absolute",
            f"left: {pos.get('x', 0.0):.2f}px",
            f"top: {pos.get('y', 0.0):.2f}px",
            f"width: {pos.get('width', 0.0):.2f}px",
            f"height: {pos.get('height', 0.0):.2f}px",
            f"z-index: {element.get('z_index', 1)}"
        ]

        transforms = []
        if pos.get('rotation'):
            transforms.append(f"rotate({pos['rotation']:.3f}deg)")
        if pos.get('flip_h'):
            transforms.append("scaleX(-1)")
        if pos.get('flip_v'):
            transforms.append("scaleY(-1)")
            
        if transforms:
            styles.append(f"transform: {' '.join(transforms)}")

        pivot_x = pos.get('pivot_x')
        pivot_y = pos.get('pivot_y')
        if pivot_x is None:
            pivot_x = pos.get('x', 0.0) + pos.get('width', 0.0) / 2.0
        if pivot_y is None:
            pivot_y = pos.get('y', 0.0) + pos.get('height', 0.0) / 2.0
        origin_x = pivot_x - pos.get('x', 0.0)
        origin_y = pivot_y - pos.get('y', 0.0)
        styles.append(f"transform-origin: {origin_x:.2f}px {origin_y:.2f}px")

        # Phase 2: 그림자 추가
        if element.get('shadow'):
            styles.append(f"box-shadow: {element['shadow']}")

        # Phase 2: 반사 추가
        if element.get('reflection'):
            styles.append(element['reflection'])

        has_custom_geometry = bool(element.get('custom_geometry'))
        svg_fragment = ''
        if has_custom_geometry:
            styles.append("background-color: transparent")
            shape_converter = ShapeGeometryConverter(self.ns, self.logger)
            
            # Pass image URL if this shape has an image fill
            image_url = element.get('image')
            
            svg_fragment = shape_converter._build_svg_fragment(
                element['custom_geometry'],
                element.get('fill', {}),
                element.get('border', {}),
                image_url=image_url
            )
        else:
            fill = element.get('fill', {'type': 'none'})
            if fill['type'] == 'solid':
                styles.append(f"background-color: {fill['color']}")
            elif fill['type'] == 'gradient' and fill.get('gradient'):
                styles.append(f"background: {self._build_gradient_css(fill)}")
            else:
                styles.append("background-color: transparent")

        border = element.get('border', {'width': 0})
        if border['width'] > 0:
            styles.append(f"border: {border['width']:.2f}px {border.get('style', 'solid')} {border.get('color', '#000000')}")

        content = []

        if svg_fragment:
            content.append(svg_fragment)

        video_payload = element.get('video')
        if video_payload:
            if isinstance(video_payload, dict):
                content.append(self._render_video_element(video_payload))
            else:
                legacy_payload = {
                    'sources': [{
                        'src': video_payload,
                        'mime': 'video/mp4',
                        'external': False
                    }],
                    'poster': element.get('image'),
                    'iframe': None,
                    'fallback': None
                }
                content.append(self._render_video_element(legacy_payload))
        elif element.get('audio'):
            content.append(f'''<audio controls style="width: 100%;">
                <source src="{element['audio']}">
                Your browser does not support the audio tag.
            </audio>''')
        elif element.get('image') and not has_custom_geometry:
            # Only add img tag if NOT using custom geometry (which handles image via SVG clipPath)
            image_styles = ["width: 100%", "height: 100%"]
            crop = element.get('image_crop') or {}
            stretch = element.get('image_stretch')
            if crop:
                top = crop.get('t', 0) * 100
                right = crop.get('r', 0) * 100
                bottom = crop.get('b', 0) * 100
                left = crop.get('l', 0) * 100
                image_styles.append("object-fit: cover")
                image_styles.append(f"clip-path: inset({top:.2f}% {right:.2f}% {bottom:.2f}% {left:.2f}%)")
            elif stretch:
                image_styles.append("object-fit: fill")
            else:
                image_styles.append("object-fit: contain")
                
            content.append(f'<img src="{element["image"]}" style="{"; ".join(image_styles)}">')

        text_props = element.get('text_props') or {}
        wrap_text = text_props.get('wrap_text', True)
        paragraphs_html = self._render_paragraphs(element.get('paragraphs', []), wrap_text=wrap_text)
        if paragraphs_html:
            anchor = element.get('text_anchor', 't')
            padding = element.get('text_padding', {})
            pad_top = padding.get('top', 0.0)
            pad_right = padding.get('right', 0.0)
            pad_bottom = padding.get('bottom', 0.0)
            pad_left = padding.get('left', 0.0)

            # Flexbox layout for vertical alignment
            justify_content = 'flex-start'
            if anchor == 'ctr':
                justify_content = 'center'
            elif anchor == 'b':
                justify_content = 'flex-end'
            elif anchor in ('dist', 'just'):
                justify_content = 'space-between'

            wrapper_styles = [
                "position: relative",
                "display: flex",
                "flex-direction: column",
                f"justify-content: {justify_content}",
                "height: 100%",
                f"padding: {pad_top:.2f}px {pad_right:.2f}px {pad_bottom:.2f}px {pad_left:.2f}px",
                "overflow: visible"  # Allow overlap if needed, or hidden if strict
            ]

            inner_styles = ["width: 100%"]
            if not wrap_text:
                inner_styles.extend(["white-space: nowrap", "width: fit-content"])
            else:
                inner_styles.extend(["white-space: normal", "overflow-wrap: break-word", "word-break: break-word"])

            inner_block = (
                f'<div class="ppt-text-inner" style="{"; ".join(inner_styles)}">{paragraphs_html}</div>'
            )
            content.append(f'<div class="ppt-text-block" style="{"; ".join(wrapper_styles)}">{inner_block}</div>')

        if not paragraphs_html and not video_payload and element.get('fill', {}).get('type') == 'none':
            styles.append("pointer-events: none")

        shape_html = f'<div class="ppt-element" data-shape-id="{element.get("shape_id", "")}" style="{"; ".join(styles)}">{"".join(content)}</div>'
        if element.get('hyperlink'):
            shape_html = f'<a href="{element["hyperlink"]}" target="_blank" class="ppt-link">{shape_html}</a>'

        return shape_html

    def _render_video_element(self, video_info: Dict) -> str:
        """비디오 요소 HTML 생성"""
        iframe_src = video_info.get('iframe')
        fallback_link = video_info.get('fallback')
        if iframe_src:
            attrs = [
                'style="width: 100%; height: 100%; border: 0;"',
                'allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"',
                'allowfullscreen'
            ]
            iframe_html = f'<iframe src="{html.escape(iframe_src)}" {" ".join(attrs)}></iframe>'
            if fallback_link:
                iframe_html += (
                    f'<p style="margin-top: 4px;">'
                    f'<a href="{html.escape(fallback_link)}" target="_blank" rel="noopener noreferrer">'
                    f'Open video'
                    f'</a></p>'
                )
            return iframe_html

        sources = video_info.get('sources') or []
        poster = video_info.get('poster')

        attr_parts = ['controls', 'preload="metadata"', 'style="width: 100%; height: 100%; object-fit: contain;"']
        primary_source = None
        for source in sources:
            src_value = source.get('src')
            if src_value:
                primary_source = src_value
                break

        if primary_source:
            attr_parts.append(f'src="{html.escape(primary_source)}"')
        if poster:
            attr_parts.append(f'poster="{html.escape(poster)}"')

        source_tags = []
        for source in sources:
            src_value = source.get('src')
            if not src_value:
                continue
            mime_value = source.get('mime') or 'video/mp4'
            source_tags.append(
                f'<source src="{html.escape(src_value)}" type="{html.escape(mime_value)}">'
            )

        video_html = (
            f'<video {" ".join(attr_parts)}>'
            f'{"".join(source_tags)}'
            f'Your browser does not support the video tag.'
            f'</video>'
        )

        if fallback_link:
            video_html += (
                f'<p style="margin-top: 4px;">'
                f'<a href="{html.escape(fallback_link)}" target="_blank" rel="noopener noreferrer">'
                f'Open video'
                f'</a></p>'
            )

        return video_html

    def generate_table_html(self, table_data, slide_width, slide_height) -> str:
        """테이블 HTML 생성"""
        pos = table_data['position']
        table_width = pos.get('width') or sum(table_data.get('col_widths', []))
        table_height = pos.get('height', 0)

        table_styles = [
            f"position: absolute",
            f"left: {pos.get('x', 0.0):.2f}px",
            f"top: {pos.get('y', 0.0):.2f}px",
            f"width: {table_width:.2f}px",
            f"height: {table_height:.2f}px",
            f"z-index: {table_data.get('z_index', 1)}",
            "border-collapse: collapse",
            "table-layout: fixed"
        ]

        colgroup_html = ''.join([
            f'<col style="width: {width:.2f}px;">'
            for width in table_data.get('col_widths', [])
        ])

        rows_html = []
        for row in table_data['rows']:
            cells_html = []

            for cell in row['cells']:
                cell_styles = []

                if cell['fill']['type'] == 'solid':
                    cell_styles.append(f"background-color: {cell['fill']['color']}")

                if cell['borders']:
                    for side, border in cell['borders'].items():
                        cell_styles.append(f"border-{side}: {border['width']:.2f}px {border['style']} {border['color']}")

                v_align_map = {'t': 'top', 'ctr': 'middle', 'b': 'bottom'}
                v_align = v_align_map.get(cell['vertical_align'], 'top')
                cell_styles.append(f"vertical-align: {v_align}")
                padding = cell.get('text_padding', {})
                pad_top = padding.get('top', 5.0)
                pad_right = padding.get('right', 5.0)
                pad_bottom = padding.get('bottom', 5.0)
                pad_left = padding.get('left', 5.0)
                cell_styles.append(f"padding: {pad_top:.2f}px {pad_right:.2f}px {pad_bottom:.2f}px {pad_left:.2f}px")
                cell_props = cell.get('text_props') or {}
                if not cell_props.get('wrap_text', True):
                    cell_styles.append("white-space: nowrap")
                    cell_styles.append("overflow: visible")
                else:
                    cell_styles.append("overflow: hidden")

                cell_html = self._render_paragraphs(cell.get('text', []), wrap_text=cell_props.get('wrap_text', True))

                colspan_attr = f' colspan="{cell["colspan"]}"' if cell['colspan'] > 1 else ''
                rowspan_attr = f' rowspan="{cell["rowspan"]}"' if cell['rowspan'] > 1 else ''
                cells_html.append(f'<td{colspan_attr}{rowspan_attr} style="{"; ".join(cell_styles)}">{cell_html}</td>')

            rows_html.append(f'<tr>{"".join(cells_html)}</tr>')

        return (
            f'<table class="ppt-element ppt-table" style="{"; ".join(table_styles)}">'
            f'<colgroup>{colgroup_html}</colgroup>'
            f'{"".join(rows_html)}'
            f'</table>'
        )

    @staticmethod
    def _escape_html(text):
        """HTML 특수 문자 이스케이프"""
        return (text
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&#39;'))


def main():
    """CLI 진입점"""
    if len(sys.argv) < 2:
        print("Usage: python convert_pptx_to_html_v2.py <input.pptx> [output_directory] [dpi]")
        print("\nExample:")
        print("  python convert_pptx_to_html_v2.py presentation.pptx ./output 150")
        sys.exit(1)

    pptx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 150

    if not Path(pptx_path).exists():
        print(f"Error: File '{pptx_path}' not found")
        sys.exit(1)

    # 로그 파일 설정
    log_file = Path(output_dir) / "conversion.log" if output_dir else None

    converter = EnhancedPPTXToHTMLV2(pptx_path, output_dir, dpi, log_file)
    result = converter.convert()

    if result:
        print(f"\n✅ HTML presentation created successfully!")
        print(f"📂 Output: {result}")
        sys.exit(0)
    else:
        print(f"\n❌ Conversion failed. Check logs for details.")
        sys.exit(1)


if __name__ == "__main__":
    main()
