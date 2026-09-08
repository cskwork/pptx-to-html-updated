#!/usr/bin/env python3
"""
Custom Shape Geometry Converter (DrawingML → SVG)
생성일: 2025-01-21
설명: PowerPoint의 커스텀 도형 경로를 SVG 경로로 변환
"""

from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET


class ShapeGeometryConverter:
    """DrawingML 커스텀 도형을 SVG로 변환하는 클래스"""

    def __init__(self, ns: Dict[str, str], logger=None):
        """
        Args:
            ns: XML 네임스페이스 딕셔너리
            logger: ConversionLogger 인스턴스
        """
        self.ns = ns
        self.logger = logger

        # 프리셋 도형 타입 매핑 (간단한 SVG 경로)
        self.preset_shapes = {
            'rect': self._create_rectangle,
            'roundRect': self._create_rounded_rectangle,
            'ellipse': self._create_ellipse,
            'triangle': self._create_triangle,
            'rightTriangle': self._create_right_triangle,
            'diamond': self._create_diamond,
            'pentagon': self._create_pentagon,
            'hexagon': self._create_hexagon,
            'octagon': self._create_octagon,
            'star5': self._create_star,
            'arrow': self._create_arrow,
            'leftArrow': self._create_left_arrow,
            'rightArrow': self._create_right_arrow,
            'upArrow': self._create_up_arrow,
            'downArrow': self._create_down_arrow,
            'leftRightArrow': self._create_left_right_arrow,
            'upDownArrow': self._create_up_down_arrow,
            'bentArrow': self._create_bent_arrow,
            'circularArrow': self._create_circular_arrow,
            'flowChartProcess': self._create_rectangle,
            'flowChartDecision': self._create_diamond,
            'flowChartData': self._create_parallelogram,
            'flowChartTerminator': self._create_terminator,
            'flowChartDocument': self._create_document,
            'flowChartManualOperation': self._create_manual_operation,
            'homePlate': self._create_home_plate,
            'smileyFace': self._create_smiley_face,
            'round2SameRect': self._create_round2_same_rect,
        }

        # avLst 조정값이 모양을 크게 바꾸는 프리셋
        self.adjustable_presets = {
            'chevron': self._create_chevron,
        }

        # 눈/입을 구멍으로 뚫어야 하는 프리셋
        self.evenodd_presets = {'smileyFace'}

    def extract_custom_geometry(self, sp_pr: ET.Element) -> Optional[Dict[str, object]]:
        """
        커스텀 도형 지오메트리를 SVG 경로로 추출

        Args:
            sp_pr: Shape Properties 엘리먼트

        Returns:
            SVG 경로 문자열 또는 None
        """
        if sp_pr is None:
            return None

        # 1. 커스텀 지오메트리 확인
        cust_geom = sp_pr.find('.//a:custGeom', self.ns)
        if cust_geom is not None:
            svg_geom = self._parse_custom_geometry(cust_geom)
            if svg_geom and self.logger:
                self.logger.increment_custom_shape()
            return svg_geom

        # 2. 프리셋 지오메트리 확인
        prst_geom = sp_pr.find('.//a:prstGeom', self.ns)
        if prst_geom is not None:
            preset_type = prst_geom.get('prst')
            adjustable = self.adjustable_presets.get(preset_type)
            if adjustable is not None:
                width, height = self._shape_extent(sp_pr)
                path = adjustable(self._parse_adjust_values(prst_geom), width, height)
            elif preset_type in self.preset_shapes:
                path = self.preset_shapes[preset_type]()
            else:
                return None

            svg_geom = {
                'path': path,
                'view_box': (0.0, 0.0, 100.0, 100.0)
            }
            if preset_type in self.evenodd_presets:
                svg_geom['fill_rule'] = 'evenodd'
            if self.logger:
                self.logger.increment_custom_shape()
                self.logger.debug(f"Converted preset shape: {preset_type}")
            return svg_geom

        return None

    def _parse_adjust_values(self, prst_geom: ET.Element) -> Dict[str, float]:
        """avLst의 조정값 파싱"""
        values: Dict[str, float] = {}
        for gd in prst_geom.findall('.//a:avLst/a:gd', self.ns):
            name = gd.get('name')
            fmla = gd.get('fmla', '')
            if not name or not fmla.startswith('val '):
                continue
            try:
                values[name] = float(fmla[4:])
            except ValueError:
                continue
        return values

    def _shape_extent(self, sp_pr: ET.Element) -> Tuple[float, float]:
        """도형 크기(EMU) 반환"""
        ext = sp_pr.find('.//a:xfrm/a:ext', self.ns)
        if ext is None:
            return 1.0, 1.0
        try:
            return float(ext.get('cx', 1)) or 1.0, float(ext.get('cy', 1)) or 1.0
        except ValueError:
            return 1.0, 1.0

    def _parse_custom_geometry(self, cust_geom: ET.Element) -> Optional[Dict[str, object]]:
        """
        커스텀 지오메트리 XML을 SVG 경로로 변환

        Args:
            cust_geom: custGeom 엘리먼트

        Returns:
            SVG 경로 문자열
        """
        path_list = cust_geom.find('.//a:pathLst', self.ns)
        if path_list is None:
            return None

        svg_commands = []
        min_x = float('inf')
        min_y = float('inf')
        max_x = float('-inf')
        max_y = float('-inf')

        def update_bounds(x_val: str, y_val: str):
            nonlocal min_x, min_y, max_x, max_y
            try:
                x = float(x_val)
                y = float(y_val)
            except (TypeError, ValueError):
                return
            min_x = min(min_x, x)
            min_y = min(min_y, y)
            max_x = max(max_x, x)
            max_y = max(max_y, y)

        for path in path_list.findall('.//a:path', self.ns):
            # 경로 명령 파싱
            for element in path:
                tag = element.tag.replace(f"{{{self.ns['a']}}}", "")

                if tag == 'moveTo':
                    # M (MoveTo)
                    pt = element.find('.//a:pt', self.ns)
                    if pt is not None:
                        x = pt.get('x', '0')
                        y = pt.get('y', '0')
                        update_bounds(x, y)
                        svg_commands.append(f"M {x} {y}")

                elif tag == 'lnTo':
                    # L (LineTo)
                    pt = element.find('.//a:pt', self.ns)
                    if pt is not None:
                        x = pt.get('x', '0')
                        y = pt.get('y', '0')
                        update_bounds(x, y)
                        svg_commands.append(f"L {x} {y}")

                elif tag == 'cubicBezTo':
                    # C (Cubic Bezier)
                    pts = element.findall('.//a:pt', self.ns)
                    if len(pts) == 3:
                        coords = []
                        for pt in pts:
                            coords.append(pt.get('x', '0'))
                            coords.append(pt.get('y', '0'))
                            update_bounds(pt.get('x', '0'), pt.get('y', '0'))
                        svg_commands.append(f"C {' '.join(coords)}")

                elif tag == 'quadBezTo':
                    # Q (Quadratic Bezier)
                    pts = element.findall('.//a:pt', self.ns)
                    if len(pts) == 2:
                        coords = []
                        for pt in pts:
                            coords.append(pt.get('x', '0'))
                            coords.append(pt.get('y', '0'))
                            update_bounds(pt.get('x', '0'), pt.get('y', '0'))
                        svg_commands.append(f"Q {' '.join(coords)}")

                elif tag == 'arcTo':
                    # A (Arc) - 근사치로 변환
                    # DrawingML arcTo는 SVG와 다른 파라미터를 사용하므로 근사화 필요
                    if self.logger:
                        self.logger.warning("arcTo command approximated (not fully supported)")

                elif tag == 'close':
                    # Z (Close path)
                    svg_commands.append("Z")

        if svg_commands:
            if min_x == float('inf') or min_y == float('inf'):
                min_x = min_y = 0.0
                max_x = max_y = 100.0

            width = max(max_x - min_x, 1.0)
            height = max(max_y - min_y, 1.0)

            return {
                'path': ' '.join(svg_commands),
                'view_box': (min_x, min_y, width, height)
            }

        return None

    # === 프리셋 도형 생성 메서드 ===

    def _create_rectangle(self) -> str:
        """직사각형 SVG 경로"""
        return "M 0 0 L 100 0 L 100 100 L 0 100 Z"

    def _create_rounded_rectangle(self) -> str:
        """둥근 직사각형 SVG 경로"""
        return "M 10 0 L 90 0 Q 100 0 100 10 L 100 90 Q 100 100 90 100 L 10 100 Q 0 100 0 90 L 0 10 Q 0 0 10 0 Z"

    def _create_ellipse(self) -> str:
        """타원 SVG 경로"""
        return "M 50 0 A 50 50 0 1 1 50 100 A 50 50 0 1 1 50 0 Z"

    def _create_triangle(self) -> str:
        """삼각형 SVG 경로"""
        return "M 50 0 L 100 100 L 0 100 Z"

    def _create_right_triangle(self) -> str:
        """직각 삼각형 SVG 경로"""
        return "M 0 0 L 100 100 L 0 100 Z"

    def _create_diamond(self) -> str:
        """마름모 SVG 경로"""
        return "M 50 0 L 100 50 L 50 100 L 0 50 Z"

    def _create_pentagon(self) -> str:
        """오각형 SVG 경로"""
        return "M 50 0 L 100 38 L 82 100 L 18 100 L 0 38 Z"

    def _create_hexagon(self) -> str:
        """육각형 SVG 경로"""
        return "M 50 0 L 93 25 L 93 75 L 50 100 L 7 75 L 7 25 Z"

    def _create_octagon(self) -> str:
        """팔각형 SVG 경로"""
        return "M 30 0 L 70 0 L 100 30 L 100 70 L 70 100 L 30 100 L 0 70 L 0 30 Z"

    def _create_star(self) -> str:
        """별 (5개 꼭지점) SVG 경로"""
        return "M 50 0 L 61 35 L 98 35 L 68 57 L 79 91 L 50 70 L 21 91 L 32 57 L 2 35 L 39 35 Z"

    def _create_arrow(self) -> str:
        """화살표 (오른쪽) SVG 경로"""
        return "M 0 30 L 70 30 L 70 10 L 100 50 L 70 90 L 70 70 L 0 70 Z"

    def _create_left_arrow(self) -> str:
        """왼쪽 화살표 SVG 경로"""
        return "M 100 30 L 30 30 L 30 10 L 0 50 L 30 90 L 30 70 L 100 70 Z"

    def _create_right_arrow(self) -> str:
        """오른쪽 화살표 SVG 경로"""
        return self._create_arrow()

    def _create_up_arrow(self) -> str:
        """위쪽 화살표 SVG 경로"""
        return "M 30 100 L 30 30 L 10 30 L 50 0 L 90 30 L 70 30 L 70 100 Z"

    def _create_down_arrow(self) -> str:
        """아래쪽 화살표 SVG 경로"""
        return "M 30 0 L 30 70 L 10 70 L 50 100 L 90 70 L 70 70 L 70 0 Z"

    def _create_left_right_arrow(self) -> str:
        """양방향 화살표 (좌우) SVG 경로"""
        return "M 30 30 L 70 30 L 70 10 L 100 50 L 70 90 L 70 70 L 30 70 L 30 90 L 0 50 L 30 10 Z"

    def _create_up_down_arrow(self) -> str:
        """양방향 화살표 (상하) SVG 경로"""
        return "M 30 30 L 30 70 L 10 70 L 50 100 L 90 70 L 70 70 L 70 30 L 90 30 L 50 0 L 10 30 Z"

    def _create_bent_arrow(self) -> str:
        """구부러진 화살표 SVG 경로"""
        return "M 0 50 L 50 50 L 50 30 L 50 10 L 100 50 L 50 90 L 50 70 L 20 70 L 20 50 Z"

    def _create_circular_arrow(self) -> str:
        """원형 화살표 SVG 경로"""
        return "M 50 10 A 40 40 0 1 1 49 10 L 60 0 L 70 10 L 60 20 Z"

    def _create_parallelogram(self) -> str:
        """평행사변형 SVG 경로"""
        return "M 20 0 L 100 0 L 80 100 L 0 100 Z"

    def _create_terminator(self) -> str:
        """터미네이터 (플로우차트) SVG 경로"""
        return "M 20 0 L 80 0 Q 100 0 100 50 Q 100 100 80 100 L 20 100 Q 0 100 0 50 Q 0 0 20 0 Z"

    def _create_document(self) -> str:
        """문서 (플로우차트) SVG 경로"""
        return "M 0 0 L 100 0 L 100 85 Q 75 100 50 85 Q 25 70 0 85 Z"

    def _create_manual_operation(self) -> str:
        """수작업 (플로우차트) SVG 경로"""
        return "M 0 0 L 100 0 L 80 100 L 20 100 Z"

    def _create_chevron(self, adjust: Dict[str, float], width: float, height: float) -> str:
        """셔브론 화살표 SVG 경로 (adj 값이 꺾이는 깊이를 결정)"""
        shortest = min(width, height) or 1.0
        max_adj = 100000.0 * width / shortest
        value = max(0.0, min(adjust.get('adj', 50000.0), max_adj))
        inset = (shortest * value / 100000.0) / width * 100.0
        return (
            f"M 0 0 L {100 - inset:.3f} 0 L 100 50 L {100 - inset:.3f} 100 "
            f"L 0 100 L {inset:.3f} 50 Z"
        )

    def _create_home_plate(self) -> str:
        """오각형 화살표 배너 SVG 경로"""
        return "M 0 0 L 75 0 L 100 50 L 75 100 L 0 100 Z"

    def _create_smiley_face(self) -> str:
        """웃는 얼굴 SVG 경로 (evenodd로 눈/입을 비운다)"""
        return (
            "M 50 0 A 50 50 0 1 1 50 100 A 50 50 0 1 1 50 0 Z "
            "M 33 25 A 6 8 0 1 1 33 41 A 6 8 0 1 1 33 25 Z "
            "M 67 25 A 6 8 0 1 1 67 41 A 6 8 0 1 1 67 25 Z "
            "M 22 60 Q 50 90 78 60 Q 50 78 22 60 Z"
        )

    def _create_round2_same_rect(self) -> str:
        """위쪽 모서리만 둥근 직사각형 SVG 경로"""
        return "M 0 100 L 0 17 Q 0 0 17 0 L 83 0 Q 100 0 100 17 L 100 100 Z"

    def _build_svg_fragment(self, svg_geom, fill: Dict, border: Dict,
                            image: Optional[Dict] = None) -> str:
        """SVG 요소 생성"""
        if isinstance(svg_geom, dict):
            path_str = svg_geom.get('path', '')
            view_box = svg_geom.get('view_box', (0.0, 0.0, 100.0, 100.0))
            fill_rule = svg_geom.get('fill_rule')
        else:
            path_str = svg_geom
            view_box = (0.0, 0.0, 100.0, 100.0)
            fill_rule = None

        fill_attr = 'none'
        if fill.get('type') == 'solid':
            fill_attr = fill.get('color', 'none')
        elif fill.get('type') == 'gradient' and fill.get('gradient'):
            fill_attr = fill['gradient'][0][1]

        stroke = border.get('color', '#000000')
        stroke_width = border.get('width', 1)

        fill_rule_attr = f' fill-rule="{fill_rule}"' if fill_rule else ''
        image_markup = self._build_clipped_image(image, path_str, view_box) if image else ''

        # 보이지도 않고 그림도 없는 도형은 SVG 자체를 생략
        if not image_markup and fill_attr == 'none' and not float(stroke_width or 0):
            return ''

        return (
            f'<svg class="ppt-shape" viewBox="{view_box[0]} {view_box[1]} {view_box[2]} {view_box[3]}" '
            'preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg">'
            f'{image_markup}'
            f'<path d="{path_str}" fill="{fill_attr}" stroke="{stroke}" stroke-width="{stroke_width}"'
            f'{fill_rule_attr} /></svg>'
        )

    @staticmethod
    def _build_clipped_image(image: Dict, path_str: str, view_box) -> str:
        """도형 경로로 잘라낸 그림 채우기 생성"""
        href = image.get('href')
        clip_id = image.get('clip_id')
        if not href or not clip_id:
            return ''

        crop = image.get('crop') or {}
        left = crop.get('l', 0.0)
        top = crop.get('t', 0.0)
        visible_width = max(0.0001, 1.0 - left - crop.get('r', 0.0))
        visible_height = max(0.0001, 1.0 - top - crop.get('b', 0.0))

        vb_x, vb_y, vb_w, vb_h = view_box
        img_w = vb_w / visible_width
        img_h = vb_h / visible_height

        return (
            f'<defs><clipPath id="{clip_id}"><path d="{path_str}" /></clipPath></defs>'
            f'<image href="{href}" x="{vb_x - left * img_w:.4f}" y="{vb_y - top * img_h:.4f}" '
            f'width="{img_w:.4f}" height="{img_h:.4f}" preserveAspectRatio="none" '
            f'clip-path="url(#{clip_id})" />'
        )

    def generate_svg_html(self, svg_geom, position: Dict, fill: Dict, border: Dict,
                         z_index: int = 1, shadow: Optional[str] = None, reflection: Optional[str] = None) -> str:
        """SVG 경로를 HTML SVG 엘리먼트로 변환"""
        # 스타일 구성 (픽셀 기반)
        styles = [
            f"position: absolute",
            f"left: {position.get('x', 0.0):.2f}px",
            f"top: {position.get('y', 0.0):.2f}px",
            f"width: {position.get('width', 0.0):.2f}px",
            f"height: {position.get('height', 0.0):.2f}px",
            f"z-index: {z_index}",
            "transform-origin: top left"
        ]

        if position.get('rotation', 0) != 0:
            styles.append(f"transform: rotate({position['rotation']}deg)")

        if shadow:
            styles.append(f"box-shadow: {shadow}")
        if reflection:
            styles.append(reflection)

        fragment = self._build_svg_fragment(svg_geom, fill, border)
        return f'<div class="ppt-element" style="{"; ".join(styles)}">{fragment}</div>'
