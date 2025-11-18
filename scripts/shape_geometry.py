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
            'line': self._create_line,
            'straightConnector1': self._create_line,
            'bentConnector2': self._create_bent_connector_2,
            'bentConnector3': self._create_bent_connector_3,
            'curvedConnector3': self._create_curved_connector_3,
        }

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
            if preset_type in self.preset_shapes:
                svg_geom = {
                    'path': self.preset_shapes[preset_type](),
                    'view_box': (0.0, 0.0, 100.0, 100.0)
                }
                if svg_geom and self.logger:
                    self.logger.increment_custom_shape()
                    self.logger.debug(f"Converted preset shape: {preset_type}")
                return svg_geom

        return None

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

    def _create_line(self) -> str:
        """직선/커넥터 SVG 경로"""
        return "M 0 0 L 100 100"

    def _create_bent_connector_2(self) -> str:
        """꺾인 연결선 (2 segments)"""
        return "M 0 0 L 50 0 L 50 100 L 100 100"

    def _create_bent_connector_3(self) -> str:
        """꺾인 연결선 (3 segments)"""
        return "M 0 0 L 50 0 L 50 100 L 100 100"

    def _create_curved_connector_3(self) -> str:
        """곡선 연결선"""
        return "M 0 0 Q 50 0 50 50 T 100 100"

    def _build_svg_fragment(self, svg_geom, fill: Dict, border: Dict, image_url: Optional[str] = None) -> str:
        """SVG 요소 생성"""
        if isinstance(svg_geom, dict):
            path_str = svg_geom.get('path', '')
            view_box = svg_geom.get('view_box', (0.0, 0.0, 100.0, 100.0))
        else:
            path_str = svg_geom
            view_box = (0.0, 0.0, 100.0, 100.0)

        stroke = border.get('color', '#000000')
        stroke_width = border.get('width', 0)
        
        content = []

        if image_url:
            # 이미지 채우기 (ClipPath 사용)
            # ID 생성 (경로 문자열 해시 사용으로 간단하게 고유성 확보 시도)
            import hashlib
            clip_id = f"clip_{hashlib.md5(path_str.encode()).hexdigest()[:8]}"
            
            content.append(f'<defs><clipPath id="{clip_id}"><path d="{path_str}" /></clipPath></defs>')
            # 이미지 비율 유지 설정은 상황에 따라 다를 수 있으나, 도형 채우기는 보통 stretch(none) 또는 cover(slice)
            # 여기서는 fill과 유사하게 전체 영역을 채우도록 설정
            content.append(f'<image href="{image_url}" x="{view_box[0]}" y="{view_box[1]}" width="{view_box[2]}" height="{view_box[3]}" preserveAspectRatio="none" clip-path="url(#{clip_id})" />')
            
            # 테두리가 있는 경우 위에 그리기
            if stroke_width > 0:
                content.append(f'<path d="{path_str}" fill="none" stroke="{stroke}" stroke-width="{stroke_width}" />')
        else:
            # 일반 채우기 (Solid/Gradient)
            fill_attr = 'none'
            if fill.get('type') == 'solid':
                fill_attr = fill.get('color', 'none')
            elif fill.get('type') == 'gradient' and fill.get('gradient'):
                fill_attr = fill['gradient'][0][1]
            
            # 테두리 기본값 처리
            if stroke_width == 0 and fill_attr == 'none':
                 stroke_width = 0 # 투명
            elif stroke_width == 0:
                 stroke_width = 0

            content.append(f'<path d="{path_str}" fill="{fill_attr}" stroke="{stroke}" stroke-width="{stroke_width}" />')

        return (
            f'<svg viewBox="{view_box[0]} {view_box[1]} {view_box[2]} {view_box[3]}" preserveAspectRatio="none" '
            'xmlns="http://www.w3.org/2000/svg" style="width: 100%; height: 100%; overflow: visible;">'
            f'{"".join(content)}'
            '</svg>'
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
        ]

        transform_ops = []
        if position.get('rotation', 0) != 0:
            transform_ops.append(f"rotate({position['rotation']}deg)")
        if position.get('flip_h'):
            transform_ops.append("scaleX(-1)")
        if position.get('flip_v'):
            transform_ops.append("scaleY(-1)")

        if transform_ops:
            styles.append(f"transform: {' '.join(transform_ops)}")

        if shadow:
            styles.append(f"box-shadow: {shadow}")
        if reflection:
            styles.append(reflection)

        fragment = self._build_svg_fragment(svg_geom, fill, border)
        return f'<div class="ppt-element" style="{"; ".join(styles)}">{fragment}</div>'
