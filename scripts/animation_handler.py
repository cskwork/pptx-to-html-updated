#!/usr/bin/env python3
"""
Animation Extraction and CSS/JavaScript Mapping
생성일: 2025-01-21
설명: PowerPoint 애니메이션을 CSS/JavaScript 애니메이션으로 변환
"""

from typing import Dict, List, Optional
from xml.etree import ElementTree as ET


class AnimationHandler:
    """PowerPoint 애니메이션을 웹 애니메이션으로 변환하는 클래스"""

    def __init__(self, ns: Dict[str, str], logger=None):
        """
        Args:
            ns: XML 네임스페이스 딕셔너리
            logger: ConversionLogger 인스턴스
        """
        self.ns = ns
        self.logger = logger

        # PowerPoint 애니메이션 → CSS 애니메이션 매핑
        self.animation_mapping = {
            'appear': 'fadeIn',
            'fade': 'fadeIn',
            'fadeIn': 'fadeIn',
            'fadeOut': 'fadeOut',
            'fly': 'slideIn',
            'flyIn': 'slideIn',
            'flyOut': 'slideOut',
            'wipe': 'wipeIn',
            'split': 'splitIn',
            'strips': 'fadeIn',
            'shape': 'scaleIn',
            'wheel': 'rotateIn',
            'randomBars': 'fadeIn',
            'grow': 'scaleIn',
            'shrink': 'scaleOut',
            'zoom': 'zoomIn',
            'swivel': 'rotateIn',
            'bounce': 'bounceIn'
        }

        # 방향 매핑
        self.direction_mapping = {
            'fromLeft': 'Left',
            'fromRight': 'Right',
            'fromTop': 'Top',
            'fromBottom': 'Bottom',
            'fromTopLeft': 'TopLeft',
            'fromTopRight': 'TopRight',
            'fromBottomLeft': 'BottomLeft',
            'fromBottomRight': 'BottomRight'
        }

    def extract_slide_animations(self, slide_xml: ET.Element) -> List[Dict]:
        """
        슬라이드에서 애니메이션 추출

        Args:
            slide_xml: 슬라이드 XML 엘리먼트

        Returns:
            애니메이션 정보 리스트
        """
        animations = []

        # timing 엘리먼트 찾기
        timing = slide_xml.find('.//p:timing', self.ns)
        if timing is None:
            return animations

        # 애니메이션 시퀀스 파싱
        tn_lst = timing.find('.//p:tnLst', self.ns)
        if tn_lst is None:
            return animations

        for par in tn_lst.findall('.//p:par', self.ns):
            anim_data = self._parse_animation_node(par)
            if anim_data:
                animations.append(anim_data)

        if animations and self.logger:
            self.logger.debug(f"Extracted {len(animations)} animation(s)")

        return animations

    def _parse_animation_node(self, node: ET.Element) -> Optional[Dict]:
        """
        애니메이션 노드 파싱

        Args:
            node: 애니메이션 노드 엘리먼트

        Returns:
            애니메이션 데이터 딕셔너리
        """
        try:
            # cTn (Common Time Node) 찾기
            c_tn = node.find('.//p:cTn', self.ns)
            if c_tn is None:
                return None

            # 지속 시간 (밀리초)
            duration = self._parse_time_value(c_tn.get('dur'), default=1000)

            # 딜레이
            delay = self._parse_time_value(c_tn.get('delay'), default=0)

            # 애니메이션 효과 찾기
            anim_effect = node.find('.//p:animEffect', self.ns)
            if anim_effect is None:
                return None

            # 효과 타입
            transition = anim_effect.get('transition', 'in')
            filter_type = anim_effect.get('filter', 'fade')

            # 타겟 찾기
            tgt_el = node.find('.//p:tgtEl', self.ns)
            if tgt_el is None:
                return None

            # 타겟 shape ID
            sp_tgt = tgt_el.find('.//p:spTgt', self.ns)
            if sp_tgt is None:
                return None

            shape_id = sp_tgt.get('spid', '')

            # CSS 애니메이션 이름 매핑
            css_animation = self.animation_mapping.get(filter_type, 'fadeIn')

            return {
                'shape_id': shape_id,
                'type': css_animation,
                'duration': duration,
                'delay': delay,
                'transition': transition
            }

        except Exception as e:
            if self.logger:
                self.logger.warning(f"Failed to parse animation node", exception=e)
            return None

    def generate_css_animations(self) -> str:
        """
        CSS 애니메이션 키프레임 생성

        Returns:
            CSS 문자열
        """
        css = """
/* PowerPoint Animation Effects */
@keyframes fadeIn {
    from { opacity: 0; }
    to { opacity: 1; }
}

@keyframes fadeOut {
    from { opacity: 1; }
    to { opacity: 0; }
}

@keyframes slideInLeft {
    from {
        opacity: 0;
        transform: translateX(-100%);
    }
    to {
        opacity: 1;
        transform: translateX(0);
    }
}

@keyframes slideInRight {
    from {
        opacity: 0;
        transform: translateX(100%);
    }
    to {
        opacity: 1;
        transform: translateX(0);
    }
}

@keyframes slideInTop {
    from {
        opacity: 0;
        transform: translateY(-100%);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

@keyframes slideInBottom {
    from {
        opacity: 0;
        transform: translateY(100%);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

@keyframes scaleIn {
    from {
        opacity: 0;
        transform: scale(0);
    }
    to {
        opacity: 1;
        transform: scale(1);
    }
}

@keyframes scaleOut {
    from {
        opacity: 1;
        transform: scale(1);
    }
    to {
        opacity: 0;
        transform: scale(0);
    }
}

@keyframes zoomIn {
    from {
        opacity: 0;
        transform: scale(0.3);
    }
    to {
        opacity: 1;
        transform: scale(1);
    }
}

@keyframes rotateIn {
    from {
        opacity: 0;
        transform: rotate(-200deg);
    }
    to {
        opacity: 1;
        transform: rotate(0);
    }
}

@keyframes bounceIn {
    0% {
        opacity: 0;
        transform: scale(0.3);
    }
    50% {
        opacity: 1;
        transform: scale(1.05);
    }
    70% {
        transform: scale(0.9);
    }
    100% {
        opacity: 1;
        transform: scale(1);
    }
}

.animated {
    animation-duration: 1s;
    animation-fill-mode: both;
}

.animated-fast {
    animation-duration: 0.5s;
}

.animated-slow {
    animation-duration: 2s;
}
"""
        return css

    def generate_animation_javascript(self, animations: List[Dict], shape_id_map: Dict[str, str]) -> str:
        """
        애니메이션 제어 JavaScript 생성

        Args:
            animations: 애니메이션 정보 리스트
            shape_id_map: Shape ID → HTML Element ID 매핑

        Returns:
            JavaScript 문자열
        """
        # if not animations:
        #     return ""

        # 애니메이션 데이터를 JSON 형식으로 변환
        anim_configs = []
        for anim in animations:
            shape_id = anim['shape_id']
            element_id = shape_id_map.get(shape_id)

            if element_id:
                anim_configs.append({
                    'elementId': element_id,
                    'type': anim['type'],
                    'duration': anim['duration'],
                    'delay': anim['delay']
                })

        # if not anim_configs:
        #     return ""

        import json
        anim_json = json.dumps(anim_configs, ensure_ascii=False)

        js = f"""
// PowerPoint Animation Controller
(function() {{
    const animations = {anim_json};

    function playAnimations() {{
        animations.forEach(function(anim) {{
            const element = document.getElementById(anim.elementId);
            if (element) {{
                // 초기 상태: 숨김
                element.style.opacity = '0';

                // 애니메이션 적용
                setTimeout(function() {{
                    element.style.animation = anim.type + ' ' + (anim.duration / 1000) + 's ease-out forwards';
                }}, anim.delay);
            }}
        }});
    }}

    // 슬라이드가 활성화될 때 애니메이션 재생
    document.addEventListener('slideChanged', playAnimations);

    // 초기 슬라이드 애니메이션 재생
    if (document.querySelector('.slide.active')) {{
        setTimeout(playAnimations, 100);
    }}
}})();
"""
        return js

    def apply_shadow_effects(self, sp_pr: ET.Element) -> Optional[str]:
        """
        그림자 효과를 CSS로 변환

        Args:
            sp_pr: Shape Properties 엘리먼트

        Returns:
            CSS box-shadow 문자열 또는 None
        """
        if sp_pr is None:
            return None

        # effectLst에서 그림자 효과 찾기
        effect_lst = sp_pr.find('.//a:effectLst', self.ns)
        if effect_lst is None:
            return None

        # 외부 그림자 (outerShdw)
        outer_shadow = effect_lst.find('.//a:outerShdw', self.ns)
        if outer_shadow is not None:
            return self._parse_outer_shadow(outer_shadow)

        # 내부 그림자 (innerShdw)
        inner_shadow = effect_lst.find('.//a:innerShdw', self.ns)
        if inner_shadow is not None:
            return self._parse_inner_shadow(inner_shadow)

        return None

    @staticmethod
    def _parse_time_value(value: Optional[str], default: int) -> int:
        """PowerPoint 시간 값을 정수 밀리초로 변환"""
        if value is None:
            return default
        value = value.strip().lower()
        if value in {'indefinite', 'media', 'auto'}:
            return 0
        try:
            return int(value)
        except (TypeError, ValueError):
            return default

    def _parse_outer_shadow(self, shadow_elem: ET.Element) -> str:
        """
        외부 그림자 파싱

        Args:
            shadow_elem: outerShdw 엘리먼트

        Returns:
            CSS box-shadow 문자열
        """
        # 블러 반경
        blur_rad = int(shadow_elem.get('blurRad', '50000'))  # EMU
        blur_px = (blur_rad / 914400) * 96 / 1000  # px로 변환

        # 거리
        dist = int(shadow_elem.get('dist', '38100'))  # EMU
        dist_px = (dist / 914400) * 96 / 1000

        # 방향 (각도)
        dir_angle = int(shadow_elem.get('dir', '2700000')) / 60000  # 도 단위로 변환
        import math
        offset_x = dist_px * math.cos(math.radians(dir_angle))
        offset_y = dist_px * math.sin(math.radians(dir_angle))

        # 색상
        color = '#000000'
        solid_fill = shadow_elem.find('.//a:srgbClr', self.ns)
        if solid_fill is not None:
            color = '#' + solid_fill.get('val', '000000')

        # 투명도
        alpha = 1.0
        alpha_elem = solid_fill.find('.//a:alpha', self.ns) if solid_fill is not None else None
        if alpha_elem is not None:
            alpha_val = int(alpha_elem.get('val', '100000'))
            alpha = alpha_val / 100000

        return f"{offset_x:.1f}px {offset_y:.1f}px {blur_px:.1f}px rgba{self._hex_to_rgba(color, alpha)}"

    def _parse_inner_shadow(self, shadow_elem: ET.Element) -> str:
        """내부 그림자 파싱 (외부 그림자와 유사, 'inset' 추가)"""
        base_shadow = self._parse_outer_shadow(shadow_elem)
        return f"inset {base_shadow}"

    @staticmethod
    def _hex_to_rgba(hex_color: str, alpha: float) -> str:
        """
        Hex 색상을 RGBA로 변환

        Args:
            hex_color: #RRGGBB 형식
            alpha: 알파 값 (0.0 ~ 1.0)

        Returns:
            (R, G, B, A) 형식 문자열
        """
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return f"({r}, {g}, {b}, {alpha:.2f})"

    def apply_reflection_effects(self, sp_pr: ET.Element) -> Optional[str]:
        """
        반사 효과를 CSS로 변환

        Args:
            sp_pr: Shape Properties 엘리먼트

        Returns:
            CSS 문자열 또는 None
        """
        if sp_pr is None:
            return None

        # effectLst에서 반사 효과 찾기
        effect_lst = sp_pr.find('.//a:effectLst', self.ns)
        if effect_lst is None:
            return None

        reflection = effect_lst.find('.//a:reflection', self.ns)
        if reflection is None:
            return None

        # 반사 효과는 CSS만으로 완벽히 구현하기 어려우므로
        # 간단한 box-reflect 사용 (WebKit 전용)
        return "-webkit-box-reflect: below 2px linear-gradient(transparent, rgba(255,255,255,0.3))"
