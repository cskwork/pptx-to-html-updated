#!/usr/bin/env python3
"""
PowerPoint Transition Extraction and Web Runtime Mapping
생성일: 2025-10-22
설명: 슬라이드 전환 효과를 파싱해 CSS/JavaScript 런타임으로 전달
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Any
from xml.etree import ElementTree as ET


@dataclass
class TransitionConfig:
    """슬라이드 전환 정보를 표현하는 데이터 클래스"""

    effect: str
    duration: int
    delay: int
    direction: Optional[str]
    orientation: Optional[str]
    thru_black: bool
    spokes: Optional[int]
    shape: Optional[str]
    speed: Optional[str]
    advance_on_click: bool
    advance_after: Optional[int]

    def to_runtime_dict(self) -> Dict[str, Any]:
        """런타임에서 사용할 수 있도록 dict 형태로 변환"""
        return {
            "effect": self.effect,
            "duration": self.duration,
            "delay": self.delay,
            "direction": self.direction,
            "orientation": self.orientation,
            "throughBlack": self.thru_black,
            "spokes": self.spokes,
            "shape": self.shape,
            "speed": self.speed,
            "advanceOnClick": self.advance_on_click,
            "advanceAfter": self.advance_after,
        }


class TransitionHandler:
    """PowerPoint 슬라이드 전환 효과 처리기"""

    MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"

    SPEED_TO_DURATION = {
        "slow": 1200,
        "med": 700,
        "fast": 350,
    }

    EFFECT_MAPPING = {
        "fade": "fade",
        "fadesmoothly": "fade",
        "dissolve": "fade",
        "cut": "cut",
        "random": "fade",
        "randombars": "bars",
        "randombar": "bars",
        "strips": "strips",
        "wipe": "wipe",
        "split": "split",
        "push": "push",
        "cover": "cover",
        "uncover": "uncover",
        "zoom": "zoom",
        "wheel": "wheel",
        "shape": "shape",
        "flash": "flash",
        "checker": "checker",
        "checkerboard": "checker",
        "blinds": "blinds",
        "comb": "comb",
        "diamond": "diamond",
        "circle": "circle",
        "plus": "plus",
        "box": "box",
        "gallery": "gallery",
        "flip": "flip",
        "switch": "switch",
        "conveyor": "conveyor",
        "pan": "pan",
        "ripple": "ripple",
        "origami": "origami",
        "window": "window",
        "door": "doors",
        "doors": "doors",
        "newsflash": "newsflash",
        "rotate": "rotate",
        "cube": "cube",
        "glitter": "bars",
        "shred": "strips",
        "wipenext": "wipe",
        "boxin": "box",
        "boxout": "box",
    }

    SUPPORTED_EFFECTS = {
        "fade",
        "cut",
        "push",
        "cover",
        "uncover",
        "wipe",
        "split",
        "bars",
        "strips",
        "zoom",
        "wheel",
        "shape",
        "flash",
        "checker",
        "blinds",
        "comb",
        "diamond",
        "circle",
        "plus",
        "box",
        "gallery",
        "flip",
        "switch",
        "conveyor",
        "pan",
        "ripple",
        "origami",
        "window",
        "doors",
        "newsflash",
        "rotate",
        "cube",
    }

    def __init__(self, ns: Dict[str, str], logger=None):
        """
        Args:
            ns: XML 네임스페이스 딕셔너리
            logger: ConversionLogger 인스턴스
        """
        self.ns = dict(ns)
        self.ns.setdefault("mc", self.MC_NS)
        self.ns.setdefault("p14", self.P14_NS)
        self.logger = logger

    # === Public API ===

    def extract_transition(self, slide_xml: ET.Element) -> TransitionConfig:
        """슬라이드 XML에서 전환 정보를 추출"""
        transition_elem = self._find_transition_element(slide_xml)
        if transition_elem is None:
            return self._default_config()

        config = self._default_config()
        config.duration = self._resolve_duration(transition_elem)
        config.delay = self._parse_int(transition_elem.get("delay"), default=0)
        config.speed = transition_elem.get("spd")
        config.advance_on_click = transition_elem.get("advClick", "1") != "0"
        config.advance_after = self._parse_int(transition_elem.get("advTm"))
        config.thru_black = self._parse_bool(transition_elem.get("thruBlk"), False)

        direction = self._normalize_direction(transition_elem.get("dir"))
        orientation = transition_elem.get("orient")
        spokes = None
        shape_type = None

        effect_token = transition_elem.get("type")

        type_elem = self._extract_type_element(transition_elem)
        if type_elem is not None:
            direction = self._normalize_direction(type_elem.get("dir")) or direction
            orientation = type_elem.get("orient") or type_elem.get("rev") or orientation
            spokes = self._parse_int(type_elem.get("spokes")) or spokes
            shape_type = type_elem.get("type") or type_elem.get("shape") or shape_type
            config.thru_black = config.thru_black or self._parse_bool(type_elem.get("thruBlk"), False)

            child_tag = self._strip_namespace(type_elem.tag)
            if child_tag == "prstTrans":
                prst_value = (type_elem.get("prst") or "").strip()
                mapped = self._map_named_effect(prst_value)
                if mapped:
                    effect_token = mapped["effect"]
                    direction = direction or mapped.get("direction")
                    orientation = orientation or mapped.get("orientation")
                    shape_type = shape_type or mapped.get("shape")
                    spokes = spokes or mapped.get("spokes")
            else:
                effect_token = child_tag or effect_token

        mapped = self._map_named_effect(effect_token)
        if mapped:
            effect_name = mapped["effect"]
            direction = direction or mapped.get("direction")
            orientation = orientation or mapped.get("orientation")
            shape_type = shape_type or mapped.get("shape")
            spokes = spokes or mapped.get("spokes")
        else:
            key = (effect_token or "").strip().lower()
            effect_name = self.EFFECT_MAPPING.get(key, key or "cut")

        if effect_name not in self.SUPPORTED_EFFECTS:
            if self.logger:
                self.logger.warning(f"지원하지 않는 전환 효과 '{effect_name}'를 컷으로 대체합니다")
            effect_name = "cut"

        config.effect = effect_name
        config.direction = self._normalize_direction(direction)
        config.orientation = orientation
        config.spokes = spokes if spokes not in {0, None} else None
        config.shape = shape_type

        return config

    def serialize_transitions(self, slides: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """슬라이드별 전환 정보를 런타임용 리스트로 직렬화"""
        serialized: List[Dict[str, Any]] = []
        for slide in slides:
            transition: Optional[TransitionConfig] = slide.get("transition")
            if transition is None:
                transition = self._default_config()
            serialized.append(transition.to_runtime_dict())
        return serialized

    def generate_transition_css(self) -> str:
        """전환 효과를 위한 CSS 정의 생성"""
        # CSS는 커스텀 프로퍼티 기반으로 작성해 방향성 효과를 재사용
        css_sections: List[str] = [
            self._base_css(),
            self._fade_css(),
            self._push_css(),
            self._cover_css(),
            self._uncover_css(),
            self._wipe_css(),
            self._split_css(),
            self._bars_css(),
            self._strips_css(),
            self._zoom_css(),
            self._wheel_css(),
            self._shape_css(),
            self._flash_css(),
            self._checker_css(),
            self._blinds_css(),
            self._comb_css(),
            self._box_css(),
            self._gallery_css(),
            self._flip_css(),
            self._switch_css(),
            self._conveyor_css(),
            self._pan_css(),
            self._ripple_css(),
            self._origami_css(),
            self._window_css(),
            self._doors_css(),
            self._newsflash_css(),
            self._rotate_css(),
            self._cube_css(),
        ]
        return "\n\n".join(filter(None, css_sections))

    # === CSS Builders ===

    def _base_css(self) -> str:
        """기본 전환 클래스 및 커스텀 프로퍼티 정의"""
        return r"""
.slide {
    --ppt-transition-duration: 700ms;
    --ppt-transition-delay: 0ms;
    --ppt-transition-easing: ease-in-out;
    --ppt-translate-in-x: 0%;
    --ppt-translate-in-y: 0%;
    --ppt-translate-out-x: 0%;
    --ppt-translate-out-y: 0%;
    --ppt-scale-in: 1;
    --ppt-scale-out: 1;
    --ppt-rotate-in: 0deg;
    --ppt-rotate-out: 0deg;
    --ppt-mask-start: 0%;
    --ppt-mask-angle: 0deg;
    --ppt-mask-shape: 100%;
    --ppt-clip-inset-top: 0%;
    --ppt-clip-inset-right: 0%;
    --ppt-clip-inset-bottom: 0%;
    --ppt-clip-inset-left: 0%;
    opacity: 0;
    visibility: hidden;
    pointer-events: none;
}

.slide.pre-active,
.slide.active,
.slide.ppt-transition-enter,
.slide.ppt-transition-exit {
    display: block;
    visibility: visible;
    pointer-events: auto;
}

.slide.active {
    opacity: 1;
    z-index: 2;
}

.slide.pre-active {
    opacity: 0;
    z-index: 1;
}

.ppt-transition-enter,
.ppt-transition-exit {
    animation-duration: var(--ppt-transition-duration);
    animation-timing-function: var(--ppt-transition-easing);
    animation-fill-mode: forwards;
}

.ppt-transition-enter {
    animation-delay: var(--ppt-transition-delay);
    z-index: 3;
}

.ppt-transition-exit {
    animation-delay: 0ms;
    z-index: 2;
}

.transition-blackout {
    position: absolute;
    inset: 0;
    background: #000;
    opacity: 0;
    pointer-events: none;
    z-index: 5;
}

.transition-blackout.active {
    transition: opacity var(--ppt-transition-duration) ease-in-out;
    opacity: 1;
}
""".strip()

    def _fade_css(self) -> str:
        return r"""
.ppt-transition-fade-in {
    animation-name: ppt-fade-in;
}

.ppt-transition-fade-out {
    animation-name: ppt-fade-out;
}

@keyframes ppt-fade-in {
    from { opacity: 0; }
    to { opacity: 1; }
}

@keyframes ppt-fade-out {
    from { opacity: 1; }
    to { opacity: 0; }
}
""".strip()

    def _push_css(self) -> str:
        return r"""
.ppt-transition-push-in {
    animation-name: ppt-push-in;
}

.ppt-transition-push-out {
    animation-name: ppt-push-out;
}

@keyframes ppt-push-in {
    from { transform: translate(var(--ppt-translate-in-x), var(--ppt-translate-in-y)); opacity: 1; }
    to { transform: translate(0%, 0%); opacity: 1; }
}

@keyframes ppt-push-out {
    from { transform: translate(0%, 0%); opacity: 1; }
    to { transform: translate(var(--ppt-translate-out-x), var(--ppt-translate-out-y)); opacity: 1; }
}
""".strip()

    def _cover_css(self) -> str:
        return r"""
.ppt-transition-cover-in {
    animation-name: ppt-cover-in;
}

@keyframes ppt-cover-in {
    from { transform: translate(var(--ppt-translate-in-x), var(--ppt-translate-in-y)); }
    to { transform: translate(0%, 0%); }
}
""".strip()

    def _uncover_css(self) -> str:
        return r"""
.ppt-transition-uncover-out {
    animation-name: ppt-uncover-out;
}

@keyframes ppt-uncover-out {
    from { transform: translate(0%, 0%); }
    to { transform: translate(var(--ppt-translate-out-x), var(--ppt-translate-out-y)); }
}
""".strip()

    def _wipe_css(self) -> str:
        return r"""
.ppt-transition-wipe-in {
    animation-name: ppt-wipe-in;
    overflow: hidden;
}

.ppt-transition-wipe-out {
    animation-name: ppt-wipe-out;
    overflow: hidden;
}

@keyframes ppt-wipe-in {
    from { clip-path: inset(var(--ppt-clip-inset-top) var(--ppt-clip-inset-right) var(--ppt-clip-inset-bottom) var(--ppt-clip-inset-left)); }
    to { clip-path: inset(0% 0% 0% 0%); }
}

@keyframes ppt-wipe-out {
    from { clip-path: inset(0% 0% 0% 0%); }
    to { clip-path: inset(var(--ppt-clip-inset-top) var(--ppt-clip-inset-right) var(--ppt-clip-inset-bottom) var(--ppt-clip-inset-left)); }
}
""".strip()

    def _split_css(self) -> str:
        return r"""
.ppt-transition-split-in {
    animation-name: ppt-split-in;
    overflow: hidden;
}

.ppt-transition-split-out {
    animation-name: ppt-split-out;
    overflow: hidden;
}

@keyframes ppt-split-in {
    from { clip-path: inset(var(--ppt-clip-inset-top) var(--ppt-clip-inset-right) var(--ppt-clip-inset-bottom) var(--ppt-clip-inset-left)); opacity: 0; }
    to { clip-path: inset(0% 0% 0% 0%); opacity: 1; }
}

@keyframes ppt-split-out {
    from { clip-path: inset(0% 0% 0% 0%); opacity: 1; }
    to { clip-path: inset(var(--ppt-clip-inset-top) var(--ppt-clip-inset-right) var(--ppt-clip-inset-bottom) var(--ppt-clip-inset-left)); opacity: 0; }
}
""".strip()

    def _bars_css(self) -> str:
        return r"""
.ppt-transition-bars-in {
    animation-name: ppt-bars-in;
    mask-image: repeating-linear-gradient(var(--ppt-mask-angle), #000 0 calc(6% - 1px), transparent calc(6% - 1px) 6%);
    -webkit-mask-image: repeating-linear-gradient(var(--ppt-mask-angle), #000 0 calc(6% - 1px), transparent calc(6% - 1px) 6%);
    mask-size: 100% 100%;
    -webkit-mask-size: 100% 100%;
}

.ppt-transition-bars-out {
    animation-name: ppt-bars-out;
    mask-image: repeating-linear-gradient(var(--ppt-mask-angle), #000 0 calc(6% - 1px), transparent calc(6% - 1px) 6%);
    -webkit-mask-image: repeating-linear-gradient(var(--ppt-mask-angle), #000 0 calc(6% - 1px), transparent calc(6% - 1px) 6%);
}

@keyframes ppt-bars-in {
    from { mask-size: var(--ppt-mask-start) 100%; -webkit-mask-size: var(--ppt-mask-start) 100%; opacity: 1; }
    to { mask-size: 100% 100%; -webkit-mask-size: 100% 100%; opacity: 1; }
}

@keyframes ppt-bars-out {
    from { mask-size: 100% 100%; -webkit-mask-size: 100% 100%; opacity: 1; }
    to { mask-size: var(--ppt-mask-start) 100%; -webkit-mask-size: var(--ppt-mask-start) 100%; opacity: 1; }
}
""".strip()

    def _strips_css(self) -> str:
        return r"""
.ppt-transition-strips-in {
    animation-name: ppt-strips-in;
    mask-image: repeating-linear-gradient(var(--ppt-mask-angle), transparent 0 calc(6% - 1px), #000 calc(6% - 1px) 6%);
    -webkit-mask-image: repeating-linear-gradient(var(--ppt-mask-angle), transparent 0 calc(6% - 1px), #000 calc(6% - 1px) 6%);
}

.ppt-transition-strips-out {
    animation-name: ppt-strips-out;
    mask-image: repeating-linear-gradient(var(--ppt-mask-angle), transparent 0 calc(6% - 1px), #000 calc(6% - 1px) 6%);
    -webkit-mask-image: repeating-linear-gradient(var(--ppt-mask-angle), transparent 0 calc(6% - 1px), #000 calc(6% - 1px) 6%);
}

@keyframes ppt-strips-in {
    from { mask-position: var(--ppt-mask-start); -webkit-mask-position: var(--ppt-mask-start); opacity: 1; }
    to { mask-position: 0%; -webkit-mask-position: 0%; opacity: 1; }
}

@keyframes ppt-strips-out {
    from { mask-position: 0%; -webkit-mask-position: 0%; opacity: 1; }
    to { mask-position: var(--ppt-mask-start); -webkit-mask-position: var(--ppt-mask-start); opacity: 1; }
}
""".strip()

    def _zoom_css(self) -> str:
        return r"""
.ppt-transition-zoom-in {
    animation-name: ppt-zoom-in;
}

.ppt-transition-zoom-out {
    animation-name: ppt-zoom-out;
}

@keyframes ppt-zoom-in {
    from { transform: scale(var(--ppt-scale-in)); opacity: 0; }
    to { transform: scale(1); opacity: 1; }
}

@keyframes ppt-zoom-out {
    from { transform: scale(1); opacity: 1; }
    to { transform: scale(var(--ppt-scale-out)); opacity: 0; }
}
""".strip()

    def _wheel_css(self) -> str:
        return r"""
.ppt-transition-wheel-in {
    animation-name: ppt-wheel-in;
    mask-image: conic-gradient(from var(--ppt-mask-angle), #000 0deg calc(360deg / var(--ppt-wheel-spokes)), transparent calc(360deg / var(--ppt-wheel-spokes)) 360deg);
    -webkit-mask-image: conic-gradient(from var(--ppt-mask-angle), #000 0deg calc(360deg / var(--ppt-wheel-spokes)), transparent calc(360deg / var(--ppt-wheel-spokes)) 360deg);
}

.ppt-transition-wheel-out {
    animation-name: ppt-wheel-out;
    mask-image: conic-gradient(from var(--ppt-mask-angle), #000 0deg calc(360deg / var(--ppt-wheel-spokes)), transparent calc(360deg / var(--ppt-wheel-spokes)) 360deg);
    -webkit-mask-image: conic-gradient(from var(--ppt-mask-angle), #000 0deg calc(360deg / var(--ppt-wheel-spokes)), transparent calc(360deg / var(--ppt-wheel-spokes)) 360deg);
}

@keyframes ppt-wheel-in {
    from { mask-rotation: 360deg; opacity: 1; }
    to { mask-rotation: 0deg; opacity: 1; }
}

@keyframes ppt-wheel-out {
    from { mask-rotation: 0deg; opacity: 1; }
    to { mask-rotation: -360deg; opacity: 1; }
}
""".strip()

    def _shape_css(self) -> str:
        return r"""
.ppt-transition-shape-in {
    animation-name: ppt-shape-in;
}

.ppt-transition-shape-out {
    animation-name: ppt-shape-out;
}

@keyframes ppt-shape-in {
    from { clip-path: var(--ppt-clip-shape); opacity: 0; }
    to { clip-path: inset(0% 0% 0% 0%); opacity: 1; }
}

@keyframes ppt-shape-out {
    from { clip-path: inset(0% 0% 0% 0%); opacity: 1; }
    to { clip-path: var(--ppt-clip-shape); opacity: 0; }
}
""".strip()

    def _flash_css(self) -> str:
        return r"""
.ppt-transition-flash-in,
.ppt-transition-flash-out {
    animation-name: ppt-flash;
}

@keyframes ppt-flash {
    0% { opacity: 0; }
    20% { opacity: 1; }
    40% { opacity: 0; }
    60% { opacity: 1; }
    100% { opacity: 1; }
}
""".strip()

    def _checker_css(self) -> str:
        return r"""
.ppt-transition-checker-in,
.ppt-transition-checker-out {
    animation-name: ppt-checker;
    mask-image: repeating-conic-gradient(#000 0deg 45deg, transparent 45deg 90deg);
    -webkit-mask-image: repeating-conic-gradient(#000 0deg 45deg, transparent 45deg 90deg);
}

@keyframes ppt-checker {
    from { mask-size: var(--ppt-mask-start) var(--ppt-mask-start); -webkit-mask-size: var(--ppt-mask-start) var(--ppt-mask-start); opacity: 0; }
    to { mask-size: 100% 100%; -webkit-mask-size: 100% 100%; opacity: 1; }
}
""".strip()

    def _blinds_css(self) -> str:
        return r"""
.ppt-transition-blinds-in,
.ppt-transition-blinds-out {
    animation-name: ppt-blinds;
    mask-image: repeating-linear-gradient(var(--ppt-mask-angle), #000 0 calc(12% - 1px), transparent calc(12% - 1px) 12%);
    -webkit-mask-image: repeating-linear-gradient(var(--ppt-mask-angle), #000 0 calc(12% - 1px), transparent calc(12% - 1px) 12%);
}

@keyframes ppt-blinds {
    from { mask-position: var(--ppt-mask-start); -webkit-mask-position: var(--ppt-mask-start); opacity: 1; }
    to { mask-position: 0%; -webkit-mask-position: 0%; opacity: 1; }
}
""".strip()

    def _comb_css(self) -> str:
        return r"""
.ppt-transition-comb-in,
.ppt-transition-comb-out {
    animation-name: ppt-comb;
    mask-image: repeating-linear-gradient(var(--ppt-mask-angle), #000 0 calc(10% - 1px), transparent calc(10% - 1px) 10%);
    -webkit-mask-image: repeating-linear-gradient(var(--ppt-mask-angle), #000 0 calc(10% - 1px), transparent calc(10% - 1px) 10%);
}

@keyframes ppt-comb {
    from { mask-position: var(--ppt-mask-start); -webkit-mask-position: var(--ppt-mask-start); opacity: 0; }
    to { mask-position: 0%; -webkit-mask-position: 0%; opacity: 1; }
}
""".strip()

    def _box_css(self) -> str:
        return r"""
.ppt-transition-box-in,
.ppt-transition-box-out {
    animation-name: ppt-box;
}

@keyframes ppt-box {
    from { clip-path: inset(50% 25% 50% 25%); opacity: 0; }
    to { clip-path: inset(0% 0% 0% 0%); opacity: 1; }
}
""".strip()

    def _gallery_css(self) -> str:
        return r"""
.ppt-transition-gallery-in,
.ppt-transition-gallery-out {
    animation-name: ppt-gallery;
    transform-origin: 50% 50%;
}

@keyframes ppt-gallery {
    0% { transform: perspective(1200px) rotateY(var(--ppt-rotate-in)); opacity: 0; }
    100% { transform: perspective(1200px) rotateY(0deg); opacity: 1; }
}
""".strip()

    def _flip_css(self) -> str:
        return r"""
.ppt-transition-flip-in,
.ppt-transition-flip-out {
    animation-name: ppt-flip;
    transform-origin: 50% 50%;
}

@keyframes ppt-flip {
    from { transform: perspective(1200px) rotateY(var(--ppt-rotate-in)); opacity: 0; }
    to { transform: perspective(1200px) rotateY(0deg); opacity: 1; }
}
""".strip()

    def _switch_css(self) -> str:
        return r"""
.ppt-transition-switch-in,
.ppt-transition-switch-out {
    animation-name: ppt-switch;
    transform-origin: 50% 50%;
}

@keyframes ppt-switch {
    from { transform: perspective(1200px) rotateX(var(--ppt-rotate-in)); opacity: 0; }
    to { transform: perspective(1200px) rotateX(0deg); opacity: 1; }
}
""".strip()

    def _conveyor_css(self) -> str:
        return r"""
.ppt-transition-conveyor-in,
.ppt-transition-conveyor-out {
    animation-name: ppt-conveyor;
}

@keyframes ppt-conveyor {
    from { transform: translate(var(--ppt-translate-in-x), var(--ppt-translate-in-y)) rotate(var(--ppt-rotate-in)); opacity: 0; }
    to { transform: translate(0%, 0%) rotate(0deg); opacity: 1; }
}
""".strip()

    def _pan_css(self) -> str:
        return r"""
.ppt-transition-pan-in,
.ppt-transition-pan-out {
    animation-name: ppt-pan;
}

@keyframes ppt-pan {
    from { transform: translate(var(--ppt-translate-in-x), var(--ppt-translate-in-y)) scale(var(--ppt-scale-in)); opacity: 0; }
    to { transform: translate(0%, 0%) scale(1); opacity: 1; }
}
""".strip()

    def _ripple_css(self) -> str:
        return r"""
.ppt-transition-ripple-in,
.ppt-transition-ripple-out {
    animation-name: ppt-ripple;
}

@keyframes ppt-ripple {
    from { clip-path: circle(0% at 50% 50%); opacity: 0; }
    to { clip-path: circle(150% at 50% 50%); opacity: 1; }
}
""".strip()

    def _origami_css(self) -> str:
        return r"""
.ppt-transition-origami-in,
.ppt-transition-origami-out {
    animation-name: ppt-origami;
    transform-origin: 50% 50%;
}

@keyframes ppt-origami {
    from { transform: perspective(1200px) rotateX(45deg) rotateY(45deg); opacity: 0; }
    to { transform: perspective(1200px) rotateX(0deg) rotateY(0deg); opacity: 1; }
}
""".strip()

    def _window_css(self) -> str:
        return r"""
.ppt-transition-window-in,
.ppt-transition-window-out {
    animation-name: ppt-window;
}

@keyframes ppt-window {
    from { clip-path: inset(50% 25% 50% 25%); opacity: 0; }
    to { clip-path: inset(0% 0% 0% 0%); opacity: 1; }
}
""".strip()

    def _doors_css(self) -> str:
        return r"""
.ppt-transition-doors-in,
.ppt-transition-doors-out {
    animation-name: ppt-doors;
    transform-origin: 50% 50%;
}

@keyframes ppt-doors {
    from { transform: perspective(1200px) rotateY(90deg); opacity: 0; }
    to { transform: perspective(1200px) rotateY(0deg); opacity: 1; }
}
""".strip()

    def _newsflash_css(self) -> str:
        return r"""
.ppt-transition-newsflash-in,
.ppt-transition-newsflash-out {
    animation-name: ppt-newsflash;
}

@keyframes ppt-newsflash {
    from { transform: scale(2); opacity: 0; filter: blur(6px); }
    to { transform: scale(1); opacity: 1; filter: blur(0); }
}
""".strip()

    def _rotate_css(self) -> str:
        return r"""
.ppt-transition-rotate-in,
.ppt-transition-rotate-out {
    animation-name: ppt-rotate;
    transform-origin: 50% 50%;
}

@keyframes ppt-rotate {
    from { transform: rotate(var(--ppt-rotate-in)); opacity: 0; }
    to { transform: rotate(0deg); opacity: 1; }
}
""".strip()

    def _cube_css(self) -> str:
        return r"""
.ppt-transition-cube-in,
.ppt-transition-cube-out {
    animation-name: ppt-cube;
    transform-origin: 50% 50%;
}

@keyframes ppt-cube {
    from { transform: perspective(1200px) rotateY(var(--ppt-rotate-in)); opacity: 0; }
    to { transform: perspective(1200px) rotateY(0deg); opacity: 1; }
}
""".strip()

    # === Utilities ===

    def _find_transition_element(self, slide_xml: ET.Element) -> Optional[ET.Element]:
        """실제 전환 요소 검색 (AlternateContent 우선)"""
        choice_elem = slide_xml.find(".//mc:Choice/p:transition", self.ns)
        if choice_elem is not None:
            return choice_elem

        return slide_xml.find(".//p:transition", self.ns)

    def _extract_type_element(self, transition_elem: ET.Element) -> Optional[ET.Element]:
        """전환 타입을 나타내는 하위 엘리먼트 추출"""
        for child in list(transition_elem):
            tag = self._strip_namespace(child.tag)
            if tag == "extLst":
                continue
            return child
        return None

    def _resolve_duration(self, transition_elem: ET.Element) -> int:
        """전환 지속시간 계산"""
        p14_duration = transition_elem.get(f"{{{self.P14_NS}}}dur")
        if p14_duration:
            return self._parse_int(p14_duration, default=700)

        raw_duration = transition_elem.get("dur")
        if raw_duration is not None:
            parsed = self._parse_int(raw_duration)
            if parsed is not None:
                return parsed

        speed = transition_elem.get("spd", "med")
        return self.SPEED_TO_DURATION.get(speed, 700)

    @staticmethod
    def _parse_int(value: Optional[str], default: Optional[int] = None) -> Optional[int]:
        """정수 파싱 유틸리티"""
        if value is None:
            return default
        try:
            return int(value)
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _strip_namespace(tag: str) -> str:
        """네임스페이스가 포함된 태그에서 로컬 이름 추출"""
        if "}" in tag:
            return tag.split("}", 1)[1]
        return tag

    @staticmethod
    def _parse_bool(value: Optional[str], fallback: bool = False) -> bool:
        """불리언 속성 파싱"""
        if value is None:
            return fallback
        lowered = value.strip().lower()
        if lowered in {"1", "true", "t", "yes", "y"}:
            return True
        if lowered in {"0", "false", "f", "no", "n"}:
            return False
        return fallback

    def _default_config(self) -> TransitionConfig:
        """전환 미정의 시 기본 컷 전환 반환"""
        return TransitionConfig(
            effect="cut",
            duration=0,
            delay=0,
            direction=None,
            orientation=None,
            thru_black=False,
            spokes=None,
            shape=None,
            speed=None,
            advance_on_click=True,
            advance_after=None,
        )

    def _normalize_direction(self, direction: Optional[str]) -> Optional[str]:
        """방향 문자열을 공통 코드로 정규화"""
        if not direction:
            return None
        key = direction.strip().lower()
        mapping = {
            "left": "l",
            "fromleft": "l",
            "l": "l",
            "right": "r",
            "fromright": "r",
            "r": "r",
            "up": "u",
            "top": "u",
            "fromtop": "u",
            "u": "u",
            "down": "d",
            "bottom": "d",
            "frombottom": "d",
            "d": "d",
            "upleft": "tl",
            "leftup": "tl",
            "topleft": "tl",
            "tl": "tl",
            "upright": "tr",
            "rightup": "tr",
            "topright": "tr",
            "tr": "tr",
            "downleft": "bl",
            "leftdown": "bl",
            "bottomleft": "bl",
            "bl": "bl",
            "downright": "br",
            "rightdown": "br",
            "bottomright": "br",
            "br": "br",
            "in": "in",
            "out": "out",
        }
        return mapping.get(key, direction)

    def _map_named_effect(self, name: Optional[str]) -> Optional[Dict[str, Any]]:
        """프리셋 전환 이름을 공통 효과 정보로 변환"""
        if not name:
            return None
        key = name.strip().lower()
        info: Dict[str, Any] = {}

        if key in {"fade", "fadesmoothly", "dissolve"}:
            info["effect"] = "fade"
            return info

        if key.startswith("push"):
            info["effect"] = "push"
            info["direction"] = self._extract_direction_from_suffix(key)
            return info

        if key.startswith("wipe"):
            info["effect"] = "wipe"
            info["direction"] = self._extract_direction_from_suffix(key)
            return info

        if key.startswith("cover"):
            info["effect"] = "cover"
            info["direction"] = self._extract_direction_from_suffix(key)
            return info

        if key.startswith("uncover"):
            info["effect"] = "uncover"
            info["direction"] = self._extract_direction_from_suffix(key)
            return info

        if key.startswith("split"):
            info["effect"] = "split"
            if "horz" in key or "horizontal" in key:
                info["orientation"] = "horz"
            elif "vert" in key or "vertical" in key:
                info["orientation"] = "vert"
            if key.endswith("in"):
                info["direction"] = "in"
            elif key.endswith("out"):
                info["direction"] = "out"
            return info

        if key.startswith("randombars"):
            info["effect"] = "bars"
            if "horizontal" in key or "horz" in key:
                info["orientation"] = "horz"
            elif "vertical" in key or "vert" in key:
                info["orientation"] = "vert"
            return info

        if key.startswith("strips"):
            info["effect"] = "strips"
            info["direction"] = self._extract_direction_from_suffix(key)
            return info

        if key.startswith("wheel"):
            info["effect"] = "wheel"
            digits = "".join(ch for ch in key if ch.isdigit())
            info["spokes"] = int(digits) if digits else None
            return info

        if key.startswith("box"):
            info["effect"] = "box"
            return info

        if key.startswith("diamond"):
            info["effect"] = "diamond"
            return info

        if key.startswith("circle"):
            info["effect"] = "circle"
            return info

        if key.startswith("plus"):
            info["effect"] = "plus"
            return info

        if key.startswith("checker"):
            info["effect"] = "checker"
            return info

        if key.startswith("blinds"):
            info["effect"] = "blinds"
            if "horizontal" in key:
                info["orientation"] = "horz"
            elif "vertical" in key:
                info["orientation"] = "vert"
            return info

        if key.startswith("comb"):
            info["effect"] = "comb"
            info["direction"] = self._extract_direction_from_suffix(key)
            return info

        if key.startswith("gallery"):
            info["effect"] = "gallery"
            return info

        if key.startswith("flip"):
            info["effect"] = "flip"
            return info

        if key.startswith("switch"):
            info["effect"] = "switch"
            return info

        if key.startswith("conveyor"):
            info["effect"] = "conveyor"
            info["direction"] = self._extract_direction_from_suffix(key)
            return info

        if key.startswith("pan"):
            info["effect"] = "pan"
            info["direction"] = self._extract_direction_from_suffix(key)
            return info

        if key.startswith("ripple"):
            info["effect"] = "ripple"
            return info

        if key.startswith("origami"):
            info["effect"] = "origami"
            return info

        if key.startswith("window"):
            info["effect"] = "window"
            return info

        if key.startswith("door"):
            info["effect"] = "doors"
            return info

        if key.startswith("newsflash"):
            info["effect"] = "newsflash"
            return info

        if key.startswith("rotate"):
            info["effect"] = "rotate"
            return info

        if key.startswith("cube"):
            info["effect"] = "cube"
            return info

        if key.startswith("shape"):
            info["effect"] = "shape"
            info["shape"] = key.replace("shape", "").strip() or None
            return info

        mapped = self.EFFECT_MAPPING.get(key)
        if mapped:
            info["effect"] = mapped
            return info

        return None

    def _extract_direction_from_suffix(self, token: str) -> Optional[str]:
        """토큰 이름의 접미사에서 방향 추출"""
        if token.endswith("left"):
            return "l"
        if token.endswith("right"):
            return "r"
        if token.endswith("up") or token.endswith("top"):
            return "u"
        if token.endswith("down") or token.endswith("bottom"):
            return "d"
        if token.endswith("upleft") or token.endswith("leftup"):
            return "tl"
        if token.endswith("upright") or token.endswith("rightup"):
            return "tr"
        if token.endswith("downleft") or token.endswith("leftdown"):
            return "bl"
        if token.endswith("downright") or token.endswith("rightdown"):
            return "br"
        return None
