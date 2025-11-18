# High-Fidelity Conversion Update (2025-11-18)

## 구현 개요
- 픽셀 기반 절대 좌표 렌더링으로 슬라이드 레이아웃을 PPTX와 1:1 매칭.
- HTML/CSS/JS 외부 번들을 생성하여 내장 스타일과 스크립트를 분리하고 유지보수성 향상.
- 임베디드 폰트 추출(ODTTF → WOFF) 및 `@font-face` 자동 등록으로 원본 폰트 보존.
- Chart.js 초기화를 위한 데이터 속성 구조와 런타임 스케일링 로직 추가.
- SmartArt/SVG 도형/테이블을 z-index 포함한 절대 위치로 재구성.

## 파일별 주요 변경 사항
- `scripts/convert_pptx_to_html_v2.py`
  - `generate_bundle` 도입으로 HTML/CSS/JS 동시 생성.
  - `FontManager` 연동 및 픽셀 좌표 기반 요소 렌더링.
  - 문단 서식(줄 간격, 들여쓰기, 번호 매기기) 세밀화.
  - 그룹 도형/커넥터 재귀 처리와 z-index 유지, 변환 체인으로 중첩 배치 대응.
  - **[NEW] 텍스트 수직 정렬 개선**: `estimated_height` 기반의 부정확한 계산을 제거하고 `display: flex`를 사용하여 브라우저 네이티브 수직 정렬(`justify-content`) 적용. 줄바꿈 시에도 정확한 중앙/하단 정렬 보장.
- `scripts/font_manager.py`
  - PPTX 임베디드 폰트 추출 및 WOFF 변환, CSS 정의 반환.
- `scripts/shape_geometry.py`
  - SVG 출력이 절대 좌표/`z-index`/그림자 속성을 지원하도록 수정.
- `scripts/smartart_parser.py`
  - SmartArt 텍스트 컨테이너가 픽셀 기준 위치를 지원.
- 테마 처리
  - `convert_pptx_to_html_v2.py`에서 PPT 테마 팔레트/배경 로딩, 레이아웃·마스터 배경 캐스케이드, 테마 기반 색상 보정 적용.
- 이미지 채움
  - `convert_pptx_to_html_v2.py`에서 `blipFill` 추출, 관계 대상 자산 복사, clip-path 기반 크롭/스트레치 반영.
- 커스텀 도형
  - `shape_geometry.py`에서 path bounding box 계산 후 viewBox 정규화, SVG와 텍스트를 동시 렌더링하여 사각형/도형 누락 해결.
- 텍스트 앵커
  - 수직 정렬(anchor) 속성을 Flex wrapper로 반영하여 텍스트 박스 위치가 PPT와 동일하게 유지되도록 개선.
- 텍스트 인셋
  - `bodyPr` inset 값(lIns/tIns/rIns/bIns)을 CSS padding으로 변환해 도형/테이블 텍스트 여백을 PPT와 일치시킴.
- 그래픽 프레임 좌표
  - `extract_shape_position`가 `a:xfrm` 요소도 직접 처리하도록 보강하여 테이블·그래픽 프레임이 원래 위치에 렌더링되도록 수정.

## 후속 작업 제안
1. Chart.js 로컬 번들화 및 lazy-loading 전략 추가.
2. SVG 그라디언트/패턴 처리 고도화.
3. 애니메이션 타임라인(JSON) → JS 재생 로직 완성.
4. 시각적 회귀 테스트(이미지 diff) 자동화.
