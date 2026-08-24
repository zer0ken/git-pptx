# git-pptx

[English](README.md)

git은 pptx를 하나의 바이너리 파일로 봅니다. 슬라이드 한 장을 고쳐도 전체를 새로 만들어도 diff는 `Binary files differ` 한 줄이라, 무엇이 바뀌었는지 확인할 수 없습니다.

`decomp`는 pptx를 디렉터리로 풀어 이 문제를 해결합니다. pptx는 슬라이드마다 XML 파일이 들어 있는 zip이므로, 풀어 두면 git이 슬라이드 단위로 텍스트 diff를 냅니다. `decomp`는 변경된 슬라이드를 렌더한 이미지도 함께 저장합니다. GitHub PR에서 이 이미지로 슬라이드를 비교할 수 있습니다.

pptx 파일을 편집한 뒤에, 커밋 전에 `decomp`를 호출하여 변경사항을 추적 가능하게 관리해보세요.

## 설치

```bash
git clone https://github.com/zer0ken/git-pptx.git
cd git-pptx
npm link
```

`npm link`는 `git-pptx` 실행 파일을 PATH에 노출합니다.

## 사용법

`git-pptx decomp`는 pptx를 압축 해제하여 이름 뒤에 `.git-pptx` 접미어가 붙은 디렉터리를 생성합니다. `a.pptx`를 넘기면 `a.git-pptx/`가 만들어집니다.

```bash
git-pptx decomp a.pptx
git add a.pptx a.git-pptx && git commit
```

pptx와 git-pptx 디렉터리를 함께 커밋합니다. clone한 사람은 pptx를 열어 편집하고, git-pptx 디렉터리를 읽어 그 커밋이 무엇을 바꿨는지 확인합니다.

`decomp`는 변경된 파일만 갱신합니다. 슬라이드 세 장을 수정하면 git-pptx 디렉터리에도 세 장만 반영하고, 미리보기도 그 세 장만 다시 렌더합니다.

`comp`는 반대 방향의 변환입니다. `pptx/` 아래 XML을 직접 수정했을 때 그 결과를 다시 pptx로 묶습니다.

```bash
git-pptx comp a.git-pptx a.pptx
```

## 유의미한 diff 제공

PowerPoint는 XML 파트 하나를 통째로 한 줄에 씁니다. 압축을 풀어도 diff는 여전히 150KB짜리 한 줄입니다. clone마다 한 번씩 diff 드라이버를 등록하면 git이 이를 태그 단위로 끊어서 보여줍니다.

```bash
git config diff.pptxml.textconv "git-pptx textconv"
```

`decomp`가 생성하는 `.gitattributes`가 XML 파트를 이 드라이버로 넘깁니다. 드라이버 등록은 git config 설정이라 커밋되지 않으므로 clone마다 반복해야 합니다.

드라이버는 git이 출력하는 화면에만 줄바꿈을 넣습니다. 파일에 저장된 바이트는 그대로이고, 줄바꿈도 태그와 태그 사이에만 들어가므로 `<a:t>` 내부 텍스트의 공백은 보존됩니다.

## 명령

```bash
git-pptx decomp a.pptx              # a.pptx -> a.git-pptx/
git-pptx diff a.pptx a.git-pptx     # 변경 내용만 출력 (쓰기 없음)
git-pptx comp a.git-pptx out.pptx   # a.git-pptx/ -> pptx
```

git-pptx 디렉터리의 구조입니다.

```
a.git-pptx/
  .gitattributes            XML 파트의 diff 규칙
  previews/   1.jpg, ...    슬라이드별 미리보기
              index.json    각 미리보기를 렌더한 시점의 슬라이드 해시
  pptx/       pptx를 압축 해제한 원본
```

`decomp` 옵션:

- `--no-preview`: 미리보기 렌더 생략
- `--format png`: 미리보기를 JPG 대신 PNG로
- `--renderer auto|powerpoint|libreoffice`: 렌더러 선택 (기본 `auto`)
- `--no-normalize`: 파트 이름을 그대로 유지

`decomp`는 기본적으로 파트 이름을 정규화합니다. 스크립트가 생성했거나 큰 덱에서 분리한 덱은 파트 번호가 비어 있는데, PowerPoint는 처음 저장할 때 이를 빈틈없는 1..N으로 다시 매깁니다. 미리 정규화하지 않으면 그 저장 한 번으로 모든 파트의 이름이 바뀌어 git-pptx 디렉터리 전체가 변경으로 잡힙니다.

## 렌더링

`decomp`는 `powerpoint`(COM, Windows) 또는 `libreoffice`(headless + poppler, 전 OS)로 미리보기를 렌더합니다. `auto`는 Windows에 PowerPoint가 설치되어 있으면 이를 사용합니다. 두 렌더러의 출력이 동일하지 않으므로, 여러 사람이 커밋하는 저장소에서는 `--renderer`로 하나를 고정합니다.

덱을 편집 중인 상태에서 실행해도 안전합니다. `decomp`는 임시 복사본에서 렌더하고, 실행 중인 PowerPoint를 종료하지 않으며, LibreOffice를 별도 프로필에서 실행합니다.

`previews/index.json`에는 각 미리보기를 렌더할 때 사용한 슬라이드 XML의 해시가 들어 있습니다. `decomp`는 이 해시와 현재 슬라이드를 대조해 달라진 슬라이드만 다시 렌더합니다. pptx와 git-pptx 디렉터리를 비교하는 방식이 아니므로, `pptx/`를 직접 수정하고 `comp`한 슬라이드도 미리보기가 갱신됩니다.

`decomp`는 결과를 stdout으로, 진행 상황을 stderr로 출력합니다. 터미널이 아니면 색 없이 출력합니다.

## 제약

- `docProps/core.xml`, `docProps/app.xml`, `docProps/thumbnail.jpeg`는 PowerPoint가 저장할 때마다 재생성하는 값입니다. `decomp`는 이 셋을 변경 감지에서 제외해 diff 노이즈를 막고, `comp`가 유효한 pptx를 만들 수 있도록 git-pptx 디렉터리에는 그대로 저장합니다.
- 관계 ID(`r:id`)는 PowerPoint가 다른 도구로 만든 덱을 처음 저장할 때 자기 순서대로 다시 매기는 값입니다. 내용이 같은 슬라이드가 변경으로 잡히고 미리보기까지 다시 렌더되지만, `decomp`의 정규화가 여기까지 처리하지 못해 막을 방법이 없습니다. 두 번째 저장부터는 나타나지 않습니다.
- SVG와 래스터 이미지를 함께 쓴 덱은 PowerPoint가 처음 저장할 때 미디어 파트 몇 개의 이름을 바꿉니다. 한 번뿐이지만 그 커밋에서는 diff 노이즈로 보이고, 이것도 막을 방법이 없습니다.
