# git-pptx

[English](README.md)

PowerPoint(pptx) 파일을 슬라이드 단위로 버전 관리하기 위한 독립 도구입니다. git/github와 결합하지 않으며, `decomp`/`comp`로 pptx ↔ git-pptx 디렉터리 형식을 상호 변환합니다. 결과는 일반 git으로 커밋·푸시합니다.

## 설치

```bash
git clone https://github.com/zer0ken/git-pptx.git
cd git-pptx
npm link
```

`npm link`는 `git-pptx` 실행 파일을 PATH에 노출합니다.

## 사용법

```bash
git-pptx decomp a.pptx              # a.pptx -> a.git-pptx/ (변경 슬라이드 갱신, 미리보기 렌더)
git-pptx comp a.git-pptx out.pptx   # a.git-pptx/ -> pptx
git-pptx diff a.pptx a.git-pptx     # 변경 슬라이드 표시 (쓰기 없이)
```

`decomp`는 `a.pptx`를 `a.git-pptx/`로 압축 해제하며(구성 아래), 번호 붙은 파트를 PowerPoint가 쓰는 1..N으로 정렬하고, 실제 내용이 바뀐 슬라이드만 감지합니다. `comp`는 `pptx/`를 다시 zip합니다.

```
a.git-pptx/
  .gitattributes            XML 파트의 diff 규칙
  previews/   1.jpg, ...    슬라이드별 미리보기(파생 산출물)
              index.json    각 미리보기를 렌더한 내용의 지문
  pptx/       pptx를 압축 해제한 원본
```

옵션:
- `--no-preview`: 미리보기 렌더 생략
- `--format png`: 미리보기를 JPG(기본) 대신 PNG로
- `--renderer auto|powerpoint|libreoffice`: 미리보기 렌더러 (기본 `auto`)
- `--no-normalize`: 파트 이름을 정규화하지 않고 덱이 갖고 온 이름을 그대로 둠

## 출력과 렌더링

- 결과는 stdout(굵게), 진행은 stderr(흐리게)로 출력되며, 터미널이 아니면 일반 텍스트로 퇴화합니다. `decomp`/`diff`는 변경 슬라이드를 요약 표시합니다.
- 렌더링은 열려 있는 편집기를 건드리지 않습니다: 미리보기는 임시 복사본에서 렌더하고, 실행 중 PowerPoint는 종료하지 않으며, LibreOffice는 격리된 프로필에서 돕니다.
- 미리보기는 자신이 렌더된 내용과 슬라이드가 더 이상 일치하지 않을 때 다시 렌더됩니다. 그 기준은 `previews/index.json`에 기록됩니다. 파트를 편집하고 `comp`한 뒤 `decomp`하면 그 시점에 pptx와 디렉터리가 서로 같아도 바뀐 슬라이드만 정확히 갱신됩니다.
- 렌더러: `powerpoint`(COM, Windows) 또는 `libreoffice`(headless + poppler, 전 OS). `auto`는 Windows에서 PowerPoint가 있으면 그걸 씁니다. 렌더러에 따라 미리보기가 다르므로 `--renderer`로 통일하세요.

## 읽을 수 있는 diff

PowerPoint는 XML 파트를 한 줄로 저장하므로, 어떤 편집이든 git에는 약 150KB짜리 한 줄이 다른 한 줄로 바뀐 것으로 보입니다. `decomp`는 그 파트들을 diff 드라이버로 보내고 줄바꿈 변환에서 제외하는 `a.git-pptx/.gitattributes`를 씁니다. 드라이버 자체는 커밋할 수 없는 git config 설정이므로, clone마다 한 번 켜야 합니다.

```bash
git config diff.pptxml.textconv "git-pptx textconv"
```

저장된 바이트는 그대로 유지됩니다. 줄바꿈은 git이 보여주는 화면에만 존재하고, 인접한 태그 사이에만 들어가므로 `<a:t>` 안의 텍스트는 파트가 가진 공백을 그대로 지킵니다.

## GitHub에서 논의하기

`previews/`를 커밋하면 GitHub가 이미지 미리보기와 PR 이미지 diff를 지원합니다. `*.pptx`는 `.gitignore`에 넣고 `a.git-pptx/`만 푸시합니다.

## 제약

- `docProps/core.xml`, `docProps/app.xml`, `docProps/thumbnail.jpeg`는 저장할 때마다 재생성되어 변경 감지에서 제외됩니다(comp가 유효하도록 폴더에는 그대로 씀).
- 정규화는 재직렬화 노이즈는 무시하지만 관계 ID 재번호(r:id)는 다루지 않아, 다른 도구가 쓴 덱의 첫 PowerPoint 저장 때 몇 개 파트가 변경으로 보일 수 있습니다.
- SVG+래스터 덱은 처음 저장 때 미디어 파트 몇 개가 한 번 이름이 바뀔 수 있습니다.
