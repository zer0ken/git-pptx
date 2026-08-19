# git-pptx

[English](README.md)

PowerPoint(pptx) 파일을 **슬라이드 단위로 버전 관리**하기 위한 독립 도구입니다. git/github와 결합하지 않으며, `decomp`/`comp`로 pptx ↔ git-pptx 디렉터리 형식을 상호 변환합니다. 변환 결과는 일반 git으로 커밋·푸시합니다.

## git-pptx 형식

`a.pptx`를 decomp하면 `a.git-pptx/` 디렉터리가 생성됩니다. 파일과 같은 폴더에 공존합니다.

```
a.git-pptx/
  previews/   1.jpg, 2.jpg, ...   슬라이드별 미리보기(파생 산출물)
  pptx/       pptx를 압축 해제한 원본 내용
```

- **decomp**는 관계 ID나 `[Content_Types].xml`을 재구성하지 않아 데이터 손실이 없습니다. 바꾸는 것은 파트 파일 이름에 붙은 번호뿐이며, PowerPoint 자신이 쓰는 1..N 형태로 맞춰 둡니다(아래 "파트 이름 정규화" 참고). `--no-normalize`를 주면 덱이 갖고 온 이름을 그대로 둡니다.
- **comp**는 `pptx/`를 다시 zip합니다 (미리보기·VCS 파일 제외).
- **변경 감지**는 XML 정규화 기반입니다. 속성 순서, 엔티티 인코딩, 빈 요소 표기, 네임스페이스 접두사처럼 PowerPoint 재저장 시 변하는 노이즈와 저장할 때마다 갱신되는 캐시를 무시하고, 진짜 내용이 바뀐 슬라이드만 감지합니다. 요소 순서(z-order)는 보존합니다.

## 파트 이름 정규화

OOXML 파트 이름에 붙은 번호(`slide7.xml`, `image240.png`, `theme4.xml`)는 문서의 내용이 아니라, 그 파일을 쓴 도구의 카운터가 어디까지 갔는지를 나타낼 뿐입니다. PowerPoint는 받은 번호를 그대로 두지 않습니다. 저장할 때마다 패키지를 루트에서부터 따라가며, 번호가 붙은 모든 파트를 방문 순서대로 1부터 촘촘하게 다시 매깁니다.

그래서 PowerPoint가 아닌 것이 쓴 덱, 즉 큰 덱에서 잘라낸 파일이나 스크립트나 python-pptx가 만든 파일은 사람이 처음 저장하는 순간 모든 파트의 이름이 바뀝니다. 추적 폴더에서는 이것이 "모든 파일이 지워지고 이름이 다른 사본이 새로 생겼다"로 보이며, 실제로 고친 슬라이드 한 장이 그 안에 묻힙니다.

`decomp`는 이 번호를 미리 같은 규칙으로 맞춰 둡니다. PowerPoint가 쓴 덱은 이미 그 형태이므로 아무것도 바뀌지 않고 `comp`도 바이트 단위로 같은 파일을 돌려주며, 그렇지 않은 덱은 PowerPoint가 결국 도달할 형태로 미리 옮겨집니다. 움직이는 것은 번호뿐이고 내용과 관계 ID와 요소 순서는 그대로입니다. 이름을 바꾸기 전에 결과를 먼저 검사해서, 끊어지는 관계가 생기거나 콘텐츠 타입을 잃는 파트가 있으면 이름을 건드리지 않고 경고만 남깁니다.

## 설치

```bash
git clone https://github.com/zer0ken/git-pptx.git
cd git-pptx
npm link
```

`npm link`는 `git-pptx` 실행 파일을 PATH에 노출합니다.

## 사용법

```bash
# a.pptx -> a.git-pptx/ 분해 (변경 슬라이드만 갱신, 미리보기 렌더)
git-pptx decomp a.pptx

# a.git-pptx/ -> a.pptx 재조립
git-pptx comp a.git-pptx a.pptx

# 변경 슬라이드만 표시 (쓰기 없이)
git-pptx diff a.pptx a.git-pptx
```

옵션:
- `--no-preview`: 미리보기 렌더 생략
- `--format png`: 미리보기를 JPG(기본) 대신 PNG로
- `--renderer auto|powerpoint|libreoffice`: 미리보기 렌더러 선택 (기본 `auto`)
- `--no-normalize`: 파트 이름을 정규화하지 않고 덱이 갖고 온 이름을 그대로 둠

## 동작 상세

| 명령 | 하는 일 |
|---|---|
| `git-pptx decomp a.pptx` | `a.git-pptx/` 생성(또는 증분 갱신). 이미 분해본이 있으면 변경 파일만 갱신하고 **변경 슬라이드만** 미리보기 재렌더 |
| `git-pptx comp <dir> <out.pptx>` | `pptx/`를 zip해 `.pptx` 재조립 |
| `git-pptx diff <deck.pptx> <dir>` | 변경 슬라이드 목록 출력 |

## GitHub에서 논의하기

`previews/1.jpg, 2.jpg, ...`를 커밋하면 GitHub가 이미지 미리보기와 PR 이미지 diff를 지원하므로, 슬라이드 번호로 지칭하며 인라인 코멘트로 논의할 수 있습니다. `.pptx` 파일은 `.gitignore`에 넣어 커밋하지 않고 `a.git-pptx/`만 푸시합니다.

## 제약

- `docProps/core.xml`, `docProps/app.xml`, `docProps/thumbnail.jpeg`는 저장할 때마다 다시 만들어지므로 변경 감지에서 제외됩니다. 다만 폴더에는 그대로 씁니다. 이들을 가리키는 파트는 기록되기 때문에, 대상 파일을 빼면 재조립한 덱을 PowerPoint가 열지 못합니다.
- 정규화는 재직렬화의 서식 노이즈를 무시하지만, 관계 ID 재번호(r:id)와 기본값 materialization은 아직 다루지 않습니다. 다른 도구가 쓴 덱을 PowerPoint로 처음 저장할 때 몇 개의 파트가 변경으로 보이며, 그다음부터는 잦아듭니다.
- 파트 이름 정규화는 시험한 모든 덱에서 PowerPoint와 같은 순서를 냈습니다. 다만 SVG 그림과 그 대체 래스터 이미지를 함께 담은 덱에서는 미디어 파트 몇 개의 순서가 달라질 수 있고, 그 경우 PowerPoint가 처음 저장할 때 한 번 이름을 바꿉니다.
- 미리보기 렌더링은 **크로스 플랫폼**입니다: `libreoffice`(LibreOffice headless + poppler, 전 OS) 또는 `powerpoint`(PowerPoint COM, Windows). 기본 `auto`는 Windows에서 PowerPoint가 있으면 그걸, 아니면 LibreOffice를 씁니다. 렌더러에 따라 미리보기가 다르므로 팀은 `--renderer`로 하나로 통일해야 합니다.

## 개발

- `scripts/make-fixture.js`: 테스트용 미니 pptx 생성
- `scripts/mutate.js`: 재직렬화 노이즈 + 실제 수정을 흉내 내 deck 변형
- `scripts/check-normalize.js`: 파트 이름 정규화 자체 검사 (`node scripts/check-normalize.js`)
