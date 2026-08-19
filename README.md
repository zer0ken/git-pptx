# git-pptx

PowerPoint(pptx) 파일을 **슬라이드 단위로 버전 관리**하기 위한 독립 도구입니다. git/github와 결합하지 않으며, `decomp`/`comp`로 pptx ↔ git-pptx 디렉터리 형식을 상호 변환합니다. 변환 결과는 일반 git으로 커밋·푸시합니다.

## git-pptx 형식

`a.pptx`를 decomp하면 `a.git-pptx/` 디렉터리가 생성됩니다. 파일과 같은 폴더에 공존합니다.

```
a.git-pptx/
  previews/   1.jpg, 2.jpg, ...   슬라이드별 미리보기(파생 산출물)
  pptx/       pptx를 압축 해제한 원본 내용
```

- **decomp**는 무손실 unzip이라 관계 ID나 `[Content_Types].xml`을 재구성하지 않아 데이터 손실이 없습니다.
- **comp**는 `pptx/`를 다시 zip합니다 (미리보기·VCS 파일 제외).
- **변경 감지**는 XML 정규화 기반입니다. 속성 순서, 엔티티 인코딩, 빈 요소 표기, 네임스페이스 접두사처럼 PowerPoint 재저장 시 변하는 노이즈를 무시하고, 진짜 내용이 바뀐 슬라이드만 감지합니다. 요소 순서(z-order)는 보존합니다.

## 설치

```bash
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

## 동작 상세

| 명령 | 하는 일 |
|---|---|
| `git-pptx decomp a.pptx` | `a.git-pptx/` 생성(또는 증분 갱신). 이미 분해본이 있으면 변경 파일만 갱신하고 **변경 슬라이드만** 미리보기 재렌더 |
| `git-pptx comp <dir> <out.pptx>` | `pptx/`를 zip해 `.pptx` 재조립 |
| `git-pptx diff <deck.pptx> <dir>` | 변경 슬라이드 목록 출력 |

## GitHub에서 논의하기

`previews/1.jpg, 2.jpg, ...`를 커밋하면 GitHub가 이미지 미리보기와 PR 이미지 diff를 지원하므로, 슬라이드 번호로 지칭하며 인라인 코멘트로 논의할 수 있습니다. `.pptx` 파일은 `.gitignore`에 넣어 커밋하지 않고 `a.git-pptx/`만 푸시합니다.

## 제약

- `docProps/core.xml`(타임스탬프 메타데이터)은 매 저장마다 바뀌므로 변경 감지에서 제외됩니다.
- 정규화는 재직렬화의 서식 노이즈를 무시하지만, 관계 ID 재번호(r:id)와 기본값 materialization은 아직 다루지 않습니다. 동일 슬라이드가 재번호 때문에 변경으로 보일 수 있습니다.
- 미리보기 렌더링은 Windows의 설치된 PowerPoint COM을 사용합니다.

## 개발

- `scripts/make-fixture.js`: 테스트용 미니 pptx 생성
- `scripts/mutate.js`: 재직렬화 노이즈 + 실제 수정을 흉내 내 deck 변형
