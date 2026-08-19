# git pptx

PowerPoint(pptx) 파일을 Git에서 **슬라이드 단위로 버전 관리**하는 도구입니다. `git`의 네이티브 서브커맨드(`git pptx`)로 동작합니다.

## 원리

- pptx는 zip이므로, "쪼개기"는 무손실 unzip이고 "합치기"는 무손실 rezip입니다. 관계 ID나 `[Content_Types].xml`을 재구성하지 않아 데이터 손실이 없습니다.
- 저장소는 pptx의 분해된 디렉터리(`slides/`)를 추적합니다. 사용자가 편집하는 `deck.pptx`는 `.gitignore`로 제외됩니다.
- 변경 감지는 **XML 정규화(canonicalization)** 기반입니다. 속성 순서, 엔티티 인코딩, 빈 요소 표기, 네임스페이스 접두사처럼 PowerPoint 재저장 시 변하는 노이즈를 무시하고, 진짜 내용이 바뀐 슬라이드만 감지해 커밋합니다. 요소 순서(z-order)는 보존합니다.

## 설치

```bash
cd pit
npm link
```

`npm link`는 `git-pptx` 실행 파일을 PATH에 노출해서 `git pptx`으로 호출되게 합니다.

## 사용법

```bash
# 분해 디렉터리(slides/)로 초기 분리 + 초기 커밋. 리모트가 있으면 push
git pptx init --deck deck.pptx --slides-dir slides --upstream <url>

# 로컬 deck.pptx의 변경 슬라이드만 감지해 커밋하고 push
git pptx push

# 리모트의 분해 디렉터리를 받아 deck.pptx로 재조립
git pptx pull

# deck.pptx와 커밋된 상태의 차이(변경 슬라이드) 확인
git pptx status
```

## 동작 상세

| 명령 | 하는 일 |
|---|---|
| `git pptx init` | `deck.pptx`를 `slides/`로 unzip하고 설정(`.pit/config.json`)과 `.gitattributes`(`slides/** -text`) 기록 |
| `git pptx push` | deck을 임시로 unzip → 정규화 비교로 변경 파일 감지 → 변경분만 `slides/`에 반영·커밋·push |
| `git pptx pull` | `git pull` 후 `slides/`를 `deck.pptx`로 rezip |
| `git pptx status` | 변경 슬라이드 목록 출력 |
| `git pptx preview` | PowerPoint COM으로 각 슬라이드를 `preview/slide-NNN.png`로 렌더링·커밋. GitHub 이미지 diff로 리뷰/논의 가능 |

## 제약

- `docProps/core.xml`(타임스탬프 메타데이터)은 매 저장마다 바뀌므로 변경 감지에서 제외됩니다.
- 정규화는 재직렬화의 서식 노이즈를 무시하지만, 관계 ID 재번호(r:id)와 기본값 materialization처럼 의미가 필요한 변화는 아직 다루지 않습니다. 동일 슬라이드가 재번호 때문에 변경으로 보일 수 있습니다.
- 리모트가 없으면 `git pptx pull`은 로컬 `slides/`에서 재조립합니다.

## 개발

- `scripts/make-fixture.js`: 테스트용 미니 pptx 생성
- `scripts/mutate.js`: 재직렬화 노이즈 + 실제 수정을 흉내 내 deck 변형

```bash
node scripts/make-fixture.js deck.pptx
node scripts/mutate.js deck.pptx
git pptx status
git pptx push
```
