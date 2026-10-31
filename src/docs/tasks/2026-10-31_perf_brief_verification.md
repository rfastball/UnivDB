# 성능 브리프 사전 검증 (2026-10-31)

대상: `지출 리팩토링\src\docs\tasks\perf-synth\univdb-perf-brief.md` (같은 날 작성).
방법: UnivDB HEAD(7fdc89b, 2.9.0) 소스 대조 + git 이력 + VBA 실측.
실측은 임시 모듈 `ZPerfProbe`를 push -> `vbaloop run` (ephemeral, 읽기 전용) 후 `push --prune`로 제거했다.
모듈과 원시 결과는 `experiments/perf-20261031/` (`ZPerfProbe.bas`, `probe_*.json`). 수치는 워밍업 1회 + 3회 중앙값.

## 0. 판정 요약

| 제안 | 판정 | 핵심 사유 |
|---|---|---|
| 1A. `UCase$` + binary Dictionary | **기각 -> 하이브리드 접기로 대체 채택** | `UCase$`만으로는 전각/반각 키가 갈라짐. 하이브리드는 의미 보존 + 91k에서 1.8배 |
| 1B. Core 백엔드 기본화 | **기각** | 2.9.0 Core는 27k/91k에서 Dict보다 느림. 브리프 수치는 2.8.0의 구 해시 측정값 |
| 2. Trim 생략 옵션 | **조건부 채택** | 이 데이터에서 약 100ms 절감. 단 Filter와 인덱스 경로의 공백 처리가 갈림 - 계약 문서화 필수 |
| 2'. `AdoptArray` 공개 | **기각 (그대로는)** | 복사가 1회 남고 중복 헤더 거부를 건너뜀. 대신 API 변경 없는 내부 복사 1회 제거 |
| 3a. UpsertFrom 매핑 호이스팅 | **채택** | 저위험, 약 0.15초 (소비자 측정) |
| 3b. 변경 행 보고 API | **방향 채택, 비교 규칙은 재설계** | 소비자 규칙이 locale 의존, "같음" 판정 행도 테이블 값은 덮어써져 테이블과 시트가 어긋남 |

## 1. 브리프의 사실 오류

1. **"핫 경로 코드는 2.8.0과 2.9.0이 같다" - 틀림.** 설치본 2.8.0은 fe5f1d9(2026-07-21) 빌드.
   e70a91d(2026-10-03)가 `Core_HashKeyText`를 LCMapStringW sort key 해시로 교체했고, clsTable 421줄,
   clsQuery 165줄이 바뀌었다. 브리프의 Core 수치와 4절 `SelectTable` 측정은 2.9.0에서 다시 재야 한다.
2. **"binary 모드면 초선형이 사라진다" - 틀림.** 아래 실측에서 binary도 91k -> 182k에 2.6배(지수 약 1.4).
   Scripting.Dictionary 자체의 성질이다. binary는 상수배 개선이다.
3. **"같은 Dictionary를 binary로만 바꿔도 3배"는 DictHashIndex 전체와 비교 불가.**
   `x.dw_dict_direct_*`는 원시 Variant 키로 Dictionary만 돈다. 실제 경로에는 `NormalizeIndexKey` + CSR + 컬럼 읽기 고정비가 붙는다
   (실측: 91k에서 real_dict 498 vs 같은 루프 복제 431ms).
4. 3절 `SameCellValue` 서술이 실제 코드와 다르다(교차 타입이면 문자열 쪽 `Trim$`, Boolean 분기 존재).

## 2. 제안 1: 인덱스 키 비교

### 2.1 실측 (VBA, 고유 키, ms)

`dict_text` 등은 `DictHashIndex.AddRow` 경로 복제(NormalizeIndexKey -> 접기 -> exists/Add -> CSR AddEntry).
`real_*`는 공개 경로 `clsTable.GetHashIndex`. 키 형태 ascii = `"R25TA" & 8자리`(합성 DW 키), hangul = `"가나상사" & 6자리`.

| 경로 | 27k | 91k | 182k | 91k hangul |
|---|---:|---:|---:|---:|
| dict_text (현행 복제) | 68 | 431 | 1,383 | 418 |
| ucase_binary | 46 | 203 | 527 | 195 |
| lcmap_inplace_binary (전 키 LCMapStringW) | 85 | 329 | 784 | 320 |
| **hybrid_bytes_binary (권장)** | **58** | **234** | **617** | **222** |
| hybrid_like_binary (`Like`로 판별) | 67 | 284 | 687 | 272 |
| strconv_binary (`StrConv` 대문자+반각) | 59 | 245 | 625 | 241 |
| real_dict (현행 공개 경로) | 80 | 498 | 1,465 | 443 |
| real_core (현행 공개 경로) | 151 | 525 | 1,108 | 510 |

하이브리드 = 문자열에 U+3000-30FF, U+3131-318E, U+FF00-FFFF 코드 유닛이 있을 때만 `LCMapStringW(UPPERCASE|HALFWIDTH)`,
아니면 `UCase$`. 판별은 `b() = s` 바이트 배열 상위 바이트 검사(키당 약 0.3us). LCMapStringW 호출은 키당 약 1.4us라
전 키에 쓰면 27k에서 현행보다 느리다.

공개 경로 기대치(고정비 약 70ms 유지): 27k 80 -> 약 70ms(소폭), 91k 498 -> 약 300ms, 182k 1,465 -> 약 690ms.
**브리프의 "마스터 27k 84ms -> ~30ms"는 성립하지 않는다.** 이득은 키가 많을수록 커진다.

### 2.2 의미론 (현행 Dictionary text 모드와 동치인가)

쌍 대조(VBA):

| 입력 쌍 | Dict text (현행) | `UCase$` | LCMapStringW / 하이브리드 | `StrConv` |
|---|---|---|---|---|
| `abc`/`ABC`, `é`/`É`, 키릴, 로마 숫자, 원문자 대소 | 같음 | 같음 | 같음 | 같음 |
| 전각 `ＡＢＣ`, 전각 숫자, 전각 소문자, 전각 공백 | 같음 | **다름** | 같음 | 같음 |
| `ㄱ`(U+3131) / 반각 `ﾡ`(U+FFA1) | 같음 | **다름** | 같음 | 같음 |
| `e` / `é` | 다름 | 다름 | 다름 | **같음 (오병합)** |
| `ガ` / 반각 `ｶﾞ` | 다름 | 다름 | 같음 | 같음 |
| `ß`/`ss`, 결합 `é`, 종지 시그마, 가타카나/히라가나 | 다름 | 다름 | 다름 | 다름 |

BMP 전 단일 문자 전수 대조(63,488자):
- 하이브리드가 현행이 가르는 문자를 합친 경우: **0건**.
- 현행이 합치는데 하이브리드가 가르는 경우: 1,400건. 전부 미할당(981), 서식 제어(ZWNJ/LRM/양방향 마크), U+FFFx 특수,
  인도계/아르메니아 문자 등이 서로 무관한 문자와 "같음"으로 판정되는 현행 쪽 오병합이다. 라틴/한글/한자/일반 구두점은 0건.
- 하이브리드 vs 전 키 LCMapStringW 차이: U+2018/2019/201C/201D(굽은 따옴표)만 - LCMAP_HALFWIDTH가 ASCII 따옴표로 접는다. 하이브리드는 현행대로 분리.

결론: `UCase$`안은 기각(전각 키를 조용히 분리). `StrConv`는 `e`/`é` 오병합으로 기각(로캘 의존이기도 함).
하이브리드는 실무 문자에서 현행과 같고, 차이는 반각 가나 탁음 결합(StrComp 쪽으로 정렬)과 현행의 미할당 문자 오병합 해소뿐.
이 차이는 Reference 6장 "Unicode 한계"에 적고 테스트로 고정할 것.

### 2.3 적용 범위

`vbTextCompare` Dictionary는 인덱스 말고도 `UpsertFrom`의 `seenKeys`, GroupBy 복합 키 `dicAccumulator`/`dicGroupParts`,
`clsSet`에 있다. 인덱스만 바꾸면 표면마다 접기가 갈린다. 접기 함수를 `modIndexKey`에 두면 UnivFilter 재스냅샷 대상
(clsSet/modIndexKey 벤더링). 원본 표기는 keyId 순서 배열로 보존(`Cursor()`가 `mDict.Keys` 순서 = keyId에 의존).

### 2.4 제안 1B와 소비자 경고

real_core는 27k 151 / 91k 525ms로 Dict(80/498)보다 느리고 182k에서만 앞선다. 기본값 전환 근거 없음.
**소비자 경고**: 지출 리팩토링이 본 Core 134ms(91k)는 2.8.0 구 해시 값이다. 그 구 해시는 전각 키 동치를 깨던 버전이다.
2.9.x를 설치하면 `IndexBackend = 1` 핀 그대로 DW 인덱스가 약 525ms로 느려진다. 하이브리드 Dict(약 300ms)를 넣어도
현재 체감(134ms)보다는 느리다 - 애드인 교체와 핀 제거를 한 묶음으로 계획하고 교체 후 `X_DwIndexBackend`로 재측정.

## 3. 제안 2: 적재 Trim / 복사

실측: 27,401 x 34(2/3 문자열 셀, 0.5% 앞뒤 공백), ms.

| 경로 | ms |
|---|---:|
| `LoadFromArray` 전체 (현행) | 241 |
| Trim 루프 (현행: 문자열이면 무조건 `Trim$` 재할당) | 103 |
| Trim 루프 (앞뒤 공백일 때만 재할당, `Left$`/`Right$` 판별) | 127 |
| Trim 루프 (같은 조건, `AscW` 판별) | 131 |

- "변경 시만 재할당" 최적화는 **기각** - 판별 비용이 재할당보다 크다(결과 일치는 확인, diff 0).
- 나머지 약 138ms는 복사 2회(`ByVal vArr`, `m_vData = vArr`) + ByVal 사본 해제 + 헤더. 문자열 셀이 많으면 복사가 비싸다
  (브리프의 27ms는 숫자 위주 데이터 값으로 보임).
- **API 변경 없는 절감**: `LoadFromSheet`/`RegisterTable`이 Private ByRef 적재 헬퍼를 타게 하면 시트 경로에서 복사 1회 + 해제가 빠진다
  (이 데이터 추정 60~70ms). 배열 경로는 `RegisterTable(ByVal Source)`까지 3회라 2회 절감.
- `AdoptArray`를 그대로 공개하면 안 된다: `m_vData = vArr`에서 여전히 복사하고, `BuildHeaders(..., False)`라 중복 헤더를 통과시킨다.
- **Trim 생략 옵션의 계약 부작용**: `Filter "="`는 비교 시 `modEngine.FastNorm`이 셀을 `Trim$`하지만 `Where`/Join/GroupBy는
  `NormalizeIndexKey`(Trim 없음)를 탄다. 지금은 적재 Trim이 둘을 일치시킨다. `bTrim:=False`면 `"abc "` 셀이 Filter로는 잡히고
  Where로는 안 잡힌다. 적재 Trim 자체가 Reference에 없는 암묵 계약이므로 이번에 명문화할 것.

## 4. 제안 3: UpsertFrom

- 3a 호이스팅: 코드 확인 결과 브리프 설명 정확(`For Each vKey In dicColMap.Keys` + 셀당 조회). 채택.
- 3b 비교 규칙을 라이브러리에 그대로 옮기면 안 되는 이유:
  1. `IsDate`/`CDate`/`IsNumeric`/`CDbl`은 로캘 의존 - 날짜 키를 locale 무관으로 만든 D3 결정과 충돌.
     `IsNumeric("&H10")`, `IsNumeric("1,000")`이 참이 되는 등 라이브러리 계약으로는 느슨하다.
  2. "같음" 판정 셀도 현 루프는 소스 문자열로 덮어쓴다. 시트는 Double 2025, 테이블은 String "2025"로 어긋난다.
- 대안 설계: 비교 모드 Enum(엄격 = VarType+값 / 시트 등가 = 문서화된 느슨 규칙) + **동등하면 쓰지 않음**(대상 값 유지).
  그러면 outChangedRowIds = 실제로 쓴 행 = 테이블과 시트 일치. 기본 `UpsertFrom`은 현행 유지.
- 추가 기회: 갱신만 있는 Upsert도 `BumpGeneration`으로 전 인덱스를 실효시킨다. `UpdateRows`처럼 변경 열만 무효화하고
  변경 0건이면 생략하면 키 인덱스를 다음 호출에서 재사용한다(값 갱신은 구조 변경이 아님).
- 3b의 비교 비용은 이번에 재지 않았다.

## 5. 구현 순서 제안

1. 적재 복사 1회 제거(API 불변) - 위험 최소.
2. UpsertFrom 호이스팅.
3. 하이브리드 키 접기: 인덱스 + `seenKeys` + GroupBy 복합 키 + `clsSet`을 한 번에. 2.2절 차이를 테스트로 고정, Reference 6장 갱신, UnivFilter 재스냅샷 메모.
4. Trim 생략 옵션 + 적재 Trim 계약 명문화.
5. 변경 행 보고 API(비교 모드 설계 확정 후).
