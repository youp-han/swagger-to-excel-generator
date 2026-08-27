# swaggerToExcel

Swagger 2.0 JSON을 태그(컨트롤러)별 API 정의서 Excel 파일로 변환하는 스크립트.

## 요구사항

- Python 3
- `openpyxl` (`pip install openpyxl`)

`urllib.request` 등 표준 라이브러리만으로 JSON을 가져오므로 `requests`는 필요 없다.

## 사용법

```
python swagger_to_excel.py --url <Swagger JSON URL> --output-dir <저장 폴더> --prefix <파일명 접두사>
```


| 옵션             | 필수  | 설명                                                                                                           |
| -------------- | --- | ------------------------------------------------------------------------------------------------------------ |
| `--url`        | O   | Swagger 2.0 JSON 엔드포인트 URL. Swagger UI 페이지 주소가 아니라 실제 JSON 응답 주소여야 함 (Swashbuckle 계열은 보통 `/swagger/docs/v1`) |
| `--output-dir` | O   | 결과 xlsx 파일들을 저장할 폴더 (없으면 자동 생성)                                                                              |
| `--prefix`     | O   | 파일명 접두사. 최종 파일명은 `{번호}.{접두사}_{태그명}_api정의서.xlsx`                                                              |
| `--tags`       | X   | 지정한 태그만 변환 (여러 개 가능) — 샘플 확인용                                                                                |
| `--limit`      | X   | 앞의 N개 API만 변환 — 샘플 확인용                                                                                       |


예시:

```
# Waven API 전체 변환
python swagger_to_excel.py --url "http://123.123.123.123:8080/swagger/docs/v1" \
  --output-dir "excel/api정의서" --prefix "asdf_adsf"

# 특정 태그만 미리 확인
python swagger_to_excel.py --url "http://123.123.123.123:8080/swagger/docs/v1" \
  --output-dir "excel/sample" --prefix "TEST" --tags BrandApi
```

## 출력 구조

Swagger의 각 operation 첫 번째 `tags` 값(=컨트롤러) 기준으로 파일이 하나씩 나뉜다.
각 파일은:

1. **업데이트 이력** 시트
2. `**<태그명>**` 목차 시트 — No / API명 / Method / URL / 설명 + 상세 시트로 이동하는 하이퍼링크
3. `**1. <API명>**` ~ `**N. <API명>**` 상세 시트 — `IF Title` / `IF 방식` / `URL` / `Method` / `설명`
 라벨 블록, `REQUEST`(header/body) 표, `RESPONSE`(status/body) 표. 중첩 필드는 들여쓰기로 표시.

모든 시트는 A열 1칸 + 1행을 여백으로 비워두고, 표(라벨 블록/REQUEST/RESPONSE/목차/이력)마다
바깥 테두리를 굵게 처리한다.

## 알려진 제약

- **Swagger 2.0 전제.** `in: body` 파라미터, `#/definitions/...` `$ref` 구조를 가정한다.
OpenAPI 3.0(`requestBody`, `#/components/schemas/...`) 소스는 파서 수정이 필요하다.
- **summary/description 커버리지에 좌우됨.** Swagger 스펙에 `summary`/`description`이 없는
API는 `IF Title`이 영문 코드명(경로 마지막 세그먼트)으로 채워지고 설명 칸은 빈 채로 나온다.

- 순환 참조 스키마는 재귀 깊이 4단계에서 끊는다 (무한루프 방지).



