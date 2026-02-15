# Excel 시트 수정 + PDF 저장 (Java)

요구사항:
1. 엑셀 파일 열기
2. 특정 시트 불러오기
3. 특정 셀에 데이터 입력
4. 특정 시트를 PDF로 저장

이 프로젝트는 Apache POI로 엑셀을 수정하고, PDFBox로 선택한 시트를 텍스트 기반 PDF로 내보냅니다.

## 저장소 클론

```bash
git clone <REPOSITORY_URL>
cd TEST
```

예시:

```bash
git clone https://github.com/your-org/your-repo.git
cd your-repo
```

## 실행 방법

```bash
mvn package
java -jar target/excel-sheet-pdf-tool-1.0.0.jar input.xlsx Sheet1 B3 "안녕하세요" output.xlsx sheet1.pdf
```

## 인자 설명

- `input.xlsx`: 원본 엑셀 파일
- `Sheet1`: 수정/내보내기 대상 시트 이름
- `B3`: 값을 입력할 셀 주소
- `"안녕하세요"`: 입력할 값
- `output.xlsx`: 수정본 엑셀 저장 경로
- `sheet1.pdf`: 대상 시트 PDF 저장 경로

## 참고

- PDF는 셀 값을 행 단위 텍스트로 출력하는 방식입니다.
- 엑셀의 복잡한 서식(폰트/테두리/병합/차트)을 1:1로 재현하지는 않습니다.
