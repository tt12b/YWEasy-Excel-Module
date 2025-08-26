# Easy Excel Module

## 📋 Tech stack

![Java](https://img.shields.io/badge/JDK_17-%23ED8B00.svg?style=flat&logo=openjdk&logoColor=white)
![Spring Boot](https://img.shields.io/badge/Spring_Boot_3.3.6-6DB33F?style=flat&logo=spring&logoColor=white)


- JDK 17 / Spring Boot 3.36 환경에서 제작되었습니다.

- EasyWorkBook은 Apache POI를 기반으로 엑셀 다운로드 기능을 더 직관적이고 간단하게 구현하기 위한 유틸 라이브러리입니다.
  복잡한 스타일 지정과 시트/행/셀 생성 로직을 캡슐화하여, 빌더 패턴으로 손쉽게 Excel 파일을 생성할 수 있습니다.


## ✨ 주요 기능

- 빌더 패턴 기반의 엑셀 Workbook 객체 생성

- 여러 개의 시트(EasySheet)를 손쉽게 추가 가능

- 시트, 행, 셀 단위의 스타일 지정 지원

- ResponseEntity<byte[]> 형태로 Spring Controller에서 바로 다운로드 응답 반환

- 행 / 시트 / Woorkbook 객체를 빌더 패턴으로 손쉽게 생성

- 기본 스타일 지원