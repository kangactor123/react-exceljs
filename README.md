# React ExcelJS

- ExcelJS를 React에서 사용한 케이스입니다.
- 커스텀 훅을 활용해 컴포넌트(페이지)에서 엑셀 파일을 다운로드 받을 수 있습니다.
- 자세한 설명은 아래 블로그를 확인해주세요.

<a href="https://kangs-develop.tistory.com/8" target="_blank">바로가기</a>

### useExcelDownload
- useExeclDownload 훅은 `src/share/hooks` 경로에서 확인할 수 있습니다.
- 해당 훅의 사용 예시는 `App.tsx` 파일에서 확인 가능합니다.

<br />

## Change log
### 08/11
- 컬럼의 너비를 계산하는 로직을 변경했습니다.
- 컬럼의 너비를 보정하는 비율을 수정했습니다. `LENGTH_CORRECTION_RATIO` 상수 값을 통해 비율을 조정할 수 있습니다.
