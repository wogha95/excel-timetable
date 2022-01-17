# 📝excel-timetable
- 수업교과목 목록을 시간표로 만드는 파일입니다.
- 수작업으로 하나씩 확인하며 만드는 건 불필요한 시간과 노동을 요구한다고 느껴져 만들게 되었습니다.

<br>

## 📌SheetJS CDN 프로세스
1. excel file 업로드하여 workbook 생성
1. workbook의 sheet별로 읽기
2. sheet를 html, json, csv 데이터로 변환
3. 데이터를 작업한 후
4. 작업된 데이터를 sheet로 변환
5. sheet를 workbook에 추가
6. workbook을 excel file로 다운로드

<br>

## ✔작동원리
### 1. Excel sheet을 json으로 만듭니다.
``` js
// line 76
const test = XLSX.utils.sheet_to_json(sheet);
```
### 2. 1교시 → 2교시 → 3교시 → .. 기준으로 정렬하고 동일한 교시는 월 → 화 → 수 → .. 기준으로 정렬합니다.
``` js
// line 77
test.sort(function (a, b) {
  return timeToIndex(a.실습시간) - timeToIndex(b.실습시간);
});
```
### 3. 정렬된 json 객체를 시간표와 수업 갯수로 맞추고
``` js
// line 87
for (let index = 0; index < 30; index++) {
  if (index !== 0 && index % 5 === 0) {
    outputs.push(output);
    output = {
      시간표: times[index / 5],
    };
  }
  const day = days[index % 5];
  if (
    test.length - 1 < testIndex ||
    test[testIndex].실습시간[0] !== day ||
    test[testIndex].실습시간.slice(2) !== times[parseInt(index / 5)]
  )
    output[day] = "-";
  else {
    output[day] =
      test[testIndex].과목명 +
      "/" +
      test[testIndex].분반 +
      "-" +
      test[testIndex].실습조 +
      "/" +
      test[testIndex].수강대상;
    testIndex++;
  }
  if (index == 29) outputs.push(output);
}
```
### 4. 작업된 json을 Excel sheet로 만들고 workBook에 추가합니다.
``` js
// line 114
const ws = XLSX.utils.json_to_sheet(outputs);
XLSX.utils.book_append_sheet(workBook, ws, sheetName[index]);
```
### 5. 1~4 과정을 입력된 excel file의 모든 sheet를 순회하며 만듭니다.
``` js
// line 66
sheetNameList.forEach((element, index) => {
  // 각 시트를 돌며 가공된 sheet를 workBook에 추가
  var sheetName = wb.Sheets[element];
  callback(sheetName, index);
});
```
### 6. 모든 sheet를 작업하고 1개의 excel file로 자동 다운로드 됩니다.
``` js
// line 71
XLSX.writeFile(workBook, "example.xlsx");
```

<br>

## 🔗참고링크
[링크1](https://eblo.tistory.com/83)
<br>
[링크2](https://ichi.pro/ko/xlsxleul-sayonghayeo-node-jseseo-excel-pail-ilg-gi-sseugi-188091786395828)
