# ๐excel-timetable
- ์์๊ต๊ณผ๋ชฉ ๋ชฉ๋ก์ ์๊ฐํ๋ก ๋ง๋๋ ํ์ผ์๋๋ค.
- ์์์์ผ๋ก ๋ชฉ๋ก์์ ํ๋์ฉ ํ์ธํ๋ฉฐ ์๊ฐํ๋ฅผ ๋ง๋๋ ๊ฑด ๋ถํ์ํ ์๊ฐ๊ณผ ๋ธ๋์ ์๊ตฌํ๋ค๊ณ  ๋๊ปด์ ธ ๋ง๋ค๊ฒ ๋์์ต๋๋ค.

<br>

## ๐SheetJS CDN ํ๋ก์ธ์ค
1. excel file ์๋ก๋ํ์ฌ workbook ์์ฑ
1. workbook์ sheet๋ณ๋ก ์ฝ๊ธฐ
2. sheet๋ฅผ html, json, csv ๋ฐ์ดํฐ๋ก ๋ณํ
3. ๋ฐ์ดํฐ๋ฅผ ์์ํ ํ
4. ์์๋ ๋ฐ์ดํฐ๋ฅผ sheet๋ก ๋ณํ
5. sheet๋ฅผ workbook์ ์ถ๊ฐ
6. workbook์ excel file๋ก ๋ค์ด๋ก๋

<br>

## โ์๋์๋ฆฌ
### 1. Excel sheet์ json์ผ๋ก ๋ง๋ญ๋๋ค.
``` js
// line 76
const test = XLSX.utils.sheet_to_json(sheet);
```
### 2. 1๊ต์ โ 2๊ต์ โ 3๊ต์ โ .. ๊ธฐ์ค์ผ๋ก ์ ๋ ฌํ๊ณ  ๋์ผํ ๊ต์๋ ์ โ ํ โ ์ โ .. ๊ธฐ์ค์ผ๋ก ์ ๋ ฌํฉ๋๋ค.
``` js
// line 77
test.sort(function (a, b) {
  return timeToIndex(a.์ค์ต์๊ฐ) - timeToIndex(b.์ค์ต์๊ฐ);
});
```
### 3. ์ ๋ ฌ๋ json ๊ฐ์ฒด๋ฅผ ์๊ฐํ์ ์์ ๊ฐฏ์๋ก ๋ง์ถ๊ณ 
``` js
// line 87
for (let index = 0; index < 30; index++) {
  if (index !== 0 && index % 5 === 0) {
    outputs.push(output);
    output = {
      ์๊ฐํ: times[index / 5],
    };
  }
  const day = days[index % 5];
  if (
    test.length - 1 < testIndex ||
    test[testIndex].์ค์ต์๊ฐ[0] !== day ||
    test[testIndex].์ค์ต์๊ฐ.slice(2) !== times[parseInt(index / 5)]
  )
    output[day] = "-";
  else {
    output[day] =
      test[testIndex].๊ณผ๋ชฉ๋ช +
      "/" +
      test[testIndex].๋ถ๋ฐ +
      "-" +
      test[testIndex].์ค์ต์กฐ +
      "/" +
      test[testIndex].์๊ฐ๋์;
    testIndex++;
  }
  if (index == 29) outputs.push(output);
}
```
### 4. ์์๋ json์ Excel sheet๋ก ๋ง๋ค๊ณ  workBook์ ์ถ๊ฐํฉ๋๋ค.
``` js
// line 114
const ws = XLSX.utils.json_to_sheet(outputs);
XLSX.utils.book_append_sheet(workBook, ws, sheetName[index]);
```
### 5. 1~4 ๊ณผ์ ์ ์๋ ฅ๋ excel file์ ๋ชจ๋  sheet๋ฅผ ์ํํ๋ฉฐ ๋ง๋ญ๋๋ค.
``` js
// line 66
sheetNameList.forEach((element, index) => {
  // ๊ฐ ์ํธ๋ฅผ ๋๋ฉฐ ๊ฐ๊ณต๋ sheet๋ฅผ workBook์ ์ถ๊ฐ
  var sheetName = wb.Sheets[element];
  callback(sheetName, index);
});
```
### 6. ๋ชจ๋  sheet๋ฅผ ์์ํ๊ณ  1๊ฐ์ excel file๋ก ์๋ ๋ค์ด๋ก๋ ๋ฉ๋๋ค.
``` js
// line 71
XLSX.writeFile(workBook, "example.xlsx");
```

<br>

## ๐์ฐธ๊ณ ๋งํฌ
- [SheetJS๋ก ํ์ผ ์ฝ๊ธฐ](https://eblo.tistory.com/83)
- [XLSX๋ฅผ ์ฌ์ฉํ์ฌ Node.js์์ Excel ํ์ผ ์ฝ๊ธฐ / ์ฐ๊ธฐ](https://ichi.pro/ko/xlsxleul-sayonghayeo-node-jseseo-excel-pail-ilg-gi-sseugi-188091786395828)
