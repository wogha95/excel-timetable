# 📝excel-timetable
- 수업교과목 목록을 시간표(Excel file)로 만드는 파일입니다.
- 수작업으로 하나씩 확인하며 만드는 건 불필요한 시간과 노동을 요구한다고 느껴져 만들게 되었습니다.

<br>

## ✔작동원리
1. Excel sheet을 json으로 만듭니다. (XLSX.utils.sheet_to_json)
2. 1교시 → 2교시 → 3교시 → .. 기준으로 정렬하고 동일한 교시는 월 → 화 → 수 → .. 기준으로 정렬합니다.
3. 정렬된 json 객체를 시간표와 수업 갯수로 맞추고
4. 작업된 json을 Excel sheet로 만듭니다. (XLSX.utils.json_to_sheet)
