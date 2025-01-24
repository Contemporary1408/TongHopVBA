# Tổng hợp các code VBA chạy Excel
*(PW for VBA project if any: 140819940)*
- Khi replace hay remove các khoảng trống trong file txt bằng VBA, khoảng trống đó có thể không thực sự là khoảng trống thông thường mà là ký tự do ấn Tab (vbTab).
- Vì sao phải Activate sheet rồi mới copy được: https://stackoverflow.com/questions/41384018/vba-copy-and-paste-only-work-if-i-activate-the-sheet
- https://stackoverflow.com/questions/27066963/scraping-data-from-website-using-vba
- Hướng dẫn scrape data website bằng VBA: https://codingislove.com/weather-app-in-excel-vba/
- Scrape data từ XML không cần thư viện, JSON thì cần thư viện từ https://github.com/VBA-tools/VBA-JSON/blob/master/JsonConverter.bas
- Get value từ 1 cell trong Excel vào Advanced Editor trong Power Query: Vào Formula > Define name cho cell đó (Ví dụ:"Cookie") sau đó sửa giá trị cần lấy vào Advanced Editor như sau:
```
Table.ToList(Excel.CurrentWorkbook(){[Name = "Cookie"]}[Content]){0}
```
- Reading text in SAP table: https://stackoverflow.com/questions/68685911/reading-text-in-table-control
- Kết hợp hàm MATCH & INDEX trong Excel:
```
=INDEX("Tọa độ cột kết quả",MATCH(Criteria,"Tọa độ cột chứa Criteria",0))
```
- Tạo lịch trong Excel, cell A1 nhập năm, cell A2 nhập:
```
=LET(x,ROW(1:13)-1,y,COLUMN(A1:AQ1),IF(x=0,IF(y=1,"",TEXT("1/1/2018"+y-2,"DDD")),IF(y=1,TEXT(DATE(A1,x,y),"mmmm"),TEXT(WORKDAY.INTL(DATE($A$1,x,1)+1,-1,"1111110")-2+y,"[<"&DATE($A$1,x,1)&"]"""";[>"&DATE($A$1,x+1,0)&"]"""";DD"))))
```
- Tính thuế TNCN:
```
=+ROUND(SUM(TEXT(C8-{0,5,10,18,32,52,80}*10^6,"0;\0")*5%),0)
```
- Tự động đánh STT, đếm lại khi sang item khác:
```
=+COUNTIF($B$6:B6,B6)
```
