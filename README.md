# Tổng hợp các code VBA chạy Excel
1. Khi replace hay remove các khoảng trống trong file txt bằng VBA, khoảng trống đó có thể không thực sự là khoảng trống thông thường mà là ký tự do ấn Tab (vbTab).
2. Vì sao phải Activate sheet rồi mới copy được: https://stackoverflow.com/questions/41384018/vba-copy-and-paste-only-work-if-i-activate-the-sheet
3. https://stackoverflow.com/questions/27066963/scraping-data-from-website-using-vba
4. Hướng dẫn scrape data website bằng VBA: https://codingislove.com/weather-app-in-excel-vba/
5. Scrape data từ XML không cần thư viện, JSON thì cần thư viện từ https://github.com/VBA-tools/VBA-JSON/blob/master/JsonConverter.bas
6. Get value từ 1 cell trong Excel vào Advanced Editor trong Power Query: Vào Formula > Define name cho cell đó (Ví dụ:"Cookie") sau đó sửa giá trị cần lấy vào Advanced Editor như sau: Table.ToList(Excel.CurrentWorkbook(){[Name = "Cookie"]}[Content]){0}
