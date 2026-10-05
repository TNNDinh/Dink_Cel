# DinkCel

Ứng dụng bảng tính desktop cho Windows. Chạy trực tiếp bằng file DinkCel.exe; không cần trình duyệt, Node.js hay máy chủ web.

## Tải bản chạy

Vào [Releases](https://github.com/TNNDinh/Dink_Cel/releases), chọn phiên bản mới nhất và tải **DinkCel.exe** trong mục Assets. Mỗi phiên bản có ghi chú những tính năng đã có và thay đổi của bản đó. Không cần tải mã nguồn để sử dụng ứng dụng.

## Chạy project

Nhấp đúp DinkCel.exe. Máy cần .NET Framework của Windows.

Nếu sửa mã nguồn, chạy build.cmd để tạo lại DinkCel.exe. Máy build cần trình biên dịch C# của .NET Framework.

## Tính năng v0.9.0 — Chart & Printing

- **Biểu đồ:** Column, Line, Pie, Bar, Area, Scatter, Stacked, 100% Stacked và Combo. Biểu đồ hiển thị nổi trên sheet; kéo tiêu đề để di chuyển, kéo góc để đổi kích thước, nhấp đúp để xem và sửa. Có thể sao chép và xuất PNG.
- **Trình sửa biểu đồ:** đổi tiêu đề, loại, vùng dữ liệu, tên trục, chú giải, nhãn dữ liệu, đường lưới, series, tên và màu của từng series, vị trí cùng kích thước.
- **In:** menu **Tệp > Thiết lập trang in** chọn hướng, khổ A4/Letter/Legal/A3, lề, tỷ lệ, vừa một trang, hàng tiêu đề lặp, vùng in, đầu/chân trang, đường lưới và trang biểu đồ. Có lệnh đặt vùng in, ngắt trang, xem trước, in vùng chọn và xuất PDF vùng chọn. Đầu/chân trang nhận `&F` (tên sheet), `&P` (số trang), `&N` (tổng trang), `&D` (ngày).
- `.dinkcel` giữ đầy đủ thiết lập biểu đồ và in. `.xlsx` trao đổi các kiểu biểu đồ mới, vị trí biểu đồ, vùng in và thiết lập trang thông dụng; PDF xuất bảng và biểu đồ thành các trang riêng.

### Tính năng từ v0.8.0 — Table & Pivot

- **Table:** chọn vùng dữ liệu rồi dùng menu **Dữ liệu > Tạo Table**. Có tên riêng, kiểu màu, hàng tiêu đề, hàng tổng, sọc hàng/cột, bộ lọc tại tiêu đề, tự mở rộng khi nhập thêm và cột công thức. Công thức hỗ trợ `=SUM(Table1[Doanh thu])` và `=[@Doanh thu]`; đổi tên Table cập nhật các công thức liên quan.
- **Pivot:** cấu hình tối đa ba trường Rows, hai Columns, ba Values và hai Filters. Hỗ trợ Sum, Count, Average, Min, Max, tổng cuối, tổng nhóm, sắp xếp, lọc, thu gọn/mở rộng nhóm, làm mới và nhóm ngày theo ngày/tháng/năm. Pivot lấy nguồn từ Table sẽ theo vùng dữ liệu mới sau khi Table mở rộng.
- **Dữ liệu lớn:** lưới mở rộng dần tới 50.000 hàng × 26 cột; đã kiểm tra nhập 20.000 dòng CSV, sửa, tổng hợp Pivot và lưu `.xlsx`.
- **Lưu tệp:** `.dinkcel` giữ cấu hình Table và Pivot để chỉnh tiếp. `.xlsx` ghi Excel Table và xuất kết quả Pivot thành sheet tổng hợp; chưa tạo Pivot cache gốc của Excel.

### Tính năng từ v0.7.1 — Data Tools

- **Sort:** tăng/giảm dần và sắp xếp tối đa ba cấp theo chữ, số, ngày hoặc màu nền ô. Giữ nguyên nội dung, kiểu và dữ liệu phụ của ô khi đổi hàng; hàng tiêu đề đầu tiên được giữ lại.
- **Filter:** lọc nhiều cột theo chữ, số hoặc ngày; hỗ trợ Contains, Begins With, Equals, Greater, Less, Between, Blank và Nonblank. Có thể bỏ toàn bộ bộ lọc từ menu **Dữ liệu**.
- **Find & Replace:** `Ctrl+F`/`Ctrl+H` mở hộp thoại có Tìm tiếp, Tìm tất cả, Thay và Thay tất cả. Chọn trang hiện tại hoặc cả workbook, giá trị hiển thị hoặc công thức gốc; có phân biệt hoa thường, khớp toàn ô và wildcard `*`/`?`. Khi tìm trong giá trị hiển thị, công thức chỉ được tìm; muốn thay công thức hãy chọn chế độ công thức gốc.
- **Data Validation:** danh sách, số nguyên, số thập phân, ngày, giờ, độ dài văn bản và công thức tùy chỉnh. Có điều kiện so sánh/khoảng, cho phép ô trống, gợi ý nhập và cảnh báo lỗi Stop/Warning/Information. Chọn vùng rồi dùng menu **Chèn > Kiểm tra dữ liệu**.
- **Conditional Formatting:** Equal, Greater, Less, Between, Duplicate, Unique, Text Contains, Blank, Formula, Color Scale, Data Bar và Icon Set. Chọn vùng rồi dùng menu **Ô > Định dạng có điều kiện**; có lệnh xóa quy tắc trên vùng chọn.
- Các điều kiện được lưu trong `.dinkcel`; `.xlsx` trao đổi các kiểu lọc, kiểm tra dữ liệu và định dạng có điều kiện thông dụng.

### Tính năng từ v0.7.0 — Giao diện DinkCel

- Giao diện desktop gọn hơn với thanh lệnh ngắn, nút theo ngữ cảnh, thanh định dạng nhỏ khi chọn vùng bằng chuột, bảng Inspector bên phải và sáu theme Light, Dark, Midnight, Paper, Solar, Mint. Lưới và biểu đồ đổi màu theo theme.
- `Ctrl+K` tìm lệnh, sheet, vùng có tên, Table, biểu đồ và hàm. `Ctrl+Shift+P` đi nhanh tới ô, sheet, vùng có tên, Table, biểu đồ hoặc ô chứa nội dung cần tìm. `Ctrl+Shift+I` mở/đóng Inspector; `Ctrl+Shift+M` bật/tắt chế độ tối giản. `Ctrl+P` vẫn dùng để in.
- Màn hình bắt đầu có bảng tính mới, mở tệp và danh sách tệp gần đây. Formula Bar tô màu tên hàm và tham chiếu khi xem công thức. Thanh trạng thái hiển thị số ô, tổng, trung bình, nhỏ nhất, lớn nhất và thanh zoom.
- Tab sheet có thể kéo đổi thứ tự, nhân bản, sao chép, ẩn/hiện và chọn màu. Trạng thái ẩn được lưu trong `.dinkcel`, `.xlsx`, `.xls`, `.ods`; màu tab được lưu trong `.dinkcel` và `.xlsx`.

### Tính năng từ v0.6.0 — Formatting

- **Font:** chọn họ font và cỡ chữ; in đậm, nghiêng, gạch chân, gạch ngang và màu chữ. Hộp chọn font cho phép đặt cỡ bất kỳ trong phạm vi hỗ trợ.
- **Ô:** màu nền, viền từng cạnh với kiểu và màu, căn ngang/dọc, xuống dòng, thu chữ vừa ô, thụt lề và xoay chữ.
- **Số:** General, Number, Currency, Accounting, Percentage, Date, Time, Scientific, Fraction và định dạng tùy chỉnh. Ngày giờ từ Excel hiển thị theo định dạng ô.
- **Công cụ:** Chổi định dạng trên thanh công cụ; Xóa định dạng, Xóa nội dung, Xóa toàn bộ. Có thể ẩn/hiện, đặt kích thước và AutoFit hàng/cột.
- **Tương thích:** lưu/mở các thuộc tính định dạng trên trong `.dinkcel`, `.xlsx` và phần lớn thuộc tính trong `.xls`; dán bảng từ Excel qua clipboard HTML giữ font, màu, căn lề, viền và định dạng số thông dụng.

### Tính năng từ v0.5.0 — Formula Engine

- **Tra cứu:** `XLOOKUP`, `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH` với khớp chính xác và gần đúng; `XLOOKUP` hỗ trợ tìm từ cuối và ký tự đại diện.
- **Điều kiện:** `IFERROR`, `IFNA`, `SUMIFS`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`; `COUNTIF` và `SUMIF` hỗ trợ điều kiện so sánh, ký tự đại diện.
- **Ngày giờ và kiểm tra:** `DATE`, `TIME`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `WEEKNUM`, `EOMONTH`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR`, `ISNA`.
- **Tính toán:** mã lỗi `#N/A`, `#VALUE!`, `#REF!`, `#DIV/0!`, `#NAME?`, `#NUM!`; vòng tham chiếu hiển thị `#CYCLE!`. Quan hệ phụ thuộc và thứ tự tính được theo dõi để chỉ tính lại công thức liên quan khi sửa một ô.
- **Tham chiếu:** công thức liên sheet hỗ trợ ô và vùng, kể cả tên sheet có dấu cách. Tham chiếu tương đối, tuyệt đối và hỗn hợp (`$A1`, `A$1`, `$A$1`) được giữ đúng khi kéo điền.

### Tính năng từ v0.4.0 — Excel Editing

- **Điều hướng:** Ctrl + mũi tên nhảy tới mép vùng dữ liệu; giữ Shift để mở rộng vùng chọn. Shift + mũi tên chỉnh vùng chọn từng ô. Ctrl + Home/End tới đầu bảng hoặc ô cuối có dữ liệu. F2 sửa ô; Enter, Shift + Enter, Tab và Shift + Tab chuyển ô; Esc hủy sửa ô hoặc thao tác cắt.
- **Chọn ô:** Ctrl + Space chọn cả cột, Shift + Space chọn cả hàng, Ctrl + A chọn toàn sheet. Giữ Ctrl và nhấp ô để chọn nhiều vùng không liền nhau. Name Box nhận địa chỉ (`A1`), vùng (`A1:C5`), nhiều vùng (`A1:C5,E1:E3`) hoặc tên vùng đã đặt.
- **Formula Bar:** hiển thị địa chỉ ô hiện tại và công thức gốc. Sửa trên thanh công thức, nhấn Enter để lưu hoặc Esc để hủy.
- **Điền dữ liệu:** kéo Fill Handle ở góc dưới bên phải vùng chọn để điền số, ngày hoặc sao chép công thức với tham chiếu được dịch. Ctrl + D điền xuống, Ctrl + R điền sang phải; menu **Chỉnh sửa > Điền chuỗi tăng** tạo dãy theo bước số hoặc ngày.
- **Clipboard:** Ctrl + C/X/V sao chép, cắt và dán nội dung cùng định dạng. Ctrl + Shift + V dán giá trị; menu **Chỉnh sửa** còn có dán công thức, định dạng và chuyển vị. Cắt/dán giữa các sheet có thể hoàn tác và làm lại.

### Tính năng từ v0.3.0

- **Table:** chọn vùng gồm hàng tiêu đề và dữ liệu, rồi tạo Table trong menu **Dữ liệu**. DinkCel tô màu tiêu đề và các hàng xen kẽ; khi lưu `.xlsx`, Table được ghi thành Excel Table thực.
- **Charts:** tạo biểu đồ cột, đường hoặc tròn từ vùng có cột đầu làm nhãn. Xem biểu đồ trong ứng dụng, lưu ảnh PNG; `.xlsx` giữ biểu đồ để mở trong Excel/LibreOffice.
- **Data validation/dropdown:** gán danh sách lựa chọn cho một vùng ô; giá trị nhập sai bị từ chối. `.xlsx` giữ danh sách nhập trực tiếp.
- **Pivot Table cơ bản:** chọn vùng có hàng tiêu đề, cột nhóm và cột số; tạo sheet tổng hợp theo Sum hoặc Count, rồi dùng **Làm mới Pivot** khi dữ liệu nguồn đổi.
- **In/PDF:** xem trước, in sheet hiện tại, hoặc xuất toàn bộ workbook ra PDF.
- **Named ranges:** đặt tên một vùng và dùng tên đó trong công thức, ví dụ `=SUM(DoanhThu)`; có thể đi tới hoặc xóa vùng có tên. Công thức liên sheet dùng `=Sheet2!A1` hoặc `='Tên sheet'!A1`.
- **Định dạng file:** mở/lưu `.dinkcel`, `.xlsx`, `.xls`, `.ods` và CSV. `.dinkcel` giữ metadata DinkCel; `.xlsx` trao đổi Table, chart, dropdown và named ranges.

### Tính năng từ v0.2.0

- Nhiều sheet; mở/lưu `.xlsx`; sắp xếp, lọc, tìm/thay thế; định dạng số, merge, AutoFit, freeze panes, conditional formatting và Paste Special.
- Các hàm số, văn bản, điều kiện và thống kê cơ bản; công thức liên sheet.

**Giới hạn:** mỗi sheet có tối đa 50.000 hàng × 26 cột. File có dữ liệu ngoài vùng này sẽ báo lỗi khi mở. CSV chỉ lưu một sheet và nội dung ô. Pivot trong `.xlsx` là sheet tổng hợp và cần làm mới trong DinkCel; chưa phải Pivot Table gốc của Excel. `.xls` và `.ods` tập trung vào dữ liệu, công thức, nhiều sheet, merge và named ranges; các thành phần nâng cao và một số kiểu định dạng có thể không được giữ khi lưu lại. Bản in và PDF chưa có biểu đồ. Với file có macro hoặc tính năng Excel nâng cao, hãy giữ bản gốc trước khi chỉnh sửa và lưu lại.

### Tính năng từ v0.1.1

- Lưới 200 hàng × 26 cột, địa chỉ ô từ A1 đến Z200; có cuộn dọc và ngang.
- Nhập dữ liệu trực tiếp vào ô hoặc qua thanh nội dung phía trên; địa chỉ ô đang chọn hiện bên trái.
- Bấm số hàng hoặc chữ cột để chọn nhanh toàn bộ hàng/cột. Nhấp chuột phải lên tiêu đề để chèn hoặc xóa hàng/cột; các lệnh này cũng có trong menu Chỉnh sửa.
- Giữ chuột trái trên tiêu đề hàng/cột rồi kéo đến vị trí mới để di chuyển cả hàng/cột. Nội dung, định dạng, kích thước và tham chiếu công thức được cập nhật; Ctrl+Z/Ctrl+Y hoàn tác hoặc làm lại thao tác kéo.
- Khi chèn/xóa, nội dung, định dạng và tham chiếu công thức được dịch theo vị trí mới. Bảng vẫn giới hạn 200 hàng × 26 cột; nếu hàng 200 hoặc cột Z đang chứa dữ liệu/định dạng, lệnh chèn sẽ dừng để tránh mất dữ liệu.
- Sao chép và dán vùng ô dạng bảng bằng Ctrl+C / Ctrl+V.
- Chọn nhiều ô rồi nhấn Delete để xóa toàn bộ nội dung vùng chọn. Ctrl+Z hoàn tác và Ctrl+Y làm lại thao tác nhập, xóa, dán, định dạng, kéo điền hoặc chèn/xóa hàng cột (tối đa 50 bước).
- Kéo ô vuông xanh ở góc dưới bên phải ô đang chọn để sao chép xuống, lên hoặc sang ngang. Công thức được dịch tham chiếu theo vị trí mới; tham chiếu có dấu $ được giữ cố định.
- Công thức hỗ trợ +, -, *, /, ^, dấu ngoặc, tham chiếu ô/vùng và các hàm SUM, AVERAGE, MIN, MAX, COUNT, IF. Ô phụ thuộc tự tính lại khi dữ liệu thay đổi.
- Nút **Giao diện** mở bảng chọn 6 mẫu màu có hình xem trước: Xanh dịu, Bạc hà, Kem ấm, Tím sương, Xám xanh và Tối dịu. Mẫu màu áp dụng cho cả cửa sổ, thanh công cụ và bảng tính.
- Giao diện đã chọn được ghi nhớ trong appearance.xml cạnh DinkCel.exe và lưu cùng file .dinkcel.
- Thanh công cụ có cỡ chữ, in đậm/nghiêng/gạch chân, màu chữ, màu nền ô và căn lề. Có thể chọn nhiều ô rồi định dạng cùng lúc.
- Menu Tệp có Mới, Mở, Lưu và Lưu thành. Ctrl+S lưu bảng tính.
- Mở CSV bằng menu Tệp > Mở hoặc nhấp phải file CSV trong Windows > Open with > chọn DinkCel.exe. Hỗ trợ CSV UTF-8, UTF-16 có BOM và mã ANSI của Windows; ô có dấu phẩy, dấu ngoặc kép hoặc xuống dòng được đọc đúng.
- CSV mở trong DinkCel được lưu trực tiếp vào chính file CSV đó bằng Lưu hoặc Ctrl+S. Lưu thành cho phép tạo bản CSV khác hoặc file .dinkcel. CSV chỉ chứa dữ liệu ô và công thức, không lưu màu hay định dạng; muốn giữ định dạng hãy dùng .dinkcel. Bảng hiện có giới hạn 200 hàng × 26 cột; file CSV vượt giới hạn sẽ hiện lỗi rõ ràng.
- Dữ liệu và định dạng được lưu thành file .dinkcel trên máy và có thể mở lại. File từ bản 20 × 10 trước vẫn mở được.

Ví dụ: nhập 10 vào A1, 20 vào A2, rồi nhập =SUM(A1:A2) vào A3; A3 sẽ hiển thị 30. Nhập =IF(A1>0,"Có","Không") để thử hàm điều kiện. Dùng dấu chấm cho số thập phân trong công thức. Có thể ngăn cách đối số hàm bằng dấu phẩy hoặc chấm phẩy.

File `.dinkcel` là định dạng riêng của DinkCel; `.xlsx`, `.xls` và `.ods` dùng để trao đổi với Excel và LibreOffice.

## Một ứng dụng Excel cơ bản cần gì?

| Mức | Tính năng | Việc cần làm |
| --- | --- | --- |
| Nền tảng | Hàng, cột, ô | Hiển thị bảng, địa chỉ như A1, chọn và sửa ô. |
| Nền tảng | Kiểu dữ liệu | Phân biệt văn bản, số, ngày và ô trống. |
| Nền tảng | Điều hướng | Di chuyển bằng phím, chọn một vùng ô, sao chép/dán. |
| Nền tảng | Lưu và mở | Lưu bảng, mở lại; hỗ trợ nhập/xuất CSV. |
| Quan trọng | Công thức | Bổ sung thêm hàm, định dạng số và ngày. |
| Quan trọng | Tính lại | Tối ưu tốc độ khi bảng có rất nhiều công thức. |
| Quan trọng | Định dạng | In đậm, màu chữ/nền, căn lề, định dạng số và ngày. |
| Quan trọng | Thao tác bảng | Thêm/xóa hàng cột, chỉnh độ rộng, hoàn tác/làm lại. |
| Mở rộng | Nhiều sheet | Tạo, đổi tên, chuyển và xóa sheet. |
| Mở rộng | Công cụ dữ liệu | Tìm kiếm, sắp xếp, lọc và biểu đồ cơ bản. |
| Mở rộng | Tệp Excel | Đọc/ghi `.xlsx` với nhiều sheet và các định dạng cơ bản. |

## Thứ tự tự làm gợi ý

1. Đọc DesktopApp.cs: SpreadsheetForm tạo giao diện, đọc/ghi file và xử lý thao tác với ô.
2. Mở rộng nhập/xuất CSV với tùy chọn dấu phân cách và mã hóa theo từng file.
3. Đọc FormulaEngine.cs và thử viết thêm hàm hoặc công thức liên sheet.
4. Mở rộng định dạng ngày và tối ưu bảng lớn hơn 200 × 26.
5. Nghiên cứu cách lưu Pivot Table gốc và giữ thêm thành phần nâng cao khi trao đổi file Excel/ODS.

## Cấu trúc

- DinkCel.exe: ứng dụng để mở trực tiếp.
- DinkCel.ico: icon được nhúng vào file .exe và dùng trên thanh tiêu đề. icon_dink_cell.png là ảnh gốc; chạy `powershell.exe -NoProfile -ExecutionPolicy Bypass -File make_icon.ps1` trong thư mục project để tạo lại .ico khi thay ảnh.
- DesktopApp.cs: mã nguồn Windows Forms.
- SpreadsheetFeatures.cs: các thao tác sheet, dữ liệu và định dạng trong giao diện.
- XlsxFile.cs, XlsxStyles.cs, XlsxCharts.cs và XlsxChartsV9.cs: đọc/ghi `.xlsx`, biểu đồ và thiết lập trang.
- ChartRendering.cs, ChartEditorUi.cs và ChartOverlayUi.cs: vẽ, chỉnh sửa và thao tác biểu đồ trên sheet.
- PrintLayout.cs, PrintSetupUi.cs, SpreadsheetOutputUi.cs và PdfFile.cs: bố cục trang, xem trước, in và PDF.
- XlsFile.cs, OdsFile.cs, PdfFile.cs: đọc/ghi `.xls`, `.ods` và xuất PDF.
- WorkbookFeatures.cs, SpreadsheetV3Ui.cs và SpreadsheetOutputUi.cs: metadata và giao diện v0.3.
- SpreadsheetEditingUi.cs: điều hướng, Name Box, Formula Bar, vùng chọn, Fill và clipboard.
- SpreadsheetFormatting.cs: công cụ định dạng, viền ô và đọc định dạng từ clipboard HTML của Excel.
- vendor/: thư viện NPOI, SharpZipLib, PDFsharp cùng giấy phép; nội dung giấy phép được nhúng trong `.exe` và xem qua menu **Trợ giúp**.
- FormulaEngine.cs và FormulaEngineV5.cs: phân tích, tính công thức, theo dõi quan hệ phụ thuộc và dịch tham chiếu khi kéo ô.
- ThemePalette.cs: các mẫu màu và hộp chọn giao diện.
- build.cmd: lệnh build lại ứng dụng.
- test.cmd: kiểm tra công thức, CSV, nhiều sheet, `.xlsx`, `.xls`, `.ods`, PDF và giao diện v0.3.

## Phát hành phiên bản mới

1. Cập nhật mã nguồn, số phiên bản trong `VERSION` và `VersionInfo.cs`.
2. Viết `releases/vX.Y.Z.md` nêu rõ phiên bản mới làm được gì, cách tải và giới hạn còn lại.
3. Chạy `test.cmd` và `build.cmd`, rồi commit và đẩy mã nguồn lên `main`.
4. Tạo và đẩy tag `vX.Y.Z`. GitHub Actions sẽ kiểm tra phiên bản, chạy test, build Windows và tạo Release có file **DinkCel.exe** để tải trực tiếp.

Đây là project riêng tại D:\projects\DinkCel, không nằm trong project Unity.
