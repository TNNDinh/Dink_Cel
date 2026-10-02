# DinkCel

Ứng dụng bảng tính desktop cho Windows. Chạy trực tiếp bằng file DinkCel.exe; không cần trình duyệt, Node.js hay máy chủ web.

## Tải bản chạy

Vào [Releases](https://github.com/TNNDinh/Dink_Cel/releases), chọn phiên bản mới nhất và tải **DinkCel.exe** trong mục Assets. Mỗi phiên bản có ghi chú những tính năng đã có và thay đổi của bản đó. Không cần tải mã nguồn để sử dụng ứng dụng.

## Chạy project

Nhấp đúp DinkCel.exe. Máy cần .NET Framework của Windows.

Nếu sửa mã nguồn, chạy build.cmd để tạo lại DinkCel.exe. Máy build cần trình biên dịch C# của .NET Framework.

## Tính năng v0.3.0

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

**Giới hạn:** mỗi sheet có 200 hàng × 26 cột. File có dữ liệu ngoài vùng này sẽ báo lỗi khi mở. CSV chỉ lưu một sheet và nội dung ô. Pivot là sheet tổng hợp phải làm mới bằng lệnh, chưa phải Pivot Table gốc của Excel. `.xls` và `.ods` tập trung vào dữ liệu, công thức, nhiều sheet, merge và named ranges; các thành phần nâng cao và một số kiểu định dạng có thể không được giữ khi lưu lại. Bản in và PDF chưa có biểu đồ. Với file có macro hoặc tính năng Excel nâng cao, hãy giữ bản gốc trước khi chỉnh sửa và lưu lại.

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
- XlsxFile.cs, XlsxStyles.cs và XlsxCharts.cs: đọc/ghi `.xlsx` và biểu đồ.
- XlsFile.cs, OdsFile.cs, PdfFile.cs: đọc/ghi `.xls`, `.ods` và xuất PDF.
- WorkbookFeatures.cs, SpreadsheetV3Ui.cs và SpreadsheetOutputUi.cs: metadata và giao diện v0.3.
- vendor/: thư viện NPOI, SharpZipLib, PDFsharp cùng giấy phép; nội dung giấy phép được nhúng trong `.exe` và xem qua menu **Trợ giúp**.
- FormulaEngine.cs: phân tích và tính công thức, dịch tham chiếu khi kéo ô.
- ThemePalette.cs: các mẫu màu và hộp chọn giao diện.
- build.cmd: lệnh build lại ứng dụng.
- test.cmd: kiểm tra công thức, CSV, nhiều sheet, `.xlsx`, `.xls`, `.ods`, PDF và giao diện v0.3.

## Phát hành phiên bản mới

1. Cập nhật mã nguồn, số phiên bản trong `VERSION` và `VersionInfo.cs`.
2. Viết `releases/vX.Y.Z.md` nêu rõ phiên bản mới làm được gì, cách tải và giới hạn còn lại.
3. Chạy `test.cmd` và `build.cmd`, rồi commit và đẩy mã nguồn lên `main`.
4. Tạo và đẩy tag `vX.Y.Z`. GitHub Actions sẽ kiểm tra phiên bản, chạy test, build Windows và tạo Release có file **DinkCel.exe** để tải trực tiếp.

Đây là project riêng tại D:\projects\DinkCel, không nằm trong project Unity.
