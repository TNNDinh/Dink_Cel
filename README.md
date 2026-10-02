# DinkCel

Ứng dụng bảng tính desktop cho Windows. Chạy trực tiếp bằng file DinkCel.exe; không cần trình duyệt, Node.js hay máy chủ web.

## Tải bản chạy

Vào [Releases](https://github.com/TNNDinh/Dink_Cel/releases), chọn phiên bản mới nhất và tải **DinkCel.exe** trong mục Assets. Mỗi phiên bản có ghi chú những tính năng đã có và thay đổi của bản đó. Không cần tải mã nguồn để sử dụng ứng dụng.

## Chạy project

Nhấp đúp DinkCel.exe. Máy cần .NET Framework của Windows.

Nếu sửa mã nguồn, chạy build.cmd để tạo lại DinkCel.exe. Máy build cần trình biên dịch C# của .NET Framework.

## Tính năng v0.2.0

- Nhiều sheet: thêm, đổi tên, xóa và chuyển sheet bằng các tab ở cuối cửa sổ. Undo/Redo được giữ riêng cho từng sheet khi chuyển tab.
- Mở, lưu và **Lưu thành** `.xlsx` trực tiếp trên máy; hỗ trợ nhiều sheet, công thức, kiểu số, màu chữ/nền, căn lề, merge, kích thước hàng/cột, freeze, filter và quy tắc tô màu có điều kiện dạng “lớn hơn”. Vẫn mở được `.dinkcel` cũ và CSV.
- Menu **Dữ liệu** có sắp xếp tăng/giảm theo cột đang chọn, lọc theo nội dung, bỏ lọc, tìm và thay thế. Nếu chọn một vùng nhiều hàng, sắp xếp chỉ áp dụng cho vùng hàng đó.
- Menu **Ô** có định dạng số (`N2`, `P1`, `C2`, `0.00`...), gộp/bỏ gộp, AutoFit hàng/cột, cố định/bỏ cố định, tô màu có điều kiện và dán đặc biệt (chỉ giá trị hoặc chỉ định dạng). Vẫn kéo mép tiêu đề để chỉnh kích thước bằng tay.
- Bổ sung công thức tham chiếu sheet (`Sheet2!A1`, `'Tên sheet'!A1`), ROUND/ROUNDUP/ROUNDDOWN, ABS, SQRT, INT, POWER, MOD, COUNTA, MEDIAN, AND, OR, NOT, LEN, LEFT, RIGHT, UPPER, LOWER, TRIM, CONCAT, COUNTIF và SUMIF.

Giới hạn hiện tại: mỗi sheet tối đa 200 hàng × 26 cột. `.xlsx` tập trung vào dữ liệu và các định dạng nêu trên; chart, pivot, macro và các tính năng Excel nâng cao chưa được giữ khi lưu lại. CSV chỉ chứa một sheet và dữ liệu ô; nếu workbook có nhiều sheet, hãy lưu thành `.xlsx` hoặc `.dinkcel`.
Lệnh **Dán chỉ định dạng** dùng vùng vừa sao chép trong DinkCel; **Dán chỉ giá trị** hỗ trợ cả văn bản từ ứng dụng khác.

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

File `.dinkcel` là định dạng riêng của DinkCel; `.xlsx` dùng để trao đổi với Excel và Google Sheets.

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
5. Bổ sung biểu đồ, in và khả năng giữ nhiều thành phần Excel nâng cao.

## Cấu trúc

- DinkCel.exe: ứng dụng để mở trực tiếp.
- DinkCel.ico: icon được nhúng vào file .exe và dùng trên thanh tiêu đề. icon_dink_cell.png là ảnh gốc; chạy `powershell.exe -NoProfile -ExecutionPolicy Bypass -File make_icon.ps1` trong thư mục project để tạo lại .ico khi thay ảnh.
- DesktopApp.cs: mã nguồn Windows Forms.
- SpreadsheetFeatures.cs: các thao tác sheet, dữ liệu và định dạng trong giao diện.
- XlsxFile.cs và XlsxStyles.cs: đọc/ghi `.xlsx`.
- FormulaEngine.cs: phân tích và tính công thức, dịch tham chiếu khi kéo ô.
- ThemePalette.cs: các mẫu màu và hộp chọn giao diện.
- build.cmd: lệnh build lại ứng dụng.
- test.cmd: kiểm tra công thức, CSV, nhiều sheet và `.xlsx`.

## Phát hành phiên bản mới

1. Cập nhật mã nguồn, số phiên bản trong `VERSION` và `VersionInfo.cs`.
2. Viết `releases/vX.Y.Z.md` nêu rõ phiên bản mới làm được gì, cách tải và giới hạn còn lại.
3. Chạy `test.cmd` và `build.cmd`, rồi commit và đẩy mã nguồn lên `main`.
4. Tạo và đẩy tag `vX.Y.Z`. GitHub Actions sẽ kiểm tra phiên bản, chạy test, build Windows và tạo Release có file **DinkCel.exe** để tải trực tiếp.

Đây là project riêng tại D:\projects\DinkCel, không nằm trong project Unity.
