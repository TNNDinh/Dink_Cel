# Data import và Data Model trong DinkCel v2.0

## Import có thể làm mới

1. Mở **Data Model > Queries > Import query...**, chọn Sheet, CSV, JSON, XML, WebCsv, WebJson, WebXml hoặc Odbc.
2. Với tệp, chọn đường dẫn. Với web/API, nhập URL GET; token Bearer nếu cần được hỏi lúc tải và không lưu. Với ODBC, nhập tên DSN và một câu `SELECT`. Chuỗi kết nối đầy đủ được hỏi khi làm mới và không lưu.
3. Với nguồn ngoài, chọn **Sheet**, **Model** hoặc **Both**. Mỗi query vào sheet riêng. Nguồn Sheet luôn tải vào Model và tính lại công thức chuẩn của ô; hàm DinkCel Script tùy chỉnh chưa được tính trong luồng này. Trong **Data Model > Queries**, **Add query step...** cho phép thêm Filter, Sort, Rename, Remove, Type hoặc Distinct theo thứ tự; **Remove last query step...** bỏ bước cuối.
4. Dùng **Refresh query...** hoặc **Refresh all queries** trong menu Queries để đọc lại nguồn, chạy các bước và cập nhật dữ liệu. Query lưu tên nguồn và bước biến đổi trong `.dinkcel`. Tệp nguồn cần còn tồn tại ở đường dẫn đã lưu; có thể dùng **Change query source...** để đổi.

JSON hỗ trợ mảng các object hoặc object có mảng `data`/`items`. XML hỗ trợ danh sách các phần tử cùng cấu trúc, lấy thuộc tính với tiền tố `@`. SQL dùng ODBC DSN của Windows và chỉ nhận một câu `SELECT` không có dấu `;`. Web/API dùng HTTP(S) GET; Bearer token chỉ được gửi qua HTTPS hoặc localhost.

## Data Model và Pivot

1. Tải ít nhất một query vào **Model**. Mỗi query tạo một bảng model.
2. Trong **Data Model > Tables and measures > Relationship...**, nối khóa ở bảng fact với khóa duy nhất ở bảng tra cứu. Nếu khóa tra cứu trùng, DinkCel báo lỗi để tránh cộng lặp dữ liệu.
3. **Calculated column...** tạo cột theo từng hàng, ví dụ `={Amount}-{Cost}`.
4. **Measure...** tạo biểu thức theo từng hàng bằng tên trường đầy đủ, ví dụ `={Orders.Amount}`, rồi chọn Sum, Count, Average, Min hoặc Max.
5. Trong **Data Model > Model Pivot**, chọn **Create...** để chọn fact table, trường Rows, tùy chọn Columns và measure. Có thể lọc bằng **Slicer...** hoặc **Timeline...**, tạo **Chart...**, rồi bấm **Refresh all** sau khi dữ liệu đổi.

Các trường từ bảng liên kết có dạng `TenBang.TenCot`, ví dụ `Customers.City`. Quan hệ hiện là một bước nối nhiều hàng fact tới một hàng tra cứu; chưa có DAX, quan hệ nhiều bước hoặc truy vấn Power Query M. Slicer/Timeline là hộp chọn trong menu. DinkCel Script vẫn là ngôn ngữ automation của ứng dụng; không thực thi VBA hay Python.

## Giới hạn và lưu file

- Nguồn tệp/web tối đa 16 MB; bảng import tối đa 100 cột, 100.000 hàng và 1.000.000 ô. Khi tải ra sheet, lưới giới hạn 26 cột, 50.000 hàng và 50.000 ô mỗi lần import. Model Pivot ghi tối đa 100.000 ô.
- `.dinkcel` giữ query, các bước, dữ liệu model, quan hệ, measure và cấu hình Model Pivot. Xuất sang `.xlsx`/`.xls`/`.ods`/CSV chỉ giữ kết quả hiện trên sheet; ứng dụng cảnh báo trước khi xuất.
- ODBC cần driver và DSN cài sẵn trên máy. DinkCel không ghi mật khẩu hoặc token vào workbook; người dùng nhập lại khi làm mới nguồn cần xác thực.
