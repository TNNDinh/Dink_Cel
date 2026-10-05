# DinkCel Script — thiết kế cho ứng dụng desktop

> **Trạng thái v1.1.0:** Editor một tệp, hàm tùy chỉnh, API đọc/ghi vùng và kết nối AI kiểu Chat Completions đã có. Các mục bên dưới về nhiều tệp, trigger, quyền file/mạng tùy biến, cache bền vững, quota chi phí và chuyển nền tảng build là định hướng tiếp theo, chưa có trong v1.1.0. Runtime hiện dùng Jint nhúng trong một EXE .NET Framework 4.6.2+.

Với hàm AI trong ô ở v1.1, dùng hàm đồng bộ và gọi `AI.generate(prompt)`. Ô tạm hiển thị `#BUSY!` trong khi request chạy nền; không khai báo `async` hoặc `await` cho custom function. Mỗi request hiển thị prompt để người dùng duyệt.

## Mục tiêu

Thêm trình viết JavaScript trong DinkCel.exe để tự động hóa workbook và tạo hàm riêng dùng trong ô. Script chạy cục bộ trên máy; không cần trình duyệt hay máy chủ. Lệnh gọi model AI là tùy chọn và chỉ thực hiện khi người dùng cấu hình endpoint, cấp quyền mạng và chủ động chạy hoặc dùng hàm AI.

## Giao diện

- Menu **Công cụ > Script Editor** mở cửa sổ riêng: danh sách tệp `.js`, vùng soạn thảo có tô cú pháp, nút Run/Stop, chọn hàm, console và lỗi kèm dòng/cột.
- Tab **Functions** liệt kê các hàm có thể gọi từ ô. Tab **Permissions** hiển thị quyền của script. Tab **AI** cấu hình model, endpoint, giới hạn chi phí và khóa API.
- Thanh trạng thái khi script chạy: tên hàm, thời gian, số ô đã sửa; có nút dừng. Mọi thay đổi vào workbook là một bước Undo.
- Mỗi workbook lưu mã script trong `.dinkcel` cùng mã định danh project. Khi nhập `.xlsx`, DinkCel không tự chạy macro/script của file. Khi xuất `.xlsx`, script DinkCel không được ghi vào file; hộp thoại lưu phải nêu rõ điều này.

## Runtime và API

Runtime nên là **Jint** (JavaScript interpreter trong tiến trình), sau khi chuyển project sang .NET Windows Desktop hiện đại và đóng gói một EXE. Không dùng `eval` của Windows Script Host hoặc chạy mã từ workbook bằng quyền hệ điều hành đầy đủ. Runtime giới hạn thời gian, bộ nhớ, độ sâu gọi và số thao tác; cung cấp API DinkCel được cho phép, không đưa trực tiếp `System.IO`, `Process` hoặc reflection cho script.

API đầu tiên:

```js
const book = DinkCel.getActiveWorkbook();
const sheet = book.getSheetByName("Sales");
const values = sheet.getRange("A2:C100").getValues();
sheet.getRange("D2:D100").setValues(values.map(row => [row[1] * row[2]]));
sheet.getRange("D1").setValue("Thành tiền");
Logger.log(`Đã cập nhật ${values.length} hàng`);
```

`getValues` trả về mảng hai chiều gồm số, chữ, boolean, ngày hoặc `null`; `setValues` yêu cầu kích thước khớp. Bổ sung `getFormulas`, `setFormulas`, `getBackgrounds`, `setBackgrounds`, `getLastRow`, `getLastColumn`, `getSheets`, `insertSheet`, `flush` theo thứ tự nhu cầu. Lệnh sửa sheet được bảo vệ phải báo lỗi. Các lệnh sửa hàng loạt chỉ phát tín hiệu tính lại một lần khi `flush`/kết thúc.

## Hàm dùng trong ô

```js
/** @customfunction */
function VAT(price, rate) {
  return Number(price) * Number(rate);
}
```

Ô có thể dùng `=VAT(A2, 0.08)`. Hàm tùy chỉnh chỉ đọc dữ liệu được truyền vào, không sửa workbook, file hay mạng. Engine ghi dependency của ô như các hàm công thức khác, cache theo input và phiên bản script, báo lỗi rõ ở ô. Mảng trả về cần kiểm tra vùng tràn trước khi ghi. Script thay đổi phải hủy cache và tính lại ô phụ thuộc.

## Kết nối AI

Cho phép khai báo provider có API kiểu OpenAI: endpoint HTTPS, model, API key, timeout, mức trần token và chi phí. Khóa được mã hóa bằng Windows DPAPI trong hồ sơ người dùng; không lưu vào workbook, log, clipboard hay Git. Không gửi toàn bộ workbook theo mặc định. UI xem trước chính xác prompt và vùng dữ liệu sẽ gửi.

```js
/** @customfunction @network */
async function AI_SUMMARY(text) {
  return await AI.generate({ prompt: `Tóm tắt ngắn: ${text}`, model: "configured" });
}
```

`=AI_SUMMARY(A2)` hiển thị trạng thái đang xử lý, trả kết quả vào ô khi hoàn tất. Không gọi lại API mỗi lần cuộn, đổi format hoặc mở file: cache theo model, prompt và input; nút **Refresh AI results** yêu cầu gọi lại. Có giới hạn số yêu cầu đồng thời và ngân sách mỗi workbook. Khi offline/hết quota, ô báo lỗi có thể hiểu được. Script hoặc file tải về không được tự gửi dữ liệu qua mạng trước khi người dùng cấp quyền.

## Quyền và an toàn dữ liệu

- Script mới: chỉ được đọc và sửa workbook đang mở khi người dùng nhấn Run. Không có quyền file/mạng mặc định.
- Quyền file hoặc mạng phải khai báo và hiện tên miền/đường dẫn trong hộp thoại cấp quyền. Thay đổi mã làm mất quyền đã cấp cho đến khi người dùng xét lại.
- Trigger `onOpen`, `onEdit`, `onChange` mặc định tắt với workbook nhập từ bên ngoài. Khi bật, dùng hàng đợi một luồng, chống gọi lặp, giới hạn thời gian và rollback nếu lỗi.
- Mỗi lần chạy có transaction và Undo; lỗi hoặc bấm Stop khôi phục dữ liệu. Có log kết quả nhưng loại bỏ API key và dữ liệu nhạy cảm.

## Kiểm thử chấp nhận

1. Script sửa 10.000 ô trong một bước Undo, tính lại công thức liên quan và lưu/mở lại `.dinkcel` đúng.
2. Custom function có tham chiếu liên sheet được tính lại khi input hoặc script đổi; vòng lặp trả lỗi rõ.
3. Workbook tải từ ngoài không tự chạy script/AI. Quyền mạng hiển thị dữ liệu sẽ gửi và từ chối thì không có request.
4. API key không xuất hiện trong `.dinkcel`, `.xlsx`, log hoặc bản release.
5. Mở Excel có VBA và lưu qua DinkCel phải cảnh báo và tạo backup trước khi ghi đè.

## Thứ tự triển khai

1. Chuyển build sang .NET Windows Desktop phù hợp với Jint và đóng gói EXE.
2. Tạo Script Editor, runtime hạn chế quyền và API đọc/sửa workbook.
3. Nối custom functions vào dependency graph và Undo transaction.
4. Thêm AI provider, cấp quyền, cache, quota và kết quả bất đồng bộ.
5. Thêm trigger và kiểm thử với workbook lớn.
