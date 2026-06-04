# Tóm tắt thay đổi Git

## Tổng quan

Project đã được nâng cấp từ cơ chế chỉ thêm dữ liệu vào file Excel local thành quy trình đồng bộ kiểu upsert vào một file Excel có sẵn trên OneDrive.

File OneDrive đích được xác định bằng `ONEDRIVE_FILE_ID`. Tool sẽ tải workbook hiện tại, thêm các khoản thanh toán mới, cập nhật những dòng đã tồn tại khi dữ liệu thay đổi, sau đó chỉ tải workbook lên lại OneDrive khi thực sự có thay đổi.

## Các file đã thay đổi

### `.env.example`

Đã bổ sung các biến cấu hình OneDrive/Microsoft Graph:

- `ONEDRIVE_TENANT_ID`
- `ONEDRIVE_CLIENT_ID`
- `ONEDRIVE_CLIENT_SECRET`
- `ONEDRIVE_USER_ID`
- `ONEDRIVE_FILE_ID`

### `.gitignore`

Đã thêm các thư mục phát sinh trong môi trường phát triển:

- `.codegraph`
- `.cursor`

### `package.json`

Đã thêm các lệnh:

```powershell
npm run once
npm run check:onedrive
```

- `npm run once`: build và chạy đúng một lượt đồng bộ, sau đó tự thoát.
- `npm run check:onedrive`: kiểm tra xác thực Microsoft Graph và quyền truy cập file OneDrive đã cấu hình.

### `src/config.ts`

- Bổ sung cấu hình OneDrive tùy chọn.
- Kiểm tra đủ cả năm biến OneDrive khi bật tích hợp.
- Cung cấp cấu hình qua `AppConfig.oneDrive`.

### `src/onedrive.ts`

Module tích hợp Microsoft Graph mới:

- Lấy access token bằng client credentials.
- Kiểm tra quyền truy cập file theo OneDrive item ID.
- Tải workbook hiện tại từ OneDrive.
- Tải workbook đã chỉnh sửa lên lại đúng file ID.

### `src/check-onedrive.ts`

Lệnh chẩn đoán mới, được sử dụng bởi `npm run check:onedrive`.

Lệnh này chỉ kiểm tra file đã cấu hình có tồn tại và truy cập được hay không, không chỉnh sửa file.

### `src/index.ts`

Các thay đổi chính trong luồng đồng bộ:

- Bổ sung chế độ `--once`, vẫn giữ nguyên scheduler chạy mỗi 30 phút.
- Thay Python/openpyxl bằng dependency `exceljs` đã có trong project.
- Đổi tên worksheet thành tên thực tế trong workbook: `VENDOR_PAYMENT`.
- Loại bỏ cơ chế chống trùng dựa trên file JSON `seenIds`.
- Bổ sung cơ chế upsert từng dòng Excel, dùng `settlementNo` làm khóa đối chiếu.
- Tải workbook hiện tại từ OneDrive trước khi xử lý.
- Chỉ tải workbook lên OneDrive khi có dòng được thêm hoặc cập nhật.

## Quy tắc upsert dữ liệu Excel

Các dòng được đối chiếu bằng `settlementNo`:

| Cột Excel | Tiêu đề | Trường API | Cách xử lý |
|---|---|---|---|
| I | `SỐ ĐNTT-FMS` | `settlementNo` | Khóa đối chiếu duy nhất |
| F | `HẠN THANH TOÁN` | `dueDate` | Thêm mới hoặc cập nhật |
| H | `NCC` | `payeeName` | Thêm mới hoặc cập nhật |
| M | `NGÀY LẬP - FMS` | `requestDate` | Thêm mới hoặc cập nhật |

Các cột Excel còn lại được giữ nguyên và không bị automation chỉnh sửa.

## Luồng chạy hiện tại

```text
Lấy access token eFMS
→ Lấy danh sách settlement payment
→ Tải workbook hiện tại từ OneDrive
→ Đối chiếu dòng theo settlementNo
→ Thêm các dòng chưa tồn tại
→ Cập nhật các trường được quản lý khi dữ liệu thay đổi
→ Chỉ tải workbook lên lại OneDrive khi có thay đổi
```

## Kết quả kiểm tra

- `npm run lint`: thành công
- `npm run build`: thành công
- `npm run check:onedrive`: thành công
- File OneDrive đã cấu hình có thể truy cập với tên `VENDOR_PAYMENT.xlsx`
- Đồng bộ đầy đủ hiện dừng tại bước lấy token vì server eFMS trả lỗi HTTP 500

## Ghi chú về trạng thái Git hiện tại

File dưới đây là artifact được tạo trong lúc kiểm tra cấu trúc workbook, không phải source code:

```text
data/Vendor_Payment_Inspect.xlsx
```

Nên xóa hoặc loại khỏi commit nếu không cần lưu.

File `.env` local chứa thông tin xác thực runtime và đã được Git bỏ qua.

