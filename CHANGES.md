# Tóm Tắt Thay Đổi Git

## Tổng Quan

Project đã được nâng cấp từ cơ chế chỉ thêm dữ liệu vào file Excel local thành quy trình đồng bộ kiểu upsert vào một file Excel có sẵn trên OneDrive.

File OneDrive đích được xác định bằng `ONEDRIVE_FILE_ID`. Tool dùng Microsoft Graph để tải workbook hiện tại về local, so sánh và cập nhật dữ liệu bằng `exceljs`, sau đó upload đè lại đúng file OneDrive đó nếu có thay đổi.

Lưu ý: logic bên trong Excel là upsert từng dòng, nhưng cách lưu lên OneDrive hiện tại là upload lại toàn bộ workbook.

## Các File Đã Thay Đổi

### `.env.example`

Đã bổ sung các biến cấu hình OneDrive/Microsoft Graph:

- `ONEDRIVE_TENANT_ID`
- `ONEDRIVE_CLIENT_ID`
- `ONEDRIVE_CLIENT_SECRET`
- `ONEDRIVE_USER_ID`
- `ONEDRIVE_USER_PRINCIPAL_NAME`
- `ONEDRIVE_DRIVE_ID`
- `ONEDRIVE_REMOTE_PATH`
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
- Kiểm tra đủ cấu hình OneDrive khi bật tích hợp.
- Hỗ trợ định danh file bằng `ONEDRIVE_FILE_ID`.
- Hỗ trợ định danh drive/user bằng `ONEDRIVE_DRIVE_ID`, `ONEDRIVE_USER_ID`, hoặc `ONEDRIVE_USER_PRINCIPAL_NAME`.

### `src/onedrive.ts`

Module tích hợp Microsoft Graph mới:

- Lấy access token bằng client credentials.
- Kiểm tra quyền truy cập file theo OneDrive item ID.
- Tải workbook hiện tại từ OneDrive.
- Upload workbook đã chỉnh sửa lên lại đúng file ID.

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
- Mở rộng mapping dữ liệu API sang nhiều cột Excel hơn.
- Tải workbook hiện tại từ OneDrive trước khi xử lý.
- Chỉ upload workbook lên OneDrive khi có dòng được thêm hoặc cập nhật.
- Sửa các chuỗi log tiếng Việt bị lỗi encoding.

## Quy Tắc Upsert Dữ Liệu Excel

Các dòng được đối chiếu bằng `settlementNo`:

| Cột Excel | Tiêu đề | Trường API | Cách xử lý |
|---|---|---|---|
| I | `SỐ ĐNTT-FMS` | `settlementNo` | Khóa đối chiếu duy nhất |
| F | `HẠN THANH TOÁN` | `dueDate` | Thêm mới hoặc cập nhật |
| G | `DỊCH VỤ` | `note` | Thêm mới hoặc cập nhật |
| H | `NCC` | `payeeName` | Thêm mới hoặc cập nhật |
| L | `SỐ TIỀN` | `amount` | Thêm mới hoặc cập nhật |
| M | `NGÀY LẬP - FMS` | `requestDate` | Thêm mới hoặc cập nhật |
| U | `TÌNH TRẠNG` | `statusApprovalName` | Thêm mới hoặc cập nhật |
| V | `CHI HỘ` | `departmentName` | Thêm mới hoặc cập nhật |

Các cột Excel còn lại được giữ nguyên và không bị automation chỉnh sửa.

## Luồng Chạy Hiện Tại

```text
Lấy access token eFMS
→ Lấy danh sách settlement payment
→ Lưu response API 1 và API 2 vào thư mục data
→ Tải workbook hiện tại từ OneDrive
→ Đối chiếu dòng theo settlementNo
→ Thêm các dòng chưa tồn tại
→ Cập nhật các trường được quản lý khi dữ liệu thay đổi
→ Ghi workbook local bằng exceljs
→ Upload đè workbook lên lại đúng file OneDrive nếu có thay đổi
```

## Kết Quả Kiểm Tra

- `npm run lint`: thành công
- `npm run build`: thành công
- `npm run check:onedrive`: thành công
- `npm run once`: chạy thành công sau khi file OneDrive không bị lock
- File OneDrive đã cấu hình có thể truy cập với tên `VENDOR_PAYMENT.xlsx`

## Lưu Ý Vận Hành

- Vì tool upload lại toàn bộ workbook, nên nên tránh chạy khi user đang mở hoặc chỉnh sửa file.
- Nếu file đang bị lock, OneDrive có thể trả lỗi `HTTP 423 resourceLocked`.
- File `.env` local chứa thông tin xác thực runtime và đã được Git bỏ qua.
