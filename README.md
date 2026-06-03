
# Tự Động Hóa eFMS Vendor

Một công cụ tự động hóa TypeScript để quản lý thanh toán quyết toán nhà cung cấp trong eFMS (Hệ Thống Quản Lý Vận Tải Điện Tử). Công cụ này lấy dữ liệu thanh toán quyết toán từ các API của eFMS và xuất chúng vào các sổ làm việc Excel với chức năng ngăn chặn trùng lặp tự động.

## Tính Năng

- 🔐 Xác thực OAuth2 với API eFMS (mô hình cấp quyền mật khẩu)
- 📊 Lấy dữ liệu thanh toán quyết toán từ eFMS
- 📝 Xuất dữ liệu vào sổ làm việc Excel
- 🔄 Ngăn chặn trùng lặp tự động sử dụng theo dõi trạng thái
- ⏱️ Thực thi theo lịch (khoảng 30 phút)
- 🛡️ Triển khai an toàn kiểu dữ liệu với TypeScript
- ⚙️ Có thể cấu hình qua các biến môi trường

## Yêu Cầu

- Node.js 18+
- npm hoặc yarn
- Thông tin xác thực eFMS hợp lệ (tên người dùng, mật khẩu)
- Quyền truy cập API eFMS

## Cài Đặt

1. Clone kho lưu trữ:
```bash
git clone <repository-url>
cd Automation_eFMS_Vendor
```

2. Cài đặt các phụ thuộc:
```bash
npm install
```

3. Thiết lập các biến môi trường:
```bash
cp .env.example .env
```

Chỉnh sửa `.env` với thông tin xác thực và cấu hình eFMS của bạn.

## Cấu Hình

Tạo một tệp `.env` ở thư mục gốc dự án với các biến sau:

```env
# Cấu Hình API eFMS
EFMS_BASE_URL=https://efms-api.sotrans.com.vn
EFMS_TOKEN_PATH=/identityserver/connect/token
EFMS_CLIENT_ID=eFMS
EFMS_SCOPE=openid profile offline_access efms_api
EFMS_USERNAME=ten_dang_nhap_cua_ban
EFMS_PASSWORD=mat_khau_cua_ban

# Cài Đặt Tùy Chọn
EFMS_COMPANY_ID=id_cong_ty_cua_ban
EFMS_AUTHORIZATION=Bearer_token_neu_can
EFMS_TIMEOUT_MS=15000
EFMS_SETTLEMENT_PAYMENT_URL=https://efms-api.sotrans.com.vn/Accounting/api/v1/en-US/AcctSettlementPayment/paging?pageNumber=1&pageSize=1000
EFMS_SETTLEMENT_PAYMENT_REQUESTER=ten_nguoi_yeu_cau
```

### Biến Môi Trường Bắt Buộc:
- `EFMS_USERNAME` - Tên người dùng eFMS của bạn
- `EFMS_PASSWORD` - Mật khẩu eFMS của bạn

### Biến Môi Trường Tùy Chọn:
- `EFMS_BASE_URL` - URL cơ sở API eFMS (mặc định: `https://efms-api.sotrans.com.vn`)
- `EFMS_TOKEN_PATH` - Điểm cuối OAuth2 token (mặc định: `/identityserver/connect/token`)
- `EFMS_CLIENT_ID` - ID ứng dụng OAuth2 (mặc định: `eFMS`)
- `EFMS_SCOPE` - Phạm vi OAuth2 (mặc định: `openid profile offline_access efms_api`)
- `EFMS_COMPANY_ID` - ID công ty cho các yêu cầu API
- `EFMS_AUTHORIZATION` - Giá trị tiêu đề ủy quyền bổ sung
- `EFMS_TIMEOUT_MS` - Hết thời gian chờ yêu cầu (mặc định: `15000`)
- `EFMS_SETTLEMENT_PAYMENT_URL` - Điểm cuối API thanh toán quyết toán
- `EFMS_SETTLEMENT_PAYMENT_REQUESTER` - Định danh người yêu cầu cho API

## Cách Sử Dụng

### Chế Độ Phát Triển
Chạy với tải lại tự động khi file thay đổi:
```bash
npm run dev
```

### Build
Biên dịch TypeScript sang JavaScript:
```bash
npm run build
```

### Khởi Động
Chạy ứng dụng đã biên dịch:
```bash
npm start
```

### Kiểm Tra Lỗi
Kiểm tra TypeScript có lỗi mà không cần build:
```bash
npm run lint
```

## Cấu Trúc Dự Án

```
.
├── src/
│   ├── index.ts           # Logic ứng dụng chính
│   ├── auth.ts            # Xác thực OAuth2
│   ├── client.ts          # Client API eFMS
│   └── config.ts          # Trình tải cấu hình
├── data/
│   ├── Vendor_Payment_Template.xlsx      # Mẫu Excel
│   ├── Vendor_Payment_Output.xlsx        # Đầu ra được tạo
│   ├── Vendor_Payment_Output.state.json  # Theo dõi trạng thái (gitignored)
│   ├── api1-response.json                # Cache phản hồi xác thực (gitignored)
│   └── api2-response.json                # Cache dữ liệu quyết toán (gitignored)
├── package.json
├── tsconfig.json
├── .env                   # Biến môi trường (gitignored)
├── .env.example           # Cấu hình ví dụ
├── .gitignore
└── README.md             # Tệp này
```

## Cách Hoạt Động

1. **Xác Thực**: Xác thực với eFMS bằng mô hình cấp quyền mật khẩu OAuth2
2. **Lấy Dữ Liệu**: Lấy dữ liệu thanh toán quyết toán từ API eFMS
3. **Xử Lý**: 
   - Theo dõi ID thanh toán đã xử lý để ngăn chặn trùng lặp
   - Xử lý các thanh toán theo thứ tự thời gian ngược (mới nhất trước)
4. **Xuất Dữ Liệu**: Cập nhật sổ làm việc Excel bằng các bản ghi thanh toán mới
5. **Quản Lý Trạng Thái**: Duy trì tệp trạng thái để theo dõi các thanh toán đã xử lý

## Các Tệp Dữ Liệu

- **Tệp Phản Hồi API** (`api1-response.json`, `api2-response.json`): Cache phản hồi API với siêu dữ liệu (lần chạy cuối, thông tin phân trang)
- **Tệp Trạng Thái** (`Vendor_Payment_Output.state.json`): Theo dõi ID thanh toán đã xử lý để tránh trùng lặp
- **Đầu Ra Excel** (`Vendor_Payment_Output.xlsx`): Báo cáo được tạo có dữ liệu thanh toán nhà cung cấp
- **Mẫu** (`Vendor_Payment_Template.xlsx`): Mẫu Excel cơ bản để tạo đầu ra

> Các tệp này được bỏ qua bởi git (`.gitignore`) vì chúng chứa dữ liệu chạy lúc sử dụng và phản hồi

## Xử Lý Lỗi

Công cụ bao gồm xử lý lỗi mạnh mẽ cho:
- Biến môi trường bị thiếu hoặc không hợp lệ
- Lỗi xác thực API
- Hết thời gian chờ mạng (có thể cấu hình qua `EFMS_TIMEOUT_MS`)
- Lỗi I/O tệp
- Lỗi xử lý sổ làm việc Excel

## Phát Triển

### Các Lệnh Khả Dụng
```bash
npm run dev      # Chế độ xem trước cho phát triển
npm run build    # Build TypeScript
npm run start    # Chạy mã đã biên dịch
npm run lint     # Kiểm tra kiểu mà không cần build
```

### Ngăn Xếp Công Nghệ
- **Ngôn Ngữ**: TypeScript 5.9
- **Runtime**: Node.js ES modules
- **HTTP**: Fetch API với AbortController
- **Excel**: ExcelJS
- **Môi Trường**: dotenv

## Khắc Phục Sự Cố

### Xác Thực Không Thành Công
- Xác minh thông tin xác thực trong `.env`
- Kiểm tra `EFMS_BASE_URL` có chính xác không
- Đảm bảo kết nối mạng với API eFMS

### Biến Môi Trường Bị Thiếu
```
Error: Missing EFMS_USERNAME or EFMS_PASSWORD in .env
```
Thêm các biến bắt buộc vào tệp `.env` của bạn

### Hết Thời Gian Chờ Yêu Cầu
Tăng `EFMS_TIMEOUT_MS` trong `.env`:
```env
EFMS_TIMEOUT_MS=30000  # 30 giây
```

### Tệp Excel Bị Khóa
Đảm bảo tệp Excel không bị mở trong ứng dụng khác trước khi chạy công cụ

## Giấy Phép

Kho lưu trữ riêng tư

## Hỗ Trợ

Nếu có vấn đề hoặc câu hỏi, vui lòng liên hệ với nhóm phát triển.
