# Trình Thu Thập Dữ Liệu Tài Chính

Dự án này tự động hóa việc trích xuất dữ liệu tài chính cho các mã chứng khoán từ trang web SSC (Ủy ban Chứng khoán Nhà nước) và ghi kết quả vào file Excel (`BCTC.xlsx`).

## Tính Năng
- Sử dụng Puppeteer để tự động hóa Chrome cho việc thu thập dữ liệu web
- Trích xuất dữ liệu tài chính từ nhiều sheet (CDKT, KQKD, LCTT-GT) cho mỗi mã chứng khoán
- Đọc mã chứng khoán từ file `codes.xlsx` (Cột E)
- Ghi dữ liệu vào nhiều sheet trong `BCTC.xlsx` với tiêu đề cột phù hợp
- Hệ thống ghi log toàn diện với các cấp độ khác nhau (INFO, ERROR, SUCCESS, WARNING, DEBUG)
- Hỗ trợ phân trang và xử lý nhiều kết quả tìm kiếm
- Có thể cấu hình năm và quý cho việc trích xuất dữ liệu

## Yêu Cầu
- **Node.js** (khuyến nghị v16 trở lên)
- **npm** (Node package manager)
- **Google Chrome** đã cài đặt (đường dẫn mặc định: `C:\Program Files\Google\Chrome\Application\chrome.exe`)
- File Excel tên `codes.xlsx` với mã chứng khoán ở Cột E
- File Excel tên `BCTC.xlsx` với ba sheet: CDKT, KQKD, và LCTT-GT

## Cài Đặt
1. Clone hoặc tải repository này về máy tính của bạn
2. Mở terminal trong thư mục dự án
3. Cài đặt các dependencies:
   ```bash
   npm install
   ```

## Cấu Hình
- Đảm bảo `codes.xlsx` tồn tại trong thư mục dự án với mã chứng khoán ở Cột E
- Đảm bảo `BCTC.xlsx` tồn tại với ba sheet: CDKT, KQKD, và LCTT-GT
- Nếu Chrome của bạn được cài đặt ở vị trí khác, cập nhật `CHROME_PATH` trong object CONFIG trong `proxy.js`:
  ```js
  CHROME_PATH: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
  ```

## Sử Dụng
Chạy script với Node.js:
```bash
node proxy.js
```

Script sẽ:
- Đọc mã chứng khoán từ `codes.xlsx` (Cột E)
- Mở Chrome và điều hướng đến trang web SSC
- Tìm kiếm mỗi mã chứng khoán và trích xuất dữ liệu tài chính
- Ghi dữ liệu vào ba sheet trong `BCTC.xlsx`:
  - **CDKT** (Cân đối kế toán): Dữ liệu bảng cân đối kế toán
  - **KQKD** (Kết quả kinh doanh): Dữ liệu báo cáo kết quả kinh doanh
  - **LCTT-GT** (Lưu chuyển tiền tệ): Dữ liệu báo cáo lưu chuyển tiền tệ
- Ghi tất cả tiến trình và lỗi vào `log.txt`

### Lưu Ý Quan Trọng: Xóa Dữ Liệu Trước Mỗi Lần Chạy
**Sau mỗi lần chạy, vui lòng xóa tất cả dữ liệu trong file BCTC.xlsx để đảm bảo bạn nhận được dữ liệu mới nhất.** Script sẽ thêm dữ liệu vào các sheet hiện có, vì vậy việc xóa dữ liệu trước mỗi lần chạy sẽ ngăn chặn thông tin trùng lặp hoặc lỗi thời được bao gồm trong kết quả của bạn.

### Lưu Ý Quan Trọng: Không Mở File Excel Khi Đang Chạy
**Không mở file BCTC.xlsx hoặc codes.xlsx khi script đang chạy.** Nếu các file này đang được mở trong Excel, script sẽ không thể ghi dữ liệu vào chúng và có thể gây ra lỗi. Hãy đóng tất cả file Excel trước khi bắt đầu chạy script.

### Tùy Chỉnh Phạm Vi, Năm và Quý

Để thu thập dữ liệu cho một phạm vi mã cụ thể hoặc thay đổi năm và quý, chỉnh sửa dòng cuối trong `proxy.js`:

```js
// Xử lý tất cả mã cho Q4 2024
main(input, 2024, 4);

// Xử lý mã từ index 1 đến cuối cho Q4 2024
main(input.slice(1), 2024, 4);

// Xử lý 10 mã đầu tiên cho Q4 2024
main(input.slice(0, 10), 2024, 4);
```

**Tham số:**
- Tham số đầu tiên: Mảng mã chứng khoán cần xử lý
- Tham số thứ hai: Năm (ví dụ: 2024)
- Tham số thứ ba: Quý (1-4)

**Ví dụ:**
- Để lấy dữ liệu cho mã từ index 50 đến 99 cho năm 2023, quý 2:
  ```js
  main(input.slice(50, 100), 2023, 2);
  ```

## Cấu Trúc Đầu Ra

Script ghi dữ liệu vào ba sheet trong `BCTC.xlsx`:

1. **Sheet CDKT**: Dữ liệu bảng cân đối kế toán với cột cho mỗi mã chứng khoán
2. **Sheet KQKD**: Dữ liệu báo cáo kết quả kinh doanh với cột cho mỗi mã chứng khoán
3. **Sheet LCTT-GT**: Dữ liệu báo cáo lưu chuyển tiền tệ với cột cho mỗi mã chứng khoán

Mỗi mã chứng khoán có hai cột: `{CODE} cuoi ky` (cuối kỳ) và `{CODE} dau ky` (đầu kỳ).

## Ghi Log

Script cung cấp hệ thống ghi log toàn diện với các cấp độ khác nhau:
- **INFO**: Thông tin tiến trình chung
- **SUCCESS**: Trích xuất dữ liệu thành công
- **WARNING**: Vấn đề không nghiêm trọng
- **ERROR**: Lỗi cần chú ý
- **DEBUG**: Thông tin debug chi tiết

Tất cả log được ghi vào `log.txt` và cũng hiển thị trong console.

## Khắc Phục Sự Cố
- **Không tìm thấy Chrome:**
  - Đảm bảo Chrome đã được cài đặt và đường dẫn trong `proxy.js` là chính xác
- **Lỗi file Excel:**
  - Đảm bảo `codes.xlsx` tồn tại với mã chứng khoán ở Cột E
  - Đảm bảo `BCTC.xlsx` tồn tại với ba sheet: CDKT, KQKD, LCTT-GT
- **Không có dữ liệu được ghi:**
  - Kiểm tra file log để tìm lỗi hoặc selector bị thiếu
  - Xác minh mã chứng khoán tồn tại trên trang web SSC
- **Script đóng quá nhanh:**
  - Đảm bảo danh sách đầu vào (Cột E trong codes.xlsx) không rỗng

## Dependencies
- `puppeteer-core`: Tự động hóa Chrome
- `xlsx`: Đọc và ghi file Excel
- `fs`: Thao tác hệ thống file
- `axios`: HTTP requests (nếu cần cho các tính năng trong tương lai)

## Giấy Phép
ISC 