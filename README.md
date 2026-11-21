# Nhập đơn Shopee / TikTok cho MISA AMIS

Ứng dụng máy tính để bàn, giao diện tiếng Việt. Tệp nguồn được xử lý trên máy. Có hai cách xuất: dùng workbook mẫu nhập đơn đặt hàng của doanh nghiệp hoặc tạo bảng đối soát độc lập.

## Cài trên Windows 10/11

1. Cài Python 3.12 từ [python.org](https://www.python.org/downloads/windows/). Trong trình cài, chọn **Add Python to PATH**, giữ thành phần **Tcl/Tk** và **Python Launcher**.
2. Tải hoặc sao chép toàn bộ dự án vào một thư mục. Mở thư mục chứa `app.py` trong File Explorer, gõ `cmd` vào thanh địa chỉ rồi nhấn Enter.
3. Chạy từng lệnh sau trong cửa sổ vừa mở:

   ```bat
   py -3.12 -m venv .venv
   .venv\Scripts\python.exe -m pip install -r requirements.txt
   .venv\Scripts\python.exe app.py
   ```

4. Để đọc PDF ảnh, cài Tesseract 5 theo [hướng dẫn Windows của Tesseract](https://tesseract-ocr.github.io/tessdoc/Installation.html). Chọn dữ liệu tiếng Việt và tiếng Anh. Thêm thư mục chứa `tesseract.exe` (thường là `C:\Program Files\Tesseract-OCR`) vào **Path** của người dùng. Mở lại cửa sổ cmd và chạy `tesseract --list-langs`. Danh sách cần có `vie` và `eng`. Excel và PDF có lớp chữ có thể đọc độc lập với bước này.
5. Những lần sau chỉ cần mở cmd ở thư mục dự án và chạy `.venv\Scripts\python.exe app.py`. Kiểm thử tự động bằng `.venv\Scripts\python.exe -m unittest discover -s tests -v`.

## Quét và đối soát

- Chọn kho mặc định. Khi bật nhận diện tên tệp, một từ khóa kho khớp duy nhất sẽ thay thế kho mặc định (ví dụ `đơn_HN_01.xlsx`). Nếu khớp cả hai hoặc không khớp, ứng dụng dùng kho mặc định. Đối chiếu **Mã kho** trong bảng kết quả.
- Nhấn **Thêm PDF / Excel** và chọn nhiều tệp bằng Ctrl/Shift. Tệp được đọc lần lượt trên luồng nền. Ứng dụng không đặt giới hạn số tệp, dung lượng thực tế phụ thuộc RAM. Tệp lỗi được báo riêng, có thể chọn lại sau khi sửa. Một đường dẫn đã nhập thành công chỉ được nạp một lần trong phiên.
- Excel được dò tiêu đề trên từng sheet. PDF ưu tiên bảng có đường kẻ, sau đó dùng vị trí chữ hoặc OCR cho trang ảnh. Bảng cần có cột mã đơn và số lượng. Nên dùng mã đơn và SKU dạng **Text** ở nguồn để bảo toàn số 0 đầu và mã dài. Hãy kiểm tra đặc biệt các trang mờ, bảng không kẻ dòng và dòng hàng bị ngắt giữa hai trang. Đơn giá được hiểu là giá một đơn vị ghi trên phiếu; dữ liệu nguồn thiếu giá được để trống.
- Gõ mã đơn, SKU hoặc tên vào ô lọc. Nhấp đúp một ô để sửa, hoặc chọn dòng rồi nhấn **Xóa dòng đã chọn**. Dòng phụ kiện do quy tắc tạo cũng sửa/xóa được. Quy tắc được áp dụng lúc quét. Khi sửa SKU, số lượng hay mã đơn sau đó, hãy đối chiếu lại dòng phụ kiện. Tránh đưa cùng đơn hàng từ hai bản xuất trùng lặp vào một phiên.
- Khi xuất, tất cả các dòng được nhóm theo mã đơn, kể cả dòng đang bị ẩn bởi bộ lọc. Nếu một đơn vượt **500 dòng**, ứng dụng hiển thị mã đơn và yêu cầu xác nhận.

## Xuất theo mẫu MISA

1. Tải **mẫu nhập đơn đặt hàng .xlsx** từ đúng phân hệ MISA AMIS của doanh nghiệp. Tạo bản sao trống, xóa các dòng ví dụ. Giữ tiêu đề, các sheet danh mục và định dạng. Trong dòng dữ liệu đầu tiên, có thể điền các hằng số chung vào cột không ánh xạ (ví dụ mã khách hàng mặc định), hoặc đặt công thức Excel. Các giá trị này được sao chép cho từng dòng.
2. Chọn **Xuất theo mẫu MISA…**, chọn bản sao mẫu. Kiểm tra sheet đích, số thứ tự **dòng dữ liệu đầu tiên** và chữ cái cột cho từng trường. Mẫu có tiêu đề gộp hoặc tên cột khác vẫn có thể ánh xạ tay. Bốn trường bắt buộc là mã đơn, mã hàng, mã kho và số lượng. Trường không sử dụng có thể để trống.
3. Lưu thành một tệp mới. Mở Excel kiểm tra chứng từ, ĐVT, ngày, giá, thuế, công thức và các mã danh mục trước khi nhập thử vào AMIS. Một số trường nguồn, chẳng hạn ngày hoặc thuế suất, có thể ở dạng văn bản: hãy chuẩn hóa theo yêu cầu mẫu thực tế. Chưa có mẫu chính thức của doanh nghiệp đi kèm dự án để xác nhận kết quả nhập trực tiếp.
4. Nút **Xuất bảng đối soát…** tạo workbook riêng để kiểm tra hoặc trao đổi dữ liệu.

## Cấu hình nghiệp vụ

Chỉnh `config/rules_config.json` bằng trình soạn thảo UTF-8 rồi khởi động lại ứng dụng.

| Thiết lập | Ý nghĩa |
|---|---|
| `warehouses` | Sửa `code` thành mã đúng trong danh mục kho AMIS. HCM và HN là giá trị ban đầu. `keywords` chứa các tên viết tắt hoặc địa danh có thể xuất hiện trong tên tệp. |
| `combo_sku_mode` | `keep` giữ SKU sàn trước khi áp dụng danh mục MISA. `map` yêu cầu có SKU lẻ tương ứng trong `combo_to_unit`. |
| `combo_to_unit` | Ví dụ `{"COMBO-BP02": "BP01"}`. Tên sản phẩm hoặc phân loại có cụm “Combo 02” với số lượng 4 sẽ xuất 8 đơn vị. Đơn giá được chia cho hệ số combo để giữ thành tiền. |
| `sku_to_misa` | Ví dụ `{"BP01": "HH0001"}`. Ưu tiên SKU phân loại, sau đó dùng SKU gốc. SKU chưa ánh xạ được giữ nguyên. |
| `adapter` | `enabled` bật/tắt, `keywords` dò trong tên hoặc phân loại, bỏ qua hoa/thường và dấu. Chỉnh `sku`, `name`, `unit`, `quantity` theo danh mục. `per_order` thêm một dòng cho mỗi đơn có máy, mặc định SL 1; `per_unit` nhân số lượng phụ kiện với tổng số máy sau quy đổi. |
| `ocr_languages` | Mặc định `vie+eng`. Dùng tên gói ngôn ngữ đã cài của Tesseract. |

## Mã nguồn

`parsers/` đọc bảng và OCR, `rules/` xử lý quy đổi và phụ kiện, `exporters/` ghi workbook, `config/` lưu cấu hình, `ui/` chứa giao diện Tkinter, `tests/` chứa kiểm thử.

## Mẫu và danh mục doanh nghiệp

Trong hộp thoại chọn mẫu, dùng `config/Mau_MISA_don_dat_hang.xlsx`. Đây là bản chuyển từ mẫu đa tiền tệ .xls đã cung cấp, đã xóa hai đơn minh họa. Sheet **Don dat hang** có tiêu đề tại dòng 8 và dữ liệu bắt đầu ở dòng 9. Cấu hình có 44 mặt hàng, tên và ĐVT từ danh mục doanh nghiệp. Chỉnh `sku_to_misa` để duyệt ánh xạ, sau đó bỏ SKU tương ứng khỏi `review_skus`; mã trong danh sách chờ duyệt sẽ để trống khi quét. `sku_multipliers` dùng cho quy cách đóng gói riêng của từng SKU.

Với cặp phiếu TikTok của cùng lô, chỉ chọn bản chi tiết SA6 để lập đơn. Picking List là bản đối chiếu riêng. Với Shopee, ứng dụng lấy SL, giá và thành tiền từng đơn từ phiếu chi tiết, đối chiếu tổng SL và tra SKU phân loại từ bảng xuất hàng đầu tệp.

## Theo dõi tồn theo ngày, tuần, tháng

1. Mở **Tồn kho… → Nạp tồn MISA…**, chọn file danh mục có các cột **Mã, Tên, Đơn vị tính chính, Số lượng tồn**. Nhập ngày chốt thực tế: số dư phải là số dư **cuối ngày** đó. Để theo dõi từ ngày 01/11, dùng số dư cuối ngày 31/10. Có thể nạp các lần chốt sau.
2. Quét đơn, đối chiếu ngày đơn, mã MISA, số lượng combo và phụ kiện. Với cùng lô, chỉ nạp một nguồn chi tiết. Chọn **Tồn kho… → Ghi đơn đã duyệt** để lưu các đơn hoàn chỉnh vào sổ. Đơn đã ghi với cùng ngày và cùng tổng số lượng theo mã sẽ được bỏ qua. Nếu cùng mã đơn nhưng khác số liệu hoặc còn SKU chưa đối chiếu, ứng dụng báo lỗi và giữ nguyên cả lô.
3. Chọn **Ngày**, **Tuần** (thứ Hai đến Chủ nhật) hoặc **Tháng**, nhập một ngày thuộc kỳ rồi chọn **Xem báo cáo**. Phần mềm lấy lần chốt gần nhất trước đầu kỳ, trừ các đơn đã ghi từ sau lần chốt đến trước đầu kỳ để tính tồn đầu kỳ. **Tồn cuối kỳ = Tồn đầu kỳ − Đã bán trong kỳ**. Số âm được đánh dấu. Có thể xuất báo cáo ra Excel.
4. Sổ lưu trên máy Windows ở `%LOCALAPPDATA%\Nhapdon\ton_kho.sqlite3`. Đóng ứng dụng trước khi sao lưu tệp. Danh mục được cung cấp có số dư tổng hợp nên báo cáo cũng tổng hợp các kho. Sổ tính theo ngày của các đơn đã ghi; nhập bổ sung, hoàn hàng hoặc điều chỉnh thực tế cần được phản ánh trong lần chốt MISA mới. Để tính trọn một kỳ đã có nhập hoặc hoàn hàng trong kỳ, cần đối chiếu thêm các phát sinh đó.

## Bản chạy Windows (.exe)

Mở [GitHub Releases](https://github.com/nguyenchicong035/Nhapdon/releases), tải **Nhapdon.exe** trong phần **Assets**, rồi chạy trên Windows 10/11 64-bit. Bản này đã đóng gói Python. Đọc PDF ảnh cần cài Tesseract như hướng dẫn ở trên. Khi chạy lần đầu, ứng dụng chép cấu hình và mẫu Excel vào `%LOCALAPPDATA%\Nhapdon\config`. Có thể chỉnh cấu hình trong thư mục này.
