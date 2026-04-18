# VR QC-TC Import Tool — Conversation Context Handoff for Codex

## Mục tiêu
Tự động crawl và nhập liệu khoảng 190 văn bản QC-TC từ:
- Nguồn: `https://www.vr.org.vn/quy-chuan-tieu-chuan/Pages/default.aspx?...`
- Đích: hệ thống quản trị nội bộ `https://quantri-vr.e-office.vn`

Workflow mong muốn:
1. Crawl metadata + tải file đính kèm từ VR.
2. Upload sẵn toàn bộ file vào **Kho dữ liệu**.
3. Khi nhập form văn bản:
   - luôn chọn cố định:
     - Loại văn bản = `Tiêu chuẩn`
     - Cơ quan ban hành = `Bộ giao thông vận tải`
     - Người ký = `Bộ giao thông vận tải`
     - Chức danh = `b`
   - Loại văn bản QC-TC phải xác định từ crawl: `Quy chuẩn` hoặc `Tiêu chuẩn`
   - Lĩnh vực phải map theo nội dung văn bản
   - Số ký hiệu, Trích yếu, Nội dung lấy từ crawl
   - Các field khác không cần
4. Nhập **từ cuối lên**
5. Đã nhập tay **6 văn bản cuối**, nên tool cần bỏ qua 6 văn bản cuối
6. Do form draft giữ state:
   - field “Lĩnh vực” cần sync theo văn bản hiện tại
   - file đính kèm cần bỏ file cũ và chọn file mới từ Kho dữ liệu

---

## Tình trạng đã xác nhận

### 1. Crawl
- Đã crawl được khoảng **192 văn bản**
- Có export `documents_summary.csv`
- Có logic phát hiện bản ghi thiếu số ký hiệu
- CSV từng bị lỗi encoding, đã hướng sửa sang `utf-8-sig`
- Nếu `so_ky_hieu` rỗng:
  - đưa xuống cuối CSV
  - log ra file riêng để xử lý thủ công

### 2. Mapping lĩnh vực
- User đã cung cấp danh sách mapping gần đủ 48 lĩnh vực
- Có file `field_mapping.json`
- Có danh sách `unmatched_fields` trước đó, sau này đã giảm bớt
- Có nhiều alias/biến thể typo cần normalize

### 3. Domain nội bộ
**Rất quan trọng**:
- Sai domain cũ: `https://quantri-vr-e-office.vn/...`
- Domain đúng: `https://quantri-vr.e-office.vn/...`

### 4. VPN / nội bộ
- Hệ thống đích chỉ truy cập được qua VPN nội bộ
- Từng gặp:
  - `ERR_NAME_NOT_RESOLVED`
  - `502 Bad Gateway`
- Sau đó user đã vào được hệ thống qua VPN

---

## Tình trạng Playwright / login
Đã debug được:
- Browser mở được
- User login thủ công
- Sau khi Enter, bot vào được:
  - `https://quantri-vr.e-office.vn/van-ban/them-moi`
- `wait_for_form_ready()` đã phải sửa để tránh bắt nhầm text ở sidebar
- Cách ổn hơn là dùng selector theo `formControlName`

---

## HTML form chính đã được user cung cấp
File:
- `apps/angular/src/app/pages/vanban-management/vanban-list/vanban-form/vanban-form.component.html`

### Những selector chắc chắn từ source code
- `nz-select[formcontrolname='idTypeOfDocument']`
- `nz-select[formcontrolname='vanBanQCTC']`
- `nz-select[formcontrolname='fields']`
- `input[formcontrolname='officialNumber']`
- `textarea[formcontrolname='source']`
- `app-editor formControlName="content"`
- Có `<app-file-attachment ...>` cho file đính kèm

### Ý nghĩa
Form cha đã rõ. Nhưng modal “Kho dữ liệu” **không nằm trong HTML này**.

---

## HTML component file attachment đã được user cung cấp
Component:
- `app-file-attachment`

Markup chính:

- Nút mở finder:
  - `<a (click)="onOpenFCKFinder()" class="btn btn-sm btn-flex btn-primary">...Chọn file</a>`
- Danh sách file đã chọn:
  - `fileAttachments`
  - icon xóa qua `deleteItem(file.id)`
  - tên file render bằng `{{ file.fileName }}`

### Ý nghĩa
- Nút `Chọn file` là đúng
- Nhưng modal / finder `Kho dữ liệu` được mở bởi `onOpenFCKFinder()`
- DOM nội bộ của finder **chưa có source code**
- Đây là điểm nghẽn hiện tại

---

## Trạng thái tool hiện tại
Tool đã đi được đến:
1. mở form thêm mới
2. click `Chọn file`
3. modal `Kho dữ liệu` mở ra

Nhưng đang kẹt ở bước:
- tìm folder `files` trong panel bên trái của modal `Kho dữ liệu`

### Quan sát UI từ ảnh
Trong modal `Kho dữ liệu`:
- cột trái có cây thư mục:
  - `files`
  - `lanhdaoduongnghiem`
  - `ldquathoiky`
  - `nv`
  - có lúc thêm `Văn bản QPPL`
- phía trên vùng file có icon/menu
- user mô tả workflow đúng là:
  1. click folder `files`
  2. bấm menu / dấu cộng cạnh folder
  3. chọn `Tải file cho thư mục này`
  4. popup `Upload file`
  5. chọn file từ `D:\Work\tinduc\Tool\data`

### Lưu ý quan trọng
Không nhất thiết phải thao tác Windows file picker bằng click chuột.
Nếu trong popup có `input[type=file]` hoặc `filechooser`, có thể dùng:
- `set_input_files(...)`
- `chooser.set_files(...)`

---

## Lỗi hiện tại mới nhất
Bot mở được modal `Kho dữ liệu`, nhưng lỗi:

`RuntimeError: Không tìm thấy folder 'files' trong Kho dữ liệu.`

Điều này xảy ra dù user thấy rõ `files` trên màn hình.

### Suy luận
- Selector hiện tại đang bắt sai DOM
- Có thể `files` nằm trong:
  - shadow-ish tree của component
  - node text không visible theo cách Playwright đang dùng
  - structure không match với `text=files`, `:text('files')`, `get_by_text(..., exact=True)` như hiện tại
- Cần debug đúng DOM của modal thật

---

## File debug nên lấy tiếp
Trong lúc debug modal finder, cần dump:
- `data/screenshots/debug_kho_root.png`
- `data/screenshots/debug_kho_root.html`
- `data/screenshots/debug_kho_root.txt`
- `data/screenshots/debug_visible_texts.txt`

Mục đích:
- xác định DOM thật của modal
- chốt selector đúng cho:
  - folder `files`
  - icon menu cạnh folder
  - item `Tải file cho thư mục này`
  - popup `Upload file`

---

## Vấn đề với code hiện có trong canvas
Canvas document hiện tại tên:
- `Vr Qc Tc Import Tool`

Nhưng nội dung canvas **không phản ánh đầy đủ các lần sửa cuối cùng** và vẫn có nhiều phần cũ/sai, ví dụ:
- docstring còn dùng domain sai `quantri-vr-e-office.vn`
- `TARGET_CREATE_URL = "https://quantri-vr-e-office.vn/van-ban/them-moi"` trong canvas hiện tại là sai
- `ensure_login()` trong canvas hiện tại là bản cũ, tự `goto()` luôn, không phải bản login tay + Enter
- `wait_for_form_ready()` trong canvas hiện tại còn bắt text quá chung (`text=Số ký hiệu`, `text=Lĩnh vực`, ...)
- logic upload file vào Kho dữ liệu trong canvas hiện tại vẫn là bản cũ, chưa phải bản debug dump finder mới nhất

=> Codex cần coi canvas hiện tại là **chưa đồng bộ hoàn toàn** với các trao đổi cuối.

---

## Những điểm chắc chắn cần giữ khi sửa tool

### Domain đúng
- `https://quantri-vr.e-office.vn`

### Selector form đúng hơn theo source code
- `input[formcontrolname='officialNumber']`
- `textarea[formcontrolname='source']`
- `nz-select[formcontrolname='idTypeOfDocument']`
- `nz-select[formcontrolname='vanBanQCTC']`
- `nz-select[formcontrolname='fields']`
- `app-file-attachment a.btn`

### Giá trị cố định
- Loại văn bản = `Tiêu chuẩn`
- Cơ quan ban hành = `Bộ giao thông vận tải`
- Người ký = `Bộ giao thông vận tải`
- Chức danh = `b`

### Dữ liệu nhập
- Số ký hiệu = `officialNumber`
- Trích yếu = `source`
- Nội dung = editor component `content`
- Lĩnh vực = multiple select `fields`

### Logic nhập
- nhập từ cuối lên
- bỏ qua 6 văn bản cuối đã nhập
- sync field “Lĩnh vực”
- remove file cũ trước khi attach file mới

---

## Commands đã dùng
Crawl:
```powershell
py .\vr_qc_tc_import_tool.py crawl
```

Upload file:
```powershell
py -u .\vr_qc_tc_import_tool.py upload-files --interactive-login
```

Import:
```powershell
py .\vr_qc_tc_import_tool.py import --interactive-login
```

---

## Các lỗi quan trọng đã gặp trước đó
1. `pip` không nhận diện trong PowerShell
2. `python` không có trong PATH
3. `PermissionError` khi đang mở `documents_summary.csv`
4. `AttributeError: Namespace object has no attribute 'headless'`
5. `ERR_NAME_NOT_RESOLVED`
6. `502 Bad Gateway`
7. selector bị bắt nhầm text ở sidebar (`Lĩnh vực`, `Loại văn bản`)
8. `textarea` không visible do bắt trúng phần tử ẩn
9. modal `Kho dữ liệu` mở được nhưng không tìm được folder `files`

---

## Hướng xử lý tiếp theo cho Codex
1. Đồng bộ lại file tool từ canvas + convo context này
2. Sửa tất cả domain sang `quantri-vr.e-office.vn`
3. Chuyển selector form sang bám `formcontrolname`
4. Giữ flow login tay + Enter
5. Giữ bước debug dump modal finder
6. Tập trung xử lý modal `Kho dữ liệu` bằng 1 trong 2 cách:
   - đọc đúng DOM dump từ `debug_kho_root.html`
   - hoặc dùng locator bám vị trí panel trái / tree node / icon ngay cạnh row `files`
7. Nếu popup `Upload file` có `input[type=file]`, dùng `set_input_files`
8. Sau khi upload file xong mới tiếp tục import văn bản

---

## Gợi ý checklist khi tiếp tục với Codex
- [ ] kiểm tra `TARGET_CREATE_URL`
- [ ] kiểm tra `TARGET_LIST_URL`
- [ ] kiểm tra `ensure_login()` là bản login tay
- [ ] kiểm tra `wait_for_form_ready()` dùng `formcontrolname`
- [ ] kiểm tra `upload_files` không dùng selector text quá chung
- [ ] dump DOM thật của `Kho dữ liệu`
- [ ] chốt selector click folder `files`
- [ ] chốt selector menu `Tải file cho thư mục này`
- [ ] chốt selector popup `Upload file`
- [ ] dùng `set_input_files` nếu có thể

---

## File user đã cung cấp trong cuộc trao đổi
- HTML form cha `vanban-form.component.html`
- HTML component `app-file-attachment`
- nhiều ảnh chụp màn hình của:
  - form thêm mới
  - modal Kho dữ liệu
  - popup Upload file
  - hộp thoại Windows file chooser
  - danh sách lĩnh vực
- danh sách mapping lĩnh vực gần đầy đủ 48 mục
- danh sách unmatched field trước đó

---

## Chốt ngắn
Tool hiện **đã đi đúng đến modal Kho dữ liệu**.
Điểm nghẽn duy nhất còn lại là **selector của finder/modal upload file**, đặc biệt là folder `files` và menu `Tải file cho thư mục này`.

Các phần crawl, mapping, login tay, vào form, và selector chính của form đã có đủ context để tiếp tục sửa.
