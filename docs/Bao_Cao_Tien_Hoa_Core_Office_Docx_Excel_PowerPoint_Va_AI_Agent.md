# BÁO CÁO KỸ THUẬT: TIẾN HÓA KIẾN TRÚC CORE XỬ LÝ DOCX, EXCEL, POWERPOINT & QUY TRÌNH TÁC NGHIỆP CỦA AI AGENT

**Hệ Thống**: Antigravity Office Studio (Phiên Bản Kiến Trúc v3.3.0)  
**Tác Giả**: Ban Phát Triển Antigravity Core & AI Agent Document Engineering  
**Thời Điểm Phát Hành**: Tháng 10/2026  
**Phạm Vi Tài Liệu**: Đặc tả kỹ thuật so sánh chuyên sâu các phiên bản Core (DOCX, Excel, PowerPoint) và quy trình tác nghiệp 5 giai đoạn của AI Agent.

---

## MỤC LỤC TỔNG QUAN

1. [TỔNG QUAN HỆ THỐNG & TRIẾT LÝ THIẾT KẾ CỐT LÕI](#1-tổng-quan-hệ-thống--triết-lý-thiết-kế-cốt-lõi)
2. [SO SÁNH CHI TIẾT CORE DOCX: TỪ TEXT THÔ ĐẾN DECOUPLED OPENXML](#2-so-sánh-chi-tiết-core-docx-từ-text-thô-đến-decoupled-openxml)
3. [SO SÁNH CHI TIẾT CORE EXCEL: TỪ BẢNG RỜI RẠC ĐẾN 12 QUALITY GATES](#3-so-sánh-chi-tiết-core-excel-từ-bảng-rời-rạc-đến-12-quality-gates)
4. [SO SÁNH CHI TIẾT CORE POWERPOINT: TỪ SLIDE CŨ ĐẾN 16:9 MASTER & IN-PLACE BRANDING](#4-so-sánh-chi-tiết-core-powerpoint-từ-slide-cũ-đến-169-master--in-place-branding)
5. [QUY TRÌNH TÁC NGHIỆP CỦA AI AGENT: CÁCH THỨC TẠO & BIẾN ĐỔI TÀI LIỆU](#5-quy-trình-tác-nghiệp-của-ai-agent-cách-thức-tạo--biến-đổi-tài-liệu)
6. [TỔNG KẾT & KẾT LUẬN KIẾN TRÚC](#6-tổng-kết--kết-luận-kiến-trúc)

---

## 1. TỔNG QUAN HỆ THỐNG & TRIẾT LÝ THIẾT KẾ CỐT LÕI

### 1.1. Bối cảnh & Thách thức trong Tự Động Hóa Tài Liệu Doanh Nghiệp
Xử lý tài liệu văn phòng (Word, Excel, PowerPoint) bằng AI và code tự động luôn là bài toán phức tạp bậc nhất trong công nghệ phần mềm:
- **Định dạng OpenXML / OOXML phức tạp**: Tài liệu văn phòng thực chất là các gói file ZIP chứa hàng chục tệp XML liên kết chéo. Việc chỉnh sửa sai một thẻ đóng (ví dụ thiếu `<w:p>` cuối ô bảng) sẽ khiến Microsoft Office từ chối mở file hoặc báo lỗi hỏng file (*"The file is corrupt and cannot be opened"*).
- **Mất mát nhận diện và đồ họa khi tạo file mới**: Các thư viện mã nguồn mở phổ thông (`python-docx`, `openpyxl`, `python-pptx`) khi khởi tạo file từ hư vô (`Workbook()`, `Presentation()`, `Document()`) sẽ **làm mất 100%** các thành phần phức tạp: logo vector trường/viện, dải màu thương hiệu (Header banner), smart art, macro VBA, biểu đồ DrawingML và định dạng in ấn khổ A4.
- **Hiện tượng xung đột file khóa (File Lock) trên Windows**: Các ứng dụng Office thường giữ quyền kiểm soát độc quyền (exclusive lock). Nếu tiến trình nền không giải phóng file handle hoặc dọn dẹp các tệp tạm (`.~lock.*`, `~$*`), hệ thống sẽ phát sinh lỗi `PermissionError: [Errno 13] Permission denied`.

### 1.2. Bốn Trụ Cột Triết Lý Thiết Kế Của Antigravity Office Studio

```
+-----------------------------------------------------------------------------------+
|                           ANTIGRAVITY OFFICE STUDIO                               |
+-----------------------------------------------------------------------------------+
          |                                       |
          v                                       v
[1. Declarative Spec Supremacy]        [2. Decoupled Architecture]
JSON Schema là nguồn chân lý duy nhất.   Tách biệt hoàn toàn Dữ liệu (.md, .json)
Không dựa vào auto-layout ngẫu nhiên.   với Kiểu dáng Trình bày (.style.yaml).
          |                                       |
          +-------------------+-------------------+
                              |
                              v
                [3. In-Place Mutation Paradigm]
                Kế thừa 100% Template gốc.
                Bảo toàn DrawingML, Header Banner, Logo, Charts.
                              |
                              v
                [4. Multi-Gate QA & Visual Guard]
                Kiểm chuẩn nghiêm ngặt: 12 Quality Gates (Excel),
                Bất biến OpenXML (Word), Khung hình 16:9 & Preview 1080p (PPTX).
```

1. **Declarative Spec Supremacy (Tối thượng Đặc tả Khai báo)**:
   Mọi tài liệu, slide, sơ đồ đều được mô tả thông qua bản đặc tả khai báo JSON chuẩn mực. AI Agent làm việc bằng cách biên dịch ý niệm người dùng thành JSON Spec trước khi gọi engine sinh mã.
2. **Decoupled Architecture (Kiến trúc Tách Rời Nội dung & Trình bày)**:
   Dữ liệu nghiệp vụ thuần túy được lưu trữ độc lập dưới dạng Markdown hoặc JSON. Kiểu dáng (màu sắc, font chữ, lề in, độ rộng viền) được quản lý qua stylesheet YAML, cho phép đổi giao diện hàng loạt mà không cần chỉnh sửa nội dung.
3. **In-Place Mutation Paradigm (Mô hình Đột biến Tại chỗ)**:
   Thay vì tạo mới tệp trắng, AI Agent nạp template gốc của người dùng, phân tích Semantic Anchors, sao chép dòng nguyên mẫu (Prototype Row Style Cloning), rồi tiêm dữ liệu trực tiếp vào đúng vị trí cần thiết.
4. **Multi-Gate Automated Verification (Kiểm chuẩn Đa Cổng Tự Động)**:
   Mọi tệp tin trước khi xuất xưởng bắt buộc phải vượt qua hệ thống kiểm tra tự động: đối chiếu 12 cổng chất lượng (Excel), kiểm tra thẻ OpenXML hợp lệ (Word), và xuất ảnh xem trước 1080p kiểm tra tràn chữ (PowerPoint).

---

## 2. SO SÁNH CHI TIẾT CORE DOCX: TỪ TEXT THÔ ĐẾN DECOUPLED OPENXML

### 2.1. Phân Tích Sự Khác Biệt Giữa Core Cũ và Core Mới

| Tiêu Chí Kỹ Thuật | Core DOCX Cũ (v1) | Core DOCX Mới (v2/v3 — Decoupled OpenXML) |
|---|---|---|
| **Thao tác Text** | Gán đè chuỗi thô `cell.text = "..."`. Xóa sạch font, kích thước, in đậm, màu chữ. | **Run Text Overwrite (`ERR_DOCX_002`)**: Thao tác sâu vào `cell.paragraphs[0].runs`, bảo tồn 100% định dạng nguyên bản. |
| **Bảng Nhiều Trang** | Hàng bị cắt đôi khi sang trang; trang sau không có tiêu đề cột. | **Bắt buộc Invariants**: Tự động chèn `<w:cantSplit/>` chống vỡ hàng và `<w:tblHeader/>` lặp lại tiêu đề. |
| **Căn Lề Ô Bảng** | Văn bản dính mép trên, lề dọc lộn xộn. | **Căn giữa dọc chuẩn mực**: Áp dụng `<w:vAlign w:val="center"/>` cho toàn bộ các ô trong tài liệu. |
| **Độ Toàn Vẹn XML** | Dễ gây lỗi hỏng file nếu cell trống hoặc thiếu thẻ đóng paragraph. | **The Last Paragraph Rule (`ERR_DOCX_001`)**: Đảm bảo mọi thẻ ô `<w:tc>` luôn kết thúc bằng tối thiểu một thẻ `<w:p>`. |
| **Chèn Sơ Đồ & Ảnh** | Hình ảnh tràn qua lề trang in ấn (> 14cm), làm vỡ khung Word A4. | **Printable Margin Lock (`ERR_DOCX_005`)**: Tự động tính toán và khóa tỷ lệ khung hình ≤ 14.0cm chiều rộng. |
| **Cơ Chế Chuyển Đổi** | Phụ thuộc công cụ chuyển đổi đóng gói sẵn, không tách được style. | **Biên dịch 2 chiều Decoupled**: Markdown + Style YAML ↔ OpenXML DOCX sạch. |
| **Vá Mẫu In-Place** | Phải viết script ad-hoc cho từng file. | `patch_docx_template`: Thay thế text tag `{{TAG}}` và block bảng `{{#BLOCK}}` linh hoạt. |

### 2.2. Chi Tiết Các Bất Biến OpenXML Bắt Buộc Trong Core DOCX Mới

Trong kiến trúc mới, mọi bảng biểu sinh ra hoặc sửa đổi đều phải tuân thủ nghiêm ngặt 3 bất biến XML cốt lõi:
```xml
<w:tbl>
  <!-- 1. Lặp lại dòng tiêu đề trên trang mới -->
  <w:tr>
    <w:trPr>
      <w:tblHeader/>
      <w:cantSplit/>
    </w:trPr>
    <w:tc>
      <w:tcPr>
        <!-- 2. Căn giữa văn bản theo chiều dọc -->
        <w:vAlign w:val="center"/>
      </w:tcPr>
      <w:p>
        <w:r><w:t>DÒNG TIÊU ĐỀ BẢNG</w:t></w:r>
      </w:p>
      <!-- 3. The Last Paragraph Rule: Luôn có w:p kết thúc cell -->
    </w:tc>
  </w:tr>
</w:tbl>
```

---

## 3. SO SÁNH CHI TIẾT CORE EXCEL: TỪ BẢNG RỜI RẠC ĐẾN 12 QUALITY GATES

### 3.1. Phân Tích Sự Khác Biệt Giữa Core Cũ và Core Mới

| Tiêu Chí Kỹ Thuật | Core Excel Cũ (v1) | Core Excel Mới (v2/v3 — 12 Quality Gates & Template-Driven) |
|---|---|---|
| **Cơ Chế Sinh File** | Khởi tạo từ `Workbook()` rỗng. Làm mất sạch DrawingML, Logo, Header vector, Charts. | **Template-Driven Token Extraction (E1)**: Nạp trực tiếp template gốc, trích xuất token và nhân bản in-place. |
| **Tính Hợp Lệ Dữ Liệu** | Gán cứng số liệu tĩnh (hardcoded numbers) vào các ô KPI và tổng kết. | **Live KPI Formulas (E3)**: Bắt buộc dùng công thức sống (`=COUNTIF`, `=SUM`, `='Sheet'!Cell`), tự động cập nhật. |
| **Xử Lý Ô Gộp (Merged)**| Gán dữ liệu làm vỡ merged ranges, rách đường viền kép (`ERR_XLSX_004`). | **Safe Merged-Cell Handling**: Chỉ gán vào ô Top-Left; đồng bộ viền toàn dải merged chống rách nét. |
| **Dịch Chuyển Công Thức**| Thêm hàng làm sai lệch các vùng tham chiếu công thức tổng kết bên dưới. | **Dynamic Multi-Axis Shifting (ERR_XLSX_002)**: Dùng Regex tịnh tiến động dải tham chiếu (ví dụ `$F$34:$H$34`). |
| **Liên Kết Biểu Đồ** | Khi chèn thêm hàng, biểu đồ vẫn trỏ vào vùng dữ liệu cũ hoặc bị lỗi đồ họa. | **Dynamic Chart Re-Anchoring (E12)**: Relink toàn bộ chuỗi dữ liệu biểu đồ và dời tọa độ neo đồ họa. |
| **Đồng Nhất Sheet Con** | Các sheet chức năng bị lệch cỡ chữ, lệch chiều rộng cột, lệch chiều cao hàng. | **Sibling Symmetry (E5 & E13)**: Đồng bộ 100% 20 sheet con về kích thước cột, font, và tự co giãn chiều cao dòng. |
| **Kiểm Định Chất Lượng** | Không có công cụ kiểm thử; chỉ dựa vào việc "mở file thử". | **12 Quality Gates Diff & Validator (E6)**: Chạy đối chiếu tự động 12 cổng format giữa Target Sheet và Reference Sheet. |
| **Trích Xuất Snapshot** | Phụ thuộc COM Excel desktop hoặc chỉ đọc được chuỗi công thức thô. | **Headless Hybrid Snapshot Inspector**: Kết hợp OpenPyXL và xlwings, đọc được cả giá trị đã tính toán (*computed values*). |
| **Quản Lý File Lock** | Bị lỗi `[Errno 13] Permission denied` khi file bị tiến trình khác mở. | **Auto Handle Release & Nonce Fallback**: Luôn gọi `wb.close()`, tự sinh tên file tạm tránh xung đột ghi đè. |

### 3.2. Bảng Đặc Tả Chi Tiết 12 Quality Gates (Cổng Kiểm Chuẩn Định Dạng)

```
[Gate 1: Header Block] ---------> Kiểm tra Font (Tahoma), Size (8pt), Nền, Viền và Căn lề dòng 2-8.
[Gate 2: Row 7 KPI Formulas] ---> Bắt buộc chứa công thức động COUNTIF/SUM, cấm ghi cứng số.
[Gate 3: Row 9 TC Headers] -----> Nền xanh Navy (#000080), chữ trắng in đậm, viền đôi đóng đáy.
[Gate 4: Col A Navy Fill] ------> Dải màu xanh hải quân cột A liền mạch, không bị đứt đoạn giữa các mục.
[Gate 5: Col B-C-D Box] --------> Khối hộp 3 cột đồng nhất: B (viền trái), C (không viền dọc), D (viền phải).
[Gate 6: Group Headers] --------> Ngăn chặn lặp lại tên nhóm trên các hàng con (chỉ hiển thị ở hàng đầu).
[Gate 7: Matrix Grid & Marks] --> Đồng nhất viền lưới ma trận, ký tự đánh dấu "O" in đậm kích cỡ 8pt.
[Gate 8: Result Footer Block] --> Font Courier 8pt không đậm, định dạng ngày tháng chuẩn mm/dd.
[Gate 9: Merged Cells] ---------> Kiểm tra độ toàn vẹn của các ô gộp tiêu đề và vùng kết quả.
[Gate 10: Outer Closing] -------> Viền đôi đáy đóng section dữ liệu trước khi chuyển vùng.
[Gate 11: Data Validations] ----> Bảo tồn danh sách dropdown lựa chọn (O/X, Pass/Fail, Boundary).
[Gate 12: Geometry & Widths] ---> Đồng nhất chiều rộng từng cột (dung sai <= 1.5) và chiều cao hàng động.
```

---

## 4. SO SÁNH CHI TIẾT CORE POWERPOINT: TỪ SLIDE CŨ ĐẾN 16:9 MASTER & IN-PLACE BRANDING

### 4.1. Phân Tích Sự Khác Biệt Giữa Core Cũ và Core Mới

| Tiêu Chí Kỹ Thuật | Core PowerPoint Cũ (v1) | Core PowerPoint Mới (v2 — 16:9 Master & In-Place Branding) |
|---|---|---|
| **Tỷ Lệ Slide** | Tỷ lệ 4:3 cổ điển hoặc 16:9 không chuẩn, viền đen khi trình chiếu trên màn hình hiện đại. | **16:9 Widescreen Master (`PPTX_INV_01`)**: Kích thước cố định $13.333 \times 7.5\text{ in}$ ($1280 \times 720\text{ pt}$). |
| **Nhận Diện Thương Hiệu** | Xóa sạch layout cũ, làm mất banner viền trên, logo trường đại học / viện nghiên cứu. | **In-Place Institutional Preservation (`preserve_branding`)**: Kế thừa 100% dải viền trên/dưới, logo, quốc huy từ template gốc. |
| **Tiêu Chuẩn Tiêu Đề** | Tiêu đề bị căn giữa ngẫu nhiên, cỡ chữ nhảy loạn (20pt, 24pt, 32pt) thiếu tính học thuật. | **Chuẩn Tiêu Đề 28pt Bold Lề Trái (`title_font_size: 28`, `title_align: "left"`)**: Chuẩn báo cáo khoa học / đồ án. |
| **Xử Lý Tràn Chữ** | Text boxes không bật `word_wrap=True`, chữ tràn qua mép đáy slide (`PPTX_INV_04`). | **Zero Text Overflow Guard**: Tự động bật `word_wrap=True` và co giãn cỡ chữ theo chiều cao khả dụng. |
| **Tỷ Lệ Sơ Đồ & Ảnh** | Ảnh bị méo hình do không khóa tỷ lệ $W/H$ khi chèn vào slide. | **Aspect-Ratio Lock (`PPTX_INV_02`)**: Giữ nguyên tỷ lệ ảnh; căn giữa 2D trong khung an toàn ($W \le 12\text{ in}, H \le 5.5\text{ in}$). |
| **Phương Thức Tạo** | Viết code thủ công tạo từng textbox, khó bảo trì và không tái sử dụng được. | **Declarative Slide Spec Supremacy (`write_pptx_from_spec`)**: Định nghĩa slide qua JSON Spec chuẩn hóa. |
| **Xem Trước Slide** | Bắt buộc phải mở file `.pptx` bằng Microsoft PowerPoint để kiểm tra. | **On-Demand 1080p Slide Preview Exporter (`export_pptx_slides`)**: Xuất ảnh PNG Full HD 1080p từng slide tức thì. |
| **Trải Nghiệm Web UI** | Chỉ có nút bấm tải file thô về máy. | **Slide Gallery & Modal Zoom**: Giao diện lưới slide 16:9, bấm phóng to toàn màn hình, duyệt phím mũi tên. |

### 4.2. Cấu Trúc Khai Báo JSON Slide Spec Chuẩn Mới

Một bài thuyết trình chuẩn Core v2 được định nghĩa thông qua cấu trúc JSON gọn gàng và tường minh:
```json
{
  "theme": "thesis_blue",
  "preserve_branding": true,
  "title_font_size": 28,
  "title_align": "left",
  "slides": [
    {
      "type": "content",
      "title": "THIẾT KẾ BỘ ĐIỀU KHIỂN HEADING PID",
      "bullets": [
        "Hàm truyền đạt vòng hở của hệ thống lái góc hướng: G(s) = K / (s * (T*s + 1))",
        "Bộ điều khiển PID heading: Kp = 1.85, Ki = 0.12, Kd = 0.45",
        "Thời gian đáp ứng xác lập t_s < 2.5s, độ vọt lố POT < 4.8%"
      ]
    },
    {
      "type": "diagram",
      "title": "SƠ ĐỒ KHỐI ĐIỀU KHIỂN ĐỘNG CƠ & HƯỚNG",
      "image_path": "diagram_assets/controller_block_diagram.png",
      "caption": "Hình 2.4: Cấu trúc điều khiển cascade động cơ đẩy và bánh lái"
    }
  ]
}
```

---

## 5. QUY TRÌNH TÁC NGHIỆP CỦA AI AGENT: CÁCH THỨC TẠO & BIẾN ĐỔI TÀI LIỆU

AI Agent trong `antigravity-doc-handler` không sinh mã ngẫu nhiên; agent vận hành như một **Kỹ Sư Hệ Thống Xử Lý Tài Liệu** tuân thủ quy trình 5 giai đoạn nghiêm ngặt:

```
+---------------------------------------------------------------------------------------------------+
|                           QUY TRÌNH 5 GIAI ĐOẠN TÁC NGHIỆP CỦA AI AGENT                           |
+---------------------------------------------------------------------------------------------------+
  [Stage 1: Ingestion & AST Parsing]
       |  AI Agent tiếp nhận tài liệu nguồn (Word, PDF, Excel thô, Thuyết minh, Prompt)
       |  Bóc tách cây cú pháp trừu tượng AST: Headings, Bullets, Math formulas, Tables, Diagrams.
       v
  [Stage 2: Declarative Spec Authoring]
       |  Biên dịch AST thành bản đặc tả khai báo JSON Spec (Slide Spec, Mutate Spec, Diagram Spec).
       |  Tách biệt hoàn toàn DỮ LIỆU và STYLING.
       v
  [Stage 3: In-Place Mutation & Engine Execution]
       |  Nạp Template mẫu doanh nghiệp, kích hoạt cơ chế clone dòng nguyên mẫu (Prototype Row).
       |  Tiêm dữ liệu in-place, tự động tịnh tiến công thức (Formula Shifting), re-anchor biểu đồ.
       |  Áp dụng các bất biến OpenXML (cantSplit, tblHeader, vAlign, 16:9 widescreen).
       v
  [Stage 4: Automated Multi-Gate Verification & Quality Assurance]
       |  Chạy script kiểm chuẩn: 12 Quality Gates (Excel), kiểm tra XML (Word).
       |  Xuất ảnh preview 1080p kiểm tra trực quan text overflow, layout.
       |  Tuân thủ quy tắc "Max 1 Fix Attempt": phân tích root-cause chính xác, không sửa mò.
       v
  [Stage 5: Dual Delivery & Multi-Channel Serving]
          Bàn giao đồng thời file gốc (.docx, .xlsx, .pptx) và file xuất bản (.pdf vector, .png 1080p).
          Phục vụ đa kênh: CLI Terminal, REST Web API Server (:8000), Web Studio UI (:5173).
```

### 5.1. Ma Trận Quyết Định Lựa Chọn Động Cơ Sơ Đồ Của AI Agent

Khi người dùng yêu cầu vẽ sơ đồ kỹ thuật, AI Agent tuân thủ chuẩn phân loại **12 loại sơ đồ kỹ thuật** để tự động điều phối tới động cơ tối ưu:

| Loại Sơ Đồ | Động Cơ Tối Ưu | Lý Do Kỹ Thuật & Chuẩn Đầu Ra |
|---|---|---|
| **Screen Flow (Tương tác Mobile)** | `canvas` Engine | Tọa độ pixel-perfect, Manhattan routing, hỗ trợ kéo thả trên `diagram_editor.py`. |
| **Database Schema (ERD Phân Hệ)** | `drawio` / `mermaid` | Nền phẳng pastel, kết nối bảng trực quan 3–15 thực thể, chuẩn MX_INV_01..06. |
| **API Sequence Diagram** | `mermaid` Engine | Chuỗi gọi hàm thanh thoát, đánh số tự động `autonumber`, tách biệt client-server. |
| **Flowchart / Thuật Toán** | `mermaid` Engine | Rẽ nhánh điều kiện `if/else` linh hoạt, styling khối màu hiện đại. |
| **Use Case Diagram (Tác nhân)** | `plantuml` Engine | Actor người que chuẩn UML, quan hệ `<<include>>`, `<<extend>>`. |
| **C4 Architecture (Hệ Thống)** | `plantuml` Engine | Thư viện C4 nội bộ `<C4/C4_Context>`, không vỡ dây container, chuẩn quốc tế. |
| **Activity Swimlane (Phân Làn)**| `plantuml` Engine | Phân làn cột dọc (`|Partition|`) tuyệt đối thẳng hàng, không đè dây logic. |
| **Hardware Schematic & Rail Bus**| `drawio` Engine | Phân tầng điện áp (+12V, +5V, +3V3, GND), bus GPIO waterfall, zero component piercing. |

---

## 6. TỔNG KẾT & KẾT LUẬN KIẾN TRÚC

Hệ thống kiến trúc Core mới của `antigravity-doc-handler` đã giải quyết triệt để các bài toán nan giải nhất trong tự động hóa văn phòng:
1. **Loại bỏ 100% rủi ro hỏng định dạng**: Nhờ chuyển dịch hoàn toàn từ mô hình sinh thô sang mô hình **In-Place Mutation** và **Declarative JSON Spec Supremacy**.
2. **Bảo tồn trọn vẹn nhận diện thương hiệu**: Giữ nguyên logo, banner, font chữ và màu sắc của các trường đại học, viện nghiên cứu và tổ chức doanh nghiệp.
3. **Kiểm chuẩn tự động không khoan nhượng**: Với hệ thống 12 Quality Gates và xuất ảnh xem trước 1080p, mọi sai lệch về viền, màu sắc hay công thức đều bị phát hiện và ngăn chặn trước khi bàn giao.
4. **AI Agent đóng vai trò Kỹ Sư Hệ Thống**: Vận hành theo quy trình 5 giai đoạn khép kín, phân tích cú pháp AST sâu và cam kết chất lượng sản phẩm đầu ra hoàn hảo cho người dùng cuối.

---
*Tài liệu được biên soạn và kiểm chứng tự động bởi Antigravity Office Studio v3.3.0.*
