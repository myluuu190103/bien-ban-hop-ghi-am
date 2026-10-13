---
name: bien-ban-hop-ghi-am
description: Dựng biên bản họp Word (.docx) từ bản ghi âm/gõ lại (ASR) rối, lộn xộn, do user dán trực tiếp vào chat. Nạp khi user dán một đoạn ghi âm cuộc họp (nhiều giọng, không dấu câu, không tên người nói) và bảo "làm biên bản" / "cuộc họp mới" / "biên bản như các biên bản trước". Công cụ ở C:\NVAI\Claude\bien-ban-hop-ghi-am.
---

# Dựng biên bản họp từ bản ghi âm rối

Công cụ: `C:\NVAI\Claude\bien-ban-hop-ghi-am\build.py`. Mẫu Word gốc: `_template.docx`
trong cùng thư mục (bản sao bền của file Á Châu 21/09/2026 — không còn phụ thuộc
scratchpad của phiên nào, có thể dùng ở phiên mới bất kỳ lúc nào).

## Khuôn chuẩn (giữ nguyên, không tự đổi)

```
BIÊN BẢN HỌP <TÊN ĐƠN VỊ>        ← căn giữa, đậm
Ngày dd/mm/yyyy

I. Thành phần
- liệt kê người dự, chủ trì ghi riêng dòng đầu

II. Thời gian, địa điểm
- Thời gian: ...
- Hình thức: ...

III. Nội dung chính
Ông X nêu ý kiến:                ← đậm, không gạch chân
- gạch đầu dòng cấp 1
  - gạch đầu dòng cấp 2 (b2, thụt sâu hơn)
Ý kiến Chủ tịch:
- ...

Kết luận:
- ...

                                   Người ghi biên bản     ← nghiêng, căn phải
                                            (tên, để trống nếu chưa biết)
```

Trình bày: **12pt toàn văn bản**, lề trái/phải 1,19cm, trên 0,76cm, dưới 1,19cm —
đây là thông số user đã tự sửa và chốt 05/10/2026, đã gắn cứng vào `build.py`
nên không cần chỉnh tay sau khi build.

## Quy trình

1. **Đọc bản ghi âm được dán vào**: đây là ASR thô, rất rối — nhiều giọng chen nhau,
   không có dấu câu, tên người nghe sai (ASR nhầm âm), có cả chuyện riêng (gia đình,
   sức khỏe, lương, lời qua tiếng lại, chuyện tâm linh/cúng bái) lẫn trong nội dung họp.
2. **Nhận diện đơn vị họp và người tham dự** từ ngữ cảnh (tên công ty, sản phẩm,
   địa danh được nhắc tới). Nếu không chắc, giữ nguyên cách gọi trong bản ghi
   (không tự thêm họ đệm), và nói rõ với user là chưa đối chiếu được.
3. **Viết lại thành spec JSON** — xem `hop_a_chau_dinh_ky.json` hoặc
   `hop_do_thi_tan_hung.json` làm ví dụ mới nhất. Mỗi phần tử là `[kind, text]`,
   `kind` thuộc `title/date/sec/h/b/b2/sign/name` (chi tiết trong đầu file `build.py`).
   Khi viết:
   - **Tóm gọn, không dịch nguyên văn** — câu nói ASR rất dài dòng, lặp; viết lại
     thành câu hành chính ngắn, giữ đúng số liệu, tên riêng, quyết định.
   - **Bỏ chuyện riêng** (sức khỏe, gia đình, lương cá nhân, lời chửi/mỉa mai,
     chuyện tâm linh) — chỉ giữ nếu nó là một quyết định/việc giao cụ thể.
   - **Giữ trung tính** khi có tranh luận gay gắt hoặc cáo buộc tiêu cực giữa các
     bên: ghi lại nội dung tranh luận (ai nói gì, kết luận gì), bỏ lời nặng/mỉa mai.
   - **Để trống chỗ không chắc** (ngày, giờ, số liệu nghe không rõ) — KHÔNG dùng
     "(?)". Viết `"……"` hoặc để trống rồi nói rõ với user trong phần trả lời.
   - Tên file spec và file docx đặt theo tên đơn vị/chủ đề, ví dụ
     `hop_song_da_truong_son_13-10.json` → `BIEN_BAN_HOP_SONG_DA_TRUONG_SON_13-10-2026.docx`.
4. **Build**: `python build.py <spec.json> <out.docx>` (chạy trong thư mục
   `bien-ban-hop-ghi-am`, hoặc dùng đường dẫn tuyệt đối — script tự tìm
   `_template.docx` cạnh nó nên chạy từ đâu cũng được).
   - Nếu file đích đang mở trong Word, script tự lưu sang `..._v2.docx` và báo
     ra `stderr` — không cần tự đặt tên khác trước.
5. **Gửi file** bằng `SendUserFile`, rồi liệt kê ngắn gọn các chỗ không chắc
   (ngày, tên, số liệu, đoạn nghe không rõ) để user xác nhận/sửa.

## Khi user tự sửa file đã gửi rồi báo "lưu lại nhé"

User thường mở file `.docx` trong Word và sửa tay (tên đơn vị, ngày, số liệu...).
Khi được bảo lưu:
1. Tìm file `.docx` vừa sửa gần nhất trong thư mục (so `mtime`).
2. Diff nội dung file đó với spec JSON gốc (đọc `paragraphs`, so với
   `[text if kind not in ('b','b2') else '- '+text for kind,text in spec]`).
3. Đồng bộ đúng các chỗ user đã đổi vào spec JSON (để lần dựng lại sau không
   mất sửa đổi của user) — **không tự đổi thêm gì khác**.
4. Nếu user còn muốn một bản sao vào `Downloads`, chép sang đó với tên rõ nghĩa.

## Bẫy đã gặp

- **Tên công ty/người dễ nhầm vì ASR nghe sai âm** (ví dụ "Trung"→Đào Xuân Chung,
  "Long Anh"/"Lâm Anh" cùng một mỏ đá, "Tiên Quang"→Tuyên Quang). Đối chiếu qua
  ngữ cảnh, không đoán bừa; nói rõ chỗ không chắc.
- **File đích đang mở trong Word** → `PermissionError`, script tự xử lý bằng cách
  lưu sang `_v2.docx` (xem mục Build ở trên).
- **Số liệu tài chính đọc lướt, mơ hồ** (ví dụ "236 triệu" và "239 triệu" cùng
  xuất hiện cho một khoản) → ghi cả hai, chú thích cần rà lại, không tự chọn một số.
- **Lẫn nội dung cá nhân/nhạy cảm** (cúng bái, sức khỏe, mâu thuẫn cá nhân, cáo
  buộc tiêu cực chưa được xác nhận hai chiều) → không đưa vào biên bản chính thức
  trừ khi nó gắn với một quyết định/việc giao; nói rõ đã bỏ gì trong phần trả lời.
- **MEMORY.md / memory**: thông số trình bày (12pt, lề) và các bẫy trên đã ghi ở
  `bien-ban-hop-tu-ghi-am.md` — khi chốt thêm quy ước mới, cập nhật file đó thay
  vì chỉ sửa trong `build.py`.
