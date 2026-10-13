# Skill: Dựng biên bản họp từ bản ghi âm

Skill cho Claude Code: dựng biên bản họp Word (`.docx`) theo khuôn chuẩn Minh Trung,
từ một đoạn ghi âm/gõ lại (ASR) rối do người dùng dán trực tiếp vào chat.

## Cấu trúc

```
.claude/skills/bien-ban-hop-ghi-am/SKILL.md   ← mô tả skill (Claude tự nạp khi phù hợp)
bien-ban-hop-ghi-am/build.py                  ← script dựng file .docx từ spec JSON
bien-ban-hop-ghi-am/_template.docx            ← mẫu Word gốc (style/font/lề chuẩn)
```

Clone repo này vào gốc project (ví dụ `C:\NVAI\Claude`) để Claude Code tự thấy skill
và công cụ đi kèm.

## Dùng thủ công

```
python bien-ban-hop-ghi-am/build.py <spec.json> <out.docx>
```

`spec.json` là một list `[kind, text]`, `kind` thuộc `title/date/sec/h/b/b2/sign/name`.
Chi tiết format và quy trình đầy đủ nằm trong `SKILL.md`.

## Trình bày

Cỡ chữ 12pt toàn văn bản, lề trái/phải 1,19cm, trên 0,76cm, dưới 1,19cm — đã gắn cứng
trong `build.py`, không cần chỉnh tay sau khi dựng file.

---
Đóng gói 13/10/2026 từ quy trình thực tế dựng hơn chục biên bản họp nội bộ.
