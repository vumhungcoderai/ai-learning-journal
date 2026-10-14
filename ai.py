"""
TRẮC NGHIỆM TỔNG QUÁT - dùng được với file .docx / .txt bất kỳ.

Cài thư viện (1 lần):  pip install python-docx
Chạy:                  python trac_nghiem.py [duong_dan_file]

File đề được nhận dạng như sau:
  - Câu hỏi:   "Câu 1." / "Câu 1:" / "1." / "1)"  (có thể nhiều dòng)
  - Đáp án:    "A. ...", "B) ..." (mỗi dòng 1 đáp án hoặc nhiều đáp án trên 1 dòng)
  - Đáp án ĐÚNG được tìm theo thứ tự ưu tiên:
      1) Dòng "Đáp án: B" / "ĐA: B" / "Answer: B" ngay sau câu hỏi
      2) Đáp án được tô ĐẬM / gạch chân / tô màu / highlight (chỉ 1 đáp án trong câu)
      3) Đáp án có dấu * ở đầu hoặc cuối
      4) Bảng đáp án ở cuối file: "1-A 2-B 3-C ..." hoặc mỗi dòng "1. B"
  - Câu nào không tìm được đáp án đúng sẽ bị bỏ qua (có thông báo).
"""
import json
import os
import random
import re
import sys
import time
from datetime import datetime

# ====================== ĐỌC FILE ======================

def _la_danh_dau(run):
    """Run có được nhấn mạnh (đậm / gạch chân / highlight / màu) không."""
    try:
        if run.bold:
            return True
        if run.underline not in (None, False):
            return True
        if run.font.highlight_color is not None:
            return True
        mau = run.font.color
        if mau is not None and mau.rgb is not None and str(mau.rgb) not in ("000000", "auto"):
            return True
    except Exception:
        pass
    return False


def _tach_dong(text, flags):
    """Tách đoạn có xuống dòng mềm (\\n) thành nhiều dòng, giữ cờ định dạng."""
    kq, bat_dau = [], 0
    for phan in text.split("\n"):
        kq.append((phan, flags[bat_dau:bat_dau + len(phan)]))
        bat_dau += len(phan) + 1
    return kq


def doc_cac_dong(duong_dan):
    """Trả về danh sách (text, flags) theo đúng thứ tự trong file."""
    if duong_dan.lower().endswith(".txt"):
        with open(duong_dan, encoding="utf-8-sig", errors="ignore") as f:
            return [(d.rstrip("\n"), [False] * len(d.rstrip("\n"))) for d in f]

    try:
        from docx import Document
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError:
        print("❌ Thiếu thư viện. Hãy chạy:  pip install python-docx")
        sys.exit(1)

    doc = Document(duong_dan)
    ket_qua = []

    def them_paragraph(p):
        chars, flags = [], []
        for r in p.runs:
            if r.text:
                chars.append(r.text)
                flags += [_la_danh_dau(r)] * len(r.text)
        ket_qua.extend(_tach_dong("".join(chars), flags))

    for phan_tu in doc.element.body.iterchildren():
        tag = phan_tu.tag.split("}")[-1]
        if tag == "p":
            them_paragraph(Paragraph(phan_tu, doc))
        elif tag == "tbl":
            for hang in Table(phan_tu, doc).rows:
                da_xet = set()
                for o in hang.cells:
                    if id(o._tc) in da_xet:
                        continue
                    da_xet.add(id(o._tc))
                    for p in o.paragraphs:
                        them_paragraph(p)
    return ket_qua


# ====================== PHÂN TÍCH CÂU HỎI ======================

RE_CAU = re.compile(r"^\s*(?:câu|cau|question|q)?\s*(\d+)\s*[\.\):\-]\s*(.*)$", re.I)
RE_OPT = re.compile(r"(?:^|(?<=\s))([A-Fa-f])[\.\)]\s*")
RE_DAPAN = re.compile(
    r"^\s*(?:đáp\s*án(?:\s*đúng)?|đ\.?\s?á\.?|answer|ans|key)\s*[:\-=\.]?\s*\(?([A-Fa-f])\b", re.I)
RE_BANG = re.compile(r"(\d+)\s*[\.\-:\)]?\s*([A-Fa-f])\b")
RE_DONG_BANG = re.compile(r"^(?:\s*\d+\s*[\.\-:\)]?\s*[A-Fa-f]\s*[,;]?)+\s*$")


def _tim_dap_an(text, ky_tu_ke_tiep):
    ket_qua = []
    for m in RE_OPT.finditer(text):
        chu = m.group(1).upper()
        if chu == ky_tu_ke_tiep:
            ket_qua.append((chu, m.start(), m.end()))
            ky_tu_ke_tiep = chr(ord(ky_tu_ke_tiep) + 1)
    return ket_qua


def _them_dap_an(cau, text, flags, tim_thay):
    for i, (_, _, kt) in enumerate(tim_thay):
        ket = tim_thay[i + 1][1] if i + 1 < len(tim_thay) else len(text)
        doan, co = text[kt:ket], flags[kt:ket]
        sao = doan.strip().startswith("") or doan.strip().endswith("")
        doan = doan.strip().strip("*").strip()
        ky_tu = [c for c, f in zip(text[kt:ket], co) if not c.isspace()]
        so_nhan = sum(1 for c, f in zip(text[kt:ket], co) if not c.isspace() and f)
        danh_dau = sao or (len(ky_tu) > 0 and so_nhan / len(ky_tu) >= 0.5)
        cau["options"].append([doan, danh_dau])


def phan_tich(cac_dong):
    cau_hoi, bang_dap_an, cur = [], {}, None

    def ket_thuc(c):
        if c is None:
            return
        opts = c["options"]
        if len(opts) >= 2:
            dung = None
            if c["ans"] is not None and c["ans"] < len(opts):
                dung = c["ans"]
            else:
                dd = [i for i, o in enumerate(opts) if o[1]]
                if len(dd) == 1:
                    dung = dd[0]
            c["dung"] = dung
            cau_hoi.append(c)
        elif not opts and re.fullmatch(r"[A-Fa-f]", c["text"].strip()):
            bang_dap_an[c["num"]] = ord(c["text"].strip().upper()) - 65

    for text, flags in cac_dong:
        if not text.strip():
            continue
        # Dòng bảng đáp án "1-A 2-B 3-C"
        if RE_DONG_BANG.match(text) and len(RE_BANG.findall(text)) >= 2:
            ket_thuc(cur)
            cur = None
            for so, chu in RE_BANG.findall(text):
                bang_dap_an[int(so)] = ord(chu.upper()) - 65
            continue
        # Dòng "Đáp án: B"
        m = RE_DAPAN.match(text)
        if m and cur is not None:
            cur["ans"] = ord(m.group(1).upper()) - 65
            continue
        # Dòng đáp án tiếp theo của câu hiện tại
        if cur is not None:
            ke_tiep = chr(65 + len(cur["options"]))
            tt = _tim_dap_an(text, ke_tiep)
            if tt and (tt[0][1] == 0 or len(tt) >= 2 or cur["options"]):
                if tt[0][1] > 0:
                    # phần chữ trước đáp án đầu tiên thuộc về câu hỏi / đáp án trước
                    truoc = text[:tt[0][1]].strip()
                    if cur["options"]:
                        cur["options"][-1][0] += " " + truoc
                    else:
                        cur["text"] += " " + truoc
                _them_dap_an(cur, text, flags, tt)
                continue
        # Câu hỏi mới
        m = RE_CAU.match(text)
        if m:
            ket_thuc(cur)
            noi_dung = m.group(2).strip()
            cur = {"num": int(m.group(1)), "text": noi_dung, "options": [], "ans": None}
            tt = _tim_dap_an(text, "A")
            tt = [t for t in tt if t[1] >= m.start(2)]
            if len(tt) >= 2:
                cur["text"] = text[m.start(2):tt[0][1]].strip()
                _them_dap_an(cur, text, flags, tt)
            continue
        # Dòng nối tiếp
        if cur is not None:
            if cur["options"]:
                cur["options"][-1][0] += " " + text.strip()
            else:
                cur["text"] += " " + text.strip()
    ket_thuc(cur)

    # Gắn đáp án từ bảng đáp án cuối file
    for c in cau_hoi:
        if c["dung"] is None and c["num"] in bang_dap_an and bang_dap_an[c["num"]] < len(c["options"]):
            c["dung"] = bang_dap_an[c["num"]]

    hop_le, bo_qua = [], []
    for c in cau_hoi:
        if c["dung"] is None:
            bo_qua.append(c["num"])
        else:
            hop_le.append({
                "num": c["num"],
                "text": c["text"].strip(),
                "options": [o[0] for o in c["options"]],
                "correct": c["options"][c["dung"]][0],
            })
    return hop_le, bo_qua


# ====================== XÁO TRỘN ======================

RE_THAM_CHIEU = re.compile(
    r"\b[A-F]\b\s*(?:,|và|and|&|hoặc|or)\s*\b[A-F]\b|"
    r"(?:tất cả|cả hai|all of|both|none of|không có).{0,25}(?:trên|above|đáp án|đúng)", re.I)


def tron_cau(q):
    """Xáo thứ tự đáp án (giữ nguyên nếu có kiểu 'Cả A và B đúng')."""
    opts = list(q["options"])
    if not any(RE_THAM_CHIEU.search(o) for o in opts):
        random.shuffle(opts)
    return {"num": q["num"], "text": q["text"], "options": opts, "correct": opts.index(q["correct"])}


def chon_de(tat_ca, so_cau, da_dung):
    """Chọn ngẫu nhiên, ưu tiên câu chưa làm; hết thì quay vòng mới."""
    con_lai = [i for i in range(len(tat_ca)) if i not in da_dung]
    if len(con_lai) >= so_cau:
        chon = random.sample(con_lai, so_cau)
        da_dung.update(chon)
    else:
        chon = con_lai[:]
        da_dung.clear()
        bu = [i for i in range(len(tat_ca)) if i not in chon]
        chon += random.sample(bu, so_cau - len(chon))
        da_dung.update(chon)
    random.shuffle(chon)
    return [tron_cau(tat_ca[i]) for i in chon]


# ====================== LÀM BÀI ======================

def lam_bai(de, phut):
    gioi_han = phut * 60
    bat_dau = time.time()
    tra_loi = {}
    print("\n" + "=" * 60)
    print(f"  BẮT ĐẦU — {len(de)} câu — {phut} phút")
    print("  Nhập A/B/C/D | Enter = bỏ qua | NOP = nộp bài")
    print("=" * 60)

    for idx, q in enumerate(de):
        con_lai = gioi_han - (time.time() - bat_dau)
        if con_lai <= 0:
            print("\n⏰ Hết giờ! Tự động nộp bài.")
            return tra_loi, time.time() - bat_dau, "hết giờ"
        print(f"\n--- Câu {idx + 1}/{len(de)} | còn {int(con_lai // 60):02d}:{int(con_lai % 60):02d} ---")
        print(q["text"])
        for i, o in enumerate(q["options"]):
            print(f"  {chr(65 + i)}. {o}")
        hop_le = [chr(65 + i) for i in range(len(q["options"]))]
        while True:
            s = input("👉 ").strip().upper()
            if s == "NOP":
                return tra_loi, time.time() - bat_dau, "nộp sớm"
            if s == "":
                break
            if s in hop_le:
                tra_loi[idx] = ord(s) - 65
                break
            print(f"⚠️ Nhập {'/'.join(hop_le)}, Enter để bỏ qua hoặc NOP")
        if time.time() - bat_dau > gioi_han:
            print("\n⏰ Hết giờ! Tự động nộp bài.")
            return tra_loi, time.time() - bat_dau, "hết giờ"
    return tra_loi, time.time() - bat_dau, "làm xong"


def ket_qua(de, tra_loi, thoi_gian, trang_thai):
    dung, sai_ds, bo_trong = 0, [], 0
    for idx, q in enumerate(de):
        if idx not in tra_loi:
            bo_trong += 1
            sai_ds.append((idx, None))
        elif tra_loi[idx] == q["correct"]:
            dung += 1
        else:
            sai_ds.append((idx, tra_loi[idx]))
    diem = dung / len(de) * 10 if de else 0

    print("\n" + "=" * 60)
    print("  KẾT QUẢ")
    print("=" * 60)
    print(f"Đúng: {dung} | Sai: {len(sai_ds) - bo_trong} | Bỏ trống: {bo_trong} | Tổng: {len(de)}")
    print(f"Điểm: {diem:.2f}/10   — Thời gian: {int(thoi_gian // 60)}p{int(thoi_gian % 60):02d}s ({trang_thai})")

    if sai_ds:
        xem = input("\nXem chi tiết các câu sai/bỏ trống? (C/K): ").strip().upper()
        if xem == "C":
            for idx, chon in sai_ds:
                q = de[idx]
                print(f"\nCâu {idx + 1}: {q['text']}")
                chon_txt = "Chưa làm" if chon is None else f"{chr(65 + chon)}. {q['options'][chon]}"
                print(f"  Bạn chọn : {chon_txt}")
                print(f"  Đáp án   : {chr(65 + q['correct'])}. {q['options'][q['correct']]}")
    return diem, [de[i] for i, _ in sai_ds]


def luu_lich_su(ten_file, diem, tong, trang_thai):
    duong_dan = os.path.join(os.path.dirname(os.path.abspath(_file_)), "lich_su_diem.json")
    try:
        du_lieu = json.load(open(duong_dan, encoding="utf-8")) if os.path.exists(duong_dan) else []
        du_lieu.append({"thoi_gian": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "file": ten_file, "so_cau": tong, "diem": round(diem, 2), "ket_thuc": trang_thai})
        json.dump(du_lieu, open(duong_dan, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    except Exception:
        pass


# ====================== CHƯƠNG TRÌNH CHÍNH ======================

def chon_file():
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        return sys.argv[1]
    cac_thu_muc = {os.getcwd(), os.path.dirname(os.path.abspath(_file_))}
    ds = []
    for d in cac_thu_muc:
        for f in sorted(os.listdir(d)):
            if f.lower().endswith((".docx", ".txt")) and not f.startswith("~$"):
                ds.append(os.path.join(d, f))
    if not ds:
        return input("Không thấy file đề. Nhập đường dẫn file .docx/.txt: ").strip().strip('"')
    if len(ds) == 1:
        return ds[0]
    print("Các file tìm thấy:")
    for i, f in enumerate(ds, 1):
        print(f"  {i}. {os.path.basename(f)}")
    while True:
        s = input("Chọn số file: ").strip()
        if s.isdigit() and 1 <= int(s) <= len(ds):
            return ds[int(s) - 1]


def hoi_so(cau_hoi, mac_dinh, thap, cao):
    while True:
        s = input(f"{cau_hoi} [{mac_dinh}]: ").strip()
        if s == "":
            return mac_dinh
        if s.isdigit() and thap <= int(s) <= cao:
            return int(s)
        print(f"⚠️ Nhập số từ {thap} đến {cao}")


def main():
    print("📚 CHƯƠNG TRÌNH TRẮC NGHIỆM")
    print("-" * 50)
    duong_dan = chon_file()
    if not os.path.isfile(duong_dan):
        print("❌ Không tìm thấy file.")
        return
    print(f"✅ Đang đọc: {os.path.basename(duong_dan)}")
    tat_ca, bo_qua = phan_tich(doc_cac_dong(duong_dan))
    if not tat_ca:
        print("❌ Không nhận dạng được câu hỏi nào có đáp án đúng.")
        print("   Hãy kiểm tra định dạng (xem phần mô tả ở đầu file .py).")
        return
    print(f"📖 Nạp được {len(tat_ca)} câu hợp lệ.")
    if bo_qua:
        print(f"⚠️ Bỏ qua {len(bo_qua)} câu không xác định được đáp án đúng: {bo_qua[:15]}{'...' if len(bo_qua) > 15 else ''}")

    so_cau = hoi_so("Số câu mỗi đề", min(60, len(tat_ca)), 1, len(tat_ca))
    phut = hoi_so("Thời gian (phút)", 60, 1, 600)

    da_dung = set()
    de = chon_de(tat_ca, so_cau, da_dung)
    while True:
        tra_loi, tg, tt = lam_bai(de, phut)
        diem, cau_sai = ket_qua(de, tra_loi, tg, tt)
        luu_lich_su(os.path.basename(duong_dan), diem, len(de), tt)

        print("\n[C] Đề mới   [L] Làm lại câu sai   [K] Thoát")
        while True:
            c = input("Chọn: ").strip().upper()
            if c == "L" and not cau_sai:
                print("Không có câu sai để làm lại.")
            elif c in ("C", "L", "K"):
                break
        if c == "K":
            print("👋 Tạm biệt!")
            return
        if c == "C":
            de = chon_de(tat_ca, so_cau, da_dung)
        else:
            de = [tron_cau(q) for q in cau_sai]
            random.shuffle(de)


if __name__ == "_main_":
    main()
