import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import openpyxl

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "matching_config.json")

DEFAULT_CONFIG = {
    "settings": {
        "default_manager": "온라인",
        "excel_columns": {
            "order_date": 0, "name": 1, "address": 2, "phone": 3,
            "product": 4, "option": 5, "quantity": 6, "zipcode": 7,
            "message": 8
        }
    },
    "vendors": {}
}


def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_CONFIG.copy()


def save_config(config):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("더해피 발주 자동분류")
        self.geometry("1100x700")
        self.config_data = load_config()
        self.orders = []

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=5, pady=5)

        self.tab_order = ttk.Frame(nb)
        self.tab_matching = ttk.Frame(nb)
        self.tab_vendor = ttk.Frame(nb)
        self.tab_settings = ttk.Frame(nb)

        nb.add(self.tab_order, text="발주 처리")
        nb.add(self.tab_matching, text="매칭 관리")
        nb.add(self.tab_vendor, text="업체 관리")
        nb.add(self.tab_settings, text="설정")

        self._build_order_tab()
        self._build_matching_tab()
        self._build_vendor_tab()
        self._build_settings_tab()

    # ── 발주 처리 탭 ──
    def _build_order_tab(self):
        top = ttk.Frame(self.tab_order)
        top.pack(fill="x", padx=5, pady=5)

        ttk.Button(top, text="엑셀 파일 열기", command=self._load_excel).pack(side="left")
        self.lbl_file = ttk.Label(top, text="파일을 선택하세요")
        self.lbl_file.pack(side="left", padx=10)

        # 주문 테이블
        cols = ("주문일시", "수취인", "주소", "전화번호", "상품명", "옵션이름", "수량", "배송메세지", "매칭업체", "품목")
        self.order_tree = ttk.Treeview(self.tab_order, columns=cols, show="headings", height=20)
        for c in cols:
            self.order_tree.heading(c, text=c)
            self.order_tree.column(c, width=100 if c not in ("주소",) else 200)
        self.order_tree.pack(fill="both", expand=True, padx=5)

        # 태그 색상
        self.order_tree.tag_configure("unmatched", background="#ffcccc")
        self.order_tree.tag_configure("matched", background="#ccffcc")

        bottom = ttk.Frame(self.tab_order)
        bottom.pack(fill="x", padx=5, pady=5)

        self.lbl_status = ttk.Label(bottom, text="")
        self.lbl_status.pack(side="left")
        ttk.Button(bottom, text="자동 매칭 실행", command=self._run_matching).pack(side="right", padx=5)
        ttk.Button(bottom, text="부반장제어 저장", command=self._save_temp_excel).pack(side="right", padx=5)
        ttk.Button(bottom, text="구글시트 전송", command=self._send_to_sheets).pack(side="right", padx=5)

    def _load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self.lbl_file.config(text=os.path.basename(path))
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        self.orders = []
        col = self.config_data["settings"]["excel_columns"]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if not any(row):
                continue
            self.orders.append({
                "date": str(row[col["order_date"]] or ""),
                "name": str(row[col["name"]] or ""),
                "address": str(row[col["address"]] or ""),
                "phone": str(row[col["phone"]] or ""),
                "product": str(row[col["product"]] or ""),
                "option": str(row[col["option"]] or ""),
                "quantity": row[col["quantity"]] or 0,
                "message": str(row[col["message"]] or ""),
                "vendor": None,
                "item_name": None,
            })
        self._refresh_order_tree()
        self.lbl_status.config(text=f"총 {len(self.orders)}건 로드됨")

    def _refresh_order_tree(self):
        self.order_tree.delete(*self.order_tree.get_children())
        for o in self.orders:
            tag = "matched" if o["vendor"] else "unmatched"
            self.order_tree.insert("", "end", values=(
                o["date"], o["name"], o["address"], o["phone"],
                o["product"], o["option"], o["quantity"], o["message"],
                o["vendor"] or "미매칭", o["item_name"] or ""
            ), tags=(tag,))

    def _build_matching_lookup(self):
        """상품명/옵션이름 → (vendor_id, item_name) 매핑 딕셔너리 생성"""
        lookup = {}
        for vid, v in self.config_data.get("vendors", {}).items():
            for keyword, pinfo in v.get("products", {}).items():
                item = pinfo["item_name"] if isinstance(pinfo, dict) else pinfo
                lookup[keyword] = (vid, v["name"], item)
        return lookup

    def _run_matching(self):
        if not self.orders:
            messagebox.showwarning("알림", "엑셀을 먼저 로드하세요")
            return
        lookup = self._build_matching_lookup()
        matched = 0
        for o in self.orders:
            key = o["option"] if o["option"] else o["product"]
            if key in lookup:
                _, o["vendor"], o["item_name"] = lookup[key]
                matched += 1
            else:
                o["vendor"] = None
                o["item_name"] = None
        self._refresh_order_tree()
        unmatched = len(self.orders) - matched
        msg = f"매칭 완료: {matched}건 성공"
        if unmatched:
            msg += f", {unmatched}건 미매칭"
        self.lbl_status.config(text=msg)
        if unmatched:
            messagebox.showwarning("미매칭 경고", f"{unmatched}건의 미매칭 상품이 있습니다.\n매칭 관리 탭에서 등록해주세요.")

    def _save_temp_excel(self):
        if not self.orders:
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="부반장제어_임시.xlsx")
        if not path:
            return
        wb = openpyxl.Workbook()
        vendors_data = {}
        for o in self.orders:
            vname = o["vendor"] or "미매칭"
            vendors_data.setdefault(vname, []).append(o)

        first = True
        for vname, rows in vendors_data.items():
            ws = wb.active if first else wb.create_sheet()
            first = False
            ws.title = vname[:31]
            ws.append(["상호", "주소", "전화번호", "품목", "수량", "택배비", "송장번호", "담당", "특이사항"])
            mgr = self.config_data["settings"].get("default_manager", "온라인")
            for o in rows:
                ws.append([o["name"], o["address"], o["phone"], o["item_name"] or "", o["quantity"], "", "", mgr, o["message"]])

        wb.save(path)
        messagebox.showinfo("저장 완료", f"부반장제어 파일 저장: {path}")

    def _send_to_sheets(self):
        messagebox.showinfo("안내", "구글시트 전송 기능은 Google API 인증 설정 후 활성화됩니다.")

    # ── 매칭 관리 탭 ──
    def _build_matching_tab(self):
        top = ttk.Frame(self.tab_matching)
        top.pack(fill="x", padx=5, pady=5)

        ttk.Label(top, text="업체 선택:").pack(side="left")
        self.match_vendor_var = tk.StringVar()
        self.match_vendor_cb = ttk.Combobox(top, textvariable=self.match_vendor_var, state="readonly", width=30)
        self.match_vendor_cb.pack(side="left", padx=5)
        self.match_vendor_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_match_tree())
        ttk.Button(top, text="새로고침", command=self._refresh_vendor_lists).pack(side="left")

        # 매칭 테이블
        cols = ("상품키워드", "품목명", "택배비유형", "택배비금액")
        self.match_tree = ttk.Treeview(self.tab_matching, columns=cols, show="headings", height=18)
        for c in cols:
            self.match_tree.heading(c, text=c)
            self.match_tree.column(c, width=200)
        self.match_tree.pack(fill="both", expand=True, padx=5)

        # 입력 폼
        form = ttk.LabelFrame(self.tab_matching, text="매칭 추가/수정")
        form.pack(fill="x", padx=5, pady=5)

        r = ttk.Frame(form)
        r.pack(fill="x", padx=5, pady=2)
        ttk.Label(r, text="상품키워드(상품명/옵션이름):").pack(side="left")
        self.match_keyword = ttk.Entry(r, width=30)
        self.match_keyword.pack(side="left", padx=5)
        ttk.Label(r, text="품목명:").pack(side="left")
        self.match_item = ttk.Entry(r, width=20)
        self.match_item.pack(side="left", padx=5)
        ttk.Label(r, text="택배비유형:").pack(side="left")
        self.match_ship_type = ttk.Combobox(r, values=["free", "fixed", "variable", "logen_calc"], width=10, state="readonly")
        self.match_ship_type.pack(side="left", padx=5)
        self.match_ship_type.set("free")
        ttk.Label(r, text="택배비금액:").pack(side="left")
        self.match_ship_amt = ttk.Entry(r, width=10)
        self.match_ship_amt.pack(side="left", padx=5)

        btns = ttk.Frame(form)
        btns.pack(fill="x", padx=5, pady=2)
        ttk.Button(btns, text="추가", command=self._add_match).pack(side="left", padx=2)
        ttk.Button(btns, text="선택 삭제", command=self._del_match).pack(side="left", padx=2)

        self.match_tree.bind("<<TreeviewSelect>>", self._on_match_select)
        self._refresh_vendor_lists()

    def _refresh_vendor_lists(self):
        names = [v["name"] for v in self.config_data.get("vendors", {}).values()]
        self.match_vendor_cb["values"] = names
        if hasattr(self, "vendor_tree"):
            self._refresh_vendor_tree()

    def _refresh_match_tree(self):
        self.match_tree.delete(*self.match_tree.get_children())
        vname = self.match_vendor_var.get()
        for vid, v in self.config_data.get("vendors", {}).items():
            if v["name"] == vname:
                for kw, pinfo in v.get("products", {}).items():
                    if isinstance(pinfo, dict):
                        self.match_tree.insert("", "end", values=(kw, pinfo.get("item_name", ""), pinfo.get("shipping_type", "free"), pinfo.get("shipping_fee", "")))
                    else:
                        self.match_tree.insert("", "end", values=(kw, pinfo, "free", ""))
                break

    def _on_match_select(self, event):
        sel = self.match_tree.selection()
        if sel:
            vals = self.match_tree.item(sel[0], "values")
            self.match_keyword.delete(0, "end"); self.match_keyword.insert(0, vals[0])
            self.match_item.delete(0, "end"); self.match_item.insert(0, vals[1])
            self.match_ship_type.set(vals[2])
            self.match_ship_amt.delete(0, "end"); self.match_ship_amt.insert(0, vals[3])

    def _find_vendor_id(self, name):
        for vid, v in self.config_data.get("vendors", {}).items():
            if v["name"] == name:
                return vid
        return None

    def _add_match(self):
        vname = self.match_vendor_var.get()
        vid = self._find_vendor_id(vname)
        if not vid:
            messagebox.showwarning("알림", "업체를 선택하세요")
            return
        kw = self.match_keyword.get().strip()
        item = self.match_item.get().strip()
        if not kw or not item:
            messagebox.showwarning("알림", "상품키워드와 품목명을 입력하세요")
            return
        self.config_data["vendors"][vid]["products"][kw] = {
            "item_name": item,
            "shipping_type": self.match_ship_type.get(),
            "shipping_fee": self.match_ship_amt.get().strip() or None
        }
        save_config(self.config_data)
        self._refresh_match_tree()

    def _del_match(self):
        sel = self.match_tree.selection()
        if not sel:
            return
        kw = self.match_tree.item(sel[0], "values")[0]
        vname = self.match_vendor_var.get()
        vid = self._find_vendor_id(vname)
        if vid and kw in self.config_data["vendors"][vid]["products"]:
            del self.config_data["vendors"][vid]["products"][kw]
            save_config(self.config_data)
            self._refresh_match_tree()

    # ── 업체 관리 탭 ──
    def _build_vendor_tab(self):
        cols = ("업체ID", "업체명", "시트이름", "구글시트URL", "담당고정값")
        self.vendor_tree = ttk.Treeview(self.tab_vendor, columns=cols, show="headings", height=18)
        for c in cols:
            self.vendor_tree.heading(c, text=c)
            w = 120 if c != "구글시트URL" else 300
            self.vendor_tree.column(c, width=w)
        self.vendor_tree.pack(fill="both", expand=True, padx=5, pady=5)

        form = ttk.LabelFrame(self.tab_vendor, text="업체 추가/수정")
        form.pack(fill="x", padx=5, pady=5)

        r1 = ttk.Frame(form); r1.pack(fill="x", padx=5, pady=2)
        ttk.Label(r1, text="업체ID:").pack(side="left")
        self.v_id = ttk.Entry(r1, width=15); self.v_id.pack(side="left", padx=5)
        ttk.Label(r1, text="업체명:").pack(side="left")
        self.v_name = ttk.Entry(r1, width=20); self.v_name.pack(side="left", padx=5)
        ttk.Label(r1, text="시트이름:").pack(side="left")
        self.v_sheet = ttk.Entry(r1, width=25); self.v_sheet.pack(side="left", padx=5)

        r2 = ttk.Frame(form); r2.pack(fill="x", padx=5, pady=2)
        ttk.Label(r2, text="구글시트URL:").pack(side="left")
        self.v_url = ttk.Entry(r2, width=60); self.v_url.pack(side="left", padx=5)
        ttk.Label(r2, text="담당고정값:").pack(side="left")
        self.v_mgr = ttk.Entry(r2, width=10); self.v_mgr.pack(side="left", padx=5)

        btns = ttk.Frame(form); btns.pack(fill="x", padx=5, pady=2)
        ttk.Button(btns, text="추가/수정", command=self._save_vendor).pack(side="left", padx=2)
        ttk.Button(btns, text="선택 삭제", command=self._del_vendor).pack(side="left", padx=2)

        self.vendor_tree.bind("<<TreeviewSelect>>", self._on_vendor_select)
        self._refresh_vendor_tree()

    def _refresh_vendor_tree(self):
        self.vendor_tree.delete(*self.vendor_tree.get_children())
        for vid, v in self.config_data.get("vendors", {}).items():
            self.vendor_tree.insert("", "end", values=(
                vid, v.get("name", ""), v.get("sheet_name", ""),
                v.get("google_sheet_url", ""), v.get("default_manager", "")
            ))

    def _on_vendor_select(self, event):
        sel = self.vendor_tree.selection()
        if sel:
            vals = self.vendor_tree.item(sel[0], "values")
            for entry, val in zip([self.v_id, self.v_name, self.v_sheet, self.v_url, self.v_mgr], vals):
                entry.delete(0, "end"); entry.insert(0, val)

    def _save_vendor(self):
        vid = self.v_id.get().strip()
        name = self.v_name.get().strip()
        if not vid or not name:
            messagebox.showwarning("알림", "업체ID와 업체명을 입력하세요")
            return
        if vid not in self.config_data["vendors"]:
            self.config_data["vendors"][vid] = {"products": {}}
        self.config_data["vendors"][vid].update({
            "name": name,
            "sheet_name": self.v_sheet.get().strip(),
            "google_sheet_url": self.v_url.get().strip(),
            "default_manager": self.v_mgr.get().strip() or self.config_data["settings"]["default_manager"]
        })
        save_config(self.config_data)
        self._refresh_vendor_tree()
        self._refresh_vendor_lists()

    def _del_vendor(self):
        sel = self.vendor_tree.selection()
        if not sel:
            return
        vid = self.vendor_tree.item(sel[0], "values")[0]
        if messagebox.askyesno("확인", f"'{vid}' 업체를 삭제하시겠습니까?"):
            self.config_data["vendors"].pop(vid, None)
            save_config(self.config_data)
            self._refresh_vendor_tree()
            self._refresh_vendor_lists()

    # ── 설정 탭 ──
    def _build_settings_tab(self):
        f = ttk.LabelFrame(self.tab_settings, text="기본 설정")
        f.pack(fill="x", padx=10, pady=10)

        r = ttk.Frame(f); r.pack(fill="x", padx=5, pady=5)
        ttk.Label(r, text="담당 고정값 (기본):").pack(side="left")
        self.set_mgr = ttk.Entry(r, width=15)
        self.set_mgr.pack(side="left", padx=5)
        self.set_mgr.insert(0, self.config_data["settings"].get("default_manager", "온라인"))

        ttk.Button(f, text="저장", command=self._save_settings).pack(padx=5, pady=5, anchor="w")

        info = ttk.LabelFrame(self.tab_settings, text="엑셀 컬럼 매핑 (0부터 시작)")
        info.pack(fill="x", padx=10, pady=10)

        self.col_entries = {}
        for i, (key, default) in enumerate(DEFAULT_CONFIG["settings"]["excel_columns"].items()):
            r = ttk.Frame(info); r.pack(fill="x", padx=5, pady=1)
            ttk.Label(r, text=f"{key}:", width=15).pack(side="left")
            e = ttk.Entry(r, width=5)
            e.pack(side="left")
            e.insert(0, str(self.config_data["settings"]["excel_columns"].get(key, default)))
            self.col_entries[key] = e

    def _save_settings(self):
        self.config_data["settings"]["default_manager"] = self.set_mgr.get().strip()
        for key, entry in self.col_entries.items():
            try:
                self.config_data["settings"]["excel_columns"][key] = int(entry.get())
            except ValueError:
                pass
        save_config(self.config_data)
        messagebox.showinfo("저장", "설정이 저장되었습니다.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
