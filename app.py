import os
import streamlit as st
import pandas as pd
import io
from itertools import groupby
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pymongo import MongoClient
from bson import ObjectId
from streamlit_cookies_controller import CookieController

# ── Load .env ─────────────────────────────────────────────────────────────────
load_dotenv("credentials/.env")

MONGO_URI    = os.getenv("MONGO_URI")
ADMIN_USER   = os.getenv("ADMIN_USER")
ADMIN_PASS   = os.getenv("ADMIN_PASS")
STUDENT_USER = os.getenv("STUDENT_USER")
STUDENT_PASS = os.getenv("STUDENT_PASS")

USERS = {
    ADMIN_USER:   ADMIN_PASS,
    STUDENT_USER: STUDENT_PASS,
}

COOKIE_NAME   = "packing_list_user"
COOKIE_EXPIRY = 30  # days

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Study Abroad Packing List", page_icon="🧳", layout="wide")

# ── Cookie manager ────────────────────────────────────────────────────────────
cookie_manager = CookieController()

# ── MongoDB ───────────────────────────────────────────────────────────────────
@st.cache_resource
def get_col():
    client = MongoClient(MONGO_URI)
    return client["packing_list_db"]["items"]

def load_items():
    items = list(get_col().find().sort([("category", 1), ("name", 1)]))
    for item in items:
        item["id"]      = str(item["_id"])
        item["checked"] = int(item.get("checked", 0))
        item["note"]    = item.get("note", "")
        item["budget"]  = item.get("budget", "")
    return items

def toggle_item(item_id, checked):
    get_col().update_one({"_id": ObjectId(item_id)}, {"$set": {"checked": 1 if checked else 0}})

def add_item(name, category, note="", budget=""):
    get_col().insert_one({"category": category, "name": name, "note": note, "budget": budget, "checked": 0})

def delete_item(item_id):
    get_col().delete_one({"_id": ObjectId(item_id)})

def update_item_fields(item_id, fields: dict):
    get_col().update_one({"_id": ObjectId(item_id)}, {"$set": fields})

def get_categories():
    return sorted(get_col().distinct("category"))

# ── Excel export ──────────────────────────────────────────────────────────────
def build_excel(items):
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing List"

    green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    cat_fill    = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    thin   = Side(border_style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers    = ["#", "Category", "Item Name", "Note", "Budget", "Packed"]
    col_widths = [5, 28, 35, 45, 20, 10]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell           = ws.cell(row=1, column=ci, value=h)
        cell.font      = Font(bold=True, color="FFFFFF", size=11)
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = border
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22

    items_sorted = sorted(items, key=lambda x: (x["category"], x["name"]))
    row = 2
    idx = 1
    for cat, group in groupby(items_sorted, key=lambda x: x["category"]):
        group_list = list(group)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        cat_cell           = ws.cell(row=row, column=1, value=f"  {cat}")
        cat_cell.fill      = cat_fill
        cat_cell.font      = Font(bold=True, size=11, color="1F4E79")
        cat_cell.alignment = Alignment(vertical="center")
        cat_cell.border    = border
        ws.row_dimensions[row].height = 18
        row += 1

        for item in group_list:
            is_checked = bool(item["checked"])
            values     = [idx, item["category"], item["name"], item["note"] or "", item.get("budget", "") or "", "✓" if is_checked else ""]
            for ci, val in enumerate(values, 1):
                cell           = ws.cell(row=row, column=ci, value=val)
                cell.border    = border
                cell.alignment = Alignment(vertical="center", wrap_text=(ci == 4))
                if is_checked:
                    cell.fill = green_fill
                    cell.font = Font(color="276221")
                if ci == 6:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[row].height = 16
            row += 1
            idx += 1

    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Excel import ──────────────────────────────────────────────────────────────
def restore_from_excel(file):
    try:
        df = pd.read_excel(file, sheet_name="Packing List")
    except Exception as e:
        return 0, f"Could not read sheet 'Packing List': {e}"

    df.columns = [c.strip().lower() for c in df.columns]
    required   = {"category", "item name", "note", "packed"}
    if not required.issubset(set(df.columns)):
        return 0, f"Missing columns. Expected: {required}. Found: {set(df.columns)}"

    df = df.dropna(subset=["item name"])
    df = df[df["item name"].astype(str).str.strip() != ""]

    col      = get_col()
    existing = set(
        (d["category"].strip().lower(), d["name"].strip().lower())
        for d in col.find({}, {"category": 1, "name": 1})
    )

    unique_rows = []
    duplicates  = []
    for _, r in df.iterrows():
        cat        = str(r["category"]).strip()
        name       = str(r["item name"]).strip()
        note       = "" if pd.isna(r["note"]) else str(r["note"]).strip()
        budget     = "" if ("budget" not in r or pd.isna(r["budget"])) else str(r["budget"]).strip()
        packed_val = str(r["packed"]).strip()
        checked    = 1 if packed_val in ("✓", "checkmark", "1", "True", "true", "yes", "Yes") else 0
        if not cat or not name:
            continue
        key = (cat.lower(), name.lower())
        if key in existing:
            duplicates.append(name)
        else:
            existing.add(key)
            unique_rows.append({"category": cat, "name": name, "note": note, "budget": budget, "checked": checked})

    if not unique_rows and not duplicates:
        return 0, "No valid data rows found in the file."

    if unique_rows:
        col.insert_many(unique_rows)

    msg = ""
    if duplicates:
        dup_list = ", ".join(f'"{d}"' for d in duplicates[:5])
        more     = f" (+{len(duplicates)-5} more)" if len(duplicates) > 5 else ""
        msg      = f"Skipped {len(duplicates)} duplicate(s): {dup_list}{more}"

    return len(unique_rows), msg

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
#MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 1rem;}
.cat-header {
    background: #EFF6FF; border-left: 4px solid #2563EB;
    padding: 8px 14px; border-radius: 0 8px 8px 0;
    font-weight: 600; font-size: 15px; color: #1E40AF;
    margin: 16px 0 6px 0;
}
.stat-box {
    background: #F0F9FF; border: 1px solid #BAE6FD;
    border-radius: 10px; padding: 10px 18px;
    text-align: center; margin: 4px;
}
.stat-num { font-size: 22px; font-weight: 700; color: #0369A1; }
.stat-lbl { font-size: 12px; color: #64748B; }
.budget-total-box {
    background: #F0FFF4; border: 1px solid #86EFAC;
    border-radius: 10px; padding: 12px 20px;
    margin-top: 20px; text-align: center;
}
.budget-total-num { font-size: 22px; font-weight: 700; color: #15803D; }
.budget-total-lbl { font-size: 13px; color: #64748B; margin-top: 2px; }
.editable-label {
    cursor: pointer; border-bottom: 1px dashed #CBD5E1;
    display: inline-block; padding: 1px 3px;
}
.editable-label:hover { background: #F1F5F9; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)

# ── Cookie-based auto-login ───────────────────────────────────────────────────
def check_remember_me_cookie():
    """Read cookie on page load and restore session if valid."""
    if st.session_state.get("logged_in"):
        return
    try:
        saved_user = cookie_manager.get(COOKIE_NAME)
        if saved_user and saved_user in USERS:
            st.session_state["logged_in"] = True
            st.session_state["username"]  = saved_user
    except Exception:
        pass

def set_remember_me_cookie(username):
    """Write cookie. Must NOT call st.rerun() in the same render cycle."""
    cookie_manager.set(COOKIE_NAME, username)

def clear_remember_me_cookie():
    try:
        cookie_manager.remove(COOKIE_NAME)
    except Exception:
        pass

# ── Login ─────────────────────────────────────────────────────────────────────
def login_screen():
    st.markdown("""
    <div style='text-align:center; padding-top: 60px;'>
        <h1 style='font-size:2.2rem;'>🧳 Study Abroad Packing List</h1>
        <p style='color:#64748B;'>Sign in to manage your packing list</p>
    </div>
    """, unsafe_allow_html=True)

    col = st.columns([1, 1.2, 1])[1]
    with col:
        with st.container(border=True):
            st.markdown("### Sign in")
            username    = st.text_input("Username", placeholder="Enter username")
            password    = st.text_input("Password", type="password", placeholder="Enter password")
            remember_me = st.checkbox("Remember me on this device", value=True)

            if st.button("Sign in", use_container_width=True, type="primary"):
                if username in USERS and USERS[username] == password:
                    # Phase 1: mark pending login in session state
                    st.session_state["pending_login"]   = True
                    st.session_state["pending_user"]    = username
                    st.session_state["pending_remember"] = remember_me
                    st.rerun()
                else:
                    st.error("Incorrect username or password.")

    # Phase 2: on the rerun after button press, write cookie then rerun again
    if st.session_state.get("pending_login"):
        username    = st.session_state.pop("pending_user")
        remember_me = st.session_state.pop("pending_remember")
        st.session_state.pop("pending_login")

        st.session_state["logged_in"] = True
        st.session_state["username"]  = username

        if remember_me:
            set_remember_me_cookie(username)
        else:
            clear_remember_me_cookie()

        st.rerun()

# ── Main App ──────────────────────────────────────────────────────────────────
def main_app():
    with st.sidebar:
        st.markdown(f"👤 **{st.session_state['username']}**")
        st.markdown("---")
        st.markdown("### 🔍 Filter")
        cats         = ["All"] + get_categories()
        selected_cat = st.selectbox("Category", cats, label_visibility="collapsed")
        search       = st.text_input("Search items", placeholder="🔍 Type to search...")
        st.markdown("---")
        show_only_packed   = st.checkbox("Show only packed items")
        show_only_unpacked = st.checkbox("Show only unpacked items")
        st.markdown("---")
        st.markdown("### 💾 Backup & Restore")

        _items_for_backup = load_items()
        _excel_backup     = build_excel(_items_for_backup)
        st.download_button(
            label="⬇️ Download checkpoint", data=_excel_backup,
            file_name="packing_list_backup.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        uploaded = st.file_uploader("⬆️ Restore from checkpoint", type=["xlsx"])
        if uploaded:
            if st.button("✅ Confirm restore", use_container_width=True, type="primary"):
                n, msg = restore_from_excel(uploaded)
                if n == 0 and msg.startswith(("Could not", "Missing", "No valid")):
                    st.error(f"Restore failed: {msg}")
                else:
                    if n > 0:
                        st.success(f"Restored {n} new item(s) successfully!")
                    if msg:
                        st.warning(msg)
                    st.rerun()

        st.markdown("---")
        if st.button("🚪 Sign out", use_container_width=True):
            clear_remember_me_cookie()
            st.session_state.clear()
            st.rerun()

    col_title, col_add, col_dl = st.columns([3, 1, 1])
    with col_title:
        st.markdown("# 🧳 Packing List")
    with col_add:
        st.markdown("<div style='margin-top:14px'>", unsafe_allow_html=True)
        if st.button("➕ Add item", use_container_width=True):
            st.session_state["show_add"] = not st.session_state.get("show_add", False)
        st.markdown("</div>", unsafe_allow_html=True)
    with col_dl:
        st.markdown("<div style='margin-top:14px'>", unsafe_allow_html=True)
        items_all = load_items()
        excel_buf = build_excel(items_all)
        st.download_button(
            label="📥 Save to Excel", data=excel_buf, file_name="packing_list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.get("show_add", False):
        with st.container(border=True):
            st.markdown("#### ➕ Add a new item")
            fc1, fc2 = st.columns(2)
            with fc1:
                new_name = st.text_input("Item name *", key="new_name", placeholder="e.g. Laptop charger")
            with fc2:
                existing_cats = get_categories()
                cat_options   = existing_cats + ["+ New category..."]
                cat_choice    = st.selectbox("Category *", cat_options, key="cat_choice")
            if cat_choice == "+ New category...":
                new_cat = st.text_input("New category name *", key="new_cat")
            else:
                new_cat = cat_choice
            new_note = st.text_input("Note (optional)", key="new_note", placeholder="e.g. 2 pcs, buy in India")
            new_budget = st.text_input("Budget (optional)", key="new_budget", placeholder="e.g. ₹500, AMAZON20")
            c1, c2 = st.columns([1, 5])
            with c1:
                if st.button("Save item", type="primary"):
                    if new_name.strip() and new_cat.strip():
                        add_item(new_name.strip(), new_cat.strip(), new_note.strip(), new_budget.strip())
                        st.session_state["show_add"] = False
                        for k in ["new_name", "new_note", "new_cat", "cat_choice", "new_budget"]:
                            st.session_state.pop(k, None)
                        st.success(f"'{new_name}' added!")
                        st.rerun()
                    else:
                        st.error("Item name and category are required.")
            with c2:
                if st.button("Cancel"):
                    st.session_state["show_add"] = False
                    st.rerun()

    items_all = load_items()
    total     = len(items_all)
    packed    = sum(1 for i in items_all if i["checked"])
    pct       = int(packed / total * 100) if total else 0
    s1, s2, s3, s4 = st.columns(4)
    with s1: st.markdown(f"<div class='stat-box'><div class='stat-num'>{total}</div><div class='stat-lbl'>Total items</div></div>", unsafe_allow_html=True)
    with s2: st.markdown(f"<div class='stat-box'><div class='stat-num'>{packed}</div><div class='stat-lbl'>Packed ✓</div></div>", unsafe_allow_html=True)
    with s3: st.markdown(f"<div class='stat-box'><div class='stat-num'>{total-packed}</div><div class='stat-lbl'>Remaining</div></div>", unsafe_allow_html=True)
    with s4: st.markdown(f"<div class='stat-box'><div class='stat-num'>{pct}%</div><div class='stat-lbl'>Complete</div></div>", unsafe_allow_html=True)
    st.progress(pct / 100)
    st.markdown("")

    items = items_all
    if selected_cat != "All":
        items = [i for i in items if i["category"] == selected_cat]
    if search:
        q     = search.lower()
        items = [i for i in items if q in i["name"].lower() or q in (i["note"] or "").lower()]
    if show_only_packed:
        items = [i for i in items if i["checked"]]
    if show_only_unpacked:
        items = [i for i in items if not i["checked"]]

    items_sorted = sorted(items, key=lambda x: (x["category"], x["name"]))
    if not items_sorted:
        st.info("No items found. Try adjusting your filters.")
        return

    # ── Helper: parse numeric budget ──────────────────────────────────────────
    def parse_numeric(val):
        """Strip currency symbols and return float if purely numeric, else None."""
        if not val:
            return None
        cleaned = val.strip().lstrip("₹$€£¥").strip().replace(",", "")
        try:
            return float(cleaned)
        except ValueError:
            return None

    total_budget_numeric = 0.0

    for cat, group in groupby(items_sorted, key=lambda x: x["category"]):
        group_list     = list(group)
        checked_in_cat = sum(1 for i in group_list if i["checked"])
        st.markdown(
            f"<div class='cat-header'>{cat} &nbsp;"
            f"<span style='font-weight:400;font-size:13px;color:#6B7280;'>{checked_in_cat}/{len(group_list)} packed</span></div>",
            unsafe_allow_html=True
        )
        for item in group_list:
            iid      = item["id"]
            is_checked = bool(item["checked"])

            # Accumulate numeric budgets for total
            bval = parse_numeric(item.get("budget", ""))
            if bval is not None:
                total_budget_numeric += bval

            editing = st.session_state.get(f"editing_{iid}", False)
            bg      = "background:#F0FDF4;" if is_checked else "background:#FAFAFA;"

            col_chk, col_info, col_del = st.columns([0.5, 9, 0.5])

            with col_chk:
                checked = st.checkbox(
                    label="packed", value=is_checked,
                    key=f"chk_{iid}", label_visibility="collapsed"
                )
                if checked != is_checked:
                    toggle_item(iid, checked)
                    st.rerun()

            with col_info:
                if editing:
                    # ── Edit mode: all 3 fields inline in one row ──────────────
                    ec_name, ec_note, ec_budget, ec_save, ec_cancel = st.columns([3, 3, 2, 1, 1])
                    with ec_name:
                        new_name_val = st.text_input(
                            "Name", value=item["name"],
                            key=f"inp_name_{iid}", label_visibility="visible"
                        )
                    with ec_note:
                        new_note_val = st.text_input(
                            "Note", value=item.get("note", ""),
                            key=f"inp_note_{iid}", label_visibility="visible",
                            placeholder="Add a note..."
                        )
                    with ec_budget:
                        new_budget_val = st.text_input(
                            "Budget", value=item.get("budget", ""),
                            key=f"inp_budget_{iid}", label_visibility="visible",
                            placeholder="e.g. ₹500"
                        )
                    with ec_save:
                        st.markdown("<div style='margin-top:24px'>", unsafe_allow_html=True)
                        if st.button("💾", key=f"save_{iid}", help="Save changes"):
                            fields = {"note": new_note_val.strip(), "budget": new_budget_val.strip()}
                            if new_name_val.strip():
                                fields["name"] = new_name_val.strip()
                            update_item_fields(iid, fields)
                            st.session_state[f"editing_{iid}"] = False
                            st.rerun()
                        st.markdown("</div>", unsafe_allow_html=True)
                    with ec_cancel:
                        st.markdown("<div style='margin-top:24px'>", unsafe_allow_html=True)
                        if st.button("✕", key=f"cancel_{iid}", help="Cancel"):
                            st.session_state[f"editing_{iid}"] = False
                            st.rerun()
                        st.markdown("</div>", unsafe_allow_html=True)
                else:
                    # ── Display mode: single row, click any text to edit ───────
                    name_style   = "color:#6B7280;text-decoration:line-through;" if is_checked else "color:#111827;font-weight:500;"
                    note_display = item.get("note") or "<span style='color:#CBD5E1'>add note</span>"
                    note_style   = "color:#6B7280;" if item.get("note") else "color:#CBD5E1;font-style:italic;"
                    budget_val   = item.get("budget") or ""
                    budget_display = budget_val if budget_val else "<span style='color:#CBD5E1'>budget</span>"
                    budget_style = "color:#059669;font-weight:500;" if budget_val else "color:#CBD5E1;font-style:italic;"

                    st.markdown(
                        f"""<div style='display:flex;align-items:center;gap:16px;padding:7px 10px;
                            border-radius:7px;{bg}cursor:pointer;border:1px solid transparent;'
                            title='Click ✏️ to edit'>
                            <span style='font-size:14px;{name_style};min-width:120px;'>{item['name']}</span>
                            <span style='font-size:12px;{note_style};min-width:100px;'>📝 {note_display}</span>
                            <span style='font-size:12px;{budget_style};min-width:80px;'>💰 {budget_display}</span>
                        </div>""",
                        unsafe_allow_html=True
                    )
                    if st.button("✏️ Edit", key=f"edit_btn_{iid}", help="Click to edit name, note or budget"):
                        st.session_state[f"editing_{iid}"] = True
                        st.rerun()

            with col_del:
                if not editing:
                    if st.button("🗑", key=f"del_{iid}", help="Delete item"):
                        delete_item(iid)
                        st.rerun()

    # ── Total Budget ───────────────────────────────────────────────────────────
    st.markdown(
        f"<div class='budget-total-box'>"
        f"<div class='budget-total-num'>₹ {total_budget_numeric:,.2f}</div>"
        f"<div class='budget-total-lbl'>Total Budget (numeric values only)</div>"
        f"</div>",
        unsafe_allow_html=True
    )

# ── Entry point ───────────────────────────────────────────────────────────────
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# Try to restore session from cookie before deciding what to render
check_remember_me_cookie()

if not st.session_state["logged_in"]:
    login_screen()
else:
    main_app()
