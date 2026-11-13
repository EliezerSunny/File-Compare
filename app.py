import streamlit as st
import streamlit.components.v1 as components

import compare_core as cc

st.set_page_config(page_title="File Compare", page_icon="🔍", layout="wide")
st.title("🔍 File Compare")
st.caption("Excel, CSV, Word, PDF and text files. Everything runs on your own machine.")

c1, c2 = st.columns(2)
types = ["xlsx", "xlsm", "xls", "csv", "tsv", "docx", "pdf", "txt", "md"]
fa = c1.file_uploader("File A (original)", type=types, key="a")
fb = c2.file_uploader("File B (new)", type=types, key="b")

if not (fa and fb):
    st.info("Upload two files to begin.")
    st.stop()

ea, eb = cc.ext(fa.name), cc.ext(fb.name)
a_tab, b_tab = ea in cc.TABULAR, eb in cc.TABULAR

try:
    if a_tab and b_tab:
        ta, tb = cc.load_tables(fa.name, fa.getvalue()), cc.load_tables(fb.name, fb.getvalue())
        s1, s2 = st.columns(2)
        sheet_a = s1.selectbox("Sheet in A", list(ta))
        # default B's sheet to the same name when it exists
        names_b = list(tb)
        sheet_b = s2.selectbox("Sheet in B", names_b,
                               index=names_b.index(sheet_a) if sheet_a in names_b else 0)
        da, db = ta[sheet_a], tb[sheet_b]
        common = [c for c in da.columns if c in db.columns]
        keys = st.multiselect(
            "Key column(s) that identify a row (e.g. ID, Email). Leave empty to compare row by row.",
            common)
        res = cc.compare_tables(da, db, keys)

        if res["identical"]:
            st.success("The two sheets are identical.")
        m = st.columns(4)
        m[0].metric("Rows A / B", f"{res['rows_a']} / {res['rows_b']}")
        m[1].metric("Added in B", len(res["added"]))
        m[2].metric("Removed from B", len(res["removed"]))
        m[3].metric("Changed cells", len(res["changed"]))
        if res["only_in_a_cols"]:
            st.warning("Columns only in A: " + ", ".join(res["only_in_a_cols"]))
        if res["only_in_b_cols"]:
            st.warning("Columns only in B: " + ", ".join(res["only_in_b_cols"]))
        if res["dup_a"] or res["dup_b"]:
            st.warning("Some key values appear more than once. They are matched in order "
                       "of appearance (shown as #1, #2...). Pick more key columns if needed.")

        t1, t2, t3 = st.tabs(["Changed cells", "Added in B", "Removed from B"])
        t1.dataframe(res["changed"], use_container_width=True, hide_index=True)
        t2.dataframe(res["added"], use_container_width=True, hide_index=True)
        t3.dataframe(res["removed"], use_container_width=True, hide_index=True)
        st.download_button("⬇️ Download Excel report", cc.build_report(res),
                           file_name="comparison_report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    elif not a_tab and not b_tab:
        la, lb = cc.load_text(fa.name, fa.getvalue()), cc.load_text(fb.name, fb.getvalue())
        res = cc.compare_text(la, lb)
        if res["identical"]:
            st.success("The two documents have identical text.")
        m = st.columns(3)
        m[0].metric("Similarity", f"{res['similarity']:.1%}")
        m[1].metric("Lines added", res["added"])
        m[2].metric("Lines removed", res["removed"])
        st.caption("Green = added, red = removed, yellow = changed. Only text is compared, "
                   "not fonts or layout.")
        components.html(
            "<style>body{font-family:sans-serif;font-size:13px}table.diff{width:100%}"
            ".diff_add{background:#aaffaa}.diff_chg{background:#ffff77}"
            ".diff_sub{background:#ffaaaa}td{padding:2px 4px}</style>" + res["html"],
            height=700, scrolling=True)
        html_bytes, txt_bytes = cc.build_text_reports(la, lb, res["html"], fa.name, fb.name)
        d1, d2 = st.columns(2)
        d1.download_button("⬇️ Download differences (HTML, opens on any phone)", html_bytes,
                           file_name="differences.html", mime="text/html")
        d2.download_button("⬇️ Download differences (TXT)", txt_bytes,
                           file_name="differences.txt", mime="text/plain")
    else:
        st.error("Please upload two files of the same kind: both spreadsheets (xlsx/csv) "
                 "or both documents (docx/pdf/txt).")
except Exception as err:  # show a readable message instead of a crash
    st.error(f"Could not compare these files: {err}")
