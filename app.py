from __future__ import annotations

import json

import pandas as pd
import streamlit as st

from trade_mapper import (
    EMAIL_COLUMNS,
    build_eml,
    build_html_table,
    build_preview_with_copy_button,
    combined_subject,
    load_trade_file,
    map_trades,
)

st.set_page_config(
    page_title="Trade Email Builder",
    page_icon="📨",
    layout="wide",
)

st.title("Trade Email Builder")
st.caption(
    "Upload a CSV/XLSX or paste cells from Excel, check the email preview, "
    "and download one Outlook-compatible email draft."
)

with st.sidebar:
    st.header("How it works")
    st.markdown(
        "1. Upload the trade file or paste cells from Excel.\n"
        "2. Check the email preview and any review notes.\n"
        "3. Copy the table or download the Outlook `.eml` file."
    )
    st.info(
        "Files are processed during the active app session. The app does not "
        "intentionally save uploaded files to permanent storage."
    )

input_method = st.radio(
    "How would you like to add trades?",
    ["Upload a file", "Paste from Excel"],
    horizontal=True,
)

if input_method == "Upload a file":
    uploaded = st.file_uploader(
        "Upload trade file",
        type=["csv", "xlsx", "xlsm"],
    )
    if uploaded is None:
        st.info(
            "Choose a file above to begin. If uploading fails, select "
            "Paste from Excel above to copy and paste your trade cells instead."
        )
        st.stop()
    input_bytes = uploaded.getvalue()
    input_filename = uploaded.name
else:
    st.info(
        "Open your workbook in Excel, select the cells containing the field "
        "names and trade values, and press Ctrl+C. Paste them below with "
        "Ctrl+V, then click outside the box to build the email."
    )
    pasted = st.text_area(
        "Paste Excel cells",
        height=260,
        help="Include both the field names and their values, keeping the Excel columns together.",
    )
    if not pasted.strip():
        st.stop()
    input_bytes = pasted.encode("utf-8")
    input_filename = "pasted_trades.tsv"

try:
    loaded_trades = load_trade_file(input_bytes, input_filename)
    results = map_trades(loaded_trades)
except Exception as exc:
    st.error(f"Could not read the trade data: {exc}")
    st.stop()

trade_labels = [
    result.trade_id or f"Trade {index + 1}"
    for index, result in enumerate(results)
]

st.success(f"Recognised {len(results)} trade{'s' if len(results) != 1 else ''}.")

field_rows: list[dict[str, str]] = [
    {column: result.fields.get(column, "") for column in EMAIL_COLUMNS}
    for result in results
]

st.subheader("Email preview")
# The preview includes a "Copy table" button that copies the table with its
# formatting so it can be pasted straight into Outlook.
preview_height = min(160 + len(results) * 35, 690)
st.components.v1.html(
    build_preview_with_copy_button(field_rows),
    height=preview_height,
    scrolling=True,
)

review_lines: list[str] = []
for label, result in zip(trade_labels, results):
    review_lines.extend(f"{label}: {note}" for note in result.review_notes)

if review_lines:
    st.warning("\n".join(f"• {line}" for line in review_lines))

recipient = st.text_input("Recipient", value="Agile@marexfp.com")
subject = st.text_input("Subject", value=combined_subject(results))

with st.expander("Recognised source fields", expanded=False):
    selected_label = st.selectbox("Choose trade", trade_labels)
    selected_index = trade_labels.index(selected_label)
    selected_result = results[selected_index]
    source_df = pd.DataFrame(
        selected_result.source.items(),
        columns=["Source field", "Value"],
    )
    st.dataframe(source_df, use_container_width=True, hide_index=True)

if any(result.unmapped_source_fields for result in results):
    with st.expander("Unmapped source fields"):
        unmapped_label = st.selectbox(
            "Choose trade to inspect",
            trade_labels,
            key="unmapped_trade",
        )
        unmapped_index = trade_labels.index(unmapped_label)
        st.json(results[unmapped_index].unmapped_source_fields)

base_name = "multi_trade" if len(results) > 1 else (trade_labels[0] or "trade")
col1, col2, col3 = st.columns(3)

with col1:
    st.download_button(
        "Download Outlook email (.eml)",
        data=build_eml(recipient, subject, field_rows),
        file_name=f"{base_name}_agile_email.eml",
        mime="message/rfc822",
        use_container_width=True,
        type="primary",
    )

with col2:
    st.download_button(
        "Download HTML",
        data=build_html_table(field_rows).encode("utf-8"),
        file_name=f"{base_name}_agile_email.html",
        mime="text/html",
        use_container_width=True,
    )

with col3:
    mapping_json = json.dumps(
        {
            "subject": subject,
            "trades": [
                {"trade_id": trade_labels[index], "fields": fields}
                for index, fields in enumerate(field_rows)
            ],
        },
        indent=2,
    )
    st.download_button(
        "Download mapping JSON",
        data=mapping_json.encode("utf-8"),
        file_name=f"{base_name}_mapping.json",
        mime="application/json",
        use_container_width=True,
    )

st.caption(
    "The cloud version creates a downloadable `.eml` file. Open it in Outlook, "
    "review the recipient and trade details, and send it manually. You can also "
    "use the Copy table button and paste the table straight into an email."
)
