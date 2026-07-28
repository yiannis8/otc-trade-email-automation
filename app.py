from __future__ import annotations

import json

import pandas as pd
import streamlit as st

from trade_mapper import (
    EMAIL_COLUMNS,
    build_eml,
    build_html_table,
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
    "Upload a single-trade or multi-trade CSV/XLSX, review the recognised rows, "
    "and download one Outlook-compatible email draft."
)

with st.sidebar:
    st.header("How it works")
    st.markdown(
        "1. Upload the trade file.\n"
        "2. Review the mapped rows.\n"
        "3. Correct any highlighted fields.\n"
        "4. Download the Outlook `.eml` file."
    )
    st.info(
        "Files are processed during the active app session. The app does not "
        "intentionally save uploaded files to permanent storage."
    )

uploaded = st.file_uploader(
    "Upload trade file",
    type=["csv", "xlsx", "xlsm"],
    help="Maximum upload size is configured at 25 MB.",
)

if uploaded is None:
    st.info(
        "Supported layouts: one field/value pair per row, multiple trades across "
        "columns, or a standard table with one trade per row."
    )
    st.stop()

try:
    loaded_trades = load_trade_file(uploaded.getvalue(), uploaded.name)
    results = map_trades(loaded_trades)
except Exception as exc:
    st.error(f"Could not read the file: {exc}")
    st.stop()

trade_labels = [
    result.trade_id or f"Trade {index + 1}"
    for index, result in enumerate(results)
]
st.success(f"Recognised {len(results)} trade{'s' if len(results) != 1 else ''}.")

with st.expander("Recognised source fields", expanded=False):
    selected_label = st.selectbox("Choose trade", trade_labels)
    selected_index = trade_labels.index(selected_label)
    selected_result = results[selected_index]
    source_df = pd.DataFrame(
        selected_result.source.items(),
        columns=["Source field", "Value"],
    )
    st.dataframe(source_df, use_container_width=True, hide_index=True)

st.subheader("Email rows")
email_df = pd.DataFrame(
    [
        {"Trade ID": trade_labels[index], **result.fields}
        for index, result in enumerate(results)
    ]
)

edited_df = st.data_editor(
    email_df,
    use_container_width=True,
    hide_index=True,
    disabled=["Trade ID"],
    num_rows="fixed",
    height=min(180 + len(results) * 38, 520),
    key=f"email_fields_{uploaded.name}_{len(results)}",
)

field_rows: list[dict[str, str]] = []
for _, row in edited_df.iterrows():
    field_rows.append(
        {
            column: "" if pd.isna(row.get(column)) else str(row.get(column))
            for column in EMAIL_COLUMNS
        }
    )

recipient = st.text_input("Recipient", value="Agile@marexfp.com")
subject = st.text_input("Subject", value=combined_subject(results))

review_lines: list[str] = []
for label, result in zip(trade_labels, results):
    review_lines.extend(f"{label}: {note}" for note in result.review_notes)
if review_lines:
    st.warning("\n".join(f"• {line}" for line in review_lines))

if any(result.unmapped_source_fields for result in results):
    with st.expander("Unmapped source fields"):
        unmapped_label = st.selectbox(
            "Choose trade to inspect",
            trade_labels,
            key="unmapped_trade",
        )
        unmapped_index = trade_labels.index(unmapped_label)
        st.json(results[unmapped_index].unmapped_source_fields)

st.subheader("Email preview")
preview_height = min(300 + len(results) * 35, 650)
st.components.v1.html(
    build_html_table(field_rows),
    height=preview_height,
    scrolling=True,
)

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
    "review the recipient and trade details, and send it manually."
)
