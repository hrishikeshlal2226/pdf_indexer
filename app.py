import contextlib
import io
import tempfile
from pathlib import Path

import streamlit as st

import pdf_indexer as indexer

st.set_page_config(
    page_title="PDF Indexer",
    page_icon="📄",
    layout="centered",
)

st.title("PDF Indexer")
st.caption("Upload a PDF to generate a clickable index prepended to the document.")

uploaded_file = st.file_uploader("Choose a PDF", type="pdf")

if uploaded_file:
    if st.button("Generate Index", type="primary", use_container_width=True):

        for key in ["indexed_pdf", "indexed_pdf_name", "excel", "html_index", "csv", "stats"]:
            st.session_state.pop(key, None)

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)

            pdf_path = tmp_path / uploaded_file.name
            pdf_path.write_bytes(uploaded_file.getvalue())

            output_folder = tmp_path / "output"
            output_folder.mkdir()

            log = io.StringIO()

            with st.status("Processing PDF...", expanded=True) as status:

                st.write("Extracting topics from each page...")

                with contextlib.redirect_stdout(log):
                    results = indexer.extract_topics(pdf_path)

                indexer.TOPIC_LOOKUP = {
                    item["page"]: item["title"]
                    for item in results
                }

                found    = sum(r["status"] == "FOUND"    for r in results)
                fallback = sum(r["status"] == "FALLBACK" for r in results)
                review   = sum(r["status"] == "REVIEW"   for r in results)

                st.write(
                    f"Processed {len(results)} pages — "
                    f"{found} found, {fallback} fallback, {review} need review"
                )

                st.write("Creating Excel, HTML, and CSV outputs...")

                with contextlib.redirect_stdout(log):
                    indexer.create_excel(
                        results,
                        output_folder / "index.xlsx",
                    )
                    indexer.create_html(
                        results,
                        output_folder / "index.html",
                        uploaded_file.name,
                    )
                    indexer.create_review_csv(
                        results,
                        output_folder / "review_pages.csv",
                    )

                st.write("Building index pages...")

                tmp_index_pdf = tmp_path / "_index_only.pdf"

                with contextlib.redirect_stdout(log):
                    index_page_count, link_data = indexer.create_index_pdf(
                        results,
                        uploaded_file.name,
                        tmp_index_pdf,
                    )

                final_pdf = output_folder / f"{pdf_path.stem}_INDEXED.pdf"

                st.write("Combining index with original PDF...")

                with contextlib.redirect_stdout(log):
                    indexer.combine_with_pypdf(
                        tmp_index_pdf,
                        pdf_path,
                        final_pdf,
                        link_data,
                    )

                # Read into session state before the temp dir is deleted
                st.session_state["indexed_pdf"]      = final_pdf.read_bytes()
                st.session_state["indexed_pdf_name"] = final_pdf.name
                st.session_state["excel"]            = (output_folder / "index.xlsx").read_bytes()
                st.session_state["html_index"]       = (output_folder / "index.html").read_bytes()
                st.session_state["csv"]              = (output_folder / "review_pages.csv").read_bytes()
                st.session_state["stats"] = {
                    "total":       len(results),
                    "found":       found,
                    "fallback":    fallback,
                    "review":      review,
                    "index_pages": index_page_count,
                }

                status.update(label="Done!", state="complete")


if "indexed_pdf" in st.session_state:
    stats = st.session_state["stats"]

    st.subheader("Results")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Pages",   stats["total"])
    c2.metric("Topics Found",  stats["found"])
    c3.metric("Fallback",      stats["fallback"])
    c4.metric("Need Review",   stats["review"])

    st.subheader("Downloads")

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "Download Indexed PDF",
            data=st.session_state["indexed_pdf"],
            file_name=st.session_state["indexed_pdf_name"],
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
        st.download_button(
            "Download Excel Index",
            data=st.session_state["excel"],
            file_name="index.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            "Download HTML Index",
            data=st.session_state["html_index"],
            file_name="index.html",
            mime="text/html",
            use_container_width=True,
        )
        st.download_button(
            "Download Review CSV",
            data=st.session_state["csv"],
            file_name="review_pages.csv",
            mime="text/csv",
            use_container_width=True,
        )
