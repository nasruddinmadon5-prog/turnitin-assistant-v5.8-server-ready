# v5.6.5 Implementation Plan

- [x] Duplicate v5.6.4 to v5.6.5
- [ ] Update config.py: add USE_AI_VERIFIER, update version
- [ ] Update app.py: version string
- [ ] template_screening_engine.py:
     - [ ] Add extract_source_text() for .docx XML parsing
     - [ ] Add source_text_section_detector()
     - [ ] Add AI verifier function (optional, gated by config)
     - [ ] Modify run_single_screening() to accept source_file
     - [ ] Add DOCX-first structure detection logic
     - [ ] Modify run_template_screening_workflow() to pass source file
     - [ ] Add debug report fields (detected_from_docx, detected_from_pdf, detected_from_ai, final_source_used)
     - [ ] Update generate_screening_pdf_report() to show debug info
- [ ] Verify version consistency across all files