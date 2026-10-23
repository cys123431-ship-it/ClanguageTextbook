# v134 — 2026-10-23 22:00 KST

- Reused existing v133 full manuscript without rewriting it.
- Fixed a structural error: original sections 4.29–4.39 had been inserted inside the frontmatter outline. Moved all 648 original paragraph/table XML elements into actual Chapter 4 after 4.28; their XML contents were preserved exactly.
- Directly inspected original Express Chapter 5 Programming 01–12 (scanned PDF physical pp.217–219).
- Added eight original beginner-focused sections 5.40–5.47 for Programming 01,02,03,04,05,07,08,11; mapped the remaining 06,09,10,12 to existing 5.33,5.36,5.34,5.35 respectively.
- Added 545 new paragraph/table XML elements. Full DOCX now contains 70,335 elements, compared with 69,790 in v133.
- GCC/Clang strict C11: 16/16 builds and 80/80 execution checks passed. GCC UBSan: 8/8 builds and 24/24 selected executions passed.
- New standalone DOCX rendered 15 pages; integrated excerpt rendered 21 pages. DOCX ZIP integrity passed.
- MSVC and whole-book page-by-page visual audit remain untested. Other Express LAB/Exercise/Mini Project and K&R goal audits remain incomplete.
- Status: First full-book draft still in progress; NOT ready for user review.