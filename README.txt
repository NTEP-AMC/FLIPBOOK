AMC NTEP — Public Health Actions Deck Builder
==============================================

FILES
-----
app.py                          - the Streamlit app
NotoSansGujarati-Regular.ttf    - required, must sit next to app.py
NotoSansGujarati-Bold.ttf       - required, must sit next to app.py
requirements.txt                - pip install -r requirements.txt

RUN LOCALLY
-----------
pip install -r requirements.txt
streamlit run app.py

DEPLOY (Streamlit Cloud / GitHub)
----------------------------------
Push all 4 files above to the SAME folder in your repo (root, or one
subfolder — just keep app.py and the two .ttf files together, the code
loads fonts from the same directory app.py is in).

USING THE APP
--------------
1. Optional: upload left/right logo images (AMC seal, NTEP logo).
2. Optional: upload equipment/reference photos. Name each file after the
   keyword that should trigger it, e.g. genexpert.png or cbnaat.jpg —
   any sub-module whose text mentions that word (case-insensitive, spaces/
   underscores ignored) automatically gets a dedicated photo slide.
3. Upload your Word document (.docx).
4. Expand "🔍 Parsed structure" to sanity-check what was extracted BEFORE
   generating — check module/sub-module counts and that all 11 fields per
   sub-module show up where you expect them.
5. Click "Generate Presentation" and download the .pptx.

IF YOUR DOC USES DIFFERENT FIELD LABELS
-----------------------------------------
The parser matches fields by the ENGLISH word in parentheses, e.g.
"... (Objective):", not by the Gujarati text. This is deliberate — it's
robust to Gujarati spelling variations in the source. The 11 fields it
looks for (edit FIELD_META near the top of app.py to change these):

  Objective, Trigger, What to do, Why, Whom, Responsible Person,
  Timeline, Documentation, Monitoring Indicators, Supervisor Checklist,
  If Not Done

If your document uses a different English word for any of these, either:
  (a) add/adjust the "en" value for that field in FIELD_META, or
  (b) add an extra English synonym to the matching pattern.

If a whole field never shows up in the parsed-structure preview, that's
the most likely cause — mismatched English term in the parentheses.

CUSTOMIZING LAYOUT / COLORS / ICONS
-------------------------------------
- Colors: NAVY, GREEN, BLUE_ICON, WARN_BG etc. near the top of app.py.
- Which fields go on the first vs. second slide of a sub-module:
  PROCESS_ORDER / QUALITY_ORDER lists.
- Gujarati text looking too tight/loose: GUJ_WIDTH_CORRECTION constant
  (currently 1.14) — raise it if you ever see text touching a card edge
  on your real document; lower it if lines look mostly empty.
