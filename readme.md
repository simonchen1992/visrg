# XT Report Generator #
### Logic of XT report generator ###
1. Find the test session id.
2. List all ACTIVE cards in database.
3. Fill position verdicts in database row by row. (Import note: If there are multiple position verdicts for one position/card/session. The tool will only pick the recent one.)
    - Test position filtering for each card.
        - For non-tested positions: the tool will skip to fill position verdicts.
        - For required-tested positions: If there's no position verdict found in database, "Pending" will be filled in. Otherwise, the tool will fill position verdict.
        - If there are at lease one FAIL in z = 3 positions, and at least one PASS in all positions, the tool will add 2N, 2W, 2E, 2S into required-tested positions.
4. Validate position verdicts row by row.
    - rules 1: if there appears FAIL and PASS position verdict for one card, all FAIL verdict will be determined as DF
    - rules 2: if there appears only FAIL verdict for one card, but it contains both CF and TF, all FAIL verdict will be determined as TF 
5. Determine card verdict
    - Position verdict contains "Pending": card verdict is considered as "Pending".
    - Position verdict contains "DF/CF/TF", and not contains "Pending": card verdict is considered as "FAIL".
    - Position verdict contains ONLY "PASS": card verdict is considered as "PASS".

### How to Make the VISA Report ###
1. Generate the VISA report from data in Roadmap.
2. Copy the VISA Report to "paintTemplate.xlsm" to paint fail positions into red.
3. Copy the painted VISA Report back to VISA template.

### How to update VISA template ###
1. Save the latest VISA template under the path: /docs/VisaTemplate/
2. It's all done. 

*Bear in mind to set proper Visa Template name in database when starting a new XT session.