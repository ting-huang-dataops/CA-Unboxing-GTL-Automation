# Fund Service - Corporate Action - GTL Boxed Position Resolve Automation (VBA)

## 📌 Business Impact & Overview
In Hedge Fund Operations, mismatched order attributes (e.g., Cover/Sell vs. Buy/Short) frequently create **Boxed Positions**—unintended offsetting inventory stuck in accounting and risk platforms. Resolving these high-volume boxed positions requires generating precise, double-entry **Generic Trade Loader (GTL) CSV files** to unbox and clean position records.

This VBA suite automates the end-to-end unboxing workflow for **Equity and Equity Swap** instruments. By dynamically generating paired `BC` (Buy Cover) and `S` (Sell) offsetting transactions with unique system identifiers, it eliminates position mismatches, restores true inventory balance, and saves hours of manual trade construction.

## 🔑 Key Technical Features
1. **Automated Double-Entry Unboxing Pair Generation:**
   - Converts raw unbox requests into precise, two-sided GTL adjusting entries (`BC` - Buy Cover and `S` - Sell).
   - Dynamically constructs collision-resistant Transaction IDs using custom pseudo-random algorithms (`GenerateRandomChars`) to ensure exact system matching.
2. **Multi-Asset Class Routing (Equity vs. Equity Swap):**
   - **Equity Module (`FillGTLSheetforEquity`):** Filters and maps cash equity adjustments into standard security trade schema (`CS`).
   - **Swap Module (`FillGTLSheetForSwap`):** Identifies synthetic swap positions, automatically populating default maturity dates (`20340101`) and equity swap codes (`EQSWAP`).
3. **Data Sanitation & String Normalization:**
   - Cleans hidden characters (e.g., non-breaking spaces `ChrW(160)` and comma delimiters) prior to numeric parsing to guarantee zero upload failures.
4. **Automated CSV Loader Export:**
   - Programmatically builds temporary workbooks, dumps validated 50+ column GTL array payloads, exports timestamped `.csv` files directly to network shares, and resets staging environments.

## 🚀 Quantifiable Operational Impact
- **Risk & Position Accuracy:** Successfully unboxed high-volume locked inventory, eliminating artificial balance sheet inflation caused by Buy/Short & Cover/Sell mismatches.
- **Operational Speed:** Reduced GTL unbox file generation time from **~hours down to seconds**.
- **Data Quality:** Achieved a **100% system upload pass rate** by automating complex 50-column layout rules.
