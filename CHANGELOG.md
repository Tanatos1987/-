# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [1.0.0] — 2026-06-24

### Added
- Initial release
- Automatic parsing of WorkFlow procurement exports (2025 and 2026 column layouts)
- 10-sheet Excel report with 15,000+ live formulas
- Monthly spend and volume comparison (YoY)
- ABC analysis by SKU revenue
- Weighted Average Price (WAP) per SKU per month
- Vendor market share and monthly breakdown
- Price volatility analysis (CV%)
- Auto-detection of file layout (9-column vs 11-column format)
- Handling of embedded sub-documents (ТРАНСФЕРЕН ПРОТОКОЛ, ФАКТУРА)
- Zero-price item flagging (yellow highlight in ITEMS sheets)
- Cross-sheet SUMIF / SUMPRODUCT formulas for full traceability
- Color-coded formula cells (blue = input, black = formula, green = cross-sheet)
- CLI interface with `--baseline`, `--ytd`, `--output` arguments
- Auto-verified totals in `⚙️ Assumptions` sheet
