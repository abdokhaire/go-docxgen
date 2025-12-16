# Changelog

## [0.2.6] - 2025-12-16
### Fixed
- Replace `<nil>` with empty string in XML output (in addition to `<no value>`)
- Graceful handling of nil interface values in templates

### Documentation
- Clarified template functions documentation - only `link` is built-in
- Updated README, CLAUDE.md, and API_REFERENCE.md to accurately document that functions like `upper`, `lower`, `formatMoney`, etc. require registering Sprig or Sprout
- Added examples for using Sprig and Sprout function libraries
- Documented Go template built-in functions (`eq`, `ne`, `len`, `printf`, etc.)

## [0.2.5] - 2025-12-16
### Fixed
- Fix null values and some merge fields not passed to the latest document
- Add sprout test fixes

## [0.2.4] - 2025-12-15
### Fixed
- Fix empty merge tags handling

## [0.2.3] - 2025-12-14
### Changed
- File structure update

## [0.2.2] - 2025-12-13
### Fixed
- Vendor go-docx with header/footer reference support

## [0.2.1] - 2025-12-12
### Added
- Header and footer template support

## [0.2.0] - 2025-12-10
### Added
- Document builder API for programmatic document creation
- Paragraph formatting (bold, italic, color, alignment, etc.)
- Table builder with cell merging and formatting
- Document properties (title, author, subject, etc.)
- Inline image support with auto-sizing
- Watermark extraction and replacement

## [0.1.2] - 2025-12-08
### Fixed
- Bug fixes and improvements

## [0.1.1] - 2025-12-06
### Fixed
- Initial bug fixes

## [0.1.0] - 2025-12-05
### Added
- Initial release
- Go template syntax support in DOCX files
- Fragmented tag merging
- Custom function registration
- Headers, footers, footnotes, endnotes processing
