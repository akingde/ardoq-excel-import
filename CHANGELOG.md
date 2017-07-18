# Changelog
As of version 1.1, all notable changes to this project will be documented in this file.

## [Unreleased]

## [1.1] - 2017-07-18
### Added
- Support for specifying leaf component types in a separate column.
- Support for specifying log level of the underlying Ardoq Client in property file (simplifies debugging)
- New example of business process using dynamic type mapping (see ./examples/business_process) example also includes a spreadsheet that uses a VBA script to auto-generate the references

### Changed
- Stopped using period as path separators in the internal component representation. This allows component names to have periods in their name.
- Ensure Excel Importer use same path separator as the Ardoq Client's sync util
- Updated to Ardoq Client 1.11
- Made organization a required property
- Don't throw NPE when component or reference sheet can't be found


### Removed
- Removed support for specifying multiple reference targets using comma (as it broke when component names contained commas).