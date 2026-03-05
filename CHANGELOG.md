# Changelog

All notable changes to this project will be documented in this file.

The format is based on Keep a Changelog, and this project uses date-based entries.

## [Unreleased]

### Added
- Added dedicated `CHANGELOG.md` for release and feature tracking.

## [2026-03-05]

### Added
- Added `Supply Eligible Analysis` mode using Orderbook rows with `Sales Status = Supply Eligible`.
- Added project priority sequence upload support for `.xlsx`, `.xlsm`, and `.csv`.
- Added upload apply modes: `Replace current sequence` and `Append / Merge into current sequence`.
- Added autosave support for project sequence state in `project_sequence_state.json`.
- Added priority-based cascading allocation views for stock and GRN by project.

### Changed
- Improved file ingestion with sheet/header auto-detection for Orderbook, GRN, and Stock files.
- Expanded analytics and drill-down outputs with safer mixed-type dataframe rendering.
