# Changelog

## 0.3.2

- **Font substitution now picks real Apple system fonts** — was incorrectly picking non-system Google Fonts (e.g. Montserrat → Barlow, which isn't installed on any Apple device). Now uses `APPLE_SYSTEM_FONT_LIST` from quicklook-pptx-renderer, ensuring all 29 candidate fonts are preinstalled on both macOS and iOS
- **Narrower substitutes preferred** — the similarity algorithm now penalizes wider fonts 3x more, preventing text overflow. E.g. Montserrat → DIN Alternate (-4.3%) instead of Barlow (+2.0%)
- **Re-exports `APPLE_SYSTEM_FONTS` and `MACOS_SYSTEM_FONTS`** from quicklook-pptx-renderer for downstream consumers
- Requires quicklook-pptx-renderer ^0.3.3

## 0.3.1

- Add embedded-fonts transform: strip and replace with safe alternatives
- Add groups transform and chart fallback generation
- Add fonts transform: replace high-risk Windows fonts
