# Popular Google Ads Scripts

Google Ads is no longer supporting [AdWords API](https://developers.google.com/adwords/api/docs/guides/awql) in favor of [Google Ads API](https://developers.google.com/google-ads/api/docs/start).

Although both APIs share similar functionality the new API is breaking a lot of the older unmaintained scripts mainly due to the schema changes.

e.g. [Keywords Performance Report](https://developers.google.com/adwords/api/docs/appendix/reports/keywords-performance-report) is now [keyword_view](https://developers.google.com/google-ads/api/fields/v11/keyword_view)

The scripts for now have essentially been converted from AWQL (AdWords Query Language) to GAQL (Google Ads Query Language)

Let me know if you run into any issues.

## Scripts

- [Advanced Quality Score Tracker by Clicteq](./scripts/advanced-quality-score-tracker-by-clicteq.ts)
- [Low Quality Score Alert](./scripts/low-quality-score-alert.ts)
- [Quality Score Analysis](./scripts/quality-score-analysis.ts)
