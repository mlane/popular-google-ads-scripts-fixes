# Popular Google Ads Scripts

## Overview

Google Ads has deprecated the [AdWords API](https://developers.google.com/adwords/api/docs/guides/awql) in favor of the more advanced [Google Ads API](https://developers.google.com/google-ads/api/docs/start). While both APIs provide similar functionality, the transition has introduced significant schema changes, rendering many older, unmaintained scripts incompatible.

For example, the [Keywords Performance Report](https://developers.google.com/adwords/api/docs/appendix/reports/keywords-performance-report) has been replaced by [`keyword_view`](https://developers.google.com/google-ads/api/fields/v11/keyword_view) in the new API.

To maintain compatibility, these scripts have been updated from AdWords Query Language (AWQL) to Google Ads Query Language (GAQL).

If you encounter any issues, please feel free to reach out.

## Available Scripts

- [Advanced Quality Score Tracker by Clicteq](./scripts/advanced-quality-score-tracker-by-clicteq.ts)
- [Low Quality Score Alert](./scripts/low-quality-score-alert.ts)
- [Quality Score Analysis](./scripts/quality-score-analysis.ts)
