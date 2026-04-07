# 📊 Demand Gen Placement Advisor

> A Google Ads Script that automatically analyses placement data from Demand Gen campaigns and recommends exclusions — directly in Google Sheets. No manual work required.

**Author:** Nadim Mahmud Sizan
**LinkedIn:** [linkedin.com/in/nmsizan98](https://www.linkedin.com/in/nmsizan98/)

---

## 🚀 What It Does

Most Google Ads advertisers running Demand Gen campaigns have no visibility into where their ads are actually showing — kids channels, spammy websites, clickbait videos, and mobile apps quietly drain budgets without converting.

This script solves that. It pulls 30 days of placement data from all your active Demand Gen campaigns, analyses every placement automatically, and writes a colour-coded report directly to your Google Sheet with clear recommendations.

**You stay in control** — the script never touches your Google Ads account. All exclusions are suggestions you apply manually.

---

## ✨ Features

- ✅ **Pulls all active Demand Gen campaigns automatically** — no manual campaign filter needed
- ✅ **30-day placement data** via Google Ads GAQL API
- ✅ **Kids channel detection** — YouTube API `madeForKids` flag + keyword matching
- ✅ **Language detection** — flags videos in languages outside your allowed list
- ✅ **Clickbait & political content detection** — scans video titles
- ✅ **Unsafe domain detection** — TLD checks, spammy keywords, suspicious domain patterns
- ✅ **Mobile app flagging** — auto-flags all mobile app placements
- ✅ **Full placement URLs** — YouTube channel/video links ready to copy-paste into exclusions
- ✅ **Colour-coded Google Sheet output** — Green / Yellow / Red rows at a glance
- ✅ **Read-only** — never modifies your Google Ads account

---

## 📋 Output Sheet Structure

The script writes a single tab named `[Account Name] - Placement Advisor` inside your Google Sheet.

| Column | Description |
|--------|-------------|
| A | Placement name |
| B | Full channel / placement URL |
| C | Placement type (YouTube channel, video, website, app) |
| D | Campaign name |
| E | Ad group name |
| F | Impressions |
| G | Clicks |
| H | Cost |
| I | Conversions |
| J | Cost / Conv. |
| K | Conv. Rate |
| L | CTR |
| M | Avg. CPC |
| N | Engagements |
| O | 👶 Kids / Made For Kids |
| P | 🌐 Language Issue |
| Q | 🔞 Unsafe TLD / Domain |
| R | 📢 Clickbait / Political |
| S | 📱 Mobile App |
| T | 💡 **Recommendation** (Keep / Review / Exclude) |
| U | 📝 Reasons |

### Colour coding

| Colour | Meaning |
|--------|---------|
| 🟢 Green | Keep — no issues detected |
| 🟡 Yellow | Review — investigate further |
| 🔴 Red | Exclude — apply manually in Google Ads UI |

---

## ⚙️ Setup

### Step 1 — Enable YouTube Advanced API
In the Google Ads Script editor:
`Advanced APIs` → tick **YouTube** → click **Save**

### Step 2 — Create a Google Sheet
Create a blank Google Sheet and copy its URL.

### Step 3 — Configure the script
Open the script and edit the configuration section at the top:

```javascript
var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit';

var LOOKBACK_WINDOW = 30;        // days of data to pull

var MIN_IMPRESSIONS = 20;        // ignore placements below this threshold

var ALLOWED_LANGUAGES = ['en'];  // videos outside these languages will be flagged

var SAFE_TLDS = ['.com', '.org', '.net', '.co.nz', '.co.uk'];  // adjust for your market
```

### Step 4 — Run the script
Click **Preview** first to test, then **Run** to write to your sheet.

> 💡 **Tip:** Schedule it to run weekly (every Monday) to keep your placement data fresh.

---

## 🔍 Detection Logic

### YouTube Channels & Videos
| Check | Method |
|-------|--------|
| Made For Kids | YouTube Data API `status.madeForKids` |
| Kids keywords in name | Keyword matching against channel/video name |
| Video language | `snippet.defaultLanguage` + `defaultAudioLanguage` |
| Clickbait title | Phrase matching against video title |
| Political content | Keyword matching against video title |
| Excessive punctuation | Regex check for `!!` `??` patterns |
| Monetary amounts | Regex check for `$` `€` `£` in titles |

### Websites
| Check | Method |
|-------|--------|
| Unsafe TLD | Checked against your `SAFE_TLDS` list |
| Spammy domain keywords | Matched against `SPAMMY_KEYWORDS` list |
| Suspicious domain length | Flags domains under 2 or over 30 characters |
| Hyphenated domains | Flags domains containing `-` |
| Punycode domains | Flags `xn--` internationalized domains |
| Repeated characters | Flags domains with 4+ repeated characters |

### Mobile Apps
All mobile app placements are automatically flagged for exclusion — they almost never convert for lead gen or service-based businesses.

---

## 🛠 Configuration Reference

```javascript
// ── Required ───────────────────────────────────────────────
var SPREADSHEET_URL = 'PASTE_YOUR_GOOGLE_SHEET_URL_HERE';

// ── Date range ─────────────────────────────────────────────
var LOOKBACK_WINDOW = 30;         // number of days to look back

// ── Minimum impressions filter ──────────────────────────────
var MIN_IMPRESSIONS = 20;         // placements below this are ignored

// ── Language filter ─────────────────────────────────────────
var ALLOWED_LANGUAGES = ['en'];   // 2-letter language codes
                                  // leave as [] to skip language check

// ── Safe TLDs ───────────────────────────────────────────────
var SAFE_TLDS = ['.com', '.edu', '.org', '.net', '.co.nz', '.nz'];

// ── Kids keywords ───────────────────────────────────────────
var KIDS_KEYWORDS = ['kids', 'children', 'nursery', ...];  // extend as needed

// ── Clickbait phrases ───────────────────────────────────────
var CLICKBAIT_PHRASES = ['shocking', 'exposed', 'you won\'t believe', ...];

// ── Political keywords ──────────────────────────────────────
var POLITICAL_KEYWORDS = ['election', 'protest', 'Trump', ...];

// ── Spammy domain keywords ──────────────────────────────────
var SPAMMY_KEYWORDS = ['free', 'cash', 'crypto', 'xxx', ...];
```

---

## 📌 How to Apply Exclusions

The script only **recommends** exclusions — it never writes to Google Ads. To apply them:

1. Filter column **T** (Recommendation) → select `Exclude`
2. Copy placement URLs from column **B**
3. In Google Ads → your Demand Gen campaign → **Placements tab** → **Exclusions** → **Add**
4. Paste the URLs → Save

---

## ⚠️ Requirements

- Google Ads account with at least one active Demand Gen campaign
- Google Ads Script access (Tools & Settings → Bulk Actions → Scripts)
- YouTube Advanced API enabled in the Script editor
- A Google Sheet (blank is fine)

---

## 📄 License

This script is free to use and modify. If you share or publish it, please credit the original author.

**Nadim Mahmud Sizan** — [linkedin.com/in/nmsizan98](https://www.linkedin.com/in/nmsizan98/)
