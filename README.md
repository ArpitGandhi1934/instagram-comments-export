# Export Instagram Comments to CSV and Excel (Python + Node.js)

[![Instagram Comments Scraper on Apify][badge]][store]

**Export Instagram comments** from any public post or reel to **Excel (.xlsx), CSV or JSONL**: username, comment text, likes, date and whether the comment is by the post's creator. Keep only the most-liked comments, or only comments that mention a keyword, have enough likes or are newer than a date. The Python and Node.js scripts in this repo call the [Instagram Comments Scraper][store] Actor on Apify. You don't need an Instagram account, cookies or a Graph API app, and comments your filters skip are never charged.

Guide: [Export Instagram comments to Excel](https://yugenox-data.vercel.app/instagram/export-comments).

## What you get

Real output from one run on 2026-10-06: the 30 most-liked comments on a public reel (`DdPsDCWRT-u`, posted by `nasaadmin`), from `--sort top --max-comments 30`. Commenters are pseudonymized and comment text is cut to one sentence in [`samples/`](samples) ([how](samples/redact.jq)).

| username | text | likes | created_at |
|---|---|---:|---|
| commenter_01 | People don’t realize It’s small things like this that inspire kids enough to make a career out of it. … | 164 | 2026-09-14T00:34:18.000Z |
| commenter_02 | | 105 | 2026-09-13T23:17:21.000Z |
| commenter_03 | Keep looking up 🇺🇸 | 101 | 2026-09-13T23:13:35.000Z |
| commenter_04 | Thank you for blessing Pittsburgh today! … | 50 | 2026-09-13T23:26:33.000Z |
| commenter_05 | T-38 is still a beautiful plane. | 21 | 2026-09-13T23:40:39.000Z |

The same 30 rows as [Excel](samples/nasa-reel-top-comments.xlsx), [CSV](samples/nasa-reel-top-comments.csv) and the Actor's [raw JSON rows](samples/nasa-reel-top-comments.json). To pick the top 30, the Actor read 189 comments on the reel; only the 30 it saved were billed, which is $0.057 at the Free-plan price.

| Column | Meaning |
|---|---|
| `username` | The commenter's Instagram handle |
| `text` | The comment text |
| `likes` | Likes on the comment |
| `created_at` | When it was posted (UTC, ISO 8601; a real date-time cell in the .xlsx) |
| `post_url`, `post_owner` | The post or reel, and the account that posted it |
| `is_post_owner` | `true` when the comment is by the post's creator |
| `is_verified` | `true` when the commenter has a verified badge |
| `comment_id` | Instagram's comment id (`--append` uses it to skip comments you already have) |

## Quick start

1. Create a free Apify account and copy your API token ([sign up][signup], then Console > Settings > API & Integrations).
2. `cp .env.example .env` and paste the token after `APIFY_TOKEN=`.

**Node.js** (20.6 or newer):

```bash
npm install
node --env-file=.env export.mjs https://www.instagram.com/reel/DdPsDCWRT-u/ --out comments.xlsx
node --env-file=.env export.mjs --file posts.example.txt --sort top --max-comments 200
node --env-file=.env export.mjs --file posts.txt --keyword price --keyword "how much" --out questions.csv
```

**Python** (3.10 or newer):

```bash
pip install -r requirements.txt
python export.py https://www.instagram.com/reel/DdPsDCWRT-u/ --out comments.xlsx
python export.py --file posts.example.txt --sort top --max-comments 200
python export.py --file posts.txt --min-likes 10 --newer-than "7 days" --out recent.csv
```

Both scripts take the same options (`--help` prints them):

| Option | What it does | Default |
|---|---|---|
| post or reel URLs | `/p/<code>/` or `/reel/<code>/` links. A profile URL collects comments from that profile's latest posts | |
| `--file posts.txt` | One URL per line; lines starting with `#` are skipped | |
| `--max-comments N` | Comments to save per post. With filters, this counts matching comments only | `100` |
| `--sort newest\|top` | `newest` is Instagram's order. `top` returns the most-liked first, picked from up to 2,000 comments per post | `newest` |
| `--min-likes N` | Only comments with at least N likes | |
| `--keyword TEXT` | Only comments containing this word or phrase (case-insensitive). Repeat it; a comment matching any of them is kept | |
| `--newer-than WHEN` | Only comments after a date (`2026-09-01`) or within a period (`7 days`, `24 hours`) | |
| `--out FILE` | `.xlsx`, `.csv` or `.jsonl` | `comments.csv` |
| `--append` | Add only comments whose `comment_id` isn't in `--out` yet (`.csv` or `.jsonl`) | |
| `--max-total N` | Hard cap on comments across all posts in one run | `1000` |
| `--max-charge USD` | Hard cost cap for the run, enforced by Apify | `2` |
| `--dataset ID` | Export the dataset of a run you already made in Apify Console. No new run starts | |
| `--from FILE` | Convert a JSON or JSONL file of the Actor's rows (Console export) or of this script's rows. Needs no token | |

The filters (`--sort`, `--min-likes`, `--keyword`, `--newer-than`) run inside the Actor, before billing, so you pay only for the comments that end up in your file. The CSV is UTF-8 with a BOM so Excel shows emoji and accents correctly, and a comment that starts with `=`, `+`, `-` or `@` gets a leading `'` in the CSV so a spreadsheet shows it as text instead of running it as a formula. The .xlsx stores every comment as plain text.

## Recipes

**Most-liked comments first.** `--sort top --max-comments 50` reads up to 2,000 comments per post and keeps the 50 with the most likes. That's the quickest way to see what an audience agrees with.

**Questions and buying signals.** `--keyword price --keyword "how much" --keyword where --keyword link` keeps only comments that mention one of those words, on any number of posts.

**Giveaway entries.** Export every comment on the giveaway post to Excel, then do the picking in the sheet: for example `=UNIQUE(A2:A)` in Google Sheets for one row per username, and `=RANDBETWEEN` or a draw your rules describe. These scripts export comments; they don't pick winners. Free Apify plans stop at 100 comments per run, so a post with more entries needs a paid plan.

**Monitor new comments every day.** Run the export on a schedule with a 24-hour window and append to one file:

```bash
# crontab -e   (every morning at 8)
0 8 * * * cd /path/to/instagram-comments-export && node --env-file=.env export.mjs --file posts.txt --newer-than "24 hours" --append --out comments.csv
```

`--newer-than` makes the Actor stop reading a post at the first older comment, so a daily run costs only the new comments. `--append` skips any comment that's already in the file. A comment the run returns twice is still billed twice, so keep the window close to your schedule. To run it on Apify instead of your machine, save the input as a task in Apify Console, attach a daily Schedule, and fetch each day's result from the task's last run:

```bash
curl -s "https://api.apify.com/v2/actor-tasks/<TASK_ID>/runs/last/dataset/items?status=SUCCEEDED&format=xlsx" \
  -H "Authorization: Bearer $APIFY_TOKEN" -o today.xlsx
```

**Comments from a profile's latest posts.** Pass a profile URL (`https://www.instagram.com/<handle>/`) instead of a post URL; `--max-comments` applies to each post and `--max-total` caps the run.

**No script at all.** One HTTP call starts a run, waits for it and returns the comments as CSV (`format=xlsx` for Excel, `format=json` for JSON). Nested fields come back flattened (`owner/username`, `owner/isVerified`, …):

```bash
curl -s -X POST "https://api.apify.com/v2/acts/yugenox~instagram-comments-scraper/run-sync-get-dataset-items?format=csv" \
  -H "Authorization: Bearer $APIFY_TOKEN" -H "Content-Type: application/json" \
  -d '{"postUrls":["https://www.instagram.com/reel/DdPsDCWRT-u/"],"maxComments":100,"sort":"top"}' -o comments.csv
```

**From Claude, ChatGPT or Cursor.** Apify's hosted MCP server can be pinned to this one Actor: `https://mcp.apify.com/?tools=fetch-actor-details,yugenox/instagram-comments-scraper`. You sign in to your own Apify account with OAuth; never put a token in the URL. Setup per client, and the same kind of pins for the Instagram and YouTube scrapers: [yugenox-mcp](https://github.com/ArpitGandhi1934/yugenox-mcp).

## Price

Pay per comment saved, no subscription and no start fee. Prices from the Actor's pricing in the Apify API on 2026-10-06; the [Store page][store] always shows the current ones.

| Comments | Free and Starter plans | Scale plan | Business plan and up |
|---:|---:|---:|---:|
| 1,000 | $1.90 | $1.50 | $1.10 |
| 10,000 | $19.00 | $15.00 | $11.00 |
| 100,000 | $190.00 | $150.00 | $110.00 |

- **Free Apify plans are limited to 100 comments per run.** Apify's Free plan includes monthly usage credit you can spend on this; the 10,000 and 100,000 rows assume a paid plan.
- Comments your filters skip are free, and so is everything the `top` sort reads but doesn't keep.
- `--max-charge` (default $2) is a hard cap Apify enforces for each run; the run stops at the number of comments it covers.

Other Instagram comment Actors on the Apify Store, per 1,000 comments, from the Apify API on 2026-10-06:

| Actor | Free | Starter | Scale | Business |
|---|---:|---:|---:|---:|
| **yugenox/instagram-comments-scraper** (this repo) | $1.90 | $1.90 | $1.50 | $1.10 |
| apify/instagram-comment-scraper | $2.60 | $2.30 | $2.10 | $1.90 |
| apidojo/instagram-comments-scraper | $0.50 | $0.50 | $0.50 | $0.50 |

apidojo's Actor has the lowest price per comment of the three. Ours includes a most-liked sort and keyword, like and date filters that run before billing, and reads up to about 4,800 comments per post.

## Why not the Instagram Graph API

The Instagram Graph API is the right tool for comments on your own posts: it returns comments on media owned by the Instagram professional accounts connected to your app, and it lets you reply and moderate. It doesn't return comments on other people's posts, and it needs a Meta developer app, a professional account, access tokens and, for most permissions, app review. For public posts you don't own, you need one Apify token here.

## Limits

- **Public posts and reels only.** Nothing here signs in to Instagram; private accounts are never accessed.
- **No replies.** Instagram shows replies, and how many replies a comment has, only to logged-in users, so they aren't included.
- **Fewer rows than the post's comment count** is normal: Instagram's count includes those replies plus hidden, restricted and deleted comments.
- Up to about 4,800 comments per post. `--sort top` picks the most-liked from up to 2,000 per post.

## Privacy

Comments are public, but usernames and comment text are personal data under PIPEDA (Canada), GDPR (EU) and similar laws. You are responsible for having a lawful reason to collect them, for how long you keep them and for what you do with them. Don't send commercial messages to commenters without consent (CASL, CAN-SPAM). See [Is web scraping legal?][legal]

## More

- Guides on [yugenox-data.vercel.app](https://yugenox-data.vercel.app): [export Instagram comments to Excel](https://yugenox-data.vercel.app/instagram/export-comments), [Instagram scraping without login](https://yugenox-data.vercel.app/instagram/scraper-no-login), [influencer vetting](https://yugenox-data.vercel.app/instagram/influencer-vetting), [use it in ChatGPT, Claude and Cursor](https://yugenox-data.vercel.app/ai/use-in-chatgpt-claude-cursor), [pricing calculator](https://yugenox-data.vercel.app/pricing-calculator)
- Posts, reels, profiles, hashtags and comments in one Actor, with examples for every input: [youtube-instagram-scraper-examples](https://github.com/ArpitGandhi1934/youtube-instagram-scraper-examples) ([Instagram Scraper][store-ig])
- Instagram Reels to text in bulk: [instagram-reels-transcript-bulk](https://github.com/ArpitGandhi1934/instagram-reels-transcript-bulk)
- The Actor itself, with its input form, output table and reviews: [Instagram Comments Scraper on Apify][store]. Other ways to call it (Python and JS clients, OpenAPI, MCP) are on its [API page][api]; n8n and Make can run it with their Apify node.

## Samples, testing and legal

- The sample comes from one run on 2026-10-06 (Actor build 0.1.4) through `export.mjs`; the .csv and .xlsx were then written from the pseudonymized JSON with `--from`. Input keys checked against the Actor's live input schema the same day.
- Not affiliated with or endorsed by Instagram or Meta. Instagram is a trademark of Meta Platforms, Inc.
- MIT licensed. Made by Yugenox Corporation.

<!-- All apify.com links for this README live below. When the Apify affiliate id exists, append ?fpr=<id> to these URLs only. -->
[store]: https://apify.com/yugenox/instagram-comments-scraper
[api]: https://apify.com/yugenox/instagram-comments-scraper/api
[badge]: https://apify.com/actor-badge?actor=yugenox/instagram-comments-scraper
[store-ig]: https://apify.com/yugenox/instagram-scraper
[signup]: https://console.apify.com/sign-up
[legal]: https://blog.apify.com/is-web-scraping-legal/
