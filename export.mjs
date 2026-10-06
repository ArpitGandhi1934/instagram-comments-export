#!/usr/bin/env node
// Export Instagram comments to CSV, Excel (.xlsx) or JSONL.
//
// Calls the yugenox/instagram-comments-scraper Actor on Apify for one or more public posts or
// reels, then writes one row per comment: username, text, likes, date, post and whether the
// comment is by the post's creator. Filters (likes, keywords, date) run inside the Actor, before
// billing, so comments they skip cost nothing.
//
// Run with --help for every option. Requires Node 20.6+; APIFY_TOKEN in the environment (or in
// .env, loaded with `node --env-file=.env`) for new runs and --dataset.

import { appendFileSync, existsSync, readFileSync, statSync, writeFileSync } from 'node:fs';

const ACTOR = 'yugenox/instagram-comments-scraper';

const HELP = `Export Instagram comments to CSV, Excel (.xlsx) or JSONL.

Usage:
  node --env-file=.env export.mjs <post or reel URL>... [options]
  node --env-file=.env export.mjs --file posts.txt --out comments.xlsx
  node --env-file=.env export.mjs --dataset <datasetId> --out comments.xlsx
  node export.mjs --from rows.json --out comments.xlsx

What to read
  <URL>...              Post or reel URLs (/p/<code>/, /reel/<code>/). A profile URL
                        collects comments from that profile's latest posts.
  --file <path>         One URL per line. Blank lines and lines starting with # are skipped.

Filters (they run on Apify before billing; comments they skip are free)
  --max-comments <n>    Comments to save per post (default 100). With filters, this
                        counts matching comments only.
  --sort <newest|top>   newest = Instagram's order (default). top = most-liked first,
                        picked from up to 2,000 comments per post.
  --min-likes <n>       Only comments with at least n likes.
  --keyword <text>      Only comments containing this word or phrase (case-insensitive).
                        Repeat it for several; a comment matching any of them is kept.
  --newer-than <when>   Only comments after a date (2026-09-01) or within a period
                        ("7 days", "24 hours").

Output
  --out <path>          .csv (default comments.csv), .xlsx or .jsonl
  --append              Add only comments whose comment_id is not in --out yet (.csv or
                        .jsonl). For a daily export into one growing file.
  --max-total <n>       Hard cap on comments across all posts in the run (default 1000).
  --max-charge <usd>    Hard cost cap for the run, enforced by Apify (default 2).

Re-export without a new run
  --dataset <id>        Export the dataset of a run you already made (Apify Console >
                        Storage > Datasets). No new run starts; filters are not applied.
  --from <file>         Convert a JSON or JSONL file of the Actor's rows (Console export)
                        or of this script's rows. Needs no token.

  -h, --help            Show this help.

Price: $1.90 per 1,000 comments saved on Apify's Free and Starter plans, $1.50 on Scale,
$1.10 on Business (checked 2026-10-06). Free plans stop at 100 comments per run.`;

// Column order for CSV, XLSX and JSONL.
const COLUMNS = ['username', 'text', 'likes', 'created_at', 'post_url', 'post_owner', 'is_post_owner',
    'is_verified', 'comment_id'];

function parseArgs(argv) {
    const o = { urls: [], keywords: [], out: 'comments.csv', maxTotal: 1000, maxCharge: 2, append: false };
    const num = (flag, v) => {
        const n = Number(v);
        if (v === undefined || !Number.isFinite(n) || n < 0) throw new Error(`${flag} needs a number`);
        return n;
    };
    const val = (flag, v) => {
        if (v === undefined || v.startsWith('--')) throw new Error(`${flag} needs a value`);
        return v;
    };
    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        if (a === '-h' || a === '--help') o.help = true;
        else if (a === '--file') o.urls.push(...readList(val(a, argv[++i])));
        else if (a === '--max-comments') o.maxComments = num(a, argv[++i]);
        else if (a === '--sort') o.sort = val(a, argv[++i]);
        else if (a === '--min-likes') o.minLikes = num(a, argv[++i]);
        else if (a === '--keyword') o.keywords.push(val(a, argv[++i]));
        else if (a === '--newer-than') o.newerThan = val(a, argv[++i]);
        else if (a === '--out') o.out = val(a, argv[++i]);
        else if (a === '--append') o.append = true;
        else if (a === '--max-total') o.maxTotal = num(a, argv[++i]);
        else if (a === '--max-charge') o.maxCharge = num(a, argv[++i]);
        else if (a === '--dataset') o.dataset = val(a, argv[++i]);
        else if (a === '--from') o.from = val(a, argv[++i]);
        else if (a.startsWith('-')) throw new Error(`Unknown option ${a} (see --help)`);
        else o.urls.push(a);
    }
    if (o.help) return o;
    if (o.sort && !['newest', 'top'].includes(o.sort)) throw new Error('--sort is newest or top');
    if (!/\.(csv|xlsx|jsonl)$/i.test(o.out)) throw new Error('--out must end in .csv, .xlsx or .jsonl');
    if (o.append && /\.xlsx$/i.test(o.out)) throw new Error('--append works with .csv or .jsonl (Excel opens the CSV)');
    const sources = [o.urls.length > 0, Boolean(o.dataset), Boolean(o.from)].filter(Boolean).length;
    if (sources === 0) throw new Error('Give post or reel URLs, --file posts.txt, --dataset <id> or --from <file> (see --help)');
    if (sources > 1) throw new Error('Use one of: URLs/--file, --dataset, --from');
    return o;
}

function readList(path) {
    return readFileSync(path, 'utf8').split('\n').map((s) => s.trim()).filter((s) => s && !s.startsWith('#'));
}

// One Actor row (or one row this script wrote earlier) -> one flat output row.
function toRow(r) {
    if ('comment_id' in r) return Object.fromEntries(COLUMNS.map((c) => [c, r[c] ?? null]));
    return {
        username: r.owner?.username ?? null,
        text: r.text ?? null,
        likes: r.likeCount ?? null,
        created_at: r.createdAt ?? null,
        post_url: r.postUrl ?? null,
        post_owner: r.postOwnerUsername ?? null,
        is_post_owner: r.isPostOwner ?? null,
        is_verified: r.owner?.isVerified ?? null,
        comment_id: r.id != null ? String(r.id) : null,
    };
}

const isComment = (r) => 'comment_id' in r || r.dataType === 'comment' || (r.dataType == null && 'text' in r && 'owner' in r);

// CSV cell. Text that starts with = + - or @ gets a leading ' so Excel and Google Sheets show it as
// text instead of running it as a formula (comments are written by strangers).
function csvCell(v, col) {
    if (v == null) return '';
    let s = String(v);
    if ((col === 'text' || col === 'username') && /^[=+\-@]/.test(s)) s = `'${s}`;
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

// Minimal RFC 4180 reader, enough to find the comment_id column of a CSV this script wrote.
function parseCsv(text) {
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (q) {
            if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
            else if (ch === '"') q = false;
            else cell += ch;
        } else if (ch === '"') q = true;
        else if (ch === ',') { row.push(cell); cell = ''; }
        else if (ch === '\n' || ch === '\r') {
            if (ch === '\r' && text[i + 1] === '\n') i++;
            row.push(cell); rows.push(row); row = []; cell = '';
        } else cell += ch;
    }
    if (cell || row.length) { row.push(cell); rows.push(row); }
    return rows;
}

function existingIds(path) {
    if (!existsSync(path) || statSync(path).size === 0) return new Set();
    const text = readFileSync(path, 'utf8').replace(/^\uFEFF/, '');
    if (/\.jsonl$/i.test(path)) {
        return new Set(text.split('\n').filter(Boolean).map((l) => JSON.parse(l).comment_id).filter(Boolean));
    }
    const [header, ...rows] = parseCsv(text);
    if (header.join(',') !== COLUMNS.join(',')) {
        throw new Error(`${path} has different columns than this script writes; use a new --out file`);
    }
    const idx = COLUMNS.indexOf('comment_id');
    return new Set(rows.map((r) => r[idx]).filter(Boolean));
}

async function writeXlsx(path, rows) {
    // Loaded only for .xlsx so CSV/JSONL users don't need it.
    const { default: writeExcelFile } = await import('write-excel-file/node');
    // Strip characters that are not allowed in XLSX (XML) cells.
    const clean = (s) => (s == null ? null : String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ''));
    const bold = (value) => ({ value, fontWeight: 'bold' });
    const data = [
        COLUMNS.map(bold),
        ...rows.map((r) => [
            { value: clean(r.username), type: String },
            { value: clean(r.text), type: String, wrap: true },
            r.likes == null ? null : { value: Number(r.likes), type: Number },
            r.created_at ? { value: new Date(r.created_at), type: Date, format: 'yyyy-mm-dd hh:mm' } : null,
            { value: clean(r.post_url), type: String },
            { value: clean(r.post_owner), type: String },
            r.is_post_owner == null ? null : { value: Boolean(r.is_post_owner), type: Boolean },
            r.is_verified == null ? null : { value: Boolean(r.is_verified), type: Boolean },
            { value: clean(r.comment_id), type: String, format: '@' },
        ]),
    ];
    const columns = [{ width: 22 }, { width: 70 }, { width: 8 }, { width: 17 }, { width: 44 }, { width: 16 },
        { width: 13 }, { width: 11 }, { width: 21 }];
    await writeExcelFile(data, { columns, stickyRowsCount: 1, sheet: 'Comments' }).toFile(path);
}

function writeOut(o, rows) {
    if (/\.xlsx$/i.test(o.out)) return writeXlsx(o.out, rows);
    if (/\.jsonl$/i.test(o.out)) {
        const body = rows.map((r) => JSON.stringify(r)).join('\n') + (rows.length ? '\n' : '');
        if (o.append) appendFileSync(o.out, body);
        else writeFileSync(o.out, body);
        return;
    }
    const lines = rows.map((r) => COLUMNS.map((c) => csvCell(r[c], c)).join(','));
    const fresh = !o.append || !existsSync(o.out) || statSync(o.out).size === 0;
    // UTF-8 with a BOM, so Excel shows emoji and accents correctly when it opens the file.
    if (fresh) writeFileSync(o.out, `\uFEFF${[COLUMNS.join(','), ...lines].join('\r\n')}\r\n`);
    else if (lines.length) appendFileSync(o.out, `${lines.join('\r\n')}\r\n`);
}

function readFromFile(path) {
    const text = readFileSync(path, 'utf8').replace(/^\uFEFF/, '').trim();
    if (!text) return [];
    const data = text.startsWith('[') ? JSON.parse(text) : text.split('\n').filter(Boolean).map((l) => JSON.parse(l));
    return Array.isArray(data) ? data : [data];
}

async function printRunReport(client, run) {
    // RUN_REPORT (in the run's key-value store) says what happened to each URL: ok, private,
    // not_found, capped... Optional: older runs may not have it.
    try {
        const rec = await client.keyValueStore(run.defaultKeyValueStoreId).getRecord('RUN_REPORT');
        const inputs = rec?.value?.queries;
        if (!Array.isArray(inputs)) return;
        for (const it of inputs) {
            const note = it.note ? ` (${it.note})` : '';
            console.error(`  ${it.status ?? '?'}  ${it.input ?? ''}  ${it.rows ?? 0} comments${note}`);
        }
    } catch { /* the report is a convenience; the export doesn't depend on it */ }
}

async function main() {
    const o = parseArgs(process.argv.slice(2));
    if (o.help) { console.log(HELP); return; }

    let items;
    if (o.from) {
        items = readFromFile(o.from);
    } else {
        if (!process.env.APIFY_TOKEN) throw new Error('Set APIFY_TOKEN (see .env.example), or use --from <file>');
        const { ApifyClient } = await import('apify-client'); // loaded here so --help and --from work before npm install
        const client = new ApifyClient({ token: process.env.APIFY_TOKEN });
        let datasetId = o.dataset;
        if (!datasetId) {
            const input = { postUrls: o.urls, maxItems: o.maxTotal };
            if (o.maxComments != null) input.maxComments = o.maxComments;
            if (o.sort) input.sort = o.sort;
            if (o.minLikes != null) input.minLikes = o.minLikes;
            if (o.keywords.length) input.keywords = o.keywords;
            if (o.newerThan) input.newerThan = o.newerThan;
            console.error(`Running ${ACTOR} on ${o.urls.length} URL(s) ...`);
            // maxTotalChargeUsd caps what this run can cost you (pay-per-event Actors only).
            const run = await client.actor(ACTOR).call(input, { maxTotalChargeUsd: o.maxCharge });
            console.error(`Run ${run.id} finished: ${run.status}`);
            await printRunReport(client, run);
            datasetId = run.defaultDatasetId;
        }
        items = [];
        for await (const it of client.dataset(datasetId).listItems()) items.push(it);
    }

    let rows = items.filter(isComment).map(toRow);
    let skipped = 0;
    if (o.append) {
        const seen = existingIds(o.out);
        const before = rows.length;
        rows = rows.filter((r) => !r.comment_id || !seen.has(r.comment_id));
        skipped = before - rows.length;
    }
    await writeOut(o, rows);

    const verb = o.append ? 'Appended' : 'Wrote';
    console.error(`${verb} ${rows.length} comments to ${o.out}${skipped ? ` (${skipped} already there)` : ''}`);
    const asked = Math.min(o.maxTotal, (o.maxComments ?? 100) * o.urls.length);
    if (!o.from && !o.dataset && items.length === 100 && asked > 100) {
        console.error('Exactly 100 comments: Free Apify plans stop at 100 comments per run; a paid plan lifts that.');
    }
}

main().catch((err) => {
    console.error(err.message);
    process.exit(1);
});
