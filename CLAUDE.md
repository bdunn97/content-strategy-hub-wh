# CLAUDE.md: Wilder Home Creator Content Strategy

## What this is
A single-page internal site ("Wilder Home Creator Content Strategy") for Wilder Home's creator and
UGC content work.
Hosted on Netlify and deployed from the GitHub repo `bdunn97/content-strategy-hub-wh`: every push
to `main` redeploys the live site.

- Live site: https://content-strategy-hub-wh.netlify.app/
- Local preview: http://localhost:4324 (`.claude/launch.json` "hub-preview", or `Start Hub.bat`) Layout and styling mirror the Content Marketing Hub
(sticky tab bar, uppercase headings, red accent, 9:16 video carousels with a pop-up player).

Tabs (in nav order; page ids in brackets):
- Competitor Analysis ▾: Ad Library [`ad-library`], Organic Content [`organic-content`]
- Creative ▾: Creative Brief [`creative-brief`], Scripts [`scripts`]
- Creators ▾: Potential [`potential-creators`], Signed [`signed-creators`]
- Content Calendar [`content-calendar`]
- Strategy [`strategy`]

Ad Library holds the competitor video carousels and is the landing page (first `<main>` in the
file). Old links `#competitor-ads` and `#strategy-proposal` redirect to `#ad-library` and `#strategy`. Pages without content
yet show a "No content yet" placeholder.

## Structure
```
index.html            The whole site: CSS, HTML and script. Tabs are <main class="page" id="...">
                      blocks plus a matching <a class="tab" data-tab="..."> link in the nav.
videos.js             List of competitor videos, grouped by brand. Generated; don't hand-edit.
update-videos.ps1     Rebuilds videos.js from the "Competitor Ads" folder.
Competitor Ads/
  <Brand>/            One folder per competitor brand = one carousel row on the Ad Library page.
    <Brand> NN.mp4    "Music Only" in the filename shows a "Music only" badge on the card.
serve.ps1             Local server for http://localhost:4324 (also used by .claude/launch.json
                      "hub-preview"). Start Hub.bat starts it with a double-click.
```

- No build step and nothing to install. Netlify serves the repo as is (publish directory `/`).
- Video thumbnails are frames from the videos themselves (no separate image files).

## Common tasks
**Add or remove competitor videos**
1. Put the files in `Competitor Ads/<Brand>/` (new brand = new folder).
2. Run `powershell -ExecutionPolicy Bypass -File .\update-videos.ps1` to rebuild `videos.js`.
3. Preview with the "hub-preview" launch config, then commit and push.

**Add a new tab**: add a `<main class="page" id="new-id" hidden>` section and a nav link with
`class="tab"`, `href="#new-id"` and `data-tab="new-id"`. For a dropdown item, add the link inside
that `.menu` (no `tab` class) and add the id to the dropdown button's `data-pages`. Routing picks
it up automatically; the first `<main>` in the file is the default page.

## Rules
- Keep every video file under 100 MB (GitHub's per-file limit). Compress anything larger first.
- Commit and push only when the user asks; pushing to `main` publishes the live site.
- Keep the existing look (Content Marketing Hub style) unless asked to change the design.
