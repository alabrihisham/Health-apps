# Deploying SutureLog with GitHub + Vercel

GitHub alone only stores and version-controls code — it doesn't run anything.
To get a live link with the AI features (photo/PDF/dictation scanning) working,
you need something that can run a small server function too, so your Anthropic
API key stays private. **Vercel** does this and connects directly to GitHub,
so that's the path below. (Netlify or Cloudflare Pages work the same way if
you'd rather use one of those.)

## 1. Get an Anthropic API key

Go to [console.anthropic.com](https://console.anthropic.com), create a key,
and add billing — calls made outside Claude's own apps are billed per-use on
your account (this differs from using Claude.ai directly).

## 2. Push this folder to GitHub

```bash
cd suturelog          # this folder
git init
git add .
git commit -m "Initial commit"
```

Then on [github.com](https://github.com), create a new empty repository (no
README/license, since you already have files), and push:

```bash
git remote add origin https://github.com/<your-username>/suturelog.git
git branch -M main
git push -u origin main
```

## 3. Connect the repo to Vercel

1. Go to [vercel.com](https://vercel.com) and sign in with your GitHub account.
2. Click **Add New → Project**, and pick the `suturelog` repo.
3. Vercel auto-detects it as a Vite project — leave the build settings as
   default (build command `vite build`, output directory `dist`).
4. Before deploying, open **Environment Variables** and add:
   - `ANTHROPIC_API_KEY` = the key from step 1
5. Click **Deploy**.

A minute or two later you'll have a live link like
`https://suturelog-yourname.vercel.app` — that's your app.

## 4. Every future update

```bash
git add .
git commit -m "describe your change"
git push
```

Vercel redeploys automatically on every push to `main`.

## Notes

- **Storage is per-device.** The `localStorage` polyfill (`src/storagePolyfill.js`)
  keeps each person's cases in their own browser, not shared across devices —
  there's no account system here. If you need cases to follow you across
  devices, that's a real backend (a database) — a bigger step than this kit
  covers.
- **HTTPS is required** for the camera, microphone, and barcode scanner to
  work in the browser. Vercel gives you HTTPS by default, so this isn't
  something you need to configure.
- **Barcode scanning and dictation** only work in Chromium browsers (Chrome,
  Edge) — this is a browser limitation, not something the deployment changes.
- If you'd rather not manage a backend at all, the earlier "static site"
  option (no photo/PDF/dictation scanning, everything else works) can go
  straight onto **GitHub Pages** with no server function needed. Ask if you'd
  like that variant instead.
