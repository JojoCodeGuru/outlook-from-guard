[README.md](https://github.com/user-attachments/files/34018867/README.md)
# From Guard – Outlook add-in

When you are in Denmark, From Guard keeps you from accidentally sending from an
`@aris-labs.com` address. When you're anywhere else, it does nothing.

| When | What happens |
| --- | --- |
| You reply, forward or write a new message | If From is @aris-labs.com and you're in Denmark, a yellow bar appears in the message: *"You're in Denmark – switch From away from aris-labs.com…"* |
| You change the From dropdown | The bar disappears when you switch to the other account, and comes back if you switch back to aris-labs.com |
| You press Send | If the bar's condition is still true, a pop-up asks you to choose **Don't Send** (go back and change From) or **Send Anyway** (manual override) |
| "Check From" button in the message toolbar | Shows the time zone the add-in sees and whether this message would trigger a warning |

**How it knows you're in Denmark:** it reads your Mac's time zone. If macOS sets
the time zone automatically, it's `Europe/Copenhagen` whenever you're in Denmark.
Turn that on in **System Settings → General → Date & Time → Set time zone
automatically using your current location**.

**Changing the rule:** edit `CONFIG` at the top of `web/launchevent.js`, then
upload the file to the website again. You don't need to reinstall the add-in.

## Files

```
manifest.template.xml   add-in manifest with {{HOST}} placeholders
make-manifest.sh        fills in your hosting address → manifest.xml
web/                    upload this folder to your HTTPS host
  commands.html         page Outlook loads to run the checks
  launchevent.js        the rules
  assets/icon-*.png     icons
```

## Setup

### 1. Host the `web/` folder on HTTPS (GitHub Pages, free)

Outlook loads add-ins from a web address, even personal ones. The files are
only static code and contain no personal data.

1. Sign in at github.com and create a new **public** repository named
   `outlook-from-guard`.
2. Choose **Add file → Upload files**, drag in the *contents* of `web/`
   (`commands.html`, `launchevent.js` and the `assets` folder), then select
   **Commit changes**.
3. Go to **Settings → Pages**, set **Source** to *Deploy from a branch*, choose
   `main` and `/ (root)`, and save.
4. After about a minute, the site is at
   `https://<your-github-username>.github.io/outlook-from-guard/`.
   Open `…/commands.html` to check that it loads (the page is blank).

### 2. Create the manifest

```bash
./make-manifest.sh https://<your-github-username>.github.io/outlook-from-guard
```

### 3. Install in Outlook

1. In Outlook, open **Get Add-ins** (in the toolbar or the **⋯** menu).
2. Go to **My add-ins → Custom add-ins → Add a custom add-in → Add from file…**
3. Choose `manifest.xml` and confirm.

### 4. Test

1. In Outlook, open an email you received on your aris-labs.com address and click
   **Reply**. The yellow bar should appear.
2. Switch From to your other account. The bar should disappear.
3. Switch back to aris-labs.com and press Send. The pop-up should appear. Choose
   **Don't Send**.

If nothing happens, click **Check From** in the reply toolbar. If even that
button is missing, the add-in didn't load; check the address in step 1.4.

## Known limits

- The add-in **can't change the From address itself**. Microsoft's add-in
  interface only lets add-ins read it.
- Microsoft only *guarantees* automatic checks like these for add-ins your
  Microsoft 365 admin deploys. Add-ins installed from a file usually run them
  too, but if the yellow bar and pop-up never appear while **Check From** works,
  ask IT to deploy this manifest via **Microsoft 365 admin center → Settings →
  Integrated apps → Upload custom apps**.
- With a VPN, your location is unaffected, because the add-in uses the Mac's
  time zone rather than your IP address. If you set the time zone by hand,
  that's what counts.
- The checks need an internet connection to load from GitHub Pages. If they
  can't load, Outlook sends normally rather than blocking you.
