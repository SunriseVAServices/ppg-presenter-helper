# PPG Presenter Helper

A small app for the Friday presentation laptop. When a presenter starts, it opens **their own file in the real app** (PowerPoint on Windows or Mac, Keynote on a Mac) or **their Canva link full screen**, and their phone's Next / Back buttons move it. When they finish, it closes and the Stage comes back.

If the helper isn't running, or a file won't open, nothing breaks: that presenter plays from the prepared slides on the Stage, exactly as before, and your Remote tells you why.

| Presenter sent | Windows PC | Mac |
|---|---|---|
| PowerPoint (.pptx) | PowerPoint | PowerPoint (or Keynote if PowerPoint isn't installed) |
| Keynote (.key) | Prepared slides | Keynote |
| Google Slides link | PowerPoint (downloaded as .pptx) | PowerPoint / Keynote |
| Canva link | Full screen on the TV | Full screen on the TV |
| PDF / pictures | Prepared slides | Prepared slides |

---

## One-time setup: build the installers on GitHub

You do this once (and again only when there's a new version).

1. **Make the repository.** On github.com, click **+ › New repository**. Name it `ppg-presenter-helper`.
   - **Public** if other leaders will download the helper from the link on the Remote (there are no passwords or keys in this code).
   - **Private** if only you will install it (the download link only works when you're signed in to GitHub).
2. **Upload the files.** Unzip `ppg-presenter-helper.zip`. In the new repository click **uploading an existing file**, drag in everything inside the folder (`src`, `build`, `.github`, `package.json`, `README.md`, `.gitignore`), and click **Commit changes**.
   - Check that a `.github` folder appears in the repository. If it doesn't (some computers hide folders starting with a dot), click **Add file › Create new file**, name it `.github/workflows/build.yml`, paste in the contents of that file from the zip, and commit.
3. **Build.** Click **Releases › Create a new release**. Under **Choose a tag** type `v1.0.0` and click **Create new tag**. Title it `1.0.0`, then **Publish release**.
4. **Wait about 10 minutes.** The **Actions** tab shows the build (yellow while it runs, green when done). The release page then has two downloads:
   - `PPG-Presenter-Helper-win.exe` (Windows)
   - `PPG-Presenter-Helper-mac.dmg` (Mac)
5. **Put the link on the Remote.** Copy the link `https://github.com/YOUR-NAME/ppg-presenter-helper/releases/latest`. In WordPress go to **Weekly Presentation › Phone remote › Settings**, paste it into **Presenter Helper download link** and save.

## Install it on the Windows PC

1. Download and run `PPG-Presenter-Helper-win.exe`.
2. Windows may say **"Windows protected your PC"** (the app isn't signed with a paid certificate). Click **More info › Run anyway**.
3. The helper opens and shows a **6-digit code**. On your phone: PPG app › **Me › Leadership › Run Friday's presentation**. Next to **Presenter Helper**, tap **Connect** and type the code.
4. Done. The helper lives by the clock in the taskbar (click the little screen icon to open it) and starts with the computer.

**PowerPoint tip:** with the TV plugged in, open any presentation once, go to **Slide Show** and set **Monitor** to the TV. Turn **Use Presenter View** on if you want the laptop to show the next slide and the TV only the slides.

## Install it on the Mac

1. Download `PPG-Presenter-Helper-mac.dmg`, open it and drag the app into **Applications**.
2. The first time, **right-click** the app › **Open** › **Open** (it isn't signed with a paid certificate).
3. Connect it with the code, the same as on the PC. It lives in the menu bar.
4. The first time a presenter starts, macOS asks **"PPG Presenter Helper wants to control Keynote"** (or PowerPoint). Click **OK**. If you clicked Don't Allow: **System Settings › Privacy & Security › Automation** and switch it on.

**Keynote tip:** Keynote › Settings › Slideshow lets you choose which display shows the slides.

## On Friday

- Open the Stage on the laptop like always, then make sure the helper is running (your Remote shows **Presenter Helper on [computer] · ready**).
- Each presenter in **Today's presenters** shows how they'll play: *PowerPoint*, *Keynote*, *Canva, full screen* or *Slides*.
- When a presenter taps **Start** on their phone, their file opens on the TV in a few seconds. Their phone, your Remote, and the laptop's own clicker all move it.
- **End presentation**, the next presenter's **Start**, or any quick screen closes it and puts the Stage back.
- In the helper window, **Canva shows on** picks the screen for Canva links. **Test** flashes a card on that screen so you can check it.

## Disconnecting a computer

From the helper: **Disconnect**. From the website: **Weekly Presentation › Phone remote › Presenter Helper › Disconnect**.

## New versions

Change `"version"` in `package.json` (for example to `1.0.1`), then publish a new release with tag `v1.0.1`. The latest-release link on the Remote picks it up by itself. Install the new download over the old one; it stays connected.

## Signing (optional, later)

Without paid certificates, Windows and macOS show the one-time warnings above. You have an Apple Developer account, so the Mac app can be signed and notarized later to remove the warning; Windows needs a code-signing certificate from a provider.

---

### For developers

- `src/core.js`: talks to the website (`/wp-json/ppg/v1/rm/h/pair`, `rm/h/poll`, `rm/h/file`). Plain Node.
- `src/ppt-win.js`: PowerPoint on Windows through COM, from one long-running PowerShell.
- `src/mac.js`: Keynote and PowerPoint on a Mac through AppleScript.
- `src/main.js`: the Electron tray app, the Canva window, and settings. The connection key is stored encrypted with the system keychain (Electron `safeStorage`).
- Run locally: `npm install` then `npm start`.
