# Equinox for Word

The Word for Mac task pane for Eureka IP's Equinox add-on (feature 13.7). It talks to the Equinox Apps Script web app (`Word.gs`, features 13.1–13.6). The contract between the two is in the project doc `claude/equinox-word-api.md`.

```
word/taskpane.html   the pane
word/taskpane.js     what it does (pane version in PANE_VERSION)
word/taskpane.css    how it looks (follows the Mac's light or dark mode)
word/assets/         icons
word/manifest.xml    tells Word where the pane lives and adds the Equinox ribbon tab (made by make-manifest.py)
templates/letters/   the letter parts as kept in Drive: Master.docx, a content template, Headers/, Sign-offs/, Emails/
make-manifest.py     writes manifest.xml for a hosting address
```

The pane holds no secrets. The web app address and token are typed into the pane's Settings and kept in Word on the Mac, so this repository can be public.

## Hosting (GitHub Pages)

Settings → Pages → Build and deployment → Source: **Deploy from a branch**, Branch: **main**, folder **/ (root)** → Save. The pane is then at `https://davidperkins2000.github.io/<repository>/word/taskpane.html`. If the repository isn't `equinox-word`, run `python3 make-manifest.py https://davidperkins2000.github.io/<repository>` and commit the new manifest.

## Installing in Word for Mac (sideloading, once)

1. Quit Word.
2. In Finder, choose Go → Go to Folder…, and paste `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`. If the `wef` folder isn't there, create it (inside `Documents`).
3. Copy `word/manifest.xml` into that folder.
4. Open Word and any document. Choose **Home → Equinox** (or Insert → Add-ins → My Add-ins → Equinox).
5. The pane opens on Settings. Paste the web app address (ends in `/exec`) and the token, both from the script editor's Execution log after running `step24_wordSetup`. Choose **Save and test**. You should see "Connected to Equinox v0.0.81".

## Updating

- **The pane** (HTML, JS or CSS): commit and push. Pages republishes within a minute or two. Word may keep the old copy cached. To clear it, quit Word, delete `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/` and `~/Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`, then reopen Word.
- **The script** (`Word.gs` etc.): Deploy → Manage deployments → Edit → Version: New version → Deploy. The address stays the same.
- **The manifest:** only when the hosting address, name, icons or button change. Bump `VER` in `make-manifest.py`, regenerate, copy to `wef` again and restart Word.

## What the pane does

- **Finds the matter** from the document's name, else from its folder in Google Drive for desktop (a matter folder from the folder sync, or a matter code in a folder name). It also takes a matter code, associate or client reference, or official number typed in the box. If several matters match, each shows as a button.
- **Matter details:** title, applicants, status, next fatal deadline, numbers linked to the registers, associate firm and reference, and links to Equinox (details, documents, family, contacts) and the Drive folder. The picture is optional (Settings).
- **File to Equinox:** sends the document as Word has it now, including unsaved changes. Each version files once; an edited document files as a new one. "Also file to" adds the other matched matters.
- **Insert summary:** the matter summary table at the cursor, optionally with the next deadline and the picture.
- **Open from Equinox:** lists the matter's Equinox documents. Choosing one saves it into the matter's Drive folder (never overwriting) and shows where it is on the Mac, so you open and edit it there. "Open a copy now" opens an unsaved copy, for reading.

## Letters (pane w0.3.0, script v0.0.101)

The **Equinox** tab on Word's ribbon has: Equinox (the pane), File to Equinox, New letter, Fill fields, Save PDF and Email PDF. Each button opens the pane, which does that job once it knows the matter.

- **Templates** are ordinary Word files in the Drive folder **Equinox letter templates**, next to the index sheet (the script makes it the first time the pane asks). Start with `templates/Blank letter.docx`. Write Equinox's merge codes in the text, e.g. `#MATTER#`, `#APPLICANT#`, `#H2CODE#`, `#ADDRESSNAME#`, `#ADDRESS#`, `#SALUTATION#`. An `.html` file with the same name (`Blank letter.html`) is the cover email used by Email PDF.
- **New letter** opens the chosen template as a new document for the matter and recipient chosen in the pane; the pane in the new window fills it. **Fill fields** fills any open document. A code with nothing in Equinox is left out, and so is a line holding only its label (e.g. "By email:"). Codes the add-on doesn't know stay in the text.
- **Save PDF** saves the document as a PDF in the matter's Drive folder (never overwriting). **Email PDF** does that and makes a Gmail draft with the PDF attached, the matter subject line, the recipient's email and the cover email. Once sent, it is filed to Equinox with the PDF (Outgoing switch on).
- The signer lines use Script Properties `LETTER_NAME`, `LETTER_TITLE` and `LETTER_PHONE` if set, else the matter's responsible user in Equinox.

After this update, copy the new `word/manifest.xml` (1.0.4.0) into the `wef` folder and restart Word, so the Equinox tab appears.

## Letters from parts (pane w0.4.0, script v0.0.102)

Like Equinox's Smart Templates, a letter is stacked from parts in the Drive folder **Equinox letter templates**:

```
Master.docx              letterhead, page, footer and the Word styles; the letter goes where it says #EQX_LETTER#
Headers/Letter.html      date, recipient, references, salutation, subject
General letter.html      the content; its first lines name the parts:  <!-- master: Master -->  <!-- header: Letter -->  <!-- signoff: Letter -->
Sign-offs/Letter.html    closing and signature
Emails/General letter.html   the cover email for Email PDF
```

A paragraph's class is its Word style (`class="letter-subject"` → "Letter Subject"); plain `<p>` is Normal, and `<h1>`–`<h3>`, lists and `<blockquote>` use Word's built-in styles. `<strong>`, `<em>`, `<u>`, `<a href>` and `<br>` (a new line in the same style) carry across. New letter has the script fill the parts, opens the master as a new document (it opens with the Equinox pane), and the pane there writes the letter in place. Whole-document .docx templates in the folder still work.
