# Custom enum repro: "One or more add-ins failed to install custom functions"

Minimal Excel add-in that shows that custom functions whose metadata contains a **custom enum**
(`enums` + `customEnumId`) fail to install in the **second** workbook opened in the same Excel
process. Excel then shows *"One or more add-ins failed to install custom functions."*

The same add-in without the enum installs fine in every workbook.

## What it contains

- `manifest.xml`: shared runtime add-in, metadata `functions.json` **with** the custom enum `STYLE`.
- `manifest-noenum.xml`: identical add-in (different id), metadata `functions-noenum.json` **without** enums.
- Two functions in the `ISSUE` namespace: `ISSUE.PLAIN` (no enum) and `ISSUE.WITHENUM` (parameter with `"customEnumId": "STYLE"`).

## Prerequisites

- Node.js.
- Trusted `localhost` certificates named `localhost.pem` + `localhost-key.pem` in this folder
  (e.g. via [`mkcert`](https://github.com/FiloSottile/mkcert): `mkcert -install && mkcert localhost`).

## Run

```bash
npm install
npm start        # serves this folder over https://localhost:3100
```

## Sideload (Excel on Windows)

Share a folder containing `manifest.xml` and register it as a trusted catalog
(**File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**), then
**Home → Add-ins → Shared Folder → customEnum repro**. Sideload only one of the two manifests at a time.

Optional: enable runtime logging to see the failure
(`HKCU\Software\Microsoft\Office\16.0\WEF\Developer\RuntimeLogging`, default value = path of a log file).

## Reproduce

1. Close Excel. Clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
2. Start Excel with a blank workbook (Book1).
3. Open a second workbook: `Workbook1.xlsx` in this folder (one cell of text) or any other `.xlsx`.

**Expected:** no message; `=ISSUE.WITHENUM(1)` works in both workbooks.

**Actual:** *"One or more add-ins failed to install custom functions."* In the runtime log,
the second workbook gets `[Installation] [Functions] [Begin]` and a successful `[Registration]` for
every function, but `[Installation] [Functions] [End] [Success]` never appears. For Book1 it does.

Repeat with `manifest-noenum.xml`: no message, and both workbooks get the `[End] [Success]` line.

## Not affected

- A single workbook (Excel started by double-clicking an `.xlsx`).
- The workbook content: reproduced with an empty workbook.
