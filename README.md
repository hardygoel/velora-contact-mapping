# Contact Mapping Viewer

This is a standalone, local-first viewer for the generated WhatsApp contact mapping. It does not contain the contact data and does not send selected files to an API.

## Run locally

From the repository root:

```powershell
python -m http.server 4173 --directory contact-mapping-app
```

Open `http://localhost:4173/` and load `C:\Users\hardy\Downloads\whatsapp_contact_mapping.xlsx`.

The workbook parser is loaded from SheetJS CDN. CSV and JSON imports are also supported. The production VPS is not modified by this app.

## GitHub Pages

This folder is deploy-ready as a standalone static repository. The included workflow publishes only these viewer files; the private workbook is intentionally outside this folder. After pushing the folder to a GitHub repository, enable Pages with the GitHub Actions source.
