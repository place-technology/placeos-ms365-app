/* eslint-disable no-undef */

// Builds the Teams / Outlook / Microsoft 365 app bar package (personal tab) as a zip.
//   npm run package:app:dev                                      -> https://localhost:3000/
//   ADDIN_URL=https://<customer>/outlook-addin/ npm run package:app
// Upload the zip in Teams (Apps > Manage your apps > Upload an app) or the Microsoft 365 admin center.

const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");

const addinUrl =
  process.argv[2] || process.env.ADDIN_URL || "https://placeos-dev.aca.im/outlook-addin/";
if (!/^https:\/\/.+\/$/.test(addinUrl)) {
  console.error(`Add-in URL must start with https:// and end with /: ${addinUrl}`);
  process.exit(1);
}
const addinHost = new URL(addinUrl).host;

const packageDir = path.join(__dirname, "..", "app-package");
const manifest = fs
  .readFileSync(path.join(packageDir, "manifest.json"), "utf8")
  .replace(/\{\{ADDIN_URL\}\}/g, addinUrl)
  .replace(/\{\{ADDIN_HOST\}\}/g, addinHost);

const zip = new AdmZip();
zip.addFile("manifest.json", Buffer.from(manifest, "utf8"));
zip.addLocalFile(path.join(packageDir, "color.png"));
zip.addLocalFile(path.join(packageDir, "outline.png"));

const outDir = path.join(__dirname, "..", "dist-app");
fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, `placeos-app-${addinHost.replace(/[^a-z0-9.-]/gi, "_")}.zip`);
zip.writeZip(outFile);
console.log(`Wrote ${outFile} for ${addinUrl}`);
