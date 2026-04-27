# Web Data Sync & Screw Specs

## Web Data Sync ("Import Matt's")

The app can fetch reference data from https://www.stohrermusic.com:

- **Screw Specs**: File > "Import Matt's Specs" fetches `/data/screw_specs.json` and imports each model with a "(Matt's)" suffix to avoid overwriting local entries.
- **Key Heights**: File > "Import Matt's Key Heights" fetches `/data/key_height_library.json` and imports all presets into a dedicated "Matt's Library".

All use `urllib.request` (stdlib) with a 10-second timeout and `get_ssl_context()` from config.py for macOS SSL compatibility (tries certifi, then system certs, then unverified fallback). The JSON formats on the website match the app's internal format exactly, so no conversion is needed.

When adding new data to the website, update the JSON files in `C:\code\stohrermusic\static\data\` and push to deploy. App users can then re-import to get the latest.

## Screw Specs Submission Form

A web form at https://www.stohrermusic.com/articles/screw-specs-library/ (Submit Specs tab) allows colleagues to submit screw thread specifications. Submissions POST to a Google Apps Script endpoint that appends rows to a Google Sheet for review.

### Integrating Submitted Data

After reviewing submissions in the Google Sheet:

1. **Update the app's screw_specs.json**: Edit the user's local config file (platform config directory) to add verified entries. Format:
   ```json
   "Manufacturer": {
     "Model": {
       "neck_screw_th": "M4x0.7",
       "neck_screw_dia": "",
       "hinge_tiny_th": "",
       "hinge_tiny_dia": "",
       "hinge_small_th": "2-56 NC",
       "hinge_small_dia": "side keys",
       ...
       "notes": "source info here"
     }
   }
   ```

2. **Update the website library**: Copy the same data to `C:\code\stohrermusic\static\data\screw_specs.json`, then commit and push to deploy.

3. **Field mapping** (Google Sheet columns → JSON keys):
   - Neck Screw Thread/Desc → `neck_screw_th`, `neck_screw_dia`
   - Hinge Rod Tiny/Small/Medium/Large → `hinge_tiny_*`, `hinge_small_*`, `hinge_med_*`, `hinge_lrg_*`
   - Pivot Small/Large → `pivot_small_*`, `pivot_lrg_*`
   - Misc 1/2 → `misc1`, `misc2`
   - Notes → `notes`

Note: The `_dia` fields in the JSON are used for descriptions despite the legacy naming.

## Related Repository

The website at https://www.stohrermusic.com is a Hugo site with the Blowfish theme, located at `C:\code\stohrermusic`. The screw specs library page there loads data from `/static/data/screw_specs.json` and must be kept in sync with this app's data.
