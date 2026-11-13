# v3.0 - You do not need LightBurn

- **Frame & Cut on the Tooling tab.** Die Inserts, Die Holders, Die Organizer, Kerf Test, Speed & Power Test and Tonehole Crescent can each be sent straight to the laser, the same way as pads.
- **G-code for the Die Organizer and the Tonehole Crescent.** Both were SVG only. Each now has a material choice and cuts at size.
- **Cork** is a Tooling material, and the crescent's default.
- **Frame & Cut no longer needs a camera calibration.** Turn on direct machine control (File > Feature Set), plug in the laser, and the button is there. The camera features still need the calibration.
- **Plug the laser in whenever.** The app notices it being connected, switched off or moved to another USB socket while it is running.
- **Adhesive rings (beta).** A ring of double-sided tape per pad, for silicone wraps. Sizing Rules > Lesser-used settings.
- **The Toner is now the Analyzer.** It opens on an intonation chart: play your horn and see how sharp or flat each note is, compare runs from your library, export as PNG or PDF. The harmonic analysis is still there, behind a switch, and experimental.
- **A new User Guide.** Shorter, with a clickable contents list, search and figures. Help > User Guide. A Welcome window after each update links to it and to a one-page Quick Start.
- **Tooling jobs are in Job History.**
- **G-code files are for the laser, not LightBurn.** LightBurn ignores the speeds and powers a G-code file carries; for LightBurn, use Generate SVG. The app now says so when it writes one.
- Die inserts now nest from the bottom-left corner of the sheet, largest first, and fit more per sheet.
- Labeled Zones are placed on the Edge Bias side, with the other pads nested against them.
- A warning when an SD card would hold more than one G-code file (the Falcon runs only a card with one).
- A microphone plugged in after the app started is found, and remembered by name.
- In German, Spanish, French or Italian, the Kerf Test and Speed & Power Test now use the material's own settings.
- Bug fixes.
