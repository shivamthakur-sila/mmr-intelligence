# LibreOffice with the fonts the decks are drawn in.
#
# The stock image has only DejaVu/Noto, so Georgia and Calibri fell back to
# fonts about 10-15% wider, and every layout constant in lib/deck-layout.ts
# was tuned against text PowerPoint never draws. This image has the real
# Georgia and Trebuchet MS, Carlito (metric-compatible with Calibri), Gelasio
# (metric-compatible with Georgia, as a fallback) and Merriweather, SILA's
# brand headline face. A line that wraps here wraps in PowerPoint.
#
# The core fonts are fetched by Microsoft's own installers at build time and
# not redistributed; confirm the licence before publishing an image built
# from this file. The others are under the SIL Open Font License.
#
# Build:  docker build -t mmr-render render/
# Use:    python render/bench.py deck.pptx   (render, contact sheets, fill)
#         PREFLIGHT_SOFFICE_IMAGE=mmr-render for lib/preflight/render.ts
FROM linuxserver/libreoffice:latest

# Georgia and Trebuchet MS themselves come from Microsoft's core-fonts
# installers (msttcorefonts-installer fetches the unmodified originals);
# Carlito and Gelasio stay as the metric clones for Calibri and as fallback.
# poppler-utils gives pdftoppm, which lib/preflight/render.ts prefers.
RUN apk add --no-cache font-carlito font-merriweather curl msttcorefonts-installer fontconfig poppler-utils \
 && update-ms-fonts \
 && mkdir -p /usr/share/fonts/gelasio \
 && curl -fsSL -o "/usr/share/fonts/gelasio/Gelasio[wght].ttf" \
      "https://github.com/google/fonts/raw/main/ofl/gelasio/Gelasio%5Bwght%5D.ttf" \
 && curl -fsSL -o "/usr/share/fonts/gelasio/Gelasio-Italic[wght].ttf" \
      "https://github.com/google/fonts/raw/main/ofl/gelasio/Gelasio-Italic%5Bwght%5D.ttf" \
 && fc-cache -f

# Map the names the decks ask for onto the metric clones, so LibreOffice
# picks them deliberately rather than through its own fallback table.
COPY fonts.conf /etc/fonts/local.conf
RUN fc-cache -f
