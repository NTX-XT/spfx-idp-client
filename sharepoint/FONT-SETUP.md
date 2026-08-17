# Font Setup — Plus Jakarta Sans (bundled)

Download the font files before building. The SCSS uses @font-face with relative
paths so no internet access is needed at runtime.

## Step 1 — Download

Go to: https://fonts.google.com/specimen/Plus+Jakarta+Sans
Click "Download family" → extracts to PlusJakartaSans/

Or from the GitHub source (static/woff2 folder):
https://github.com/tokotype/PlusJakartaSans/tree/master/fonts/webfonts

## Step 2 — Copy these 4 files into BOTH webpart asset folders

  src/webparts/nacTaskActions/assets/fonts/
  src/webparts/nacForms/assets/fonts/

Files needed:
  PlusJakartaSans-Regular.woff2      (weight 400)
  PlusJakartaSans-SemiBold.woff2     (weight 600)
  PlusJakartaSans-Bold.woff2         (weight 700)
  PlusJakartaSans-ExtraBold.woff2    (weight 800)

These filenames are what the @font-face declarations in the SCSS expect.
Rename the downloaded files to match exactly if they differ.

## Step 3 — Logo replacement

In: src/webparts/nacTaskActions/assets/
    src/webparts/nacForms/assets/

Replace:
  nintex_RGB_color_500.png    → NintexLogos_Screen_RGB/Main/Nintex_Logo_Main_FullColor_RGB.png
  nintex_RGB_reversed_500.png → NintexLogos_Screen_RGB/Main/Nintex_Logo_Main_White_Orange_RGB.png

Keep the ORIGINAL filenames (nintex_RGB_color_500.png etc) — just overwrite
the files. The TSX imports will continue to work with no code changes.

## Step 4 — Build

  cd sharepoint/sp-only
  npm install --legacy-peer-deps
  npm run build
