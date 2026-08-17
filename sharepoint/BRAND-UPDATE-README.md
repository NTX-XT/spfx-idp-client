# Nintex SPFx Brand Update — SCSS files

Drop these two files into the `sp-only/` tree and rebuild.

## Files delivered

| File | Destination |
|------|-------------|
| `NacTaskActions.module.scss` | `src/webparts/nacTaskActions/components/` |
| `NacForms.module.scss` | `src/webparts/nacForms/components/` |

No other source files changed — all SCSS class names are unchanged so no TSX edits are needed.

---

## What changed and why

### Typography
- **Font**: `Plus Jakarta Sans` loaded via Google Fonts `@import` at the top of each SCSS file. This is Nintex's brand typeface (brand guidelines p.35). Falls back to `Arial` per brand spec (p.36).
- Font sizes and weights now align with the brand hierarchy: headings 700 weight, CTAs 800 weight, body 400 weight.

### Colour palette
All hard-coded SharePoint/Fluent blue (`#0078d4`) and orange (`orange`) values replaced with brand-approved tokens:

| Old value | New value | Brand token |
|-----------|-----------|-------------|
| `#0078d4` | `#8439A6` | Purple (primary) |
| `#106ebe` | `#3D0456` | Deep Purple |
| `orange` / `#FF6D00` button | `#BE0075` | Pink 300 (CTA per styleguide `.nintex-btn`) |
| `#107c10` (green approve) | `#8439A6` / `#3D0456` | Purple / Deep Purple |
| `#a4262c` (red reject) | `#BE0075` | Pink 300 |
| `#f3f2f1` (grey bg) | `#FFF9F4` | Ivory |
| `#e8f3ff` (blue badge bg) | `#F1D3FF` | Pale Purple |
| `#dff6dd` (green badge bg) | `#FEE9E3` | Pale Orange |

### Buttons (Login / Confirm / Cancel / Outcome)
- Primary buttons (Login, Confirm) → **Pink 300 fill**, white text, extra-bold, **no border-radius** — matching `.nintex-btn` from `nintex-styleguide-formatted.css`.
- Secondary buttons (Cancel, Reject, neutral outcomes) → **purple outline**, transparent fill — matching `.nintex-btn--secondary`.
- Approve outcome → **Purple fill**, white text.
- Hover states use Pink 100 (`#ED2891`) on primary, Deep Purple on secondary, per styleguide hover definitions.

### Cards
- Left accent border (`3px solid #8439A6`) on task and form cards — a common brand application pattern.
- Hover shadow uses a purple-tinted `rgba(132, 57, 166, 0.12)` instead of neutral black.
- Border colour `#e1dfdd` → `#F1D3FF` (Pale Purple).

### Filter pills
- Rounded pill shape (`border-radius: 100px`) with purple outline/text when inactive, purple fill when active.

### Status badges
- Active/pending → Pale Purple bg + Deep Purple text.
- Completed/approved → Pale Orange bg + Deep Navy text.
- Rejected/expired → kept a legible red-tinted pair (no brand-approved red is in the primary palette, closest is the specialty Red `#B22E00` but that's agency-only; the bg is `#FDE7E9` which was already present).

### Welcome / logo area
- Horizontal divider using Pale Purple border instead of none.
- `welcomeImage` constrained to `height: 32px` so the logo renders at a sensible webpart size alongside SharePoint chrome.

### Dialogs and banners
- Confirmation dialog → top accent border in Purple.
- Success banner → Pale Orange bg with Nintex Orange left stripe.
- Error banner → pink-tinted bg with Pink 300 left stripe.

---

## Logo files (manual step required)

The `NintexLogos_Print_Screen.zip` contains SVG/PNG files. SPFx imports logos as ES modules inside the TSX files:

```ts
import nintexColorLogo from '../assets/nintex-color.svg';
import nintexReversedLogo from '../assets/nintex-reversed.svg';
```

**Recommended replacements from the zip:**

| Asset slot | File to use |
|---|---|
| `nintexColorLogo` (light bg) | `NintexLogos_Screen_RGB/Main/Nintex_Logo_Main_FullColor_RGB.svg` |
| `nintexReversedLogo` (dark bg) | `NintexLogos_Screen_RGB/Main/Nintex_Logo_Main_White_Orange_RGB.svg` |

Copy the chosen SVGs into each webpart's `assets/` folder and rename to match the existing import paths (or update the import paths). No TypeScript changes are needed beyond updating the file names.

---

## Build

```bash
cd sharepoint/sp-only
npm install --legacy-peer-deps
npm run build
```

The Google Fonts `@import` in the SCSS will be bundled as a CSS `@import` in the output. SharePoint pages need outbound internet access to `fonts.googleapis.com` for the font to load — if that's blocked on-premises, host the font files locally and update the `@import` URL accordingly.
