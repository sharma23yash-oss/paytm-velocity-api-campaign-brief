# Paytm Velocity API — Mock B2B Campaign Brief

A **mock / illustrative** go-to-market advertising campaign brief for a *fictional* product, **Paytm Velocity API** — an ultra-fast, drop-in UPI payments API for e-commerce platforms and online businesses.

> ⚠️ This is a fictional exercise. "Paytm Velocity API" is not a real product. All figures are illustrative, modelled on 2025–26 Indian B2B / fintech advertising benchmarks. Paytm brand colours are used for realism only.

## What's inside

| File | Description |
|------|-------------|
| `Paytm_Velocity_API_Campaign_Brief.pptx` | The 13-slide campaign brief deck (the deliverable). |
| `build.js` | Node script that generates the deck from scratch with [PptxGenJS](https://gitbrent.github.io/PptxGenJS/). |
| `package.json` | Dependencies for the build. |

## The brief covers

1. **Audience research** — the 18–60 B2B audience split into three decision-maker personas (the Conversion-Obsessed Founder, the Pragmatic Tech Lead, the Digitizing Retail Owner), each with pain points and buying triggers.
2. **Ad creatives** — a brand kit (palette + type + motif) plus three detailed creative spec sheets ready to design in Canva/Figma.
3. **Ad copy & CTAs** — primary text, headlines, and CTA buttons for each creative.
4. **Targeting strategy** — a full-funnel plan across Meta, LinkedIn, and Google Search (titles, interests, keywords, formats).
5. **Mock campaign results** — a realistic 30-day report (₹30L spend): CTR 1.15%, CPC ₹67, ROAS 4.2× (LTV-based/projected), with the rationale for each number.

## Rebuild the deck

```bash
npm install
node build.js
# writes Paytm_Velocity_API_Campaign_Brief.pptx
```

Requires Node 18+. To render slides to images for review you'll also need LibreOffice (`soffice`) and Poppler (`pdftoppm`).

## Key realism note

UPI person-to-merchant payments are **zero-MDR** in India, so the modelled ROAS is **not** built on transaction fees. It is LTV-driven — platform/subscription fees plus cross-sell to cards, EMI, and value-added services — and projected over 12 months. In-month realised ROAS is ~0.4× because fintech payback typically lands at month 4–6.

---

*Generated as a Cowork exercise. Not affiliated with or endorsed by Paytm / One97 Communications.*
