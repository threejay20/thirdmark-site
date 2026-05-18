"use strict";

const PptxGenJS = require("C:/Users/3Jay/AppData/Roaming/npm/node_modules/pptxgenjs");
const path = require("path");
const fs = require("fs");

// ─── Design Tokens ───────────────────────────────────────────────────────────
const BG        = "0d0d1a";
const BG_LIGHT  = "12122a";
const SURFACE   = "1a1a2e";
const BORDER    = "2d2d45";
const WHITE     = "ffffff";
const MUTED     = "a0a0b8";
const ACCENT    = "818cf8";
const ACCENT2   = "c084fc";

const PAD = 0.6;            // slide margin
const SW  = 13.33;          // slide width
const SH  = 7.5;            // slide height
const CONTENT_W = SW - PAD * 2;

// ─── Helpers ─────────────────────────────────────────────────────────────────

function sectionLabel(slide, text, y = 0.42) {
  slide.addText(text, {
    x: PAD, y, w: 3, h: 0.22,
    fontSize: 9, bold: true, color: ACCENT,
    fontFace: "Calibri", charSpacing: 2,
    margin: 0,
  });
}

function headline(slide, text, y = 0.72, size = 24, w = CONTENT_W) {
  slide.addText(text, {
    x: PAD, y, w, h: 0.9,
    fontSize: size, bold: true, color: WHITE,
    fontFace: "Calibri", margin: 0,
    wrap: true,
  });
}

// ─── Build deck ──────────────────────────────────────────────────────────────

async function buildDeck() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_WIDE";   // 13.3" × 7.5"
  pres.author  = "Thirdmark";
  pres.title   = "Thirdmark Pitch Deck";

  // ════════════════════════════════════════════════════════════════════════════
  // SLIDE 1 — Cover
  // ════════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: BG };

    // Subtle radial glow (large transparent circle behind headline area)
    s.addShape(pres.shapes.OVAL, {
      x: 3.5, y: 1.5, w: 6, h: 4,
      fill: { color: ACCENT, transparency: 92 },
      line: { color: BG, width: 0 },
    });

    // Top-left wordmark
    s.addText("THIRDMARK", {
      x: PAD, y: 0.42, w: 3.5, h: 0.32,
      fontSize: 14, bold: true, color: ACCENT,
      fontFace: "Calibri", charSpacing: 5, margin: 0,
    });

    // Center headline
    s.addText("Senior-level output\nacross every business function.", {
      x: PAD, y: 2.1, w: CONTENT_W, h: 1.8,
      fontSize: 40, bold: true, color: WHITE,
      fontFace: "Calibri", align: "center", valign: "middle",
      margin: 0,
    });

    // Sub-headline
    s.addText("In days, not months.", {
      x: PAD, y: 4.05, w: CONTENT_W, h: 0.65,
      fontSize: 28, bold: true, color: ACCENT,
      fontFace: "Calibri", align: "center", valign: "middle",
      margin: 0,
    });

    // Bottom-left: URL
    s.addText("thirdmark.co", {
      x: PAD, y: SH - 0.45, w: 3, h: 0.25,
      fontSize: 12, color: ACCENT, fontFace: "Calibri", margin: 0,
    });

    // Bottom-right: decorative accent line
    s.addShape(pres.shapes.RECTANGLE, {
      x: SW - PAD - 1.5, y: SH - 0.55, w: 1.5, h: 0.04,
      fill: { color: ACCENT }, line: { color: ACCENT, width: 0 },
    });
  }

  // ════════════════════════════════════════════════════════════════════════════
  // SLIDE 2 — The Problem
  // ════════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: BG };

    sectionLabel(s, "THE PROBLEM");
    headline(s, "Agencies are too slow. Freelancers are too narrow.\nHiring is too expensive.", 0.72, 22, CONTENT_W);

    // Three content boxes
    const boxW  = 3.6;
    const boxH  = 2.9;
    const boxY  = 1.9;
    const gap   = (CONTENT_W - boxW * 3) / 2;
    const boxes = [
      {
        title: "Traditional Agency",
        bullets: [
          "6 weeks before work starts",
          "Junior team on your account",
          "Large retainer before anything ships",
        ],
      },
      {
        title: "Freelancers",
        bullets: [
          "One function at a time",
          "You become the project manager",
          "Inconsistent quality",
        ],
      },
      {
        title: "In-House Team",
        bullets: [
          "$22K to $43K USD/month",
          "Before benefits or tools",
          "Months to hire and onboard",
        ],
      },
    ];

    boxes.forEach((box, i) => {
      const bx = PAD + i * (boxW + gap);

      // Card background
      s.addShape(pres.shapes.RECTANGLE, {
        x: bx, y: boxY, w: boxW, h: boxH,
        fill: { color: SURFACE },
        line: { color: BORDER, width: 1 },
      });

      // Card title
      s.addText(box.title, {
        x: bx + 0.22, y: boxY + 0.22, w: boxW - 0.44, h: 0.3,
        fontSize: 11, bold: true, color: MUTED,
        fontFace: "Calibri", margin: 0,
      });

      // Divider under title
      s.addShape(pres.shapes.RECTANGLE, {
        x: bx + 0.22, y: boxY + 0.58, w: boxW - 0.44, h: 0.02,
        fill: { color: BORDER }, line: { color: BORDER, width: 0 },
      });

      // Bullet list
      const bulletItems = box.bullets.map((b, idx) => ({
        text: b,
        options: {
          bullet: { code: "2013" },
          color: WHITE,
          fontSize: 10,
          fontFace: "Calibri",
          breakLine: idx < box.bullets.length - 1,
          paraSpaceAfter: 6,
        },
      }));

      s.addText(bulletItems, {
        x: bx + 0.22, y: boxY + 0.72, w: boxW - 0.44, h: boxH - 0.95,
        valign: "top", margin: 0,
      });
    });

    // Bottom separator line
    const sepY = boxY + boxH + 0.22;
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD, y: sepY, w: CONTENT_W, h: 0.02,
      fill: { color: BORDER }, line: { color: BORDER, width: 0 },
    });

    // Bottom callout
    s.addText("The gap: early-stage companies need broad, senior output now.", {
      x: PAD, y: sepY + 0.1, w: CONTENT_W, h: 0.28,
      fontSize: 11, color: MUTED, fontFace: "Calibri",
      italic: true, margin: 0,
    });
  }

  // ════════════════════════════════════════════════════════════════════════════
  // SLIDE 3 — The Solution
  // ════════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: BG };

    sectionLabel(s, "THE SOLUTION");
    headline(s, "One account lead. 184 specialized AI agents.\nEvery function covered.", 0.72, 22, CONTENT_W);

    // Diagram center
    const diagCenterX = SW / 2;
    const topBoxW = 3.2;
    const topBoxH = 0.8;
    const topBoxX = diagCenterX - topBoxW / 2;
    const topBoxY = 1.95;

    // Top box — JUSTIN
    s.addShape(pres.shapes.RECTANGLE, {
      x: topBoxX, y: topBoxY, w: topBoxW, h: topBoxH,
      fill: { color: ACCENT, transparency: 80 },
      line: { color: ACCENT, width: 1.5 },
    });
    s.addText("THIRDMARK", {
      x: topBoxX, y: topBoxY + 0.04, w: topBoxW, h: 0.34,
      fontSize: 14, bold: true, color: WHITE,
      fontFace: "Calibri", align: "center", margin: 0,
    });
    s.addText("Account lead: intake, strategy, QA, delivery", {
      x: topBoxX, y: topBoxY + 0.41, w: topBoxW, h: 0.28,
      fontSize: 9.5, color: MUTED,
      fontFace: "Calibri", align: "center", margin: 0,
    });

    // Arrow down
    const arrowX = diagCenterX - 0.02;
    const arrowTopY = topBoxY + topBoxH;
    s.addShape(pres.shapes.RECTANGLE, {
      x: arrowX, y: arrowTopY, w: 0.04, h: 0.38,
      fill: { color: ACCENT }, line: { color: ACCENT, width: 0 },
    });
    // Arrow head (triangle via thin overlapping rects approximation)
    s.addShape(pres.shapes.RECTANGLE, {
      x: arrowX - 0.1, y: arrowTopY + 0.32, w: 0.24, h: 0.04,
      fill: { color: ACCENT }, line: { color: ACCENT, width: 0 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: arrowX - 0.07, y: arrowTopY + 0.36, w: 0.18, h: 0.04,
      fill: { color: ACCENT }, line: { color: ACCENT, width: 0 },
    });

    // Bottom box — AGENTS
    const botBoxW = 5.6;
    const botBoxH = 0.92;
    const botBoxX = diagCenterX - botBoxW / 2;
    const botBoxY = arrowTopY + 0.44;

    s.addShape(pres.shapes.RECTANGLE, {
      x: botBoxX, y: botBoxY, w: botBoxW, h: botBoxH,
      fill: { color: SURFACE },
      line: { color: ACCENT, width: 1.5 },
    });
    s.addText("184 SPECIALIZED AI AGENTS", {
      x: botBoxX, y: botBoxY + 0.06, w: botBoxW, h: 0.36,
      fontSize: 14, bold: true, color: WHITE,
      fontFace: "Calibri", align: "center", charSpacing: 1, margin: 0,
    });
    s.addText("Strategy  ·  Content  ·  Product  ·  Engineering  ·  Marketing  ·  Operations", {
      x: botBoxX, y: botBoxY + 0.48, w: botBoxW, h: 0.28,
      fontSize: 9.5, color: MUTED,
      fontFace: "Calibri", align: "center", margin: 0,
    });

    // Stat pills row
    const pills = [
      "3 to 7 day delivery",
      "14 delivery divisions",
      "184 agents",
      "Single point of contact",
    ];
    const pillY   = botBoxY + botBoxH + 0.34;
    const pillH   = 0.38;
    const totalPillW = CONTENT_W;
    const pillGap = 0.18;
    const pillW   = (totalPillW - pillGap * (pills.length - 1)) / pills.length;

    pills.forEach((pill, i) => {
      const px = PAD + i * (pillW + pillGap);
      s.addShape(pres.shapes.RECTANGLE, {
        x: px, y: pillY, w: pillW, h: pillH,
        fill: { color: SURFACE },
        line: { color: ACCENT, width: 1 },
      });
      s.addText(pill, {
        x: px, y: pillY, w: pillW, h: pillH,
        fontSize: 10, color: WHITE, bold: true,
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
      });
    });
  }

  // ════════════════════════════════════════════════════════════════════════════
  // SLIDE 4 — What We Build / Capabilities
  // ════════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: BG };

    sectionLabel(s, "CAPABILITIES");
    headline(s, "Six delivery areas. One point of contact.", 0.72, 24, CONTENT_W);

    const tiles = [
      { title: "Marketing",   body: "Messaging, content strategy, SEO foundations, campaign briefs" },
      { title: "Product",     body: "Roadmap support, feature prioritization, launch planning" },
      { title: "Sales",       body: "Outbound sequences, pitch decks, proposal strategy" },
      { title: "Engineering", body: "Technical writing, architecture docs, code review, spec writing" },
      { title: "Operations",  body: "Process documentation, workflow design, reporting structures" },
      { title: "Strategy",    body: "Competitive analysis, pricing models, financial narrative" },
    ];

    const cols   = 2;
    const rows   = 3;
    const tileH  = 1.38;
    const tileW  = (CONTENT_W - 0.25) / cols;
    const startY = 1.82;
    const gapY   = 0.2;
    const accentBarW = 0.055;

    tiles.forEach((tile, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const tx  = PAD + col * (tileW + 0.25);
      const ty  = startY + row * (tileH + gapY);

      // Card background
      s.addShape(pres.shapes.RECTANGLE, {
        x: tx, y: ty, w: tileW, h: tileH,
        fill: { color: SURFACE },
        line: { color: BORDER, width: 1 },
      });

      // Left accent bar
      s.addShape(pres.shapes.RECTANGLE, {
        x: tx, y: ty, w: accentBarW, h: tileH,
        fill: { color: ACCENT }, line: { color: ACCENT, width: 0 },
      });

      // Title
      s.addText(tile.title, {
        x: tx + accentBarW + 0.18, y: ty + 0.18, w: tileW - accentBarW - 0.26, h: 0.3,
        fontSize: 13, bold: true, color: WHITE,
        fontFace: "Calibri", margin: 0,
      });

      // Body
      s.addText(tile.body, {
        x: tx + accentBarW + 0.18, y: ty + 0.54, w: tileW - accentBarW - 0.26, h: 0.72,
        fontSize: 10, color: MUTED,
        fontFace: "Calibri", wrap: true, margin: 0,
      });
    });
  }

  // ════════════════════════════════════════════════════════════════════════════
  // SLIDE 5 — Service Packages
  // ════════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: BG };

    sectionLabel(s, "SERVICES");
    headline(s, "Three ways to work together.", 0.72, 24, CONTENT_W);

    const cards = [
      {
        tier:    "AI Sprint",
        price:   "$1,500 to $5,000 USD",
        sub:     "One-time engagement",
        bullets: [
          "Single deliverable",
          "3 to 7 business days",
          "No ongoing commitment",
        ],
        footer:   "Best for: a specific output you need now.",
        popular: false,
        borderW:  1,
      },
      {
        tier:    "AI Retainer",
        price:   "$5,000 to $15,000 USD/mo",
        sub:     "Ongoing",
        bullets: [
          "Monthly output, clear deliverables",
          "One or more functions",
          "Flexible scope each cycle",
        ],
        footer:   "Best for: consistent delivery without a team.",
        popular: true,
        borderW:  2,
      },
      {
        tier:    "AI Studio",
        price:   "$15,000 to $35,000 USD/mo",
        sub:     "Embedded",
        bullets: [
          "Full-stack across all functions",
          "Operates as your external team",
          "Multiple workstreams in parallel",
        ],
        footer:   "Best for: running hard across all functions.",
        popular: false,
        borderW:  1,
      },
    ];

    const cardH    = 4.1;
    const cardW    = (CONTENT_W - 0.3) / 3;
    const cardGap  = 0.15;
    const cardY    = 1.52;

    cards.forEach((card, i) => {
      const cx = PAD + i * (cardW + cardGap);

      // Card background
      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: cardY, w: cardW, h: cardH,
        fill: { color: SURFACE },
        line: { color: card.popular ? ACCENT : BORDER, width: card.borderW },
      });

      // "MOST POPULAR" pill
      if (card.popular) {
        const pillW2 = 1.2;
        const pillH2 = 0.24;
        s.addShape(pres.shapes.RECTANGLE, {
          x: cx + cardW - pillW2 - 0.14, y: cardY + 0.14, w: pillW2, h: pillH2,
          fill: { color: ACCENT }, line: { color: ACCENT, width: 0 },
        });
        s.addText("MOST POPULAR", {
          x: cx + cardW - pillW2 - 0.14, y: cardY + 0.14, w: pillW2, h: pillH2,
          fontSize: 7.5, bold: true, color: WHITE,
          fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
        });
      }

      // Tier label
      s.addText(card.tier, {
        x: cx + 0.22, y: cardY + 0.18, w: cardW - 0.44, h: 0.3,
        fontSize: 14, bold: true, color: ACCENT,
        fontFace: "Calibri", margin: 0,
      });

      // Price
      s.addText(card.price, {
        x: cx + 0.22, y: cardY + 0.54, w: cardW - 0.44, h: 0.48,
        fontSize: 18, bold: true, color: WHITE,
        fontFace: "Calibri", margin: 0, wrap: true,
      });

      // Sub
      s.addText(card.sub, {
        x: cx + 0.22, y: cardY + 1.08, w: cardW - 0.44, h: 0.26,
        fontSize: 11, color: MUTED,
        fontFace: "Calibri", margin: 0,
      });

      // Divider
      s.addShape(pres.shapes.RECTANGLE, {
        x: cx + 0.22, y: cardY + 1.4, w: cardW - 0.44, h: 0.02,
        fill: { color: BORDER }, line: { color: BORDER, width: 0 },
      });

      // Bullets
      const bulletItems = card.bullets.map((b, idx) => ({
        text: b,
        options: {
          bullet: { code: "2013" },
          color: WHITE,
          fontSize: 10,
          fontFace: "Calibri",
          breakLine: idx < card.bullets.length - 1,
          paraSpaceAfter: 5,
        },
      }));
      s.addText(bulletItems, {
        x: cx + 0.22, y: cardY + 1.5, w: cardW - 0.44, h: 1.5,
        valign: "top", margin: 0,
      });

      // Footer italic
      s.addText(card.footer, {
        x: cx + 0.22, y: cardY + cardH - 0.44, w: cardW - 0.44, h: 0.32,
        fontSize: 9.5, color: MUTED, italic: true,
        fontFace: "Calibri", margin: 0, wrap: true,
      });
    });

    // Bottom comparison callout
    const callY = cardY + cardH + 0.2;
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD, y: callY, w: CONTENT_W, h: 0.5,
      fill: { color: SURFACE },
      line: { color: ACCENT, width: 1 },
    });
    s.addText([
      { text: "Market rate for a 3-specialist team: $22,000 to $43,000 USD/month. Thirdmark Retainer: $5,000 to $15,000 USD/month.  ", options: { color: WHITE, fontSize: 10.5, fontFace: "Calibri" } },
      { text: "50 to 65% cheaper.", options: { color: ACCENT, fontSize: 10.5, bold: true, fontFace: "Calibri" } },
    ], {
      x: PAD + 0.2, y: callY, w: CONTENT_W - 0.4, h: 0.5,
      valign: "middle", margin: 0,
    });
  }

  // ════════════════════════════════════════════════════════════════════════════
  // SLIDE 6 — Proof
  // ════════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: BG };

    sectionLabel(s, "RECENT WORK");
    headline(s, "CloudHarbor needed a full marketing\nfoundation. Delivered in 48 hours.", 0.72, 22, 8.5);

    // Client tag
    s.addText("Cloud cost management  ·  Seed stage", {
      x: PAD, y: 1.82, w: 7, h: 0.26,
      fontSize: 11, color: MUTED, fontFace: "Calibri", margin: 0,
    });

    // Divider
    s.addShape(pres.shapes.RECTANGLE, {
      x: PAD, y: 2.16, w: CONTENT_W, h: 0.02,
      fill: { color: BORDER }, line: { color: BORDER, width: 0 },
    });

    // Context paragraph
    s.addText("No marketing team. No existing infrastructure. Needed to launch before their next funding conversation.", {
      x: PAD, y: 2.26, w: 7.4, h: 0.5,
      fontSize: 11, color: MUTED, fontFace: "Calibri", wrap: true, margin: 0,
    });

    // Deliverables list
    const deliverables = [
      "18-term keyword strategy across four intent levels with 90-day content roadmap",
      "Homepage messaging framework for dual audience: technical and finance buyers",
      "30-day content calendar: 12 blog posts, 12 LinkedIn posts, 4 email newsletters, 2 explainers",
      "6 LinkedIn thought leadership angles with full production briefs",
    ];

    const dlItems = deliverables.map((d, idx) => ({
      text: d,
      options: {
        bullet: { code: "2713" },
        color: WHITE,
        fontSize: 10.5,
        fontFace: "Calibri",
        breakLine: idx < deliverables.length - 1,
        paraSpaceAfter: 8,
      },
    }));

    s.addText(dlItems, {
      x: PAD, y: 2.84, w: 7.4, h: 3.2,
      valign: "top", margin: 0,
    });

    // Right stat box
    const statX = SW - PAD - 3.4;
    const statY = 2.0;
    const statW = 3.4;
    const statH = 4.5;

    s.addShape(pres.shapes.RECTANGLE, {
      x: statX, y: statY, w: statW, h: statH,
      fill: { color: SURFACE },
      line: { color: ACCENT, width: 1.5 },
    });

    // Stat 1: 48hrs
    s.addText("48hrs", {
      x: statX, y: statY + 0.4, w: statW, h: 0.9,
      fontSize: 52, bold: true, color: ACCENT,
      fontFace: "Calibri", align: "center", margin: 0,
    });
    s.addText("Delivery time", {
      x: statX, y: statY + 1.36, w: statW, h: 0.28,
      fontSize: 12, color: MUTED, fontFace: "Calibri", align: "center", margin: 0,
    });

    // Divider
    s.addShape(pres.shapes.RECTANGLE, {
      x: statX + 0.3, y: statY + 1.76, w: statW - 0.6, h: 0.02,
      fill: { color: BORDER }, line: { color: BORDER, width: 0 },
    });

    // Stat 2: 30+
    s.addText("30+", {
      x: statX, y: statY + 1.96, w: statW, h: 0.72,
      fontSize: 36, bold: true, color: WHITE,
      fontFace: "Calibri", align: "center", margin: 0,
    });
    s.addText("Individual pieces\nplanned and briefed", {
      x: statX, y: statY + 2.76, w: statW, h: 0.52,
      fontSize: 10, color: MUTED, fontFace: "Calibri", align: "center", margin: 0,
    });
  }

  // ════════════════════════════════════════════════════════════════════════════
  // SLIDE 7 — CTA
  // ════════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: BG_LIGHT };

    // Top decorative gradient bar (simulated with two rects blending accent to accent2)
    const barW = 3.0;
    const barX = (SW - barW) / 2;
    s.addShape(pres.shapes.RECTANGLE, {
      x: barX, y: 0.38, w: barW / 2, h: 0.06,
      fill: { color: ACCENT }, line: { color: ACCENT, width: 0 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: barX + barW / 2, y: 0.38, w: barW / 2, h: 0.06,
      fill: { color: ACCENT2 }, line: { color: ACCENT2, width: 0 },
    });

    // Section label
    s.addText("READY TO START?", {
      x: 0, y: 0.82, w: SW, h: 0.28,
      fontSize: 10, bold: true, color: ACCENT,
      fontFace: "Calibri", align: "center", charSpacing: 2, margin: 0,
    });

    // Headline
    s.addText("Tell us what you need.", {
      x: 0, y: 1.2, w: SW, h: 0.9,
      fontSize: 40, bold: true, color: WHITE,
      fontFace: "Calibri", align: "center", margin: 0,
    });

    // Subhead
    s.addText("We respond within 24 hours with a scope and a price.", {
      x: 0, y: 2.18, w: SW, h: 0.42,
      fontSize: 16, color: MUTED,
      fontFace: "Calibri", align: "center", margin: 0,
    });

    // Two option boxes
    const optW  = 4.6;
    const optH  = 1.9;
    const optY  = 3.0;
    const gap   = 0.4;
    const optX1 = (SW - optW * 2 - gap) / 2;
    const optX2 = optX1 + optW + gap;

    const opts = [
      {
        title: "Send a Brief",
        body:  "Describe the project or goal. We'll respond with scope and pricing within 24 hours.",
      },
      {
        title: "Send a Message",
        body:  "Not sure where to start? Message us and we will scope it together and match you to the right tier.",
      },
    ];

    [optX1, optX2].forEach((ox, i) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: ox, y: optY, w: optW, h: optH,
        fill: { color: SURFACE },
        line: { color: ACCENT, width: 1.5 },
      });
      s.addText(opts[i].title, {
        x: ox + 0.24, y: optY + 0.22, w: optW - 0.48, h: 0.34,
        fontSize: 14, bold: true, color: WHITE,
        fontFace: "Calibri", margin: 0,
      });
      s.addText(opts[i].body, {
        x: ox + 0.24, y: optY + 0.62, w: optW - 0.48, h: 1.1,
        fontSize: 11, color: MUTED,
        fontFace: "Calibri", wrap: true, margin: 0,
      });
    });

    // Bottom contact line
    s.addText("hello@thirdmark.co  ·  thirdmark.co", {
      x: 0, y: SH - 0.52, w: SW, h: 0.3,
      fontSize: 12, color: MUTED, fontFace: "Calibri",
      align: "center", margin: 0,
    });
  }

  // ─── Save ────────────────────────────────────────────────────────────────
  const outPath = "C:/Users/3Jay/agents/deliverables/agency-marketing/thirdmark-pitch-deck.pptx";
  await pres.writeFile({ fileName: outPath });

  const stat = fs.statSync(outPath);
  const kb   = (stat.size / 1024).toFixed(1);
  console.log(`Saved: ${outPath}`);
  console.log(`File size: ${kb} KB (${stat.size} bytes)`);
}

buildDeck().catch(err => {
  console.error("Build failed:", err);
  process.exit(1);
});
