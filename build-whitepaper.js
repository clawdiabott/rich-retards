const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function p(text, opts = {}) {
  return new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text, ...opts })] });
}
function bullet(text, ref = "bullets") {
  return new Paragraph({ numbering: { reference: ref, level: 0 }, children: [new TextRun(text)] });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "FF6B9D" },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "00D4AA" },
        paragraph: { spacing: { before: 180, after: 180 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "RICH RETARDS ($RETRD) — WHITEPAPER v1.0", bold: true, size: 18, color: "888888" })]
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Not financial advice. NFA. ", size: 18, color: "888888" }),
          new TextRun({ children: [PageNumber.CURRENT], size: 18 })
        ]
      })] })
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text: "RICH RETARDS", bold: true, size: 72, color: "FFD700" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: "$RETRD Whitepaper v1.0", size: 32, italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
        children: [new TextRun({ text: "Solana Summer 2026 | June 25, 2026", size: 24, color: "FF6B9D" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text: "\"We didn't read the whitepaper. We bought the yacht.\"", italics: true, size: 28 })]
      }),

      h1("Abstract"),
      p("Rich Retards ($RETRD) is a community-driven meme token on the Solana blockchain. It has no utility, no roadmap, and no pretense of fundamental value. Its sole function is to serve as a cultural mirror for participants in the Solana meme coin ecosystem who have profited — or aspire to profit — through collective stupidity, good timing, and unbreakable conviction."),

      h1("1. Introduction"),
      p("The Solana meme coin market has generated billions in trading volume through tokens with zero revenue, zero product, and zero employees. Participants in this market share a common experience: they have made life-changing money by making decisions that would be considered irrational in any other financial context."),
      p("Rich Retards names this experience and builds a community around it."),

      h1("2. Problem Statement"),
      h2("The Honesty Gap"),
      p("Every meme coin whitepaper pretends to be something it's not. AI integrations. Gaming partnerships. Ecosystem tokens. Nobody believes it. Everyone apes anyway."),
      h2("The Identity Gap"),
      p("Traders who profit in meme markets have no symbol for their shared experience. They are called degens, apes, or worse. None capture the specific pride of being stupid and rich simultaneously."),
      h2("The Cultural Gap"),
      p("Pop-culture moments (GTA VI, elections, viral videos) drive meme coin volume, but most themed tokens are lazy IP knockoffs with no narrative depth."),

      h1("3. Solution"),
      p("$RETRD solves all three gaps by being exactly what it claims:"),
      bullet("Honest: No fake utility. The token is the joke AND the product."),
      bullet("Identity: Rich Retard is a badge earned through participation."),
      bullet("Cultural: Vice City aesthetic + Solana trench culture = a meme that writes itself."),

      h1("4. Token Specifications"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 6240],
        rows: [
          ["Name", "Rich Retards"],
          ["Symbol", "RETRD"],
          ["Chain", "Solana (SPL Token)"],
          ["Total Supply", "1,000,000,000"],
          ["Decimals", "6"],
          ["Launch Platform", "pump.fun"],
          ["Mint Authority", "Revoked at launch"],
          ["Freeze Authority", "Revoked at launch"],
        ].map(([k, v]) => new TableRow({
          children: [
            new TableCell({ borders, width: { size: 3120, type: WidthType.DXA },
              shading: { fill: "FFF0F5", type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: k, bold: true })] })] }),
            new TableCell({ borders, width: { size: 6240, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun(v)] })] }),
          ]
        }))
      }),

      new Paragraph({ children: [new PageBreak()] }),

      h1("5. Tokenomics"),
      p("Total Supply: 1,000,000,000 RETRD (1 Billion)"),
      bullet("Bonding Curve: 800,000,000 (80%) — pump.fun public trading"),
      bullet("Liquidity Pool: 200,000,000 (20%) — auto-seeded to Raydium on migration"),
      bullet("Team Allocation: 0%"),
      bullet("Presale: 0%"),
      bullet("Tax: 0%"),
      p("Fair Launch Principles:", { bold: true }),
      bullet("100% of tokens enter through the bonding curve"),
      bullet("No insider wallets, no team tokens, no presale"),
      bullet("Creator buys like everyone else"),

      h1("6. Technology"),
      p("$RETRD is a standard SPL token on Solana. No custom programs. No staking contracts. No NFT integrations. No AI agents."),
      p("The technology is Solana itself: sub-second finality, near-zero fees, and a pump.fun bonding curve that turns vibes into liquidity."),

      h1("7. Governance"),
      p("There is no governance. Rich retards don't vote. They ape."),

      h1("8. Roadmap"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 7020],
        rows: [
          ["Q2 2026", "Launch. Meme. Bond."],
          ["Q3 2026", "???"],
          ["Q4 2026", "Profit."],
          ["Q1 2027", "Richer. More retarded."],
        ].map(([q, m]) => new TableRow({
          children: [
            new TableCell({ borders, width: { size: 2340, type: WidthType.DXA },
              shading: { fill: "E0FFF8", type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: q, bold: true })] })] }),
            new TableCell({ borders, width: { size: 7020, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun(m)] })] }),
          ]
        }))
      }),
      p("This roadmap is intentionally useless. If you need a roadmap to ape, you might be a smart poor.", { italics: true }),

      h1("9. The Rich Retard Creed"),
      new Paragraph({ spacing: { before: 200, after: 200 }, indent: { left: 720 },
        children: [new TextRun({ text: "I shall not read.\nI shall not research.\nI shall not FUD my own bags.\nI shall buy the top and call it conviction.\nI shall die on this hill and the hill shall be made of money.", italics: true, size: 26 })]
      }),

      h1("10. Legal Disclaimer"),
      p("$RETRD is a meme token with no intrinsic value. It is not an investment, security, or financial instrument. No promises of profit are made or implied. The term Rich Retard is used satirically within crypto culture and is not intended to target any individual or group."),
      p("This token is not affiliated with Rockstar Games, Take-Two Interactive, or Grand Theft Auto. Do not spend money you cannot afford to vaporize."),

      h1("11. Conclusion"),
      p("The market doesn't need another utility token. It needs a mirror."),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 },
        children: [new TextRun({ text: "Be stupid. Get rich. Repeat.", bold: true, size: 36, color: "FFD700" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/patrick/rich-retards/docs/RICH-RETARDS-WHITEPAPER.docx", buffer);
  console.log("Whitepaper written.");
});