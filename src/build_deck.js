// Builds "Beneful_Industry_Research.pptx" — Unit 2 Individual Industry Research Presentation.
// Run: node src/build_deck.js  (needs pptxgenjs, react, react-dom, react-icons, sharp, jszip)
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const JSZip = require("jszip");
const fs = require("fs");
const fa = require("react-icons/fa");

const OUT = path.join(__dirname, "..", "Beneful_Industry_Research.pptx");
const APPLY_THEME = process.env.APPLY_THEME_JS;

// ---------- Theme ----------
const THEME = {
  name: "Beneful Red",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "2A1A1C", // near-black (body text)
    lt1: "FFFFFF",
    dk2: "5E4A4D", // warm gray (secondary text)
    lt2: "FCEBED", // pale red tint (cards)
    accent1: "C8102E", // Beneful red (primary)
    accent2: "F9D3D8", // light pink (text on red)
    accent3: "E04A5F", // coral red
    accent4: "8E0C20", // deep red
    accent5: "F4B6BF", // soft pink
    accent6: "5C0A16", // maroon
    hlink: "C8102E",
    folHlink: "8E0C20",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333" x 7.5"
pres.author = "Greydan Bruckner";
pres.title = "Beneful: The Purina Beneful Product Line";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333;

// ---------- Icons ----------
async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(fa[name], { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Icon centered in a filled circle (the deck's recurring motif).
async function iconCircle(slide, name, x, y, d, fill, iconColor = HEX.lt1, objectName = "Icon") {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: objectName + " Circle" });
  const pad = d * 0.25;
  slide.addImage({ data: await icon(name, iconColor), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: objectName + " Glyph" });
}

const shadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.12 });

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.accent1 },
  objects: [
    { placeholder: { options: { name: "headshot", type: "pic", x: 8.55, y: 1.45, w: 4.1, h: 4.1 }, text: "Click the icon to add your headshot" } },
  ],
});

const contentObjects = (dark) => [
  { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.35, w: 9, h: 0.35, fontSize: 13, bold: true, color: dark ? C.accent2 : C.accent1, charSpacing: 3, margin: 0, valign: "middle" }, text: "" } },
  { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.72, w: 12.1, h: 0.9, fontSize: 36, bold: true, color: dark ? C.background1 : C.text1, fontFace: THEME.headFontFace, align: "left", margin: 0, valign: "middle" }, text: "" } },
  { placeholder: { options: { name: "cite", type: "body", x: 6.2, y: 6.92, w: 6.53, h: 0.35, fontSize: 11, italic: true, color: dark ? C.accent2 : C.text2, align: "right", margin: 0, valign: "middle" }, text: "" } },
];

pres.defineSlideMaster({
  title: "CONTENT_LIGHT",
  background: { color: C.background1 },
  objects: contentObjects(false),
  slideNumber: { x: 0.6, y: 6.92, w: 0.6, h: 0.35, fontSize: 11, color: C.text2, fontFace: THEME.bodyFontFace },
});

pres.defineSlideMaster({
  title: "CONTENT_DARK",
  background: { color: C.accent1 },
  objects: contentObjects(true),
  slideNumber: { x: 0.6, y: 6.92, w: 0.6, h: 0.35, fontSize: 11, color: C.accent2, fontFace: THEME.bodyFontFace },
});

pres.defineSlideMaster({
  title: "HOOK",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "cite", type: "body", x: 6.2, y: 6.92, w: 6.53, h: 0.35, fontSize: 11, italic: true, color: C.text2, align: "right", margin: 0, valign: "middle" }, text: "" } },
  ],
});

const tb = (extra) => Object.assign({ isTextBox: true, margin: 0 }, extra);

function header(slide, kicker, title, cite) {
  slide.addText(kicker, { placeholder: "kicker" });
  slide.addText(title, { placeholder: "title" });
  if (cite) slide.addText(cite, { placeholder: "cite" });
}

// Transition per slide number (1-based); applied in post-processing.
const transitions = {};

async function build() {
  // ===== 1. Title =====
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
  transitions[1] = "fade";
  s.addText("BUS 102  ·  FYIC INDUSTRY RESEARCH", tb({ x: 0.75, y: 1.3, w: 7.2, h: 0.4, fontSize: 14, bold: true, color: C.accent2, charSpacing: 3, objectName: "Course Label" }));
  s.addText("Beneful", tb({ x: 0.75, y: 1.8, w: 7.4, h: 1.3, fontSize: 72, bold: true, fontFace: THEME.headFontFace, color: C.background1, objectName: "Deck Title" }));
  s.addText("The Purina Beneful Product Line: History, Products & Marketing Strategy", tb({ x: 0.75, y: 3.15, w: 7.2, h: 1.0, fontSize: 22, color: C.accent2, objectName: "Subtitle" }));
  s.addText([
    { text: "Greydan Bruckner", options: { fontSize: 22, bold: true, color: C.background1, breakLine: true } },
    { text: "October 2026", options: { fontSize: 16, color: C.accent2 } },
  ], tb({ x: 0.75, y: 4.75, w: 6, h: 0.9, paraSpaceAfter: 4, objectName: "Presenter" }));
  // Headshot frame: ring behind the circular picture placeholder.
  s.addShape(pres.shapes.OVAL, { x: 8.35, y: 1.25, w: 4.5, h: 4.5, fill: { color: C.accent4 }, line: { color: C.background1, width: 5 }, objectName: "Headshot Ring" });
  s.addShape(pres.shapes.OVAL, { x: 12.75, y: 5.55, w: 0.4, h: 0.4, fill: { color: C.background1 }, line: { type: "none" }, objectName: "Accent Dot White" });
  s.addShape(pres.shapes.OVAL, { x: 8.15, y: 1.35, w: 0.35, h: 0.35, fill: { color: C.accent5 }, line: { type: "none" }, objectName: "Accent Dot Pink" });
  s.addImage({ data: await icon("FaPaw", HEX.lt1), x: 9.9, y: 6.1, w: 0.55, h: 0.55, transparency: 30, objectName: "Paw" });
  s.addText("Client Partner: Purina US", tb({ x: 10.55, y: 6.1, w: 2.4, h: 0.55, fontSize: 13, color: C.accent2, valign: "middle", objectName: "Client" }));
  s.addNotes(
    "Good morning, everyone! My name is Greydan Bruckner, and today I'll be walking you through Beneful — one of Purina's best-known dog food brands. " +
    "My purpose is to give you a quick but solid foundation on Beneful's history, its products, its top sellers, and how it markets to different groups of dog owners."
  );

  // ===== 2. Hook =====
  s = pres.addSlide({ masterName: "HOOK", sectionTitle: "Introduction" });
  transitions[2] = "fade";
  s.addText("7,300", tb({ x: 0.75, y: 1.9, w: 6.6, h: 2.1, fontSize: 140, bold: true, fontFace: THEME.headFontFace, color: C.accent1, objectName: "Hook Number" }));
  s.addText("meals the average dog eats over a 10-year life*", tb({ x: 0.8, y: 4.1, w: 6.2, h: 0.9, fontSize: 24, color: C.text1, objectName: "Hook Label" }));
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.75, y: 1.9, w: 4.95, h: 3.6, rectRadius: 0.2, fill: { color: C.accent1 }, line: { type: "none" }, shadow: shadow(), objectName: "Question Card" });
  await iconCircle(s, "FaDog", 8.15, 2.25, 1.0, HEX.lt1, HEX.accent1, "Dog");
  s.addText("So why should every bowl taste the same?", tb({ x: 8.15, y: 3.45, w: 4.2, h: 1.6, fontSize: 26, italic: true, bold: true, fontFace: THEME.headFontFace, color: C.background1, valign: "top", objectName: "Hook Question" }));
  s.addText("*2 meals a day × 365 days × 10 years", tb({ x: 0.8, y: 6.92, w: 5, h: 0.35, fontSize: 11, italic: true, color: C.text2, objectName: "Footnote" }));
  s.addText("Source: (Purina, n.d.)", { placeholder: "cite" });
  s.addNotes(
    "Quick question: what if you had to eat the same dinner every night for the next 20 years? " +
    "That's roughly what a dog faces — about 7,300 meals over a 10-year life. Beneful was built on the idea that every one of those meals should be something a dog looks forward to."
  );

  // ===== 3. Roadmap =====
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Introduction" });
  transitions[3] = "push";
  header(s, "ROADMAP", "Today’s Roadmap", "Source: (Purina, n.d.)");
  const agenda = [
    ["FaHistory", "01", "History", "From a 2001 launch to a household name"],
    ["FaBone", "02", "Products", "Dry food, wet food & treats for every dog"],
    ["FaTrophy", "03", "Top Sellers", "The flagship products driving the brand"],
    ["FaBullhorn", "04", "Marketing", "How Beneful reaches each consumer group"],
  ];
  const aw = 2.75, ag = 0.367;
  for (let i = 0; i < agenda.length; i++) {
    const [ic, num, t, d] = agenda[i];
    const x = 0.6 + i * (aw + ag);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.0, w: aw, h: 4.2, rectRadius: 0.15, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Agenda Card ${num}` });
    await iconCircle(s, ic, x + 0.35, 2.35, 1.0, i % 2 ? HEX.accent4 : HEX.accent1, HEX.lt1, `Agenda ${num}`);
    s.addText(num, tb({ x: x + aw - 1.15, y: 2.35, w: 0.8, h: 0.6, fontSize: 28, bold: true, fontFace: THEME.headFontFace, color: C.accent1, align: "right", objectName: `Agenda Num ${num}` }));
    s.addText(t, tb({ x: x + 0.35, y: 3.65, w: aw - 0.7, h: 0.55, fontSize: 24, bold: true, fontFace: THEME.headFontFace, color: C.text1, objectName: `Agenda Title ${num}` }));
    s.addText(d, tb({ x: x + 0.35, y: 4.3, w: aw - 0.7, h: 1.4, fontSize: 16, color: C.text2, valign: "top", objectName: `Agenda Desc ${num}` }));
  }
  s.addNotes(
    "Here's our roadmap. First, Beneful's history. Second, the product line. Third, the top-selling flagship products. " +
    "And finally, the marketing strategies Beneful uses to reach different groups of dog owners. Let's start at the beginning."
  );

  // ===== 4. History =====
  pres.addSection({ title: "History" });
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "History" });
  transitions[4] = "push";
  header(s, "01  ·  HISTORY", "A Brand Built on Taste, Nutrition & Variety", "Source: (Purina, n.d.; “Purina Launches,” 2015)");
  const milestones = [
    ["FaFlag", "2001", "Launch", "Nestlé Purina PetCare introduces Beneful, centered on taste, nutrition & variety"],
    ["FaDog", "Growth", "Tailored recipes", "Formulas adapt to dog size & needs — like IncrediBites and Healthy Weight"],
    ["FaUtensils", "Expansion", "Beyond the bag", "Wet meals, chopped blends & treats join the dry line"],
    ["FaShieldAlt", "2015", "Brand trust", "“I Stand Behind Beneful” features the employees who make it"],
  ];
  const lineY = 2.95;
  s.addShape(pres.shapes.LINE, { x: 1.4, y: lineY, w: 10.5, h: 0, line: { color: C.accent5, width: 3 }, objectName: "Timeline" });
  const mw = 2.75, mg = 0.367;
  for (let i = 0; i < milestones.length; i++) {
    const [ic, when, t, d] = milestones[i];
    const x = 0.6 + i * (mw + mg);
    const cx = x + mw / 2;
    s.addText(when, tb({ x, y: 1.85, w: mw, h: 0.45, fontSize: 20, bold: true, fontFace: THEME.headFontFace, color: C.accent1, align: "center", objectName: `Milestone When ${i + 1}` }));
    await iconCircle(s, ic, cx - 0.5, lineY - 0.5, 1.0, i === 3 ? HEX.accent4 : HEX.accent1, HEX.lt1, `Milestone ${i + 1}`);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 3.85, w: mw, h: 2.55, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Milestone Card ${i + 1}` });
    s.addText(t, tb({ x: x + 0.25, y: 4.05, w: mw - 0.5, h: 0.45, fontSize: 18, bold: true, color: C.text1, objectName: `Milestone Title ${i + 1}` }));
    s.addText(d, tb({ x: x + 0.25, y: 4.55, w: mw - 0.5, h: 1.7, fontSize: 15, color: C.text2, valign: "top", objectName: `Milestone Desc ${i + 1}` }));
  }
  s.addNotes(
    "Beneful launched in 2001 under Nestlé Purina PetCare, built around three ideas: taste, nutrition, and variety. " +
    "As it grew, it adapted to what owners wanted — recipes for small dogs, for weight control — and expanded beyond dry food into wet meals and treats. " +
    "In 2015, Purina even launched the 'I Stand Behind Beneful' campaign, featuring the employees who make it."
  );

  // ===== 5. Product line =====
  pres.addSection({ title: "Products" });
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Products" });
  transitions[5] = "push";
  header(s, "02  ·  PRODUCTS", "One Brand, Three Ways to Feed", "Source: (Purina, n.d.)");
  const cols = [
    ["FaBone", HEX.accent1, "Dry Food", ["Originals", "IncrediBites (small dogs)", "Healthy Weight", "Grain Free"]],
    ["FaUtensils", HEX.accent4, "Wet Food", ["Prepared Meals stews", "Chopped Blends", "IncrediBites wet"]],
    ["FaCookieBite", HEX.accent3, "Treats", ["Baked Delights", "Healthy Smile dental chews"]],
  ];
  const cw = 3.83, cg = 0.3;
  for (let i = 0; i < cols.length; i++) {
    const [ic, col, t, items] = cols[i];
    const x = 0.6 + i * (cw + cg);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.95, w: cw, h: 3.75, rectRadius: 0.15, fill: { color: C.background1 }, line: { color: C.background2, width: 1 }, shadow: shadow(), objectName: `Product Card ${t}` });
    await iconCircle(s, ic, x + 0.35, 2.25, 0.85, col, HEX.lt1, t);
    s.addText(t, tb({ x: x + 1.4, y: 2.25, w: cw - 1.6, h: 0.85, fontSize: 24, bold: true, fontFace: THEME.headFontFace, color: C.text1, valign: "middle", objectName: `Product Title ${t}` }));
    s.addText(items.map((it, j) => ({ text: it, options: { bullet: { indent: 18 }, breakLine: j < items.length - 1 } })),
      tb({ x: x + 0.4, y: 3.35, w: cw - 0.7, h: 2.15, fontSize: 17, color: C.text1, paraSpaceAfter: 8, valign: "top", objectName: `Product List ${t}` }));
  }
  s.addText("Formulas tailored by:", tb({ x: 0.6, y: 6.0, w: 2.6, h: 0.5, fontSize: 16, bold: true, color: C.text2, valign: "middle", objectName: "Tailored Label" }));
  const chips = ["Dog size", "Life stage", "Weight", "Taste & texture"];
  for (let i = 0; i < chips.length; i++) {
    const x = 3.2 + i * 2.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 6.0, w: 2.15, h: 0.5, rectRadius: 0.25, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Chip ${chips[i]}` });
    s.addText(chips[i], tb({ x, y: 6.0, w: 2.15, h: 0.5, fontSize: 15, bold: true, color: C.accent4, align: "center", valign: "middle", objectName: `Chip Text ${chips[i]}` }));
  }
  s.addNotes(
    "Today, Beneful feeds dogs three ways. Dry food is the core — Originals, IncrediBites for small dogs, Healthy Weight, and Grain Free. " +
    "Wet food includes Prepared Meals stews and Chopped Blends. And treats like Baked Delights and Healthy Smile round it out. " +
    "The common thread is variety: recipes tailored by size, life stage, weight, and taste."
  );

  // ===== 6. Top sellers =====
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Products" });
  transitions[6] = "fade";
  header(s, "03  ·  TOP SELLERS", "The Flagships Driving the Brand", "Source: (Owen, 2026; Purina, n.d.)");
  const sellers = [
    ["FaStar", HEX.accent1, "Beneful Originals", "The core dry recipe with real meat and visible veggies — the brand’s foundation"],
    ["FaPaw", HEX.accent4, "IncrediBites", "Small, crunchy kibble sized for small and toy breeds"],
    ["FaUtensils", HEX.accent3, "Prepared Meals", "Ready-to-serve stews owners use as a full meal or a topper"],
  ];
  for (let i = 0; i < sellers.length; i++) {
    const [ic, col, t, d] = sellers[i];
    const y = 1.95 + i * 1.6;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 7.3, h: 1.35, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Seller Card ${i + 1}` });
    await iconCircle(s, ic, 0.85, y + 0.2, 0.95, col, HEX.lt1, `Seller ${i + 1}`);
    s.addText(t, tb({ x: 2.05, y: y + 0.15, w: 5.6, h: 0.45, fontSize: 20, bold: true, color: C.text1, objectName: `Seller Title ${i + 1}` }));
    s.addText(d, tb({ x: 2.05, y: y + 0.6, w: 5.6, h: 0.65, fontSize: 15, color: C.text2, valign: "top", objectName: `Seller Desc ${i + 1}` }));
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.3, y: 1.95, w: 4.43, h: 4.55, rectRadius: 0.15, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Insight Panel" });
  s.addText("MARKET INSIGHT", tb({ x: 8.7, y: 2.3, w: 3.7, h: 0.35, fontSize: 13, bold: true, color: C.accent2, charSpacing: 3, objectName: "Insight Label" }));
  s.addText("Dry food is still the backbone of U.S. dog food — and more owners are adding wet food as a mealtime upgrade.", tb({ x: 8.7, y: 2.8, w: 3.65, h: 2.6, fontSize: 21, fontFace: THEME.headFontFace, italic: true, color: C.background1, valign: "top", objectName: "Insight Text" }));
  s.addText("Beneful sells both.", tb({ x: 8.7, y: 5.6, w: 3.65, h: 0.5, fontSize: 18, bold: true, color: C.accent2, objectName: "Insight Kicker" }));
  s.addNotes(
    "So what actually sells? Beneful's flagships are Originals — the core dry recipe — IncrediBites for small dogs, and Prepared Meals wet stews. " +
    "That lineup fits the market: according to Mintel, dry food is still the backbone of U.S. dog food, while more owners are adding wet food as an upgrade. Beneful covers both."
  );

  // ===== 7. Consumer groups =====
  pres.addSection({ title: "Marketing" });
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Marketing" });
  transitions[7] = "push";
  header(s, "04  ·  MARKETING", "Matching Products to Consumer Groups", "Source: (Owen, 2026; Purina, n.d.)");
  const groups = [
    ["FaPaw", HEX.accent1, "Small-dog owners", "IncrediBites", "Kibble and portions sized for little mouths"],
    ["FaWeight", HEX.accent4, "Weight-conscious owners", "Healthy Weight", "Fewer calories, same great taste"],
    ["FaLeaf", HEX.accent3, "Ingredient-focused “pet parents”", "Grain Free", "Real meat first, no corn, wheat or soy"],
    ["FaShoppingCart", HEX.accent6, "Value-minded families", "Originals", "Affordable variety at grocery & mass retailers"],
  ];
  const gw = 5.9, gh = 2.05;
  for (let i = 0; i < groups.length; i++) {
    const [ic, col, g, prod, msg] = groups[i];
    const x = 0.6 + (i % 2) * (gw + 0.333);
    const y = 1.95 + Math.floor(i / 2) * (gh + 0.35);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: gw, h: gh, rectRadius: 0.12, fill: { color: C.background1 }, line: { color: C.background2, width: 1 }, shadow: shadow(), objectName: `Group Card ${i + 1}` });
    await iconCircle(s, ic, x + 0.3, y + 0.3, 0.9, col, HEX.lt1, `Group ${i + 1}`);
    s.addText(g, tb({ x: x + 1.45, y: y + 0.3, w: gw - 1.7, h: 0.45, fontSize: 19, bold: true, color: C.text1, objectName: `Group Name ${i + 1}` }));
    s.addText([{ text: "→  ", options: { color: C.accent1, bold: true } }, { text: prod, options: { color: C.accent1, bold: true } }],
      tb({ x: x + 1.45, y: y + 0.8, w: gw - 1.7, h: 0.4, fontSize: 17, objectName: `Group Product ${i + 1}` }));
    s.addText(msg, tb({ x: x + 1.45, y: y + 1.25, w: gw - 1.7, h: 0.55, fontSize: 15, color: C.text2, valign: "top", objectName: `Group Message ${i + 1}` }));
  }
  s.addNotes(
    "That variety lets Beneful speak to different consumer groups. Small-dog owners get IncrediBites. Weight-conscious owners get Healthy Weight. " +
    "Ingredient-focused 'pet parents' — a growing group Mintel tracks — get Grain Free. And value-minded families get Originals at the grocery store. " +
    "So how does Beneful actually reach these people? Four strategies."
  );

  // ===== 8. Emotional marketing =====
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Marketing" });
  transitions[8] = "fade";
  header(s, "STRATEGY 1", "Emotional Marketing: The Dog–Owner Bond", "Source: (Owen, 2026; Schoenherr, 2011)");
  const emo = [
    ["FaHeart", HEX.accent1, "Celebrates the relationship", "Ads center on happy, playful dogs and the people who love them"],
    ["FaSmile", HEX.accent4, "Makes mealtime a moment of joy", "Food is framed as a way owners show love every day"],
    ["FaHome", HEX.accent3, "Speaks to “pets as family”", "Owners who see dogs as family want food that reflects it"],
  ];
  for (let i = 0; i < emo.length; i++) {
    const [ic, col, t, d] = emo[i];
    const y = 2.0 + i * 1.5;
    await iconCircle(s, ic, 0.6, y, 0.95, col, HEX.lt1, `Emotion ${i + 1}`);
    s.addText(t, tb({ x: 1.85, y: y + 0.02, w: 5.9, h: 0.45, fontSize: 20, bold: true, color: C.text1, objectName: `Emotion Title ${i + 1}` }));
    s.addText(d, tb({ x: 1.85, y: y + 0.5, w: 5.9, h: 0.6, fontSize: 16, color: C.text2, valign: "top", objectName: `Emotion Desc ${i + 1}` }));
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.3, y: 1.95, w: 4.43, h: 4.55, rectRadius: 0.15, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Quote Panel" });
  s.addImage({ data: await icon("FaQuoteLeft", HEX.accent5), x: 8.7, y: 2.35, w: 0.7, h: 0.7, objectName: "Quote Mark" });
  s.addText("Dogs in ads grab attention — and the warm feelings they create carry over to the brand.", tb({ x: 8.7, y: 3.25, w: 3.65, h: 2.4, fontSize: 22, italic: true, fontFace: THEME.headFontFace, color: C.background1, valign: "top", objectName: "Quote Text" }));
  s.addText("— Schoenherr, 2011", tb({ x: 8.7, y: 5.75, w: 3.65, h: 0.4, fontSize: 15, bold: true, color: C.background2, objectName: "Quote Attribution" }));
  s.addNotes(
    "Strategy one is emotional marketing. Beneful's ads aren't really about kibble — they're about the bond between a dog and its owner. " +
    "Research from Washington University shows dogs in ads grab attention and pass those warm feelings on to the brand. And since so many owners now see their dogs as family, that message lands."
  );

  // ===== 9. Product-focused + digital =====
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Marketing" });
  transitions[9] = "fade";
  header(s, "STRATEGIES 2 & 3", "Product Proof Meets Digital Reach", "Source: (Purina, n.d.; “Purina Launches,” 2015)");
  const two = [
    ["FaSearch", HEX.accent1, "Product-Focused", ["Real meat as a headline ingredient", "Visible veggies you can see in the bowl", "Variety of flavors, textures & formats"]],
    ["FaMobileAlt", HEX.accent4, "Digital & Social Media", ["Social posts & online video starring real dogs", "Product pages, reviews & e-commerce retailers", "2015: employees & their dogs answer online critics"]],
  ];
  for (let i = 0; i < two.length; i++) {
    const [ic, col, t, items] = two[i];
    const x = 0.6 + i * (5.9 + 0.333);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.95, w: 5.9, h: 2.95, rectRadius: 0.15, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Strategy Card ${t}` });
    await iconCircle(s, ic, x + 0.35, 2.25, 0.85, col, HEX.lt1, t);
    s.addText(t, tb({ x: x + 1.45, y: 2.25, w: 4.2, h: 0.85, fontSize: 24, bold: true, fontFace: THEME.headFontFace, color: C.text1, valign: "middle", objectName: `Strategy Title ${t}` }));
    s.addText(items.map((it, j) => ({ text: it, options: { bullet: { indent: 18 }, breakLine: j < items.length - 1 } })),
      tb({ x: x + 0.4, y: 3.35, w: 5.2, h: 1.4, fontSize: 16, color: C.text1, paraSpaceAfter: 6, valign: "top", objectName: `Strategy List ${t}` }));
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.2, w: 12.13, h: 1.35, rectRadius: 0.12, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Campaign Quote Panel" });
  s.addImage({ data: await icon("FaQuoteLeft", HEX.accent5), x: 0.95, y: 5.5, w: 0.55, h: 0.55, objectName: "Campaign Quote Mark" });
  s.addText([
    { text: "“After working here, there’s no other food that I would feed my pets.”", options: { fontSize: 20, italic: true, fontFace: THEME.headFontFace, color: C.background1, breakLine: true } },
    { text: "Purina employee, “I Stand Behind Beneful” campaign (2015)", options: { fontSize: 14, bold: true, color: C.accent2 } },
  ], tb({ x: 1.75, y: 5.3, w: 10.7, h: 1.15, valign: "middle", objectName: "Campaign Quote" }));
  s.addNotes(
    "Strategies two and three work together. Product-focused messaging shows the proof — real meat, visible veggies, lots of variety. " +
    "Digital and social media spread that message. In 2015, when critics attacked Beneful online, Purina answered with real employees and their own dogs. As one put it: 'there's no other food I would feed my pets.'"
  );

  // ===== 10. Community =====
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Marketing" });
  transitions[10] = "fade";
  header(s, "STRATEGY 4", "Community Engagement: Dream Dog Parks", "Source: (Purina, n.d.)");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.95, w: 5.2, h: 4.55, rectRadius: 0.15, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Stat Panel" });
  await iconCircle(s, "FaTree", 1.0, 2.3, 1.0, HEX.lt1, HEX.accent1, "Park");
  s.addText("$500K", tb({ x: 1.0, y: 3.55, w: 4.5, h: 1.3, fontSize: 80, bold: true, fontFace: THEME.headFontFace, color: C.background1, objectName: "Stat Number" }));
  s.addText("dog park makeovers awarded through the Beneful Dream Dog Park Project", tb({ x: 1.0, y: 4.9, w: 4.4, h: 1.2, fontSize: 18, color: C.background1, valign: "top", objectName: "Stat Label" }));
  const park = [
    ["FaVoteYea", HEX.accent1, "Communities vote", "Public voting picks the parks that win makeovers"],
    ["FaHandsHelping", HEX.accent4, "Local partnerships", "Grants, matching donations & crowdfunding help towns build parks"],
    ["FaHeart", HEX.accent3, "Lasting goodwill", "The brand shows up where dog owners actually spend time"],
  ];
  for (let i = 0; i < park.length; i++) {
    const [ic, col, t, d] = park[i];
    const y = 2.0 + i * 1.5;
    await iconCircle(s, ic, 6.3, y, 0.95, col, HEX.lt1, `Park ${i + 1}`);
    s.addText(t, tb({ x: 7.55, y: y + 0.02, w: 5.2, h: 0.45, fontSize: 20, bold: true, color: C.text1, objectName: `Park Title ${i + 1}` }));
    s.addText(d, tb({ x: 7.55, y: y + 0.5, w: 5.2, h: 0.65, fontSize: 16, color: C.text2, valign: "top", objectName: `Park Desc ${i + 1}` }));
  }
  s.addNotes(
    "Finally, community engagement. Through the Beneful Dream Dog Park Project, communities voted for local dog parks, and winners received makeovers worth up to 500,000 dollars. " +
    "Beneful also used grants and crowdfunding partnerships to help towns build parks — earning goodwill exactly where dog owners spend their time."
  );

  // ===== 11. Conclusion =====
  pres.addSection({ title: "Conclusion" });
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Conclusion" });
  transitions[11] = "push";
  header(s, "WRAP-UP", "Key Takeaways", "Source: (Owen, 2026; Purina, n.d.)");
  const take = [
    ["FaHistory", HEX.lt1, "History", "Launched in 2001, Beneful grew from one dry line into a broad, needs-based portfolio"],
    ["FaBone", HEX.lt1, "Products", "Dry food, wet food & treats — led by Originals, IncrediBites & Prepared Meals"],
    ["FaBullhorn", HEX.lt1, "Marketing", "Emotion, product proof, digital reach & community build loyalty across groups"],
  ];
  for (let i = 0; i < take.length; i++) {
    const [ic, col, t, d] = take[i];
    const y = 2.0 + i * 1.5;
    await iconCircle(s, ic, 0.6, y, 0.95, col, HEX.accent1, `Takeaway ${i + 1}`);
    s.addText(t, tb({ x: 1.85, y: y + 0.02, w: 5.9, h: 0.45, fontSize: 20, bold: true, color: C.accent2, objectName: `Takeaway Title ${i + 1}` }));
    s.addText(d, tb({ x: 1.85, y: y + 0.5, w: 5.9, h: 0.75, fontSize: 16, color: C.background1, valign: "top", objectName: `Takeaway Desc ${i + 1}` }));
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.3, y: 1.95, w: 4.43, h: 4.55, rectRadius: 0.15, fill: { color: C.background1 }, line: { type: "none" }, shadow: shadow(), objectName: "Thanks Panel" });
  await iconCircle(s, "FaPaw", 9.97, 2.4, 1.1, HEX.accent1, HEX.lt1, "Thanks Paw");
  s.addText("Thank You!", tb({ x: 8.5, y: 3.75, w: 4.03, h: 0.9, fontSize: 40, bold: true, fontFace: THEME.headFontFace, color: C.accent1, align: "center", objectName: "Thanks" }));
  s.addText("Questions?", tb({ x: 8.5, y: 4.7, w: 4.03, h: 0.6, fontSize: 22, color: C.text1, align: "center", objectName: "Questions" }));
  s.addText("Greydan Bruckner", tb({ x: 8.5, y: 5.6, w: 4.03, h: 0.45, fontSize: 16, bold: true, color: C.text2, align: "center", objectName: "Thanks Name" }));
  s.addNotes(
    "To wrap up: Beneful started in 2001 and grew into a broad portfolio built around variety. Its dry food, wet food, and treats — led by Originals, IncrediBites, and Prepared Meals — fit dogs of every size and need. " +
    "And it wins owners through emotion, product proof, digital reach, and community. So remember those 7,300 meals — Beneful's goal is to make every one count. Thank you so much for your time!"
  );

  // ===== 12. References =====
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Conclusion" });
  transitions[12] = "fade";
  s.addText("References", { placeholder: "title", align: "center" });
  const I = { italic: true };
  const refs = [
    [{ t: "Owen, J. (2026, September 25). " }, { t: "Pet food – US – 2026", o: I }, { t: ". Mintel. https://clients.mintel.com/report/pet-food-us-2026" }],
    [{ t: "Purina. (n.d.). " }, { t: "Purina: Nutritious dog and cat food for your pet", o: I }, { t: ". Retrieved October 5, 2026, from https://www.purina.com" }],
    [{ t: "Purina launches ‘I Stand Behind Beneful’ campaign. (2015, June 11). " }, { t: "PetfoodIndustry", o: I }, { t: ". https://www.petfoodindustry.com/pet-food-market/article/15461104/purina-launches-i-stand-behind-beneful-campaign" }],
    [{ t: "Schoenherr, N. (2011, October 13). Advertising goes to the dogs. " }, { t: "The Source", o: I }, { t: ". https://source.washu.edu/2011/10/advertising-goes-to-the-dogs/" }],
  ];
  const runs = [];
  refs.forEach((parts, ri) => {
    parts.forEach((p, pi) => {
      const o = Object.assign({}, p.o || {});
      if (pi === parts.length - 1 && ri < refs.length - 1) o.breakLine = true;
      runs.push({ text: p.t, options: o });
    });
  });
  s.addText(runs, tb({ x: 0.9, y: 1.75, w: 11.5, h: 5.0, fontSize: 16, fontFace: "Times New Roman", color: C.text1, lineSpacingMultiple: 2.0, valign: "top", objectName: "References List" }));
  s.addNotes("References slide — APA 7th edition, alphabetical, double-spaced with hanging indents. Includes the Mintel report from the Miami University Libraries databases.");

  await pres.writeFile({ fileName: OUT });
  if (APPLY_THEME) {
    const { applyTheme } = require(APPLY_THEME);
    await applyTheme(OUT, THEME);
  }
  await postProcess(OUT);
  console.log("Wrote", OUT);
}

// ---------- Post-processing: transitions, hanging indents, circular headshot placeholder ----------
async function postProcess(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const slideFiles = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  for (const f of slideFiles) {
    const n = parseInt(f.match(/slide(\d+)\.xml/)[1], 10);
    let xml = await zip.file(f).async("string");
    const kind = transitions[n] || "fade";
    const trans = kind === "push"
      ? '<p:transition spd="med"><p:push dir="u"/></p:transition>'
      : '<p:transition spd="med"><p:fade/></p:transition>';
    xml = xml.replace(/<p:transition[\s\S]*?<\/p:transition>/, "");
    xml = xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + trans);

    if (n === 12) {
      // Hanging indent (0.5") on every reference paragraph.
      xml = xml.replace(/(<p:sp>(?:(?!<\/p:sp>)[\s\S])*?name="References List"[\s\S]*?<\/p:sp>)/, (sp) =>
        sp.replace(/<a:pPr([^>]*?)(\/?)>/g, (m, attrs, selfClose) => {
          attrs = attrs.replace(/\s(marL|indent)="[^"]*"/g, "");
          return `<a:pPr${attrs} marL="457200" indent="-457200"${selfClose}>`;
        })
      );
    }

    if (n === 1) {
      // Add an empty slide-level picture placeholder so PowerPoint shows "click to add picture".
      const rels = await zip.file("ppt/slides/_rels/slide1.xml.rels").async("string");
      const layoutTarget = rels.match(/Target="\.\.\/slideLayouts\/(slideLayout\d+\.xml)"/)[1];
      let layoutXml = await zip.file("ppt/slideLayouts/" + layoutTarget).async("string");
      // pptxgenjs writes the headshot placeholder without a type; mark it as a picture placeholder.
      const idx = layoutXml.match(/<p:ph\s+idx="(\d+)"/)[1];
      layoutXml = layoutXml.replace(/<p:ph\s+idx="(\d+)"/, '<p:ph type="pic" idx="$1"');
      zip.file("ppt/slideLayouts/" + layoutTarget, layoutXml);
      const phTag = `<p:ph type="pic" idx="${idx}"/>`;
      const sp = `<p:sp><p:nvSpPr><p:cNvPr id="900" name="Headshot" descr="Presenter headshot"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phTag}</p:nvPr></p:nvSpPr><p:spPr/></p:sp>`;
      // Insert after the ring (last shapes are drawn on top, so append at end of spTree).
      xml = xml.replace("</p:spTree>", sp + "</p:spTree>");
    }
    zip.file(f, xml);
  }
  // Make the headshot placeholder circular in the layout.
  const layoutFiles = Object.keys(zip.files).filter((f) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f));
  for (const f of layoutFiles) {
    let xml = await zip.file(f).async("string");
    if (!/type="pic"/.test(xml)) continue;
    xml = xml.replace(/(<p:sp>(?:(?!<\/p:sp>)[\s\S])*?type="pic"[\s\S]*?<\/p:sp>)/, (sp) =>
      sp.replace(/<a:prstGeom prst="\w+">/, '<a:prstGeom prst="ellipse">')
    );
    zip.file(f, xml);
  }
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(file, buf);
}

build().catch((e) => { console.error(e); process.exit(1); });
