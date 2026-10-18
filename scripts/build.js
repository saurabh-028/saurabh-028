// Single source of truth: resume/resume.json
// Outputs: dist/resume.docx, README.md, dist/linkedin.md, dist/portfolio.json
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink,
  AlignmentType, BorderStyle, TabStopType, LevelFormat,
} = require("docx");

const ROOT = path.join(__dirname, "..");
const data = JSON.parse(fs.readFileSync(path.join(ROOT, "resume", "resume.json"), "utf8"));
const DIST = path.join(ROOT, "dist");
fs.mkdirSync(DIST, { recursive: true });

const plain = (t) => t.replace(/\*\*/g, "");
const bio = data.basics;

// ---------------------------------------------------------------- DOCX
const FONT = "Calibri", NAVY = "1F3864", DARK = "222222", GRAY = "555555";
const RIGHT = 10800;

const runs = (t) => t.split("**").map((x, i) => new TextRun({
  text: x, bold: i % 2 === 1, size: 19, color: i % 2 === 1 ? "000000" : DARK, font: FONT,
}));
const heading = (t) => new Paragraph({
  spacing: { before: 170, after: 70 }, keepNext: true,
  border: { bottom: { color: NAVY, space: 2, style: BorderStyle.SINGLE, size: 8 } },
  children: [new TextRun({ text: t.toUpperCase(), bold: true, size: 22, color: NAVY, font: FONT })],
});
const roleLine = (title, org, date) => new Paragraph({
  spacing: { before: 100, after: 0 }, keepNext: true,
  tabStops: [{ type: TabStopType.RIGHT, position: RIGHT }],
  children: [
    new TextRun({ text: title, bold: true, size: 21, color: DARK, font: FONT }),
    ...(org ? [new TextRun({ text: "  |  " + org, size: 19, color: GRAY, font: FONT })] : []),
    ...(date ? [new TextRun({ text: "\t" + date, bold: true, size: 19, color: GRAY, font: FONT })] : []),
  ],
});
const projLine = (p) => new Paragraph({
  spacing: { before: 100, after: 0 }, keepNext: true,
  tabStops: [{ type: TabStopType.RIGHT, position: RIGHT }],
  children: [
    new TextRun({ text: `${p.name}  |  ${p.subtitle}`, bold: true, size: 21, color: DARK, font: FONT }),
    ...(p.link ? [
      new TextRun({ text: "   ", size: 18, font: FONT }),
      new ExternalHyperlink({ link: p.link.url, children: [new TextRun({ text: p.link.label, size: 17, color: NAVY, font: FONT, underline: {} })] }),
    ] : []),
    ...(p.dates ? [new TextRun({ text: "\t" + p.dates, bold: true, size: 19, color: GRAY, font: FONT })] : []),
  ],
});
const bullet = (t) => new Paragraph({ numbering: { reference: "bl", level: 0 }, spacing: { after: 25 }, children: runs(t) });
const techLine = (t) => new Paragraph({
  spacing: { before: 10, after: 40 }, indent: { left: 300 },
  children: [
    new TextRun({ text: "Tech: ", bold: true, size: 18, color: NAVY, font: FONT }),
    new TextRun({ text: t, size: 18, color: GRAY, font: FONT }),
  ],
});
const skillLine = (l, v) => new Paragraph({
  spacing: { after: 45 },
  children: [
    new TextRun({ text: l + ": ", bold: true, size: 19, color: NAVY, font: FONT }),
    new TextRun({ text: v, size: 19, color: DARK, font: FONT }),
  ],
});
const link = (t, u) => new ExternalHyperlink({ link: u, children: [new TextRun({ text: t, size: 18, color: NAVY, font: FONT, underline: {} })] });
const sep = () => new TextRun({ text: "   |   ", size: 18, color: GRAY, font: FONT });

function buildDocx() {
  const c = [];
  c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [new TextRun({ text: bio.name.toUpperCase(), bold: true, size: 40, color: NAVY, font: FONT })] }));
  c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 10 }, children: [new TextRun({ text: bio.label.replace(" | ", "  |  "), size: 20, color: GRAY, font: FONT })] }));
  c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 10 }, children: [new TextRun({ text: `${bio.location}   |   ${bio.phone}   |   ${bio.email}`, size: 18, color: DARK, font: FONT })] }));
  c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 30 }, children: [
    link("LinkedIn", bio.links.linkedin), sep(), link("GitHub", bio.links.github), sep(), link("Portfolio", bio.links.portfolio),
  ] }));

  c.push(heading("Professional Summary"));
  c.push(new Paragraph({ spacing: { after: 40 }, alignment: AlignmentType.JUSTIFIED, children: runs(bio.summary) }));

  c.push(heading("Education"));
  for (const e of data.education) {
    c.push(roleLine(e.degree, e.school, e.dates));
    e.bullets.forEach((b) => c.push(bullet(b)));
  }

  c.push(heading("Professional Experience"));
  for (const w of data.work) {
    c.push(roleLine(w.title, w.org, w.dates));
    w.bullets.forEach((b) => c.push(bullet(b)));
    if (w.tech) c.push(techLine(w.tech));
  }

  c.push(heading("Key Projects"));
  for (const p of data.projects) {
    c.push(projLine(p));
    if (p.award) c.push(new Paragraph({ spacing: { after: 30 }, keepNext: true, children: [
      new TextRun({ text: p.award.prefix, italics: true, size: 19, color: GRAY, font: FONT }),
      new TextRun({ text: p.award.bold, italics: true, bold: true, size: 19, color: "000000", font: FONT }),
    ] }));
    p.bullets.forEach((b) => c.push(bullet(b)));
    if (p.tech) c.push(techLine(p.tech));
  }

  c.push(heading("Certifications"));
  data.certifications.forEach((x) => c.push(bullet(x)));

  c.push(heading("Technical Skills"));
  data.skills.forEach((s) => c.push(skillLine(s.group, s.items.join(", "))));

  const doc = new Document({
    creator: bio.name,
    title: `${bio.name} - Resume`,
    styles: { default: { document: { run: { font: FONT, size: 19, color: DARK } } } },
    numbering: { config: [{ reference: "bl", levels: [{
      level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { run: { color: NAVY }, paragraph: { indent: { left: 300, hanging: 200 } } },
    }] }] },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 600, bottom: 600, left: 720, right: 720 } } },
      children: c,
    }],
  });
  return Packer.toBuffer(doc);
}

// ---------------------------------------------------------------- README
function buildReadme() {
  const fp = data.projects.find((p) => p.featured) || data.projects[0];
  const typing = encodeURIComponent(bio.typingLines.join(";")).replace(/%20/g, "+");
  const edu = data.education[0];
  const cur = data.work[0], prev = data.work[1];
  const L = bio.links;
  const bullets = (arr) => arr.map((b) => `- ${b}`).join("\n");

  const skillsMd = data.skills.map((s) => `**${s.group}:** ${s.items.join(" · ")}`).join("  \n");
  const expRows = data.work.map((w) => `| ${w.title} | ${w.company} | ${w.dates} |`).join("\n");
  const projMd = data.projects.filter((p) => p !== fp).map((p) => {
    const first = plain(p.bullets[0]);
    const lk = p.link ? ` ([${p.link.label}](${p.link.url}))` : "";
    const aw = p.award ? ` 🏆 *${p.award.prefix.trim()} ${p.award.bold}*` : "";
    return `- **${p.name}**, ${p.subtitle}${lk}${aw}  \n  ${first}`;
  }).join("\n");

  return `<!-- GENERATED FILE. Edit resume/resume.json, not this file. Built by scripts/build.js -->
<div align="center">

# Hi, I'm ${bio.name} 👋

### ${bio.label.replace(" | ", ": ")}

![Typing SVG](https://readme-typing-svg.demolab.com?font=Fira+Code&size=20&pause=1200&color=2563EB&center=true&vCenter=true&width=560&lines=${typing})

[![Portfolio](https://img.shields.io/badge/Portfolio-2563EB?style=flat-square&logo=googlechrome&logoColor=white)](${L.portfolio})
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0A66C2?style=flat-square&logo=linkedin&logoColor=white)](${L.linkedin})
[![Email](https://img.shields.io/badge/Email-D14836?style=flat-square&logo=gmail&logoColor=white)](mailto:${bio.email})

</div>

---

### About Me

${bio.summary}

- 🎓 **${edu.degree}**, ${edu.school.split(",")[0]} (${edu.dates})
- 💼 Currently **${cur.title}** at ${cur.company}; previously **${prev.title}** at ${prev.company}
- 📍 Based in **${bio.location}**
- 📫 Reach me at **${bio.email}**

---

### 🚀 Featured Project: ${fp.name}

**${fp.subtitle}** → [${fp.link.url.replace("https://", "")}](${fp.link.url})

${bullets(fp.bullets)}

${fp.repo ? `[![Repo](https://img.shields.io/badge/Source-${encodeURIComponent(fp.repo.split("/").pop()).replace(/-/g, "--")}-181717?style=flat-square&logo=github&logoColor=white)](${fp.repo}) ` : ""}[![Live](https://img.shields.io/badge/Live-${fp.link.url.replace("https://", "").replace(/-/g, "--")}-10B981?style=flat-square&logo=vercel&logoColor=white)](${fp.link.url})

---

### 🧪 More Projects

${projMd}

---

### 🧰 Tech Stack

${skillsMd}

---

### 💼 Experience

| Role | Company | Period |
|---|---|---|
${expRows}

---

### 📊 GitHub Stats

<div align="center">

<img height="165" src="https://github-readme-stats.vercel.app/api?username=saurabh-028&show_icons=true&theme=default&hide_border=true&count_private=true" alt="GitHub stats" />
<img height="165" src="https://github-readme-stats.vercel.app/api/top-langs/?username=saurabh-028&layout=compact&theme=default&hide_border=true" alt="Top languages" />

</div>

---

<div align="center">

**Open to AI Engineer / ML Engineer roles in Ireland** · [Portfolio](${L.portfolio}) · [LinkedIn](${L.linkedin})

</div>
`;
}

// ---------------------------------------------------------------- LinkedIn copy-paste file
function buildLinkedin() {
  const out = [];
  out.push("# LinkedIn update sheet\n");
  out.push("_Generated from resume/resume.json. Paste each block into the matching LinkedIn field._\n");
  out.push("## Headline (max 220 chars)\n");
  out.push(bio.headline + `\n\n_(${bio.headline.length} chars)_\n`);
  out.push("## About\n");
  out.push(plain(bio.summary) + "\n");
  out.push("## Experience\n");
  for (const w of data.work) {
    out.push(`### ${w.title} | ${w.company}`);
    out.push(`Dates: ${w.dates}\n`);
    out.push(w.bullets.map((b) => `• ${plain(b)}`).join("\n"));
    if (w.tech) out.push(`\nSkills: ${w.tech}`);
    out.push("");
  }
  out.push("## Projects\n");
  for (const p of data.projects) {
    out.push(`### ${p.name}: ${p.subtitle}`);
    if (p.dates) out.push(`Dates: ${p.dates}`);
    if (p.link) out.push(`Link: ${p.link.url}`);
    if (p.award) out.push(`Award: ${p.award.prefix.trim()} ${p.award.bold}`);
    out.push("");
    out.push(p.bullets.map((b) => `• ${plain(b)}`).join("\n"));
    if (p.tech) out.push(`\nSkills: ${p.tech}`);
    out.push("");
  }
  out.push("## Education\n");
  for (const e of data.education) out.push(`• ${e.degree}, ${e.school} (${e.dates})`);
  out.push("\n## Licenses & Certifications\n");
  data.certifications.forEach((c) => out.push(`• ${c}`));
  out.push("\n## Skills (LinkedIn allows 50; pin the top 3)\n");
  const flat = [...new Set(data.skills.flatMap((s) => s.items))].slice(0, 50);
  out.push(flat.join(", ") + "\n");
  return out.join("\n");
}

// ---------------------------------------------------------------- run
(async () => {
  const strip = (o) => JSON.parse(JSON.stringify(o, (k, v) => (typeof v === "string" ? plain(v) : v)));
  fs.writeFileSync(path.join(DIST, "resume.docx"), await buildDocx());
  fs.writeFileSync(path.join(ROOT, "README.md"), buildReadme());
  fs.writeFileSync(path.join(DIST, "linkedin.md"), buildLinkedin());
  fs.writeFileSync(path.join(DIST, "portfolio.json"), JSON.stringify(strip(data), null, 2) + "\n");
  console.log("built: dist/resume.docx, README.md, dist/linkedin.md, dist/portfolio.json");
})();
