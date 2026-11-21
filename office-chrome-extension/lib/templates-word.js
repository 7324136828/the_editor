/**
 * Word document templates: ready-to-fill word models (see lib/model-word.js
 * for the model contract). Used by "New -> Word document from template…" and
 * viewer.html?new=word&template=<id>.
 *
 *   WORD_TEMPLATES            [{ id, name, description }] in menu order
 *   wordTemplate(id, opts?)   a fresh model (deep copy) or null for an unknown id;
 *                             opts.date: Date used for date lines (default: today)
 *
 * Placeholders are written in [square brackets] so they are easy to spot and
 * replace (Find & Replace finds them too).
 */

export const WORD_TEMPLATES = [
  { id: "blank", name: "Blank document", description: "An empty page." },
  { id: "letter", name: "Letter", description: "A formal letter with sender, recipient and signature blocks." },
  { id: "resume", name: "Résumé", description: "A one-page CV with experience, education and skills." },
  { id: "report", name: "Report", description: "Title page, sections, a data table and page numbers." },
  { id: "meeting", name: "Meeting notes", description: "Attendees, agenda, decisions and an action-item table." },
  { id: "memo", name: "Memo", description: "An internal memorandum with To / From / Date / Subject lines." },
];

const ACCENT = "#1F3864";
const MUTED = "#595959";

const t = (text, marks = {}) => ({ t: text, ...marks });
const p = (runs, align) => {
  const b = { type: "p", runs: typeof runs === "string" ? [t(runs)] : runs };
  if (align) b.align = align;
  return b;
};
const h = (level, text, align) => {
  const b = { type: `h${level}`, runs: [t(text)] };
  if (align) b.align = align;
  return b;
};
const li = (text, list = "ul", level = 0) => ({ type: "list", list, level, runs: typeof text === "string" ? [t(text)] : text });
const cell = (runs) => ({ blocks: [p(runs)] });
const table = (rows) => ({ type: "table", rows: rows.map((r) => r.map((c) => (c && c.blocks ? c : cell(c)))) });
const empty = () => p([]);

function longDate(d) {
  try {
    return new Intl.DateTimeFormat(undefined, { year: "numeric", month: "long", day: "numeric" }).format(d);
  } catch {
    return d.toDateString();
  }
}

const BUILDERS = {
  blank: () => ({ kind: "word", blocks: [empty()] }),

  letter: ({ date }) => ({
    kind: "word",
    meta: { title: "Letter" },
    blocks: [
      p([t("[Your Name]", { b: true, size: 14, color: ACCENT })]),
      p([t("[Street Address]\n[City, State ZIP]\n[Email] · [Phone]", { color: MUTED })]),
      empty(),
      p(longDate(date)),
      empty(),
      p("[Recipient Name]\n[Title]\n[Company]\n[Street Address]\n[City, State ZIP]"),
      empty(),
      p("Dear [Recipient Name],"),
      p("I am writing to [state the purpose of your letter in one or two sentences]. [Add the most important detail the reader needs to know.]"),
      p("[Use the second paragraph for supporting details: background, facts, dates or figures. Keep sentences short and specific.]"),
      p("[Close with what you would like to happen next, and how the reader can reach you.] Thank you for your time and consideration."),
      empty(),
      p("Sincerely,"),
      empty(),
      empty(),
      p([t("[Your Name]", { b: true })]),
      p([t("[Your Title]", { color: MUTED })]),
    ],
  }),

  resume: () => ({
    kind: "word",
    meta: { title: "Résumé", page: { size: "Letter", orientation: "portrait", margins: { top: 1080, right: 1080, bottom: 1080, left: 1080, header: 720, footer: 720 } } },
    blocks: [
      { type: "h1", align: "center", runs: [t("[YOUR NAME]", { color: ACCENT })] },
      p([t("[City, Country] · [phone] · [email@example.com] · [linkedin.com/in/you]", { color: MUTED })], "center"),
      h(2, "Summary"),
      p("[Two or three sentences about who you are, what you are good at and the kind of role you are looking for.]"),
      h(2, "Experience"),
      p([t("[Job Title]", { b: true }), t(" — [Company], [City]")]),
      p([t("[Month Year] – Present", { i: true, color: MUTED })]),
      li("[Achievement with a number: e.g. cut processing time by 30% by …]"),
      li("[Responsibility or project you led, and its result]"),
      li("[Tool, process or team you improved]"),
      p([t("[Job Title]", { b: true }), t(" — [Company], [City]")]),
      p([t("[Month Year] – [Month Year]", { i: true, color: MUTED })]),
      li("[Achievement with a measurable result]"),
      li("[Key responsibility]"),
      h(2, "Education"),
      p([t("[Degree], [Field of study]", { b: true }), t(" — [School], [Year]")]),
      h(2, "Skills"),
      table([
        [[t("Technical", { b: true })], "[Skill], [Skill], [Skill]"],
        [[t("Languages", { b: true })], "[Language (level)], [Language (level)]"],
        [[t("Other", { b: true })], "[Certification], [Volunteer work], [Interest]"],
      ]),
    ],
  }),

  report: ({ date }) => ({
    kind: "word",
    meta: {
      title: "Report",
      header: { runs: [t("[Report title]", { color: MUTED })], align: "right" },
      footer: { runs: [t("Page {PAGE} of {NUMPAGES}", { color: MUTED })], align: "center" },
    },
    blocks: [
      empty(),
      empty(),
      { type: "h1", align: "center", runs: [t("[Report Title]", { size: 28, color: ACCENT })] },
      p([t("[Subtitle or one-line description]", { size: 14, color: MUTED })], "center"),
      empty(),
      p([t("Prepared by [Author]", { b: true })], "center"),
      p(longDate(date), "center"),
      { type: "pagebreak" },
      h(1, "Executive summary"),
      p("[Summarize the purpose, the key findings and your main recommendation in one short paragraph. Many readers stop here, so make it count.]"),
      h(1, "1. Introduction"),
      p("[Explain the background, the question this report answers and how the information was gathered.]"),
      h(1, "2. Findings"),
      h(2, "2.1 [First finding]"),
      p("[Describe what you found and why it matters.]"),
      table([
        [[t("Metric", { b: true })], [t("Last period", { b: true })], [t("This period", { b: true })], [t("Change", { b: true })]],
        ["[Metric A]", "[0]", "[0]", "[+0%]"],
        ["[Metric B]", "[0]", "[0]", "[+0%]"],
        ["[Metric C]", "[0]", "[0]", "[+0%]"],
      ]),
      h(2, "2.2 [Second finding]"),
      li("[Supporting point]"),
      li("[Supporting point]"),
      h(1, "3. Recommendations"),
      li("[Recommended action and expected benefit]", "ol"),
      li("[Recommended action and expected benefit]", "ol"),
      li("[Recommended action and expected benefit]", "ol"),
      h(1, "4. Conclusion"),
      p("[Wrap up with the single most important takeaway.]"),
    ],
  }),

  meeting: ({ date }) => ({
    kind: "word",
    meta: { title: "Meeting notes" },
    blocks: [
      h(1, "[Meeting name] — Notes"),
      table([
        [[t("Date", { b: true })], longDate(date)],
        [[t("Time", { b: true })], "[10:00 – 11:00]"],
        [[t("Location", { b: true })], "[Room / video link]"],
        [[t("Facilitator", { b: true })], "[Name]"],
        [[t("Note taker", { b: true })], "[Name]"],
      ]),
      h(2, "Attendees"),
      li("[Name], [Role]"),
      li("[Name], [Role]"),
      li("[Name], [Role]"),
      h(2, "Agenda"),
      li("[Topic] — [owner], [10 min]", "ol"),
      li("[Topic] — [owner], [15 min]", "ol"),
      li("[Topic] — [owner], [5 min]", "ol"),
      h(2, "Discussion"),
      p("[Key points raised for each agenda item.]"),
      h(2, "Decisions"),
      li("[Decision made, and who agreed]"),
      h(2, "Action items"),
      table([
        [[t("Action", { b: true })], [t("Owner", { b: true })], [t("Due", { b: true })], [t("Status", { b: true })]],
        ["[What needs to happen]", "[Name]", "[Date]", "[Open]"],
        ["[What needs to happen]", "[Name]", "[Date]", "[Open]"],
      ]),
      h(2, "Next meeting"),
      p("[Date, time and place]"),
    ],
  }),

  memo: ({ date }) => ({
    kind: "word",
    meta: {
      title: "Memo",
      header: { runs: [t("[Company Name]", { b: true, color: ACCENT }), t("\tInternal", { color: MUTED })] },
    },
    blocks: [
      { type: "h1", runs: [t("MEMORANDUM", { color: ACCENT })] },
      p([t("To:\t", { b: true }), t("[Recipients]")]),
      p([t("From:\t", { b: true }), t("[Your Name, Title]")]),
      p([t("CC:\t", { b: true }), t("[Others who should be informed]")]),
      p([t("Date:\t", { b: true }), t(longDate(date))]),
      p([t("Subject:\t", { b: true }), t("[Short, specific subject]")]),
      p([t("________________________________________________________________________________", { color: "#BFBFBF" })]),
      p("[Opening: state the purpose of the memo and the action you need, in one or two sentences.]"),
      p("[Context: give the background the readers need. Keep it short.]"),
      li("[Key point or requirement]"),
      li("[Key point or requirement]"),
      li("[Deadline or next step]"),
      p("[Closing: who to contact with questions.]"),
    ],
  }),
};

/** A fresh model for template `id`, or null when there is no such template. */
export function wordTemplate(id, opts = {}) {
  const build = Object.prototype.hasOwnProperty.call(BUILDERS, id) ? BUILDERS[id] : null;
  if (!build) return null;
  const date = opts.date instanceof Date && !Number.isNaN(opts.date.getTime()) ? opts.date : new Date();
  return JSON.parse(JSON.stringify(build({ date })));
}
