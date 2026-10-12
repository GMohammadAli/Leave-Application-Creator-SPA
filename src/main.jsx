import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import { Document, Packer, Paragraph, TextRun, AlignmentType } from "docx";
import { Download, Printer, RotateCcw, FileText, CalendarDays, Eye, PencilLine } from "lucide-react";
import "./styles.css";

const DEFAULTS = {
  recipient: "The Head Master",
  school: "Dr. Zakir Hussain High School",
  city: "Parbhani",
  salutation: "Respected Madam,",
  subject: "Application for six days E.L.",
  startDate: "2024-07-03",
  endDate: "2024-07-08",
  reason:
    "Due to some personal reason, I am not able to attend the school. I kindly request you to grant me leave for the dates mentioned above.",
  teacherName: "Mrs. Shabana Parveen",
  additionalName: "A. Wadhab Siddhewalam",
  designation: "Assistant Teacher",
  closing: "Yours faithfully,",
};

const FIELD_GROUPS = [
  {
    title: "School details",
    icon: FileText,
    fields: [
      ["recipient", "Recipient / designation"],
      ["school", "School name"],
      ["city", "City"],
      ["salutation", "Salutation"],
    ],
  },
  {
    title: "Leave details",
    icon: CalendarDays,
    fields: [
      ["subject", "Subject"],
      ["startDate", "Leave from", "date"],
      ["endDate", "Leave until", "date"],
    ],
  },
  {
    title: "Applicant details",
    icon: PencilLine,
    fields: [
      ["teacherName", "Teacher name"],
      ["additionalName", "Additional name / signature detail (optional)"],
      ["designation", "Designation"],
      ["closing", "Closing"],
    ],
  },
];

function formatDate(value) {
  if (!value) return "__________";
  const [year, month, day] = value.split("-");
  return `${day}/${month}/${year}`;
}

function makeDates(data) {
  return `${formatDate(data.startDate)} to ${formatDate(data.endDate)}`;
}

function App() {
  const [form, setForm] = useState(DEFAULTS);
  const [notice, setNotice] = useState("");

  const dates = useMemo(() => makeDates(form), [form.startDate, form.endDate]);

  function update(key, value) {
    setForm((current) => ({ ...current, [key]: value }));
    setNotice("");
  }

  function resetForm() {
    setForm({ ...DEFAULTS });
    setNotice("Original values restored.");
  }

  function validate() {
    if (form.startDate && form.endDate && form.endDate < form.startDate) {
      setNotice("The end date must be on or after the start date.");
      return false;
    }
    if (!form.teacherName.trim() || !form.school.trim()) {
      setNotice("Please enter at least the teacher name and school name.");
      return false;
    }
    return true;
  }

  async function downloadDocx() {
    if (!validate()) return;
    setNotice("Preparing your Word document…");
    const p = (text, { bold = false, align, size = 24, after = 120 } = {}) =>
      new Paragraph({
        children: [new TextRun({ text, bold, size, font: "Times New Roman" })],
        alignment: align,
        spacing: { after },
      });

    const children = [
      p("LEAVE APPLICATION", { bold: true, size: 30, align: AlignmentType.CENTER, after: 420 }),
      p("To,"),
      p(`${form.recipient},`),
      p(`${form.school},`),
      p(form.city, { after: 260 }),
      p(`Subject: ${form.subject}`, { after: 160 }),
      p(`Leave dates: ${dates}`, { after: 300 }),
      p(form.salutation, { after: 220 }),
      p(form.reason.replaceAll("{DATES}", dates), { after: 180 }),
      p("Thanking you.", { after: 520 }),
      p(form.closing, { align: AlignmentType.RIGHT }),
      p(form.teacherName, { bold: true, align: AlignmentType.RIGHT }),
      ...(form.additionalName.trim()
        ? [p(form.additionalName, { align: AlignmentType.RIGHT })]
        : []),
      p(form.designation, { align: AlignmentType.RIGHT }),
      p(form.school, { align: AlignmentType.RIGHT }),
    ];

    try {
      const doc = new Document({
        sections: [{
          properties: {
            page: { margin: { top: 900, right: 1080, bottom: 900, left: 1080 } },
          },
          children,
        }],
      });
      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "leave-application.docx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1500);
      setNotice("Word document downloaded.");
    } catch (error) {
      console.error(error);
      setNotice("Could not create the Word document. Please try Print / Save as PDF.");
    }
  }

  function printLetter() {
    if (!validate()) return;
    window.print();
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand-mark"><FileText size={23} /></div>
        <div>
          <h1>Leave Application Maker</h1>
          <p>Edit the details, preview the letter, and download a Word document.</p>
        </div>
      </header>

      <main className="workspace">
        <section className="editor-panel panel">
          <div className="section-heading">
            <div>
              <span className="eyebrow">CUSTOMIZE</span>
              <h2>Application details</h2>
            </div>
            <button type="button" className="icon-button" onClick={resetForm} title="Restore defaults" aria-label="Restore defaults">
              <RotateCcw size={17} />
            </button>
          </div>

          {FIELD_GROUPS.map((group) => {
            const Icon = group.icon;
            return (
              <fieldset className="field-group" key={group.title}>
                <legend><Icon size={16} /> {group.title}</legend>
                <div className="field-grid">
                  {group.fields.map(([key, label, type]) => (
                    <label className={`field ${key === "additionalName" || key === "school" || key === "subject" ? "wide" : ""}`} key={key}>
                      <span>{label}</span>
                      <input
                        type={type || "text"}
                        value={form[key]}
                        onChange={(event) => update(key, event.target.value)}
                      />
                    </label>
                  ))}
                </div>
              </fieldset>
          );})}

          <div className="field-group reason-group">
            <label className="field wide">
              <span>Reason / application text</span>
              <textarea rows={5} value={form.reason} onChange={(event) => update("reason", event.target.value)} />
            </label>
          </div>

          {notice && <p className={`notice ${notice.includes("must") || notice.includes("Please") || notice.includes("Could not") ? "error" : ""}`} role="status">{notice}</p>}

          <div className="actions">
            <button type="button" className="button primary" onClick={downloadDocx}><Download size={17} /> Download Word (.docx)</button>
            <button type="button" className="button secondary" onClick={printLetter}><Printer size={17} /> Print / Save as PDF</button>
            <button type="button" className="button subtle" onClick={resetForm}><RotateCcw size={16} /> Restore defaults</button>
          </div>
          <p className="privacy-note">Your entries stay in this browser session. No backend or account is required.</p>
        </section>

        <section className="preview-panel panel">
          <div className="section-heading preview-heading">
            <div>
              <span className="eyebrow"><Eye size={13} /> LIVE PREVIEW</span>
              <h2>Your letter</h2>
            </div>
            <span className="live-pill"><span /> Live</span>
          </div>
          <article className="paper" aria-label="Leave application preview">
            <h3>LEAVE APPLICATION</h3>
            <p>To,<br />{form.recipient},<br />{form.school},<br />{form.city}</p>
            <p className="letter-subject"><strong>Subject:</strong> {form.subject}</p>
            <p><strong>Leave dates:</strong> {dates}</p>
            <p className="salutation">{form.salutation}</p>
            <p className="letter-body">{form.reason.replaceAll("{DATES}", dates)}</p>
            <p>Thanking you.</p>
            <div className="signature">
              <p>{form.closing}</p>
              <p><strong>{form.teacherName}</strong></p>
              {form.additionalName.trim() && <p>{form.additionalName}</p>}
              <p>{form.designation}</p>
              <p>{form.school}</p>
            </div>
          </article>
          <p className="preview-footnote">Please verify names and dates before submitting the application.</p>
        </section>
      </main>
      <footer>Leave Application Maker <span>•</span> React SPA <span>•</span> Ready for static hosting</footer>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<React.StrictMode><App /></React.StrictMode>);
