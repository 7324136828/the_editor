from __future__ import annotations

from .models import DocumentState

_BASE = (
    '<!DOCTYPE html><html><head><meta charset="utf-8"></head>'
    '<body style="font-family:Calibri;font-size:11pt;color:#263244;">'
)
_END = "</body></html>"

_TITLE = 'style="font-size:32pt;color:#17365d;margin-top:0px;margin-bottom:24px;"'
_SUBTITLE = 'style="font-size:16pt;color:#64748b;margin-top:0px;margin-bottom:21px;"'
_H1 = 'style="font-size:22pt;font-weight:bold;color:#17365d;margin-top:27px;margin-bottom:13px;"'
_H2 = 'style="font-size:16pt;font-weight:bold;color:#245c9f;margin-top:21px;margin-bottom:11px;"'
_BODY = 'style="font-size:11pt;color:#263244;margin-top:0px;margin-bottom:11px;"'
_QUOTE = ('style="font-size:12pt;font-style:italic;color:#64748b;'
          'margin-top:13px;margin-bottom:13px;margin-left:24px;"')
_TABLE = ('style="border-collapse:collapse;width:100%;margin-top:8px;margin-bottom:12px;" '
          'border="1" cellspacing="0" cellpadding="6"')
_TH = 'style="background-color:#dce8f9;border:1px solid #b8c4d4;text-align:left;"'
_TD = 'style="border:1px solid #b8c4d4;"'


def _welcome() -> str:
    return _BASE + f"""
<p {_TITLE}>Welcome to Folio</p>
<p {_SUBTITLE}>A local-first word processor that stays out of your way.</p>
<h1 {_H1}>Getting started</h1>
<p {_BODY}>Everything you type stays on this computer. Folio saves real
.docx files you can open in Microsoft Word, keeps automatic local
recovery snapshots, and never needs an account or a network connection.</p>
<ul>
<li {_BODY}>Pick a style such as Heading 1 from the Home tab to structure your document.</li>
<li {_BODY}>Use the Review tab to track text changes and attach comments to a selection.</li>
<li {_BODY}>Open the Layout tab for page size, margins, columns, and running headers.</li>
</ul>
<h1 {_H1}>Handy shortcuts</h1>
<table {_TABLE}>
<tr><th {_TH}>Shortcut</th><th {_TH}>What it does</th></tr>
<tr><td {_TD}>Ctrl+1 / Ctrl+2 / Ctrl+3</td><td {_TD}>Apply Heading 1, 2, or 3</td></tr>
<tr><td {_TD}>Ctrl+0</td><td {_TD}>Return to Normal text</td></tr>
<tr><td {_TD}>Ctrl+Enter</td><td {_TD}>Insert a page break</td></tr>
<tr><td {_TD}>Ctrl+F / Ctrl+H</td><td {_TD}>Find / Replace</td></tr>
</table>
<p {_QUOTE}>Tip: the References tab can generate a table of contents from
your headings, and File &gt; Export PDF produces a publication-ready
layout with columns, headers, and page numbers.</p>
""" + _END


def _report() -> str:
    return _BASE + f"""
<p {_TITLE}>Project Report</p>
<p {_SUBTITLE}>{{{{Project}}}} &mdash; status summary</p>
<h1 {_H1}>Overview</h1>
<p {_BODY}>Summarize the goal, scope, and current state of the project in a
few sentences. This template is a starting point &mdash; replace each
section with your own content.</p>
<h1 {_H1}>Milestones</h1>
<table {_TABLE}>
<tr><th {_TH}>Milestone</th><th {_TH}>Owner</th><th {_TH}>Status</th><th {_TH}>Due</th></tr>
<tr><td {_TD}>Kickoff</td><td {_TD}>{{{{Owner}}}}</td><td {_TD}>Complete</td><td {_TD}>&nbsp;</td></tr>
<tr><td {_TD}>Draft deliverable</td><td {_TD}>&nbsp;</td><td {_TD}>In progress</td><td {_TD}>&nbsp;</td></tr>
<tr><td {_TD}>Final review</td><td {_TD}>&nbsp;</td><td {_TD}>Planned</td><td {_TD}>&nbsp;</td></tr>
</table>
<h1 {_H1}>Risks and next steps</h1>
<ul>
<li {_BODY}>List the top risk and its mitigation.</li>
<li {_BODY}>List the next concrete action and its owner.</li>
</ul>
""" + _END


def _resume() -> str:
    return _BASE + f"""
<p {_TITLE}>{{{{Name}}}}</p>
<p {_SUBTITLE}>{{{{City}}}} &middot; {{{{Email}}}} &middot; {{{{Phone}}}}</p>
<h1 {_H1}>Experience</h1>
<h2 {_H2}>Job Title &mdash; Company</h2>
<p {_BODY}><i>2021 &ndash; Present</i></p>
<ul>
<li {_BODY}>Describe a measurable accomplishment.</li>
<li {_BODY}>Describe a responsibility or project.</li>
</ul>
<h1 {_H1}>Education</h1>
<p {_BODY}><b>Degree, Field</b> &mdash; School, Year</p>
<h1 {_H1}>Skills</h1>
<ul>
<li {_BODY}>Skill one</li>
<li {_BODY}>Skill two</li>
<li {_BODY}>Skill three</li>
</ul>
""" + _END


def _invoice() -> str:
    return _BASE + f"""
<p {_TITLE}>Invoice</p>
<p {_SUBTITLE}>{{{{Company}}}}</p>
<p {_BODY}><b>Bill to:</b> {{{{CustomerName}}}}<br>
<b>Invoice #:</b> {{{{InvoiceNumber}}}} &nbsp;&nbsp; <b>Date:</b> {{{{Date}}}}</p>
<table {_TABLE}>
<tr><th {_TH}>Item</th><th {_TH}>Qty</th><th {_TH}>Unit price</th><th {_TH}>Amount</th></tr>
<tr><td {_TD}>{{{{ItemOne}}}}</td><td {_TD}>1</td><td {_TD}>{{{{PriceOne}}}}</td><td {_TD}>{{{{PriceOne}}}}</td></tr>
<tr><td {_TD}>{{{{ItemTwo}}}}</td><td {_TD}>1</td><td {_TD}>{{{{PriceTwo}}}}</td><td {_TD}>{{{{PriceTwo}}}}</td></tr>
<tr><td {_TD}><b>Total</b></td><td {_TD}></td><td {_TD}></td><td {_TD}><b>{{{{Total}}}}</b></td></tr>
</table>
<p {_QUOTE}>Amounts are not calculated automatically &mdash; fill in each
placeholder, for example with the Mail Merge tool on the Insert tab.</p>
""" + _END


def _essay() -> str:
    return _BASE + f"""
<p {_TITLE}>Essay Title</p>
<p {_SUBTITLE}>{{{{Student}}}} &mdash; {{{{Course}}}}</p>
<p {_BODY}>Introduce your topic and state your thesis in the opening
paragraph. A strong introduction gives the reader a map of the argument
that follows.</p>
<h1 {_H1}>Background</h1>
<p {_BODY}>Provide the context a reader needs. Cite sources inline and keep
each paragraph focused on a single idea.</p>
<p {_QUOTE}>Use the Quote style for short excerpts or pull quotes that
deserve emphasis.</p>
<h1 {_H1}>Analysis</h1>
<p {_BODY}>Develop your argument with evidence. Finish with a conclusion
that answers the "so what" question.</p>
""" + _END


_BUILDERS = {
    "blank": lambda: _BASE + "<p></p>" + _END,
    "welcome": _welcome,
    "report": _report,
    "resume": _resume,
    "invoice": _invoice,
    "essay": _essay,
}

_TITLES = {
    "blank": "Untitled",
    "welcome": "Welcome",
    "report": "Project report",
    "resume": "Resume",
    "invoice": "Invoice",
    "essay": "Academic essay",
}

TEMPLATE_NAMES = tuple(_BUILDERS.keys())


def make_template(name: str) -> DocumentState:
    builder = _BUILDERS.get(name)
    if builder is None:
        raise ValueError(f"Unknown template: {name}")
    state = DocumentState()
    state.title = _TITLES[name]
    state.html = builder()
    return state
