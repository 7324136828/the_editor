import {
  WordDocumentModel,
  ExcelDocumentModel,
  PptDocumentModel,
  CodeDocumentModel,
} from "../types/office";

export const sampleWordDocument: WordDocumentModel = {
  id: "doc-word-01",
  title: "Executive_Strategy_2026.docx",
  author: "Sarah Chen, VP Strategy",
  createdAt: "2026-09-15 09:30 AM",
  modifiedAt: "2026-09-18 04:15 PM",
  headings: [
    { id: "h-1", level: 1, text: "1. Executive Summary", pageIndex: 1 },
    {
      id: "h-2",
      level: 2,
      text: "1.1 Vision & Market Opportunity",
      pageIndex: 1,
    },
    {
      id: "h-3",
      level: 1,
      text: "2. Financial Performance & Targets",
      pageIndex: 2,
    },
    {
      id: "h-4",
      level: 2,
      text: "2.1 Revenue Growth by Territory",
      pageIndex: 2,
    },
    {
      id: "h-5",
      level: 1,
      text: "3. Technical Architecture & Synthesizer",
      pageIndex: 3,
    },
    {
      id: "h-6",
      level: 2,
      text: "3.1 Adaptive Accordion Sidebar",
      pageIndex: 3,
    },
    {
      id: "h-7",
      level: 1,
      text: "4. Strategic Milestones & Next Steps",
      pageIndex: 4,
    },
  ],
  sections: [
    {
      id: "sec-1",
      title: "Executive Overview",
      pageNumber: 1,
      orientation: "portrait",
    },
    {
      id: "sec-2",
      title: "Financial Forecast",
      pageNumber: 2,
      orientation: "portrait",
    },
    {
      id: "sec-3",
      title: "Platform Engineering",
      pageNumber: 3,
      orientation: "portrait",
    },
    {
      id: "sec-4",
      title: "Roadmap & Rollout",
      pageNumber: 4,
      orientation: "portrait",
    },
  ],
  paragraphs: [
    {
      id: "p-1",
      type: "heading-1",
      text: "1. Executive Summary",
    },
    {
      id: "p-2",
      type: "body",
      runs: [
        {
          text: "In fiscal year 2026, our core objective is to deliver the industry’s first unified ",
        },
        { text: "VS Code-based Office Synthesizer", bold: true },
        {
          text: ". By bridging developer workflows with document processing across Word (.docx), Excel (.xlsx), and PowerPoint (.pptx), teams can inspect, edit, and orchestrate documents with code-editor ergonomics.",
        },
      ],
    },
    {
      id: "p-3",
      type: "heading-2",
      text: "1.1 Vision & Market Opportunity",
    },
    {
      id: "p-4",
      type: "body",
      text: "Modern knowledge workers switch between up to seven disparate productivity tools daily. Integrating rich document inspectors into a familiar IDE surface reduces task switching latency by an estimated 38% while providing scriptable automation.",
    },
    {
      id: "p-5",
      type: "callout",
      text: "Key Directive: Ensure zero-install web and node module distribution, allowing instant integration into custom internal engineering portals and desktop wrappers.",
    },
    {
      id: "p-6",
      type: "heading-1",
      text: "2. Financial Performance & Targets",
    },
    {
      id: "p-7",
      type: "body",
      text: "The projected growth across enterprise tiers demonstrates strong capital efficiency and high net dollar retention.",
    },
    {
      id: "p-8",
      type: "table",
      tableData: [
        [
          "Quarter",
          "Projected ARR ($M)",
          "Growth YoY",
          "Gross Margin",
          "Status",
        ],
        ["Q1 2026", "$18.4M", "+42%", "78.2%", "Exceeded"],
        ["Q2 2026", "$24.6M", "+54%", "80.5%", "On Track"],
        ["Q3 2026", "$32.1M", "+61%", "81.4%", "Projected"],
        ["Q4 2026", "$41.8M", "+68%", "83.0%", "Target"],
      ],
    },
    {
      id: "p-9",
      type: "heading-2",
      text: "2.1 Revenue Growth by Territory",
    },
    {
      id: "p-10",
      type: "body",
      text: "North America continues to account for 58% of enterprise commitments, followed closely by EMEA (26%) and APAC (16%), where adoption of developer productivity tooling is accelerating.",
    },
    {
      id: "p-11",
      type: "heading-1",
      text: "3. Technical Architecture & Synthesizer",
    },
    {
      id: "p-12",
      type: "body",
      runs: [
        { text: "The core innovation lies in the " },
        { text: "Adaptive Left Accordion Layout", bold: true },
        {
          text: ". Rather than presenting static file explorer nodes, selecting an office file automatically swaps the sidebar into specialized domain panels.",
        },
      ],
    },
    {
      id: "p-13",
      type: "heading-2",
      text: "3.1 Adaptive Accordion Sidebar",
    },
    {
      id: "p-14",
      type: "body",
      text: "For Word documents, the sidebar reveals Outline Navigation, Page Breaks, Review Comments, and Document Statistics. For Excel, it offers Sheet Management, Formula Auditing, and Data Pivot Fields. For PowerPoint, it shows Slide Filmstrips and Object Layers.",
    },
    {
      id: "p-15",
      type: "heading-1",
      text: "4. Strategic Milestones & Next Steps",
    },
    {
      id: "p-16",
      type: "body",
      text: "Beta rollout begins October 2026 across 50 select pilot organizations. General Availability with full desktop packaging is slated for Q1 2027.",
    },
  ],
  comments: [
    {
      id: "c-1",
      author: "David Kim (Lead Architect)",
      avatarColor: "#007acc",
      timestamp: "Yesterday at 3:42 PM",
      selectedText: "VS Code-based Office Synthesizer",
      comment:
        "Consider documenting the memory footprint when loading workbooks larger than 50MB.",
      resolved: false,
    },
    {
      id: "c-2",
      author: "Elena Rostova (CFO)",
      avatarColor: "#107c41",
      timestamp: "Sep 17 at 11:20 AM",
      selectedText: "$41.8M",
      comment:
        "Q4 numbers reflect upside from expansion in the financial services sector.",
      resolved: true,
    },
    {
      id: "c-3",
      author: "Marcus Vance (Design)",
      avatarColor: "#c43e1c",
      timestamp: "Sep 18 at 2:05 PM",
      selectedText: "Adaptive Left Accordion Layout",
      comment:
        "Visual transitions between document accordions feel very smooth.",
      resolved: false,
    },
  ],
  stats: {
    words: 1840,
    characters: 11450,
    charactersNoSpaces: 9680,
    paragraphs: 34,
    pages: 4,
    readingTimeMinutes: 7,
    readabilityScore: "Grade 11 (Business Professional)",
  },
  typography: {
    fontFamily: "Segoe UI",
    fontSize: 14,
    lineHeight: 1.65,
    marginSize: "normal",
    textColor: "#24292f",
  },
};

export const sampleExcelDocument: ExcelDocumentModel = {
  id: "doc-excel-01",
  title: "Global_Sales_Model_2026.xlsx",
  activeSheetId: "sheet-1",
  sheets: [
    {
      id: "sheet-1",
      name: "Regional Sales",
      rowCount: 12,
      colCount: 7,
      columns: [
        { key: "A", label: "Region", width: 130, type: "string" },
        { key: "B", label: "Product Tier", width: 150, type: "string" },
        { key: "C", label: "Units Sold", width: 110, type: "number" },
        { key: "D", label: "Unit Price", width: 120, type: "currency" },
        { key: "E", label: "Gross Revenue", width: 140, type: "currency" },
        { key: "F", label: "Op Margin", width: 110, type: "currency" },
        { key: "G", label: "Net Profit", width: 140, type: "currency" },
      ],
      cells: {
        A1: { value: "Region", format: { bold: true, fill: "#f2f4f8" } },
        B1: { value: "Product Tier", format: { bold: true, fill: "#f2f4f8" } },
        C1: {
          value: "Units Sold",
          format: { bold: true, align: "right", fill: "#f2f4f8" },
        },
        D1: {
          value: "Unit Price",
          format: { bold: true, align: "right", fill: "#f2f4f8" },
        },
        E1: {
          value: "Gross Revenue",
          format: { bold: true, align: "right", fill: "#f2f4f8" },
        },
        F1: {
          value: "Op Margin %",
          format: { bold: true, align: "right", fill: "#f2f4f8" },
        },
        G1: {
          value: "Net Profit",
          format: { bold: true, align: "right", fill: "#f2f4f8" },
        },

        A2: { value: "North America" },
        B2: { value: "Enterprise" },
        C2: { value: 1450, format: { align: "right", numberFormat: "number" } },
        D2: {
          value: 1200,
          format: { align: "right", numberFormat: "currency" },
        },
        E2: {
          value: 1740000,
          formula: "=C2*D2",
          formatted: "$1,740,000",
          format: { align: "right", numberFormat: "currency" },
        },
        F2: {
          value: 0.42,
          format: { align: "right", numberFormat: "percent" },
        },
        G2: {
          value: 730800,
          formula: "=E2*F2",
          formatted: "$730,800",
          format: { align: "right", numberFormat: "currency", bold: true },
        },

        A3: { value: "North America" },
        B3: { value: "Professional" },
        C3: { value: 3800, format: { align: "right", numberFormat: "number" } },
        D3: {
          value: 450,
          format: { align: "right", numberFormat: "currency" },
        },
        E3: {
          value: 1710000,
          formula: "=C3*D3",
          formatted: "$1,710,000",
          format: { align: "right", numberFormat: "currency" },
        },
        F3: {
          value: 0.38,
          format: { align: "right", numberFormat: "percent" },
        },
        G3: {
          value: 649800,
          formula: "=E3*F3",
          formatted: "$649,800",
          format: { align: "right", numberFormat: "currency" },
        },

        A4: { value: "EMEA" },
        B4: { value: "Enterprise" },
        C4: { value: 920, format: { align: "right", numberFormat: "number" } },
        D4: {
          value: 1250,
          format: { align: "right", numberFormat: "currency" },
        },
        E4: {
          value: 1150000,
          formula: "=C4*D4",
          formatted: "$1,150,000",
          format: { align: "right", numberFormat: "currency" },
        },
        F4: { value: 0.4, format: { align: "right", numberFormat: "percent" } },
        G4: {
          value: 460000,
          formula: "=E4*F4",
          formatted: "$460,000",
          format: { align: "right", numberFormat: "currency" },
        },

        A5: { value: "EMEA" },
        B5: { value: "Professional" },
        C5: { value: 2400, format: { align: "right", numberFormat: "number" } },
        D5: {
          value: 450,
          format: { align: "right", numberFormat: "currency" },
        },
        E5: {
          value: 1080000,
          formula: "=C5*D5",
          formatted: "$1,080,000",
          format: { align: "right", numberFormat: "currency" },
        },
        F5: {
          value: 0.36,
          format: { align: "right", numberFormat: "percent" },
        },
        G5: {
          value: 388800,
          formula: "=E5*F5",
          formatted: "$388,800",
          format: { align: "right", numberFormat: "currency" },
        },

        A6: { value: "APAC" },
        B6: { value: "Enterprise" },
        C6: { value: 680, format: { align: "right", numberFormat: "number" } },
        D6: {
          value: 1150,
          format: { align: "right", numberFormat: "currency" },
        },
        E6: {
          value: 782000,
          formula: "=C6*D6",
          formatted: "$782,000",
          format: { align: "right", numberFormat: "currency" },
        },
        F6: {
          value: 0.45,
          format: { align: "right", numberFormat: "percent" },
        },
        G6: {
          value: 351900,
          formula: "=E6*F6",
          formatted: "$351,900",
          format: { align: "right", numberFormat: "currency" },
        },

        A7: { value: "APAC" },
        B7: { value: "Professional" },
        C7: { value: 1950, format: { align: "right", numberFormat: "number" } },
        D7: {
          value: 420,
          format: { align: "right", numberFormat: "currency" },
        },
        E7: {
          value: 819000,
          formula: "=C7*D7",
          formatted: "$819,000",
          format: { align: "right", numberFormat: "currency" },
        },
        F7: {
          value: 0.34,
          format: { align: "right", numberFormat: "percent" },
        },
        G7: {
          value: 278460,
          formula: "=E7*F7",
          formatted: "$278,460",
          format: { align: "right", numberFormat: "currency" },
        },

        A8: {
          value: "TOTAL / SUMMARY",
          format: { bold: true, fill: "#e6f4ea" },
        },
        B8: { value: "All Tiers", format: { bold: true, fill: "#e6f4ea" } },
        C8: {
          value: 11200,
          formula: "=SUM(C2:C7)",
          formatted: "11,200",
          format: { bold: true, align: "right", fill: "#e6f4ea" },
        },
        D8: {
          value: 820,
          formula: "=AVERAGE(D2:D7)",
          formatted: "$820",
          format: { bold: true, align: "right", fill: "#e6f4ea" },
        },
        E8: {
          value: 7281000,
          formula: "=SUM(E2:E7)",
          formatted: "$7,281,000",
          format: { bold: true, align: "right", fill: "#e6f4ea" },
        },
        F8: {
          value: 0.39,
          formula: "=AVERAGE(F2:F7)",
          formatted: "39.0%",
          format: { bold: true, align: "right", fill: "#e6f4ea" },
        },
        G8: {
          value: 2859760,
          formula: "=SUM(G2:G7)",
          formatted: "$2,859,760",
          format: { bold: true, align: "right", fill: "#e6f4ea" },
        },
      },
    },
    {
      id: "sheet-2",
      name: "Quarterly Summary",
      rowCount: 8,
      colCount: 5,
      columns: [
        { key: "A", label: "Quarter", width: 120, type: "string" },
        { key: "B", label: "Budget Target", width: 140, type: "currency" },
        { key: "C", label: "Actual Revenue", width: 140, type: "currency" },
        { key: "D", label: "Variance ($)", width: 130, type: "currency" },
        { key: "E", label: "Pacing %", width: 110, type: "number" },
      ],
      cells: {
        A1: { value: "Quarter", format: { bold: true, fill: "#f2f4f8" } },
        B1: {
          value: "Budget Target",
          format: { bold: true, align: "right", fill: "#f2f4f8" },
        },
        C1: {
          value: "Actual Revenue",
          format: { bold: true, align: "right", fill: "#f2f4f8" },
        },
        D1: {
          value: "Variance ($)",
          format: { bold: true, align: "right", fill: "#f2f4f8" },
        },
        E1: {
          value: "Pacing %",
          format: { bold: true, align: "right", fill: "#f2f4f8" },
        },

        A2: { value: "Q1" },
        B2: {
          value: 1600000,
          formatted: "$1,600,000",
          format: { align: "right" },
        },
        C2: {
          value: 1740000,
          formatted: "$1,740,000",
          format: { align: "right" },
        },
        D2: {
          value: 140000,
          formula: "=C2-B2",
          formatted: "+$140,000",
          format: { align: "right" },
        },
        E2: {
          value: 1.087,
          formula: "=C2/B2",
          formatted: "108.7%",
          format: { align: "right" },
        },

        A3: { value: "Q2" },
        B3: {
          value: 1700000,
          formatted: "$1,700,000",
          format: { align: "right" },
        },
        C3: {
          value: 1810000,
          formatted: "$1,810,000",
          format: { align: "right" },
        },
        D3: {
          value: 110000,
          formula: "=C3-B3",
          formatted: "+$110,000",
          format: { align: "right" },
        },
        E3: {
          value: 1.064,
          formula: "=C3/B3",
          formatted: "106.4%",
          format: { align: "right" },
        },

        A4: { value: "Total FY26", format: { bold: true, fill: "#e6f4ea" } },
        B4: {
          value: 3300000,
          formula: "=SUM(B2:B3)",
          formatted: "$3,300,000",
          format: { bold: true, align: "right", fill: "#e6f4ea" },
        },
        C4: {
          value: 3550000,
          formula: "=SUM(C2:C3)",
          formatted: "$3,550,000",
          format: { bold: true, align: "right", fill: "#e6f4ea" },
        },
        D4: {
          value: 250000,
          formula: "=C4-B4",
          formatted: "+$250,000",
          format: { bold: true, align: "right", fill: "#e6f4ea" },
        },
        E4: {
          value: 1.075,
          formula: "=C4/B4",
          formatted: "107.5%",
          format: { bold: true, align: "right", fill: "#e6f4ea" },
        },
      },
    },
  ],
  formulasAudit: [
    {
      cellCoord: "E2",
      formula: "=C2*D2",
      evaluatedValue: "$1,740,000",
      dependencies: ["C2", "D2"],
    },
    {
      cellCoord: "G2",
      formula: "=E2*F2",
      evaluatedValue: "$730,800",
      dependencies: ["E2", "F2"],
    },
    {
      cellCoord: "E8",
      formula: "=SUM(E2:E7)",
      evaluatedValue: "$7,281,000",
      dependencies: ["E2:E7"],
    },
    {
      cellCoord: "G8",
      formula: "=SUM(G2:G7)",
      evaluatedValue: "$2,859,760",
      dependencies: ["G2:G7"],
    },
    {
      cellCoord: "C8",
      formula: "=SUM(C2:C7)",
      evaluatedValue: "11,200",
      dependencies: ["C2:C7"],
    },
    {
      cellCoord: "D8",
      formula: "=AVERAGE(D2:D7)",
      evaluatedValue: "$820",
      dependencies: ["D2:D7"],
    },
  ],
};

export const samplePptDocument: PptDocumentModel = {
  id: "doc-ppt-01",
  title: "Product_Launch_Keynote.pptx",
  activeSlideId: "slide-1",
  themeName: "Midnight Executive",
  accentColor: "#c43e1c",
  slides: [
    {
      id: "slide-1",
      slideNumber: 1,
      title: "Antigravity Synthesizer 2026",
      layout: "title",
      background: "linear-gradient(135deg, #181b22 0%, #222734 100%)",
      notes:
        "Welcome everyone to the annual 2026 Developer & Knowledge Worker Summit. Pause for applause.",
      transition: "fade",
      objects: [
        {
          id: "s1-tag",
          kind: "shape",
          shapeKind: "badge",
          x: 80,
          y: 60,
          width: 200,
          height: 32,
          text: "KEYNOTE PRESENTATION",
          fill: "rgba(196, 62, 28, 0.25)",
          stroke: "#c43e1c",
          strokeWidth: 1,
          borderRadius: 16,
          color: "#ff8566",
          fontSize: 11,
          fontWeight: "bold",
          align: "center",
          zIndex: 1,
        },
        {
          id: "s1-title",
          kind: "text",
          x: 80,
          y: 110,
          width: 720,
          height: 90,
          text: "Next-Generation Multi-Format Document Studio",
          fill: "transparent",
          color: "#ffffff",
          fontSize: 34,
          fontWeight: "bold",
          zIndex: 2,
        },
        {
          id: "s1-subtitle",
          kind: "text",
          x: 80,
          y: 215,
          width: 680,
          height: 60,
          text: "A unified VS Code environment seamlessly integrating Word, Excel, and PowerPoint files with dynamic sidebar accordions and bottom search tools.",
          fill: "transparent",
          color: "#a0aec0",
          fontSize: 16,
          zIndex: 3,
        },
        {
          id: "s1-card1",
          kind: "metric",
          shapeKind: "card",
          x: 80,
          y: 310,
          width: 220,
          height: 100,
          metricValue: "10x Faster",
          metricLabel: "Document Workflow Switching",
          fill: "rgba(255, 255, 255, 0.05)",
          stroke: "rgba(255, 255, 255, 0.1)",
          strokeWidth: 1,
          borderRadius: 8,
          color: "#ffffff",
          fontSize: 14,
          zIndex: 4,
        },
        {
          id: "s1-card2",
          kind: "metric",
          shapeKind: "card",
          x: 320,
          y: 310,
          width: 220,
          height: 100,
          metricValue: "100% Client",
          metricLabel: "Zero Latency Node Module",
          fill: "rgba(255, 255, 255, 0.05)",
          stroke: "rgba(255, 255, 255, 0.1)",
          strokeWidth: 1,
          borderRadius: 8,
          color: "#ffffff",
          fontSize: 14,
          zIndex: 5,
        },
        {
          id: "s1-card3",
          kind: "metric",
          shapeKind: "card",
          x: 560,
          y: 310,
          width: 220,
          height: 100,
          metricValue: "Adaptive",
          metricLabel: "Dynamic Left Accordions",
          fill: "rgba(255, 255, 255, 0.05)",
          stroke: "rgba(255, 255, 255, 0.1)",
          strokeWidth: 1,
          borderRadius: 8,
          color: "#ffffff",
          fontSize: 14,
          zIndex: 6,
        },
      ],
    },
    {
      id: "slide-2",
      slideNumber: 2,
      title: "Adaptive Left Accordion Mechanics",
      layout: "content",
      background: "linear-gradient(135deg, #161a23 0%, #1f232e 100%)",
      notes:
        "Highlight how the left sidebar automatically morphs when the user clicks between tabs.",
      transition: "slide",
      objects: [
        {
          id: "s2-header",
          kind: "text",
          x: 60,
          y: 40,
          width: 700,
          height: 50,
          text: "Dynamic Left Accordion Architecture",
          fill: "transparent",
          color: "#ffffff",
          fontSize: 26,
          fontWeight: "bold",
          zIndex: 1,
        },
        {
          id: "s2-card-word",
          kind: "shape",
          shapeKind: "card",
          x: 60,
          y: 110,
          width: 220,
          height: 280,
          title: "📄 Word Mode",
          subtitle:
            "• Document Outline Tree\n• Page Thumbnails & Breaks\n• Review Comments & Notes\n• Word/Reading Statistics",
          fill: "rgba(43, 87, 154, 0.15)",
          stroke: "#2b579a",
          strokeWidth: 1.5,
          borderRadius: 8,
          color: "#90cdf4",
          fontSize: 13,
          zIndex: 2,
        },
        {
          id: "s2-card-excel",
          kind: "shape",
          shapeKind: "card",
          x: 310,
          y: 110,
          width: 220,
          height: 280,
          title: "📊 Excel Mode",
          subtitle:
            "• Worksheet Switcher\n• Formulas & Function Audit\n• Columns & Pivot Fields\n• Quick Sort & Filter",
          fill: "rgba(16, 124, 65, 0.15)",
          stroke: "#107c41",
          strokeWidth: 1.5,
          borderRadius: 8,
          color: "#9ae6b4",
          fontSize: 13,
          zIndex: 3,
        },
        {
          id: "s2-card-ppt",
          kind: "shape",
          shapeKind: "card",
          x: 560,
          y: 110,
          width: 220,
          height: 280,
          title: "🖼️ PowerPoint Mode",
          subtitle:
            "• Slide Filmstrip Navigator\n• Objects & Layer Hierarchy\n• Slide Layout Templates\n• Presenter Speaker Notes",
          fill: "rgba(196, 62, 28, 0.15)",
          stroke: "#c43e1c",
          strokeWidth: 1.5,
          borderRadius: 8,
          color: "#feb2b2",
          fontSize: 13,
          zIndex: 4,
        },
      ],
    },
    {
      id: "slide-3",
      slideNumber: 3,
      title: "Find Tools & Secondary Inspector",
      layout: "two-column",
      background: "linear-gradient(135deg, #1b1c24 0%, #262732 100%)",
      notes:
        "Demonstrate the bottom drawer Find in Files with regex and the right inspector.",
      transition: "fade",
      objects: [
        {
          id: "s3-header",
          kind: "text",
          x: 60,
          y: 40,
          width: 700,
          height: 50,
          text: "Comprehensive Bottom Tools & Right View",
          fill: "transparent",
          color: "#ffffff",
          fontSize: 26,
          fontWeight: "bold",
          zIndex: 1,
        },
        {
          id: "s3-box-bottom",
          kind: "shape",
          shapeKind: "card",
          x: 60,
          y: 120,
          width: 350,
          height: 260,
          title: "🔍 Bottom Panel Tools",
          subtitle:
            "• Find in Files: Deep search across docx text, xlsx cells, and pptx slides\n• Interactive Shell Terminal\n• Problems & Lint Diagnostics\n• Parser & Build Output Logs\n• Formula Debug Console",
          fill: "rgba(0, 122, 204, 0.12)",
          stroke: "#007acc",
          strokeWidth: 1,
          borderRadius: 8,
          color: "#cbd5e0",
          fontSize: 13,
          zIndex: 2,
        },
        {
          id: "s3-box-right",
          kind: "shape",
          shapeKind: "card",
          x: 440,
          y: 120,
          width: 350,
          height: 260,
          title: "🛠️ Right Secondary Sidebar",
          subtitle:
            "• Dynamic Property Inspector (Word typography, Excel number formatting, PPT shapes)\n• Copilot AI Assistant Chat\n• Document Minimap Overview\n• Metadata & File Statistics",
          fill: "rgba(128, 90, 213, 0.12)",
          stroke: "#805ad5",
          strokeWidth: 1,
          borderRadius: 8,
          color: "#e2e8f0",
          fontSize: 13,
          zIndex: 3,
        },
      ],
    },
    {
      id: "slide-4",
      slideNumber: 4,
      title: "Rollout Roadmap 2026-2027",
      layout: "dashboard",
      background: "linear-gradient(135deg, #181e28 0%, #202738 100%)",
      notes:
        "Timeline from prototype to open source NPM package and enterprise pilots.",
      transition: "zoom",
      objects: [
        {
          id: "s4-title",
          kind: "text",
          x: 60,
          y: 40,
          width: 700,
          height: 40,
          text: "Delivery Roadmap & Milestones",
          fill: "transparent",
          color: "#ffffff",
          fontSize: 26,
          fontWeight: "bold",
          zIndex: 1,
        },
        {
          id: "s4-m1",
          kind: "shape",
          shapeKind: "card",
          x: 60,
          y: 110,
          width: 220,
          height: 120,
          title: "Phase 1: Synthesizer Core",
          subtitle:
            "Dynamic accordions, VS Code shell, tabbed editors, SheetJS integration.",
          fill: "rgba(255, 255, 255, 0.05)",
          stroke: "#48bb78",
          strokeWidth: 2,
          borderRadius: 8,
          color: "#cbd5e0",
          fontSize: 12,
          zIndex: 2,
        },
        {
          id: "s4-m2",
          kind: "shape",
          shapeKind: "card",
          x: 310,
          y: 110,
          width: 220,
          height: 120,
          title: "Phase 2: Universal Search",
          subtitle:
            "Cross-document regex search, replace in files, jump-to-location.",
          fill: "rgba(255, 255, 255, 0.05)",
          stroke: "#4299e1",
          strokeWidth: 2,
          borderRadius: 8,
          color: "#cbd5e0",
          fontSize: 12,
          zIndex: 3,
        },
        {
          id: "s4-m3",
          kind: "shape",
          shapeKind: "card",
          x: 560,
          y: 110,
          width: 220,
          height: 120,
          title: "Phase 3: Node Module GA",
          subtitle:
            "NPM library packaging, Electron desktop builds, Copilot integrations.",
          fill: "rgba(255, 255, 255, 0.05)",
          stroke: "#ed8936",
          strokeWidth: 2,
          borderRadius: 8,
          color: "#cbd5e0",
          fontSize: 12,
          zIndex: 4,
        },
      ],
    },
  ],
};

export const sampleCodeDocument: CodeDocumentModel = {
  id: "doc-code-01",
  title: "office_synthesizer.ts",
  language: "typescript",
  content: `/**
 * VS Code Office Studio - Document Synthesizer Kernel
 * Coordinates multi-file document indexing, search queries, and dynamic view mapping.
 */

export interface SynthesizerConfig {
  enableWordParsing: boolean;
  enableExcelFormulas: boolean;
  enablePptxLayers: boolean;
  maxFileSizeMB: number;
}

export class OfficeSynthesizer {
  private config: SynthesizerConfig;
  private loadedDocuments: Map<string, any> = new Map();

  constructor(config: Partial<SynthesizerConfig> = {}) {
    this.config = {
      enableWordParsing: true,
      enableExcelFormulas: true,
      enablePptxLayers: true,
      maxFileSizeMB: 100,
      ...config,
    };
  }

  public registerDocument(id: string, doc: any): void {
    this.loadedDocuments.set(id, doc);
  }

  public searchAcrossDocuments(query: string, regex: boolean = false): any[] {
    const results: any[] = [];
    for (const [id, doc] of this.loadedDocuments.entries()) {
      // Cross-document search across Word, Excel, and PPT
    }
    return results;
  }
}
`,
  symbols: [
    { name: "SynthesizerConfig", kind: "interface", line: 6 },
    { name: "OfficeSynthesizer", kind: "class", line: 13 },
    { name: "constructor", kind: "function", line: 17 },
    { name: "registerDocument", kind: "function", line: 27 },
    { name: "searchAcrossDocuments", kind: "function", line: 31 },
  ],
};
