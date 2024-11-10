Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    const doc = context.document;
    const docBody = doc.body;

    // Generate additional content
    generateProcurementDocument(docBody);

    // Load all paragraphs in the document
    const paragraphs = docBody.paragraphs;
    paragraphs.load("items/style");     // Load paragraph styles

    await context.sync();

    // Center-align 'Title' or 'Heading' paragraphs; justify others
    paragraphs.items.forEach((paragraph) => {
      if (paragraph.style === "Heading 1" || paragraph.style === "Title" || paragraph.style === "Quote") {
        paragraph.alignment = Word.Alignment.centered;
      } else if(paragraph.style === "Heading 2" || paragraph.style === "Heading 3" || paragraph.style === "Subheading 1" || paragraph.style === "Subheading 2" || paragraph.style === "Subheading 3") {
      paragraph.alignment = Word.Alignment.left;
      } else {
        paragraph.alignment = Word.Alignment.justified;
      }
    });

    // Set document font color and type
    docBody.font.color = "black";
    docBody.font.name = "Times New Roman";

    await context.sync();
  });
}



export function generateProcurementDocument(docBody) {

    // Variables for styling
    const titleStyle = "Title";
    const boldStyle = "Heading 1";
    const italicStyle = "Heading 2";
    const normalStyle = "Normal";
    const boldItalicStyle = "Heading 3";

    // Insert the title page with centered text
    docBody.insertParagraph("STANDARD BIDDING DOCUMENTS", Word.InsertLocation.start).style = titleStyle;
    docBody.insertParagraph("Procurement of Goods", Word.InsertLocation.end).style = titleStyle;
    docBody.insertParagraph("Open National Bidding", Word.InsertLocation.end).style = titleStyle;
    docBody.insertParagraph(formatDateToMonthYear(new Date()), Word.InsertLocation.end).style = titleStyle;
    docBody.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    // Insert "Foreword" section
    docBody.insertParagraph("Foreword", Word.InsertLocation.end).style = boldStyle;
    docBody.insertParagraph(
        "These Bidding Documents for Procurement of Goods have been prepared by the Zambia Public Procurement Authority to be used for the procurement of goods through Open National Bidding (ONB) in projects that are financed in whole or in part by the Government of the Republic of Zambia.",
        Word.InsertLocation.end
    ).style = normalStyle;

    // Page break for the next section
    docBody.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    // Insert "SBD for Procurement of Goods" section
    docBody.insertParagraph("SBD for Procurement of Goods", Word.InsertLocation.end).style = boldStyle;
    docBody.insertParagraph("Summary", Word.InsertLocation.end).style = italicStyle;
    docBody.insertParagraph("PART 1 – BIDDING PROCEDURES", Word.InsertLocation.end).style = italicStyle;

    // Table of contents for bidding procedures
    const tableOfContents = [
        {
            title: "Section I. Instructions to Bidders (ITB)",
            description: "This Section provides information to help Bidders prepare their bids. Information is also provided on the submission, opening, and evaluation of bids and on the award of Contracts. Section I contains provisions that are to be used without modification."
        },
        {
            title: "Section II. Bidding Data Sheet (BDS)",
            description: "This Section includes provisions that are specific to each procurement and that supplement Section I, Instructions to Bidders."
        },
        {
            title: "Section III. Evaluation and Qualification Criteria",
            description: "This Section specifies the criteria to be used to determine the best-evaluated bid, and the Bidder’s qualification requirements to perform the contract."
        },
        {
            title: "Section IV. Bidding Forms",
            description: "This Section includes the forms for the Bid Submission, Price Schedules, and Bid Security to be submitted with the Bid."
        },
        {
            title: "Section V. Eligible Countries",
            description: "This Section contains information regarding eligible countries."
        }
    ];

    // Insert each section with title in bold, and description starting on the next line indented
    tableOfContents.forEach(section => {
        // Insert section title with bold formatting
        const sectionTitle = docBody.insertParagraph(section.title, Word.InsertLocation.end);
        sectionTitle.style = boldItalicStyle;

        // Insert description on the next line, indented
        const sectionDescription = docBody.insertParagraph(section.description, Word.InsertLocation.end);
        sectionDescription.style = normalStyle;
        sectionDescription.leftIndent = 40;  // Indent the description to align with the section number's end
    });

    // Continue with other parts as per your layout
    docBody.insertParagraph("PART 2 – SUPPLY REQUIREMENTS", Word.InsertLocation.end).style = italicStyle;
    const sectionVI = docBody.insertParagraph("Section I. Schedule of Requirements", Word.InsertLocation.end);
    sectionVI.style = boldItalicStyle;

    docBody.insertParagraph(
        "This Section includes the List of Goods and Related Services, the Delivery and Completion Schedules, the Technical Specifications and the Drawings that describe the Goods and Related Services to be procured.",
        Word.InsertLocation.end
    ).style = normalStyle;

    docBody.insertParagraph("PART 3 – EVALUATION CRITERIA", Word.InsertLocation.end).style = italicStyle;
    const sectionVII = docBody.insertParagraph("Section I. Evaluation Criteria", Word.InsertLocation.end);
    sectionVII.style = boldItalicStyle;

    docBody.insertParagraph(
        "This Section includes the Evaluation Criteria, which are used to determine the best-evaluated bid, and the Bidder’s qualification requirements to perform the contract.",
        Word.InsertLocation.end
    ).style = normalStyle;

    docBody.insertParagraph("PART 4 – BID SECURITY", Word.InsertLocation.end).style = italicStyle;
    const sectionVIII = docBody.insertParagraph("Section I. Bid Security", Word.InsertLocation.end);
    sectionVIII.style = boldItalicStyle;

    docBody.insertParagraph(
        "This Section includes the Bid Security, which is required to be submitted with the Bid.",
        Word.InsertLocation.end
    ).style = normalStyle;
}

// Function to format the current date as "Month Year"
function formatDateToMonthYear(date) {
  const options = { year: "numeric", month: "long" };
  return date.toLocaleDateString("en-US", options);
}
