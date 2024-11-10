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

    // variables for styling
    const titleStyle = "Title";
    const boldStyle = "Heading 1";
    const italicStyle = "Heading 2";
    const boldItalicStyle = "Heading 3";


    const subBoldStyle = "Subheading 1";
    const subItalicStyle = "Subheading 2";
    const subBoldItalicStyle = "Subheading 3";

    const normalStyle = "Normal";

    // Insert the title page with centered text
    const title = docBody.insertParagraph("STANDARD BIDDING DOCUMENTS", Word.InsertLocation.start).style = titleStyle;


    // Insert additional paragraphs
    const titleParagraph = docBody.insertParagraph("Procurement of Goods", Word.InsertLocation.end).style = titleStyle;
    const openNationalBidding = docBody.insertParagraph("Open National Bidding", Word.InsertLocation.end).style = titleStyle;
  // Insert the current date
  docBody.insertParagraph(formatDateToMonthYear(new Date()), Word.InsertLocation.end).style = titleStyle;

    // Page break to start the next section on a new page
    docBody.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

  // Insert "Foreword" as a heading
  const foreword = docBody.insertParagraph("Foreword", Word.InsertLocation.end).style = boldStyle;

  // Insert the long paragraph for foreword with justified alignment
  docBody.insertParagraph(
    "These Bidding Documents for Procurement of Goods have been prepared by the Zambia Public Procurement Authority to be used for the procurement of goods through Open National Bidding (ONB) in projects that are financed in whole or in part by the Government of the Republic of Zambia.",
    Word.InsertLocation.end
  ).style = normalStyle;  // Justify the content

  docBody.insertParagraph(
    "These Standard Bidding Documents are based on the Master Bidding Documents for Procurement of Goods and User’s Guide, prepared by the Multilateral Development Banks and International Financing Institutions, while they are customised to be consistent with the Public Procurement Act No. 12 of 2008 of the Laws of Zambia and the Public Procurement Regulations, Statutory Instrument No. 63 of 2011. The Master Bidding Documents reflect “international best practices”.",
    Word.InsertLocation.end
  ).style = normalStyle;

  docBody.insertParagraph(
    "These Bidding Documents for Procurement of Goods assumes that no prequalification has taken place before bidding.",
    Word.InsertLocation.end
  ).style = normalStyle;

  docBody.insertParagraph(
    "Those wishing to submit comments or questions on these Bidding Documents or to obtain additional information on procurement in Zambia projects are encouraged to contact:",
    Word.InsertLocation.end
  ).style = normalStyle;
  docBody.insertParagraph("The Director General, Zambia Public Procurement Authority, Red Cross House, P.O. Box 31009, Plot 2837, Los Angeles Boulevard, Longacres, Lusaka, ZAMBIA, http://www.ppa.org.zm", Word.InsertLocation.end).style = "Quote";


  // Page break to start the next section on a new page
  docBody.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
  // Insert "SBD for Procurement of Goods" as a Heading 1 with center alignment
  const sbd = docBody.insertParagraph("SBD for Procurement of Goods", Word.InsertLocation.end);
  sbd.style = boldStyle;

  // Insert "Summary" as a Heading 2 with center alignment
  const summary = docBody.insertParagraph("Summary", Word.InsertLocation.end);
  summary.style = italicStyle;

  // Insert "PART 1 – BIDDING PROCEDURES"
  const part1 = docBody.insertParagraph("PART 1 – BIDDING PROCEDURES", Word.InsertLocation.end).style = italicStyle;

  // Define the table of contents for bidding procedures
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

  // Create the table of contents as a list with justified descriptions
    tableOfContents.forEach(section => {
      const sectionTitle = docBody.insertParagraph(section.title, Word.InsertLocation.end);
      sectionTitle.style = boldItalicStyle;
      sectionTitle.alignment = Word.Alignment.center;  // Center-align the section title

      const sectionDescription = docBody.insertParagraph(section.description, Word.InsertLocation.end);
    })

  // Insert "PART 2 – SUPPLY REQUIREMENTS"
  const part2 = docBody.insertParagraph("PART 2 – SUPPLY REQUIREMENTS", Word.InsertLocation.end);
  part2.style = italicStyle;

  // Insert "Section VI. Schedule of Requirements"
  const sectionVI = docBody.insertParagraph("Section I. Schedule of Requirements", Word.InsertLocation.end);
  sectionVI.style = boldItalicStyle;

  // Insert the section description for "Schedule of Requirements"
  docBody.insertParagraph(
    "This Section includes the List of Goods and Related Services, the Delivery and Completion Schedules, the Technical Specifications and the Drawings that describe the Goods and Related Services to be procured.",
    Word.InsertLocation.end
  );  // Justify the section description

  // Insert "PART 3 – EVALUATION CRITERIA"
  const part3 = docBody.insertParagraph("PART 3 – EVALUATION CRITERIA", Word.InsertLocation.end);
  part3.style = italicStyle;

  // Insert "Section VII. Evaluation Criteria"
  const sectionVII = docBody.insertParagraph("Section I. Evaluation Criteria", Word.InsertLocation.end);
  sectionVII.style = boldItalicStyle;

  // Insert the section description for "Evaluation Criteria"
  docBody.insertParagraph(
    "This Section includes the Evaluation Criteria, which are used to determine the best-evaluated bid, and the Bidder’s qualification requirements to perform the contract.",
    Word.InsertLocation.end
  );  // Justify the section description

  // Insert "PART 4 – BID SECURITY"
  const part4 = docBody.insertParagraph("PART 4 – BID SECURITY", Word.InsertLocation.end);
  part4.style = italicStyle;

  // Insert "Section VIII. Bid Security"
  const sectionVIII = docBody.insertParagraph("Section I. Bid Security", Word.InsertLocation.end);
  sectionVIII.style = boldItalicStyle;

  // Insert the section description for "Bid Security"
  docBody.insertParagraph(
    "This Section includes the Bid Security, which is required to be submitted with the Bid.",
    Word.InsertLocation.end
  );  // Justify the section description
}

// Function to format the current date as "Month Year"
function formatDateToMonthYear(date) {
  const options = { year: "numeric", month: "long" };
  return date.toLocaleDateString("en-US", options);
}
