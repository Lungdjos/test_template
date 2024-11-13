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
      if (paragraph.style === "Heading 1" || paragraph.style === "Title" || paragraph.style === "Subtle Reference") {
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
    docBody.insertParagraph("", Word.InsertLocation.end).style = titleStyle;
    docBody.insertParagraph("", Word.InsertLocation.end).style = titleStyle;
    docBody.insertParagraph("Procurement of Goods", Word.InsertLocation.end).style = titleStyle;
    docBody.insertParagraph("", Word.InsertLocation.end).style = titleStyle;
    docBody.insertParagraph("Open National Bidding", Word.InsertLocation.end).style = titleStyle;
    docBody.insertParagraph("", Word.InsertLocation.end).style = titleStyle;
    docBody.insertParagraph("", Word.InsertLocation.end).style = titleStyle;
    docBody.insertParagraph(formatDateToMonthYear(new Date()), Word.InsertLocation.end).style = titleStyle;
    docBody.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    // Insert "Foreword" section
    docBody.insertParagraph("Foreword", Word.InsertLocation.end).style = boldStyle;
    // Insert the Foreword content line by line
    docBody.insertParagraph(
      "These Bidding Documents for Procurement of Goods have been prepared by the Zambia Public Procurement Authority to be used for the procurement of goods through Open National Bidding (ONB) in projects that are financed in whole or in part by the Government of the Republic of Zambia.",
      Word.InsertLocation.end
    ).style = normalStyle;

    docBody.insertParagraph(
      "These Standard Bidding Documents are based on the Master Bidding Documents for Procurement of Goods and User’s Guide, prepared by the Multilateral Development Banks and International Financing Institutions, while they are customised to be consistent with the Public Procurement Act No. 12 of 2008 of the Laws of Zambia and the Public Procurement Regulations, Statutory Instrument No. 63 of 2011. The Master Bidding Documents reflect “international best practices”.",
      Word.InsertLocation.end
    ).style = normalStyle;

    docBody.insertParagraph(
      "These Bidding Documents for Procurement of Goods assume that no prequalification has taken place before bidding.",
      Word.InsertLocation.end
    ).style = normalStyle;

    docBody.insertParagraph(
      "Those wishing to submit comments or questions on these Bidding Documents or to obtain additional information on procurement in Zambia projects are encouraged to contact:",
      Word.InsertLocation.end
    ).style = normalStyle;

    // Insert the centered contact information for "The Director General" without quote style
    let contactInfo = [
      "The Director General",
      "Zambia Public Procurement Authority",
      "Red Cross House, P.O. Box 31009",
      "Plot 2837, Los Angeles Boulevard",
      "Longacres, Lusaka",
      "ZAMBIA",
      "http://www.ppa.org.zm"
    ];

    contactInfo.forEach(line => {
      let paragraph = docBody.insertParagraph(line, Word.InsertLocation.end);
      paragraph.style = "Subtle Reference"; // Use normal style instead of quote
    });

    // Page break for the next section
    docBody.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    // Insert "SBD for Procurement of Goods" section
    docBody.insertParagraph("SBD for Procurement of Goods", Word.InsertLocation.end).style = boldStyle;
    docBody.insertParagraph("Summary", Word.InsertLocation.end).style = italicStyle;
    docBody.insertParagraph("PART 1 – BIDDING PROCEDURES", Word.InsertLocation.end).style = italicStyle;

    // part 1 of contents for bidding procedures
    const part1Contents = [
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
    part1Contents.forEach(section => {
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
    const sectionVI = docBody.insertParagraph("Section VI. Schedule of Requirements", Word.InsertLocation.end);
    sectionVI.style = boldItalicStyle;

    docBody.insertParagraph(
        "This Section includes the List of Goods and Related Services, the Delivery and Completion Schedules, the Technical Specifications and the Drawings that describe the Goods and Related Services to be procured.",
        Word.InsertLocation.end
    ).style = normalStyle;


    docBody.insertParagraph("PART 3 – EVALUATION CRITERIA", Word.InsertLocation.end).style = italicStyle;
    // part 3 of contents for bidding procedures
    const part3Contents = [
        {
            title: "Section VII. General Conditions of Contract (GCC)",
            description: "This Section includes the general clauses to be applied in all contracts.  The text of the clauses in this Section shall not be modified."
        },
        {
            title: "Section VIII.	Special Conditions of Contract (SCC)",
            description: "This Section includes clauses specific to each contract that modify or supplement Section VII, General Conditions of Contract."
        },
        {
            title: "Section IX:	Contract Forms",
            description: "This Section includes the form for the Agreement, which, once completed, incorporates corrections or modifications to the accepted bid that are permitted under the Instructions to Bidders, the General Conditions of Contract, and the Special Conditions of Contract.\u000D\u000AThe forms for Performance Security and Advance Payment Security, when required, shall only be completed by the successful Bidder after contract award."
        },
        {
            title: "Attachment:	 Invitation for Bids ",
            description: "An “Invitation for Bids” form is provided at the end of the Bidding Documents for information."
        }
    ];

    // Insert each section with title in bold, and description starting on the next line indented
    part3Contents.forEach(section => {
        // Insert the title for the section
          const sectionTitle = docBody.insertParagraph(section.title, Word.InsertLocation.end);
          sectionTitle.style = boldItalicStyle;

          // Insert each line of the description as a new paragraph
          const descriptionLines = section.description.split("\u000D\u000A");
          descriptionLines.forEach(line => {
            const descriptionParagraph = docBody.insertParagraph(line, Word.InsertLocation.end);
            descriptionParagraph.style = normalStyle; // Apply normal style for the description text
            descriptionParagraph.leftIndent = 40; // Adjust left indentation if needed
          });
    });
}

// Function to format the current date as "Month Year"
function formatDateToMonthYear(date) {
  const options = { year: "numeric", month: "long" };
  return date.toLocaleDateString("en-US", options);
}
