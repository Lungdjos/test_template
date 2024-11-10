/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

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
//    doc_header = doc.head;
    const doc_body = doc.body;
//    const doc_footer = doc.foot;

    // insert the title page.
    const title = doc_body.insertParagraph("STANDARD BIDDING DOCUMENTS", Word.InsertLocation.start);
    title.bold = true;
    title.style = "Title";
    title.alignment = Word.Alignment.center;

    generateProcurementDocument(doc_body);
    // change the paragraph color to blue.
    doc_body.font.color = "black";

    await context.sync();
  });
}

export function generateProcurementDocument(docBody) {

    // Insert the title "Procurement of Goods"
    docBody.insertParagraph("Procurement of Goods", Word.InsertLocation.end).bold = true;

    // Insert "Open National Bidding"
    docBody.insertParagraph("Open National Bidding", Word.InsertLocation.end).bold = true;

    // Insert the date
    docBody.insertParagraph(formatDateToMonthYear(new Date()), Word.InsertLocation.end);

    // Insert "Foreword" as a heading
    docBody.insertParagraph("Foreword", Word.InsertLocation.end).style = "Heading 1";

    // Insert the long paragraph for foreword
    docBody.insertParagraph(
      "These Bidding Documents for Procurement of Goods have been prepared by the Zambia Public Procurement Authority to be used for the procurement of goods through Open National Bidding (ONB) in projects that are financed in whole or in part by the Government of the Republic of Zambia.\n\nThese Standard Bidding Documents are based on the Master Bidding Documents for Procurement of Goods and User’s Guide, prepared by the Multilateral Development Banks and International Financing Institutions, while they are customised to be consistent with the Public Procurement Act No. 12 of 2008 of the Laws of Zambia and the Public Procurement Regulations, Statutory Instrument No. 63 of 2011. The Master Bidding Documents reflect “international best practices”.\n\nThese Bidding Documents for Procurement of Goods, assumes that no prequalification has taken place before bidding.\n\nThose wishing to submit comments or questions on these Bidding Documents or to obtain additional information on procurement in Zambia projects are encouraged to contact:\n\nThe Director General\nZambia Public Procurement Authority\nRed Cross House, P.O. Box 31009\nPlot 2837, Los Angeles Boulevard\nLongacres, Lusaka\nZAMBIA\nhttp://www.ppa.org.zm",
      Word.InsertLocation.end
    );

    // Insert the heading for the table of contents "SBD for Procurement of Goods"
    docBody.insertParagraph("SBD for Procurement of Goods", Word.InsertLocation.end).style = "Heading 1";

    // Insert the section title "Summary"
    docBody.insertParagraph("Summary", Word.InsertLocation.end).style = "Heading 2";

    // Insert "PART 1 – BIDDING PROCEDURES"
    docBody.insertParagraph("PART 1 – BIDDING PROCEDURES", Word.InsertLocation.end).style = "Heading 2";

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

    // Create the table of contents as a list
    tableOfContents.forEach(section => {
      docBody.insertParagraph(section.title, Word.InsertLocation.end).style = "Heading 3";
      docBody.insertParagraph(section.description, Word.InsertLocation.end);
    });

    // Insert "PART 2 – SUPPLY REQUIREMENTS"
    docBody.insertParagraph("PART 2 – SUPPLY REQUIREMENTS", Word.InsertLocation.end).style = "Heading 2";

    // Insert "Section VI. Schedule of Requirements"
    docBody.insertParagraph("Section I. Schedule of Requirements", Word.InsertLocation.end).style = "Heading 3";

    // Insert the section description for "Schedule of Requirements"
    docBody.insertParagraph(
      "This Section includes the List of Goods and Related Services, the Delivery and Completion Schedules, the Technical Specifications and the Drawings that describe the Goods and Related Services to be procured.",
      Word.InsertLocation.end
    );

    // Insert "PART 3 – EVALUATION CRITERIA"
    docBody.insertParagraph("PART 3 – EVALUATION CRITERIA", Word.InsertLocation.end).style = "Heading 2";

    // Insert "Section VII. Evaluation Criteria"
    docBody.insertParagraph("Section I. Evaluation Criteria", Word.InsertLocation.end).style = "Heading 3";

    // Insert the section description for "Evaluation Criteria"
    docBody.insertParagraph(
      "This Section includes the Evaluation Criteria, which are used to determine the best-evaluated bid, and the Bidder’s qualification requirements to perform the contract.",
      Word.InsertLocation.end
    );

    // Insert "PART 4 – BID SECURITY"
    docBody.insertParagraph("PART 4 – BID SECURITY", Word.InsertLocation.end).style = "Heading 2";

    // Insert "Section VIII. Bid Security"
    docBody.insertParagraph("Section I. Bid Security", Word.InsertLocation.end).style = "Heading 3";

    // Insert the section description for "Bid Security"
    docBody.insertParagraph(
      "This Section includes the Bid Security, which is required to be submitted with the Bid.",
      Word.InsertLocation.end
    );
}

// Function to format a date as "Month Year"
function formatDateToMonthYear(date) {
  const options = { year: "numeric", month: "long" };
  return date.toLocaleDateString("en-US", options);
}
