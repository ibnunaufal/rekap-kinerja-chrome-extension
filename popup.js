document.getElementById("capture-docx").addEventListener("click", () => {
  chrome.tabs.query({ active: true, currentWindow: true }, async (tabs) => {
    const tabId = tabs[0].id;
    try {
      // First, inject the docx library
      await chrome.scripting.executeScript({
        target: { tabId: tabId },
        files: ["libs/docx.js"],
      });

      // Then execute your function
      await chrome.scripting.executeScript({
        target: { tabId: tabId },
        function: captureAndDownload,
        args: ["docx"],
      });
    } catch (error) {
      console.error("Error:", error);
      alert(`Failed to capture. Check console for details. ${error.message}`);
    }
  });
});

document.getElementById("capture-docx-coba").addEventListener("click", () => {
  console.log("Generating DOCX...", docx);
});

document.getElementById("capture-pdf").addEventListener("click", () => {
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    chrome.scripting.executeScript({
      target: { tabId: tabs[0].id },
      function: captureAndDownload,
      args: ["pdf"],
    });
  });
});

console.log("Popup script loaded.");
// Function runs inside page context
async function captureAndDownload(format) {
  console.log("captureAndDownload called with format:", format);
  // Helper functions must be defined inside this function because
  // `chrome.scripting.executeScript({ function: captureAndDownload })`
  // serializes only this function's source when running in the page
  // context. Defining the helpers here ensures they exist where the
  // script runs and avoids ReferenceError for undefined helpers.
  function giveMeWeekdaysList(weekNumber, year = new Date().getFullYear()) {
    const simple = new Date(year, 0, 1 + (weekNumber - 2) * 7);
    const dayOfWeek = simple.getDay();
    const ISOweekStart = new Date(simple);
    const diff = dayOfWeek <= 1 ? 1 - dayOfWeek : 8 - dayOfWeek;
    ISOweekStart.setDate(simple.getDate() + diff);
    const format = (date) =>
      `${String(date.getDate()).padStart(2, "0")}-${String(
        date.getMonth() + 1
      ).padStart(2, "0")}-${date.getFullYear()}`;
    let days = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(ISOweekStart);
      d.setDate(ISOweekStart.getDate() + i);
      days.push(format(d));
    }
    return days;
  }

  function extractWeekNumber(text) {
    // Matches "Minggu 46" → captures 46
    const match = text.match(/Minggu\s+(\d+)/i);
    return match ? parseInt(match[1], 10) : null;
  }
  // week 46
  // list of date which was on week 46, 46 will be dynamic value that can be inputted by user
  try {
    const {
      Document,
      Packer,
      Paragraph,
      Table,
      TableRow,
      TableCell,
      TextRun,
      WidthType,
      AlignmentType,
    } = docx;
    const btnEl = document.querySelector(".vuecal__title button");
    const text = btnEl.innerText; // "Minggu 46 (November 2025)"
    const currentWeek = extractWeekNumber(text);

    /**
     * let data = [{date: 'dd-mm-yyyy', events: ["event1", "event2"]}]
     */
    var data = [];

    var weekDates = giveMeWeekdaysList(currentWeek); // it should has format dd-mm-yyyy
    console.log(`Week Dates (${currentWeek}):`, weekDates);

    var nameHrefElement = document.querySelector(".nav-link");
    // get the div element inside nameHrefElement that has class "d-none d-md-block d-lg-inline-block", then get the innerText and get the text after "Hi, "
    var name = nameHrefElement.querySelector(".d-none.d-md-block.d-lg-inline-block").innerText.replace("Hi, ", "").trim();
    console.log("Captured Name:", name);
    // capitialize each word in name
    name = String(name).toLowerCase();
    name = name
      .split(" ")
      .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
      .join(" ");

    console.log("Capitalized Name:", name);

    const cellHasEvents = document.querySelectorAll(
      ".vuecal__cell.vuecal__cell--has-events"
    );
    cellHasEvents.forEach((cell, index) => {
      data.push({
        date: weekDates[index],
        //   capture all item that has class vuecal__event-title inside the cell
        activities: Array.from(
          cell.querySelectorAll(".vuecal__event-title")
        ).map((eventEl) => eventEl.innerText),
      });
    });

    console.log("Captured Data:", data);

    const table = new Table({
  width: {
    size: 100,
    type: WidthType.PERCENTAGE,
  },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({ 
                  text: "No", 
                  bold: true, 
                  size: 24, 
                  font: "Arial" 
                })
              ],
            }),
          ],
          margins: {
            top: 100,
            bottom: 100,
            left: 100,
            right: 100,
          },
        }),
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({ 
                  text: "Tanggal", 
                  bold: true, 
                  size: 24, 
                  font: "Arial" 
                })
              ],
            }),
          ],
          margins: {
            top: 100,
            bottom: 100,
            left: 100,
            right: 100,
          },
        }),
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Kegiatan yang dilakukan",
                  bold: true,
                  size: 24,
                  font: "Arial",
                }),
              ],
            }),
          ],
          margins: {
            top: 100,
            bottom: 100,
            left: 100,
            right: 100,
          },
        }),
      ],
    }),
    ...data.map(
      (entry, index) =>
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ 
                      text: String(index + 1), 
                      size: 24, 
                      font: "Arial" 
                    })
                  ],
                }),
              ],
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ 
                      text: entry.date, 
                      size: 24, 
                      font: "Arial" 
                    })
                  ],
                }),
              ],
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: entry.activities.map(
                (activity) =>
                  new Paragraph({
                    children: [
                      new TextRun({ 
                        text: `- ${activity}`, 
                        size: 24, 
                        font: "Arial" 
                      })
                    ],
                  })
              ),
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
            }),
          ],
        })
    ),
  ],
});

    console.log("Constructed Table:", table);
    // Build a simple document and include the table
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Laporan Kinerja",
                  bold: true,
                  size: 28, // You can adjust this for the title
                  font: "Arial",
                }),
              ],
              alignment: AlignmentType.CENTER, // Center the title
              spacing: {
                after: 240, // One line spacing after (240 twips = 1 line at 12pt)
              },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Nama\t\t: ",
                  bold: false,
                  size: 24, // 12pt in Word = 24 half-points
                  font: "Arial",
                }),
                new TextRun({
                  text: name,
                  size: 24,
                  font: "Arial",
                }),
              ],
              spacing: {
                after: 100,
              },
            }),
            // NIP field
            new Paragraph({
              children: [
                new TextRun({
                  text: "NIP\t\t: ",
                  bold: false,
                  size: 24,
                  font: "Arial",
                }),
              ],
              spacing: {
                after: 100,
              },
            }),
            // Tanggal field
            new Paragraph({
              children: [
                new TextRun({
                  text: "Tanggal\t: ",
                  bold: false,
                  size: 24,
                  font: "Arial",
                }),
                new TextRun({
                  text: `${weekDates[0].replaceAll("-","/")} - ${weekDates[weekDates.length - 1].replaceAll("-","/")}`,
                  size: 24,
                  font: "Arial",
                }),
              ],
              spacing: {
                after: 100,
              },
            }),
            // Unit field
            new Paragraph({
              children: [
                new TextRun({
                  text: "Unit\t\t: ",
                  bold: false,
                  size: 24,
                  font: "Arial",
                }),
              ],
              spacing: {
                after: 200,
              },
            }),
            table,
          ],
        },
      ],
    });

    try {
      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      let firstAndLastDate =
        weekDates[0] + "_to_" + weekDates[weekDates.length - 1];
      a.download = `kinerja_${firstAndLastDate}.docx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Failed to generate docx", err);
      alert("Error generating document. See console for details.");
    }
  } catch (error) {
    console.error("Error capturing week dates or events:", error);
  }

  return;
}
