document.getElementById("export-to-docx").addEventListener("click", () => {
  chrome.tabs.query({ active: true, currentWindow: true }, async (tabs) => {
    const tabId = tabs[0].id;
    try {
      const currentUrl = tabs[0].url;
      console.log("Current URL:", currentUrl);
      if (
        currentUrl.includes("chrome://extensions") ||
        currentUrl.includes("chrome://newtab")
      ) {
        alert(
          "Mohon tidak mengakses Ekstensi dari halaman tab Baru (chrome://newtab) atau Ekstensi (chrome://extensions)."
        );
        return;
      }
      // First, inject the docx library
      await chrome.scripting.executeScript({
        target: { tabId: tabId },
        files: ["libs/docx.js"],
      });

      // ================ Get Identity ==================
      const includeIdentity =
        document.getElementById("toggle-identity").checked;
      let identityData = {
        name: "",
        nip: "",
        position: "",
        unit: "",
      };
      if (includeIdentity) {
        const name = document.getElementById("nama").value.trim();
        const nip = document.getElementById("nip").value.trim();
        const position = document.getElementById("jabatan").value.trim();
        const unit = document.getElementById("unit-kerja").value.trim();
        identityData = { name, nip, position, unit };
        console.log("Identity Data to inject:", identityData);
        await chrome.storage.local.set({
          identityData: JSON.stringify(identityData),
        });
      }

      // ================ Get Signature Option ==================
      const includeSignature =
        document.getElementById("toggle-signature").checked;
      const signatureCount = document.getElementById("signature-count").value;
      let signatureData = {};
      if (includeSignature) {
        console.log("Signature Data to inject:", signatureData);
        signatureData = {
          cityName: document.getElementById("city-name").value.trim(),
          type: signatureCount, //or "double"
          useAnchor: document.getElementById("toggle-anchor").checked, //or false
          signer:
            signatureCount === "single"
              ? // single signer
                [
                  {
                    name: document
                      .getElementById("signer1-name-single")
                      .value.trim(),
                    nip: document
                      .getElementById("signer1-nip-single")
                      .value.trim(),
                    title: document
                      .getElementById("signer1-title-single")
                      .value.trim(),
                    anchor: document
                      .getElementById("signer1-anchor-single")
                      .value.trim(),
                  },
                ]
              : // double signer
                [
                  {
                    name: document
                      .getElementById("signer1-name-double")
                      .value.trim(),
                    nip: document
                      .getElementById("signer1-nip-double")
                      .value.trim(),
                    title: document
                      .getElementById("signer1-title-double")
                      .value.trim(),
                    anchor: document
                      .getElementById("signer1-anchor-double")
                      .value.trim(),
                  },
                  {
                    name: document
                      .getElementById("signer2-name-double")
                      .value.trim(),
                    nip: document
                      .getElementById("signer2-nip-double")
                      .value.trim(),
                    title: document
                      .getElementById("signer2-title-double")
                      .value.trim(),
                    anchor: document
                      .getElementById("signer2-anchor-double")
                      .value.trim(),
                  },
                ],
        };
        await chrome.storage.local.set({
          signatureData: JSON.stringify(signatureData),
        });
      }
      else {
        await chrome.storage.local.set({
          signatureData: JSON.stringify({}),
        });
      }

      // ================ Get savedKinerjaPage from chrome.storage.local ==================

      const savedKinerjaPage = await chrome.storage.local
        .get("savedKinerjaPage")
        .then((result) => result.savedKinerjaPage);
      if (!savedKinerjaPage) {
        alert("No saved eKinerja data found in local storage.");
        return;
      }
      const parsedData = JSON.parse(savedKinerjaPage);
      console.log("Parsed savedKinerjaPage for export:", parsedData);

      // get selected dates from checkboxes
      const selectedDates = [];
      parsedData.forEach((entry) => {
        const checkbox = document.getElementById(
          `week-${entry.week}-date-${entry.date}`
        );
        console.log(
          `week-${entry.week}-date-${entry.date} checkbox:`,
          checkbox
        );
        if (checkbox && checkbox.checked) {
          selectedDates.push(entry);
        }
      });

      // sort selectedDates by selectedDates.date ascending
      selectedDates.sort((a, b) => {
        const [aYear, aMonth, aDay] = a.date.split("-").map(Number);
        const [bYear, bMonth, bDay] = b.date.split("-").map(Number);
        const aDate = new Date(aYear, aMonth - 1, aDay);
        const bDate = new Date(bYear, bMonth - 1, bDay);
        return aDate - bDate;
      });

      // check the completeness of identityData if includeIdentity is true
      // also check the completeness of signatureData if includeSignature is true
      if (includeIdentity) {
        if (
          !identityData.name ||
          !identityData.nip ||
          !identityData.position ||
          !identityData.unit
        ) {
          alert("Mohon lengkapi data identitas sebelum mengekspor.");
          return;
        }
      }
      if (includeSignature) {
        for (const signer of signatureData.signer) {
          if (!signer.name || !signer.nip || !signer.title) {
            alert("Mohon lengkapi data tanda tangan sebelum mengekspor.");
            return;
          }
          if (signatureData.useAnchor && !signer.anchor) {
            alert(
              "Mohon lengkapi data anchor tanda tangan sebelum mengekspor."
            );
            return;
          }
        }
      }

      // if signatureData is empty object, set to null
      if (signatureData && Object.keys(signatureData).length === 0) {
        signatureData = null;
      }
      // Then execute your function
      await chrome.scripting.executeScript({
        target: { tabId: tabId },
        function: exportToFile,
        args: ["docx", selectedDates, identityData, signatureData],
      });
    } catch (error) {
      console.error("Error:", error);
      alert(`Failed to export. Check console for details. ${error.message}`);
    }
  });
});

document.getElementById("toggle-help").addEventListener("click", () => {
  const instruction = document.getElementById("instruction");
  instruction.classList.toggle("visible");
  const isVisible = instruction.classList.contains("visible");
  document.getElementById("toggle-help").innerText = isVisible
    ? "Sembunyikan"
    : "Lihat";
});

document.getElementById("toggle-identity").addEventListener("change", (e) => {
  const identityInputs = document.getElementById("identity-inputs");
  identityInputs.classList.toggle("visible", e.target.checked);
});

document.getElementById("toggle-signature").addEventListener("change", (e) => {
  const signatureInputs = document.getElementById("signature-inputs");
  signatureInputs.classList.toggle("visible", e.target.checked);
});

document.getElementById("toggle-anchor").addEventListener("change", (e) => {
  const anchorContainers = [
    document.getElementById("signer1-anchor-double-container"),
    document.getElementById("signer2-anchor-double-container"),
    document.getElementById("signer1-anchor-single-container"),
  ];
  anchorContainers.forEach((anchorContainer) => {
    anchorContainer.classList.toggle("show-anchor", e.target.checked);
  });
  const anchorInputs = [
    document.getElementById("signer1-anchor-double"),
    document.getElementById("signer2-anchor-double"),
    document.getElementById("signer1-anchor-single"),
  ];
  if (!e.target.checked) {
    // clear all anchor input values
    anchorInputs.forEach((anchorInput) => {
      anchorInput.value = "";
    });
  }
});

document.getElementById("signature-count").addEventListener("change", (e) => {
  // if value is "single"
  const singleSignerContainer = document.getElementById("single-signer");
  const doubleSignerContainer = document.getElementById("double-signer");

  if (e.target.value === "single") {
    singleSignerContainer.style.display = "block";
    doubleSignerContainer.style.display = "none";
  } else {
    singleSignerContainer.style.display = "none";
    doubleSignerContainer.style.display = "block";
  }
});

document.getElementById("save-current").addEventListener("click", () => {
  chrome.tabs.query({ active: true, currentWindow: true }, async (tabs) => {
    const tabId = tabs[0].id;
    try {
      const currentUrl = tabs[0].url;
      console.log("Current URL:", currentUrl);
      // First, inject the docx library
      // await chrome.scripting.executeScript({
      //   target: { tabId: tabId },
      //   files: ["libs/docx.js"],
      // });

      // Then execute your function
      await chrome.scripting.executeScript({
        target: { tabId: tabId },
        function: save,
        args: [currentUrl],
      });
    } catch (error) {
      console.error("Error:", error);
      alert(`Failed to capture. Check console for details. ${error.message}`);
    }
  });
});

async function saveCurrentPageInLocalStorage(currentUrl) {
  console.log("Current URL inside saveCurrentPageInLocalStorage:", currentUrl);
  if (!currentUrl.includes("kinerja.bkn.go.id/kinerja_harian")) {
    alert("Silakan buka halaman eKinerja pada tab Mingguan terlebih dahulu.");
    return;
  }
  try {
    localStorage.setItem("savedKinerjaPage", "test");
    alert(
      "Halaman eKinerja berhasil disimpan di penyimpanan lokal browser Anda."
    );
  } catch (error) {
    console.error("Error saving page content:", error);
    alert("Gagal menyimpan halaman. Periksa konsol untuk detailnya.");
  }
}

async function save(currentUrl) {
  console.log("Current URL inside save:", currentUrl);

  if (!currentUrl.includes("kinerja.bkn.go.id/kinerja_harian")) {
    alert("Silakan buka halaman eKinerja pada tab Mingguan terlebih dahulu.");
    return;
  }
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
    console.log("Saving current week's data...");
    const btnEl = document.querySelector(".vuecal__title button");
    const text = btnEl.innerText; // "Minggu 46 (November 2025)"
    const currentWeek = extractWeekNumber(text);

    console.log(`Current Week Number: ${currentWeek}`);
    var data = [];

    // get chrome local storage
    const savedKinerjaPage = await chrome.storage.local
      .get("savedKinerjaPage")
      .then((result) => result.savedKinerjaPage);

    var localData = localStorage.getItem("savedKinerjaPage");
    var parsedLocalData = savedKinerjaPage ? JSON.parse(savedKinerjaPage) : [];

    console.log("Parsed Local Data:", parsedLocalData);

    var weekDates = giveMeWeekdaysList(currentWeek); // it should has format dd-mm-yyyy
    console.log(`Week Dates (${currentWeek}):`, weekDates);

    var nameHrefElement = document.querySelector(".nav-link");
    // get the div element inside nameHrefElement that has class "d-none d-md-block d-lg-inline-block", then get the innerText and get the text after "Hi, "
    var name = nameHrefElement
      .querySelector(".d-none.d-md-block.d-lg-inline-block")
      .innerText.replace("Hi, ", "")
      .trim();
    console.log("Captured Name:", name);
    // capitialize each word in name
    name = String(name).toLowerCase();
    name = name
      .split(" ")
      .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
      .join(" ");

    console.log("Capitalized Name:", name);

    // const allCells = document.querySelectorAll(".vuecal__cell");
    // allCells.forEach((cell, index) => {
    //   // show all class in the cell
    //   console.log(`Cell ${index} classes:`, cell.className);
    // });

    const cellHasEvents = document.querySelectorAll(
      ".vuecal__cell"
      // "vuecal__event-title"
    );
    cellHasEvents.forEach((cell, index) => {
      if (cell.classList.contains("vuecal__cell--has-events")) {
        data.push({
          date: weekDates[index],
          //   capture all item that has class vuecal__event-title inside the cell
          activities: Array.from(
            cell.querySelectorAll(".vuecal__event-title")
          ).map((eventEl) => eventEl.innerText),
        });
      }
    });

    var monthlyPeriod = document
      .querySelectorAll(".col-lg-8")[1]
      .querySelector("h6").innerText;

    console.log("Captured Monthly Period:", monthlyPeriod);

    console.log("Captured Data:", data);

    /**
     * {
        "week": 46,
        "year": 2023,
        "month": 11,
        "date": "2023-11-13",
        "tasks": [
            "Menyelesaikan fitur ekspor DOCX di ekstensi Chrome Rekap Kinerja",
            "Memperbarui antarmuka pengguna popup.html untuk menambahkan tombol ekspor dan navigasi"
        ]
    }
     */

    var newData = [];

    data.forEach((entry) => {
      const [day, month, year] = entry.date.split("-").map(Number);
      const formattedDate = `${year}-${String(month).padStart(2, "0")}-${String(
        day
      ).padStart(2, "0")}`;
      newData.push({
        week: currentWeek,
        year: year,
        month: month,
        monthlyPeriod: monthlyPeriod,
        date: formattedDate,
        tasks: entry.activities,
      });
    });

    console.log("New Data to be saved:", newData);
    // merge parsedLocalData with newData, but replace duplicate same date entry
    console.log("Previous data", parsedLocalData);
    var mergedData = [...parsedLocalData];
    console.log("mergedData before forEach", mergedData);
    newData.forEach((newEntry) => {
      const exists = mergedData.some(
        (existingEntry) => existingEntry.date === newEntry.date
      );
      if (!exists) {
        mergedData.push(newEntry);
      }
      // if exists, replace the tasks
      else {
        mergedData = mergedData.map((existingEntry) => {
          if (existingEntry.date === newEntry.date) {
            return {
              ...existingEntry,
              tasks: newEntry.tasks,
            };
          }
          return existingEntry;
        });
      }
    });

    console.log("Merged data", mergedData);

    // localStorage.setItem("savedKinerjaPage", JSON.stringify(mergedData));
    // save into chrome.storage.local as well
    await chrome.storage.local.set(
      { savedKinerjaPage: JSON.stringify(mergedData) },
      function () {
        console.log("Data saved to chrome.storage.local");
      }
    );

    alert(
      "Halaman eKinerja berhasil disimpan di penyimpanan lokal browser Anda."
    );
  } catch (error) {
    console.error("Error capturing week dates or events:", error);
    alert("Gagal menyimpan halaman. Periksa konsol untuk detailnya.");
  }
}

// export to file from localStorage
async function exportToFile(
  format,
  selectedDates,
  identityData,
  signatureData
) {
  console.log("exportToFile called with format:", format);
  // get from chrome.storage.local as well

  if (selectedDates.length === 0) {
    alert("Mohon pilih minimal satu tanggal untuk diekspor.");
    return;
  }
  console.log("Selected Dates for export:", selectedDates);
  console.log("Identity Data for export:", identityData);

  if (format === "docx") {
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
        BorderStyle,
      } = docx;

      const includeSignature = signatureData && signatureData.signer.length > 0;

      /** create two column table like this:
       *  if double signature:
       *  |                               | Semarang, [date now in DD MMM YYYY] |
       *  |Mengetahui,                    |                                     |
       *  |[signer1 title]                | [signer2 title]                     |
       *  |[anchor (if useAnchor)]        |                                     |
       *  |[signer1 name]                 | [signer2 name]                      |
       *  |[signer1 nip]                  | [signer2 nip]                       |
       *
       *   if single signature:
       *  |                               | Semarang, [date now in DD MMM YYYY] |
       *  |                               |[signer1 title]                      |
       *  |                               |[anchor (if useAnchor)]              |
       *  |                               |[signer1 name]                       |
       *  |                               |[signer1 nip]                        |
       */

      const FONT = "Arial";
      const FONT_SIZE = 24;
      console.log("declared NO_BORDER");
      const NO_BORDER = {
        top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      };
      const CELL_MARGINS = {
        top: 100,
        bottom: 100,
        left: 100,
        right: 100,
      };
      const CELL_MARGINS_ANCHOR = {
        top: 1200,
        bottom: 100,
        left: 100,
        right: 100,
      };

      console.log("declared tableBorders");
      const tableBorders = {
        top: NO_BORDER.top,
        bottom: NO_BORDER.bottom,
        left: NO_BORDER.left,
        right: NO_BORDER.right,
        insideHorizontal: NO_BORDER.top,
        insideVertical: NO_BORDER.left,
      };

      const textRun = (text = "") =>
        new TextRun({
          text,
          size: FONT_SIZE,
          font: FONT,
        });

      const paragraph = (text = "", align = AlignmentType.LEFT) =>
        new Paragraph({
          children: [textRun(text)],
          alignment: align,
        });

      console.log("declared cell");
      const cell = (text = "", align = AlignmentType.LEFT) =>
        new TableCell({
          width: {
            size: 50,
            type: WidthType.PERCENTAGE,
          },
          children: [paragraph(text, align)],
          margins: CELL_MARGINS,
          borders: NO_BORDER,
        });

      const row = (cells) =>
        new TableRow({
          children: cells.map((text) => cell(text)),
        });

      const cellAnchor = (text = "", align = AlignmentType.LEFT) =>
        new TableCell({
          width: {
            size: 50,
            type: WidthType.PERCENTAGE,
          },
          children: [paragraph(text, align)],
          margins: CELL_MARGINS_ANCHOR,
          borders: NO_BORDER,
        });
      const rowAnchor = (cells) =>
        new TableRow({
          children: cells.map((text) => cellAnchor(text)),
        });

      const formattedDate = new Date().toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      });

      console.log(
        "declared buildSingleSignatureTable and buildDoubleSignatureTable"
      );
      const buildSingleSignatureTable = (signatureData) =>
        new Table({
          borders: tableBorders,
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            row(["", `${signatureData.cityName}, ${formattedDate}`]),
            row(["", signatureData.signer[0].title]),
            rowAnchor(
              signatureData.useAnchor
                ? ["", signatureData.signer[0].anchor]
                : ["", ""]
            ),
            row(["", signatureData.signer[0].name]),
            row(["", `NIP. ${signatureData.signer[0].nip}`]),
          ],
        });

      const buildDoubleSignatureTable = (signatureData) =>
        new Table({
          borders: tableBorders,
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            row(["", `${signatureData.cityName}, ${formattedDate}`]),
            row(["Mengetahui,", ""]),
            row([signatureData.signer[0].title, signatureData.signer[1].title]),
            rowAnchor(
              signatureData.useAnchor
                ? [
                    signatureData.signer[0].anchor,
                    signatureData.signer[1].anchor,
                  ]
                : ["", ""]
            ),
            row([signatureData.signer[0].name, signatureData.signer[1].name]),
            row([
              `NIP. ${signatureData.signer[0].nip}`,
              `NIP. ${signatureData.signer[1].nip}`,
            ]),
          ],
        });

      let signatureTable = null;

      if (includeSignature) {
        signatureTable =
          signatureData.type === "single"
            ? buildSingleSignatureTable(signatureData)
            : buildDoubleSignatureTable(signatureData);
      }

      const contentTable = new Table({
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
              new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: "Tanggal",
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
          ...selectedDates.map(
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
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: entry.date,
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
                  new TableCell({
                    children: entry.tasks.map(
                      (activity) =>
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: `- ${activity}`,
                              size: 24,
                              font: "Arial",
                            }),
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
                    text: identityData.name || "",
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
                  new TextRun({
                    text: identityData.nip || "",
                    size: 24,
                    font: "Arial",
                  }),
                ],
                spacing: {
                  after: 100,
                },
              }),

              // Jabatan field
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Jabatan\t: ",
                    bold: false,
                    size: 24,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: identityData.position || "",
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
                  new TextRun({
                    text: identityData.unit || "",
                    size: 24,
                    font: "Arial",
                  }),
                ],
                spacing: {
                  after: 100,
                },
              }),
              contentTable,
              new Paragraph({}), // empty paragraph for spacing
              ...(includeSignature && signatureTable
                ? [signatureTable]
                : []), // add signature table if needed
            ],
          },
        ],
      });

      try {
        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        // current date time for filename
        const currentDateTime = new Date();
        const a = document.createElement("a");
        a.href = url;

        a.download = `kinerja_${currentDateTime.toISOString()}.docx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      } catch (err) {
        console.error("Failed to generate docx", err);
        alert("Error generating document. See console for details.");
      }
    } catch (error) {}
  }
}

// call when popup loaded
document.addEventListener("DOMContentLoaded", async function () {
  console.log("Popup loaded!");

  // const savedKinerjaPage = localStorage.getItem("savedKinerjaPage");
  // get from chrome.storage.local as well
  const container = document.getElementById("saved-data-container");
  const savedKinerjaPage = await chrome.storage.local
    .get("savedKinerjaPage")
    .then((result) => result.savedKinerjaPage);
  if (savedKinerjaPage) {
    console.log("Found savedKinerjaPage in localStorage.");
    // parse it
    const parsedData = JSON.parse(savedKinerjaPage);
    console.log("Parsed savedKinerjaPage:", parsedData);
    // you can use parsedData to pre-fill checkboxes or other UI elements
  } else {
    console.log("No savedKinerjaPage found in localStorage.");
    // you can show a message to user that no data found
    // center the text horizontally and vertically in the container
    container.style.height = "100px";
    container.style.textAlign = "center";
    container.style.display = "flex";
    container.style.justifyContent = "center";
    container.style.alignItems = "center";
    container.style.color = "red";
    container.style.fontSize = "14px";
    container.innerText =
      "Belum ada data eKinerja yang disimpan.\nSilahkan ikuti langkah 1 untuk menyimpan data.";
  }

  /*
  [ ] Pekan 47  | [ ] Pekan 48  | [ ] Pekan 49
  ______________________________________________________
  [ ] 17-10-2025| [ ] 24-10-2025| [ ] 01-12-2025|
  [ ] 18-10-2025| [ ] 25-10-2025| [ ] 02-12-2025|
  [ ] 19-10-2025| [ ] 26-10-2025| [ ] 03-12-2025|
  [ ] 20-10-2025| [ ] 27-10-2025| [ ] 04-12-2025|
  [ ] 21-10-2025| [ ] 28-10-2025| [ ] 05-12-2025|
   */

  // create checkboxes based on the savedKinerjaPage data into div with id="saved-data-container"

  if (savedKinerjaPage) {
    const parsedData = JSON.parse(savedKinerjaPage);
    // group by week number
    const groupedByWeek = {};
    parsedData.forEach((entry) => {
      if (!groupedByWeek[entry.week]) {
        groupedByWeek[entry.week] = [];
      }
      groupedByWeek[entry.week].push(entry);
    });
    console.log("Grouped Data by Week:", groupedByWeek);

    // create table
    // const table = document.createElement("table");
    // const tbody = document.createElement("tbody");
    // const weeks = Object.keys(groupedByWeek);
    // const maxDays = Math.max(
    //   ...weeks.map((week) => groupedByWeek[week].length)
    // );
    // for (let i = 0; i < maxDays; i++) {
    //   const tr = document.createElement("tr");
    //   weeks.forEach((week) => {
    //     const td = document.createElement("td");
    //     const dayEntry = groupedByWeek[week][i];
    //     if (dayEntry) {
    //       const checkbox = document.createElement("input");
    //       checkbox.type = "checkbox";
    //       checkbox.id = `week-${week}-date-${dayEntry.date}`;
    //       const label = document.createElement("label");
    //       label.htmlFor = checkbox.id;
    //       label.innerText = dayEntry.date;
    //       td.appendChild(checkbox);
    //       td.appendChild(label);
    //     }
    //     tr.appendChild(td);
    //   });
    //   tbody.appendChild(tr);
    // }
    // table.appendChild(tbody);

    // new format
    const table = document.createElement("table");
    table.style.textAlign = "left";
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    /**
     * <table>
        <thead>
          <tr>
            <th><input type="checkbox" id="week-47" /> <label htmlFor="week-47"> Pekan 47 </label></th>
            <th><input type="checkbox" id="week-48" /> <label htmlFor="week-48"> Pekan 48 </label></th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><input type="checkbox" id="week-47-date-17-10-2025" /> <label htmlFor="week-47-date-17-10-2025"> 17-10-2025 </label></td>
            <td><input type="checkbox" id="week-48-date-24-10-2025" /> <label htmlFor="week-48-date-24-10-2025"> 24-10-2025 </label></td>
          </tr>
          ...
        </tbody>
      </table>

      if cb week-47 is checked, all date under week 47 will be checked as well

     */

    const weeks = Object.keys(groupedByWeek);
    thead.appendChild(headerRow);
    weeks.forEach((week) => {
      const th = document.createElement("th");
      const weekCheckbox = document.createElement("input");
      weekCheckbox.type = "checkbox";
      weekCheckbox.id = `week-${week}`;
      weekCheckbox.addEventListener("change", () => {
        selectAllWeekDates(week);
      });
      const weekLabel = document.createElement("label");
      weekLabel.htmlFor = weekCheckbox.id;
      weekLabel.innerText = ` Pekan ${week} `;
      th.appendChild(weekCheckbox);
      th.appendChild(weekLabel);
      th.style.paddingRight = "50px";
      headerRow.appendChild(th);
    });

    const tbody = document.createElement("tbody");
    const maxDays = Math.max(
      ...weeks.map((week) => groupedByWeek[week].length)
    );
    for (let i = 0; i < maxDays; i++) {
      const tr = document.createElement("tr");
      weeks.forEach((week) => {
        const td = document.createElement("td");
        const dayEntry = groupedByWeek[week][i];
        if (dayEntry) {
          const checkbox = document.createElement("input");
          checkbox.type = "checkbox";
          checkbox.id = `week-${week}-date-${dayEntry.date}`;
          const label = document.createElement("label");
          label.htmlFor = checkbox.id;
          label.innerText = dayEntry.date;
          td.appendChild(checkbox);
          td.appendChild(label);
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    table.appendChild(thead);
    table.style.whiteSpace = "nowrap";

    const divOfTable = document.createElement("div");
    divOfTable.style.overflowX = "auto";
    divOfTable.style.width = "100%";
    divOfTable.appendChild(table);
    container.appendChild(divOfTable);

    // input identity data from chrome.storage.local if exists
    const identityData = await chrome.storage.local
      .get("identityData")
      .then((result) => JSON.parse(result.identityData));
    if (identityData) {
      console.log("Found identityData in chrome.storage.local:", identityData);
      document.getElementById("toggle-identity").checked = true;
      document.getElementById("identity-inputs").classList.add("visible");
      document.getElementById("nama").value = identityData.name || "";
      document.getElementById("nip").value = identityData.nip || "";
      document.getElementById("jabatan").value = identityData.position || "";
      document.getElementById("unit-kerja").value = identityData.unit || "";
    }

    const signatureData = await chrome.storage.local
      .get("signatureData")
      .then((result) => JSON.parse(result.signatureData));
    // check if signatureData != {}
    if (signatureData && Object.keys(signatureData).length > 0) {
      console.log(
        "Found signatureData in chrome.storage.local:",
        signatureData
      );
      document.getElementById("toggle-signature").checked = true;
      document.getElementById("signature-inputs").classList.add("visible");
      const typeSelect = document.getElementById("signature-count");
      const useAnchor = signatureData.useAnchor || false;
      const toggleAnchor = document.getElementById("toggle-anchor");
      typeSelect.value = signatureData.type || "single";
      // trigger change event to show/hide signer inputs
      typeSelect.dispatchEvent(new Event("change"));
      // fill city name and useAnchor
      document.getElementById("city-name").value = signatureData.cityName || "";
      toggleAnchor.checked = useAnchor;
      toggleAnchor.dispatchEvent(new Event("change"));
      // fill signer data
      if (signatureData.type === "single") {
        // document.getElementById("single-signer").style.display = "block";
        // document.getElementById("double-signer").style.display = "none";
        document.getElementById("signer1-title-single").value =
          signatureData.signer[0].title || "";
        document.getElementById("signer1-anchor-single").value =
          signatureData.signer[0].anchor || "";
        document.getElementById("signer1-name-single").value =
          signatureData.signer[0].name || "";
        document.getElementById("signer1-nip-single").value =
          signatureData.signer[0].nip || "";
      } else if (signatureData.type === "double") {
        // document.getElementById("single-signer").style.display = "none";
        // document.getElementById("double-signer").style.display = "block";
        document.getElementById("signer1-title-double").value =
          signatureData.signer[0].title || "";
        document.getElementById("signer1-anchor-double").value =
          signatureData.signer[0].anchor || "";
        document.getElementById("signer1-name-double").value =
          signatureData.signer[0].name || "";
        document.getElementById("signer1-nip-double").value =
          signatureData.signer[0].nip || "";
        document.getElementById("signer2-title-double").value =
          signatureData.signer[1].title || "";
        document.getElementById("signer2-anchor-double").value =
          signatureData.signer[1].anchor || "";
        document.getElementById("signer2-name-double").value =
          signatureData.signer[1].name || "";
        document.getElementById("signer2-nip-double").value =
          signatureData.signer[1].nip || "";
      }
    }
  }

  function selectAllWeekDates(week) {
    const weekCheckbox = document.getElementById(`week-${week}`);
    const isChecked = weekCheckbox.checked;
    const checkboxes = container.querySelectorAll(
      `input[id^="week-${week}-date-"]`
    );
    checkboxes.forEach((checkbox) => {
      checkbox.checked = isChecked;
    });
  }
});
