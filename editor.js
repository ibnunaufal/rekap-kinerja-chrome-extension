let period2025 = [
  {
    id: 1,
    title: "Januari 2025",
    start: "2025-01-01",
    end: "2025-01-15",
    text: "01 Januari - 15 Januari 2025",
  },
  {
    id: 2,
    title: "Januari - Februari 2025",
    start: "2025-01-16",
    end: "2025-02-15",
    text: "16 Januari - 15 Februari 2025",
  },
  {
    id: 3,
    title: "Februari - Maret 2025",
    start: "2025-02-16",
    end: "2025-03-15",
    text: "16 Februari - 15 Maret 2025",
  },
  {
    id: 4,
    title: "Maret - April 2025",
    start: "2025-03-16",
    end: "2025-04-15",
    text: "16 Maret - 15 April 2025",
  },
  {
    id: 5,
    title: "April - Mei 2025",
    start: "2025-04-16",
    end: "2025-05-15",
    text: "16 April - 15 Mei 2025",
  },
  {
    id: 6,
    title: "Mei - Juni 2025",
    start: "2025-05-16",
    end: "2025-06-15",
    text: "16 Mei - 15 Juni 2025",
  },
  {
    id: 7,
    title: "Juni - Juli 2025",
    start: "2025-06-16",
    end: "2025-07-15",
    text: "16 Juni - 15 Juli 2025",
  },
  {
    id: 8,
    title: "Juli - Agustus 2025",
    start: "2025-07-16",
    end: "2025-08-15",
    text: "16 Juli - 15 Agustus 2025",
  },
  {
    id: 9,
    title: "Agustus - September 2025",
    start: "2025-08-16",
    end: "2025-09-15",
    text: "16 Agustus - 15 September 2025",
  },
  {
    id: 10,
    title: "September - Oktober 2025",
    start: "2025-09-16",
    end: "2025-10-15",
    text: "16 September - 15 Oktober 2025",
  },
  {
    id: 11,
    title: "Oktober - November 2025",
    start: "2025-10-16",
    end: "2025-11-15",
    text: "16 Oktober - 15 November 2025",
  },
  {
    id: 12,
    title: "November - Desember 2025",
    start: "2025-11-16",
    end: "2025-12-15",
    text: "16 November - 15 Desember 2025",
  },
  {
    id: 13,
    title: "Desember 2025",
    start: "2025-12-16",
    end: "2025-12-31",
    text: "16 Desember - 31 Desember 2025",
  },
];

// Populate month select
const monthSelect = document.getElementById("month-select");
const months = [
  "Januari",
  "Februari",
  "Maret",
  "April",
  "Mei",
  "Juni",
  "Juli",
  "Agustus",
  "September",
  "Oktober",
  "November",
  "Desember",
];

period2025.forEach((month, index) => {
  const option = document.createElement("option");
  option.value = index + 1;
  option.textContent = month.text;
  monthSelect.appendChild(option);
});

// Populate week select
const weekSelect = document.getElementById("week-select");
for (let i = 1; i <= 53; i++) {
  const option = document.createElement("option");
  option.value = i;
  option.textContent = `Minggu ${i}`;
  weekSelect.appendChild(option);
}

// Period type toggle
const periodType = document.getElementById("period-type");
const monthlyPeriod = document.getElementById("monthly-period");
const weeklyPeriod = document.getElementById("weekly-period");

function togglePeriod() {
  if (periodType.value === "monthly") {
    monthlyPeriod.classList.remove("hidden");
    weeklyPeriod.classList.add("hidden");
  } else {
    monthlyPeriod.classList.add("hidden");
    weeklyPeriod.classList.remove("hidden");
  }
}

// Initialize on page load
togglePeriod();

// Listen for changes
periodType.addEventListener("change", togglePeriod);

// Export button handler
document.getElementById("export-docx").addEventListener("click", function () {
  console.log("Export button clicked");
  const judul = document.getElementById("judul-dokumen").value;
  const nama = document.getElementById("nama").value;
  const nip = document.getElementById("nip").value;
  const unitKerja = document.getElementById("unit-kerja").value;
  const periode = periodType.value;

  let periodeValue;
  if (periode === "monthly") {
    const monthIndex = document.getElementById("month-select").value;
    periodeValue = months[monthIndex - 1];
  } else {
    periodeValue = `Minggu ${document.getElementById("week-select").value}`;
  }

  try {
    window.localStorage.setItem(
      "rekapKinerjaData",
      JSON.stringify({
        judul,
        nama,
        nip,
        unitKerja,
        periode,
        periodeValue,
      })
    );
  } catch (e) {
    console.error("Gagal menyimpan data ke localStorage:", e);
  }

  alert(
    `Ekspor DOCX:\n\nJudul: ${judul}\nNama: ${nama}\nNIP: ${nip}\nUnit Kerja: ${unitKerja}\nPeriode: ${periodeValue}`
  );
});

// Load saved data
function loadSavedData() {
  const savedData = JSON.parse(window.localStorage.getItem("rekapKinerjaData"));
  const savedKinerjaPage = window.localStorage.getItem("savedKinerjaPage");
  console.log("Loaded savedKinerjaPage from localStorage:", savedKinerjaPage);
  if (savedData) {
    document.getElementById("judul-dokumen").value = savedData.judul || "";
    document.getElementById("nama").value = savedData.nama || "";
    document.getElementById("nip").value = savedData.nip || "";
    document.getElementById("unit-kerja").value = savedData.unitKerja || "";
    document.getElementById("period-type").value =
      savedData.periode || "monthly";
    togglePeriod();
    if (savedData.periode === "monthly" && savedData.periodeValue) {
      const monthIndex = months.indexOf(savedData.periodeValue) + 1;
      document.getElementById("month-select").value = monthIndex || "";
    } else if (savedData.periode === "weekly" && savedData.periodeValue) {
      const weekNumber = parseInt(
        savedData.periodeValue.replace("Minggu ", ""),
        10
      );
      document.getElementById("week-select").value = weekNumber || "";
    }
  }
}
loadSavedData();
