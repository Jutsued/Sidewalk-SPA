let input = document.querySelector("input");
const btnActive = document.querySelector(".btnActive");
const btnReset = document.querySelector(".btnReset");

let wordAbbreviations = {
  ".": "",
  "&": "and",
  "@": "",
  And: "and",
  At: "at",
  Ave: "Ave",
  AVE: "Ave",
  Avenue: "Ave",
  AVENUE: "Ave",
  Between: "b/t",
  Boulevard: "Blvd",
  Drive: "Dr",
  Eastbound: "E/B",
  Expressway: "Expwy",
  East: "E",
  From: "",
  Lane: "Ln",
  North: "N",
  Northbound: "N/B",
  Parkway: "Pkwy",
  Place: "Pl",
  Road: "Rd",
  Roadway: "Rway",
  South: "S",
  Southbound: "S/B",
  Street: "St",
  Turnpike: "Tpke",
  Of: "of",
  West: "W",
};

btnActive.addEventListener("click", function () {
  let inputFile = document.getElementById("fileInput").files[0];
  if (inputFile) {
    formatText(inputFile);
  } else {
    alert("Please select a file.");
  }
});

async function formatText(inputFile) {
  let data = await readFile(inputFile);
  let workbook = await XlsxPopulate.fromDataAsync(data);
  let sheet = workbook.sheet(0);

  for (let i = 2; i <= sheet.usedRange().endCell().rowNumber(); i++) {
    let cell = sheet.cell(`D${i}`);
    let line = cell.value();

    if (line) {
      let ignorePatterns = [
        /^HWPR22QX\s+/i,
        /^HWPR22KR\s+/i,
        /^HWPR22MQ\s+/i,
        /^HWPR22K1\s+/i,
        /^HWPR21QX\s+/i,
        /^HWPR21KR\s+/i,
        /^HWPR21MQ\s+/i,
        /^HWPR21K1\s+/i,
        /^HWPR20Q2\s+/i,
        /^HWPR20K3\s+/i,
        /^HWPR20K2\s+/i,
        /^BLCOMPLNT\s+\d+/i,
        /^HWS\d+[QR]/,
        /^HWP\d+[QRX]/,
      ];

      if (ignorePatterns.some((pattern) => pattern.test(line))) {
        continue;
      }

      // Separate any word directly attached to a dash (e.g., "Street-" => "Street -")
      line = line.replace(/\./g, "");

      let words = line.split(/\s+/);

      let formattedWords = words.map((word, index) => {
        let nextWord = words[index + 1];

        if (word === "/") {
          return nextWord
            ? `/${nextWord.charAt(0).toUpperCase()}${nextWord
                .slice(1)
                .toLowerCase()}`
            : "/";
        } else if (index > 0 && words[index - 1] === "/") {
          return "";
        } else {
          const cleaned = word.replace(/^[(-]+|[.,)\-]+$/g, "");
          const leadingMatch = word.match(/^[(-]+/);
          const trailingMatch = word.match(/[.,)\-]+$/);

          let capitalizedCleaned =
            cleaned.charAt(0).toUpperCase() + cleaned.slice(1).toLowerCase();

          // Expand "Ave" to "Avenue" if followed by single letter
          if (
            capitalizedCleaned.toLowerCase() === "ave" &&
            nextWord &&
            /^[A-Z]$/i.test(nextWord)
          ) {
            return `${leadingMatch ? leadingMatch[0] : ""}Avenue${
              trailingMatch ? trailingMatch[0] : ""
            }`;
          }

          // Abbreviate "Avenue" to "Ave" only if not followed by single letter
          if (
            capitalizedCleaned.toLowerCase() === "avenue" &&
            (!nextWord || !/^[a-zA-Z]$/.test(nextWord))
          ) {
            return `${leadingMatch ? leadingMatch[0] : ""}Ave${
              trailingMatch ? trailingMatch[0] : ""
            }`;
          }

          const formatted =
            wordAbbreviations[capitalizedCleaned] || capitalizedCleaned;
          return `${leadingMatch ? leadingMatch[0] : ""}${formatted}${
            trailingMatch ? trailingMatch[0] : ""
          }`;
        }
      });

      let formattedLine = formattedWords.join(" ").replace(/\s+/g, " ").trim();
      formattedLine = formattedLine.replace(/\s*\/\s*/g, "/");

      // Normalize double dashes to a single spaced dash
      formattedLine = formattedLine.replace(/--+/g, " - ");

      // Ensure dashes are surrounded by spaces (fixes stuck-on dash cases)
      formattedLine = formattedLine.replace(/\s*-\s*/g, " - ");

      const specificEntries = [
        "ne",
        "nc",
        "se",
        "nw",
        "sw",
        "nb",
        "sec",
        "nec",
        "swc",
        "nwc",
        "usa",
        "ny",
        "nycha",
        "dpr",
        "doe",
      ];
      specificEntries.forEach((entry) => {
        const regex = new RegExp(`\\b${entry}\\b`, "gi");
        formattedLine = formattedLine.replace(regex, entry.toUpperCase());
      });

      cell.value(formattedLine);
    }
  }

  let updatedData = await workbook.outputAsync();
  let blob = new Blob([updatedData], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  let url = window.URL.createObjectURL(blob);
  let a = document.createElement("a");
  a.href = url;
  a.download = "updated_file.xlsx";
  a.click();
  const notification = document.getElementById("notificationOutput");
  if (notification) {
    notification.classList.add("show");
  }
  window.URL.revokeObjectURL(url);
}

function readFile(file) {
  return new Promise((resolve, reject) => {
    let reader = new FileReader();
    reader.onload = (event) => resolve(event.target.result);
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}

btnReset.addEventListener("click", () => {
  input.value = "";
  const notification = document.getElementById("notificationOutput");
  if (notification) {
    notification.classList.remove("show");
  }
});

