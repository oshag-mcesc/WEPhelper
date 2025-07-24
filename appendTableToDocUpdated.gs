/**
 * Appends tables to Google Documents with formatted text.
 * @param {Array<Array<string>>} info Data array (titles in first row, doc ID in first element of subsequent rows).
 * @return {Object} Result object with `done` (true/false) and optional `info`.
 * @customfunction
 */
function appendTable1(info) {
  const result = { done: false };

  try {
    const [titles, ...rows] = info;
    const numTitles = titles.length - 1;

    rows.forEach(rowData => {
      const docId = rowData[0];
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();

      // Create cells array using a for loop (more efficient here)
      const cells = [];
      for (let j = 0; j < numTitles; j++) {
        cells[j] = [titles[j + 1], rowData[j + 1]];
      }

      const tbl = body.appendTable(cells);
      tbl.setColumnWidth(0, 75);

      // Style title column
      for (let k = 0; k < tbl.getNumRows(); k++) {
        tbl.getCell(k, 0).setBackgroundColor("#E0E0E0");
      }

      // Format cells with newlines
      for (let k = 0; k < tbl.getNumRows(); k++) {
        const cell = tbl.getCell(k, 1);
        const text = cell.getText();

        if (typeof text === 'string' && text.includes('\n')) {
          const lines = text.split('\n');
          let formattedText = lines.length > 0 ? lines[0] : "";

          if (lines.length > 1) {
            formattedText += "\n";
            formattedText += lines.slice(1).map(line => `• ${line}\n`).join("");
          }
          cell.setText(formattedText);

          const paragraph = cell.getChild(0).asParagraph();
          if (paragraph && paragraph.getNumChildren() > 0) {
            paragraph.getChild(0).asText().setItalic(true);
          }
        }
      }

      doc.saveAndClose();
    });

    result.done = true;
  } catch (err) {
    console.error("Error:", err);
    result.info = err;
  }

  return result;
}