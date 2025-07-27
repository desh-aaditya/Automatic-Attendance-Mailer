function onEdit(e) {
  const sheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();

  if (sheet.getName() !== "Attendance" || editedRow === 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const approvedCol = headers.indexOf("Approved (YES/NO)") + 1;
  const emailSentCol = headers.indexOf("Email Sent") + 1;
  const classTeacherCol = headers.indexOf("Class Teacher Email") + 1;
  const nameCol = headers.indexOf("Member Name") + 1;
  const dateCol = headers.indexOf("Date") + 1;
  const timeCol = headers.indexOf("Time Worked") + 1;

  if (editedCol !== approvedCol) return;

  const approvedValue = String(sheet.getRange(editedRow, approvedCol).getValue()).trim().toLowerCase();
  if (approvedValue !== "yes") return;

  const memberName = sheet.getRange(editedRow, nameCol).getDisplayValue();

  const data = sheet.getDataRange().getValues();
  const matchingRows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameCol - 1];
    const approved = String(row[approvedCol - 1]).toLowerCase();
    const emailSent = row[emailSentCol - 1];

    if (name === memberName && approved === "yes" && emailSent !== true) {
      matchingRows.push({
        rowIndex: i + 1,
        dateStr: sheet.getRange(i + 1, dateCol).getDisplayValue(),
        timeStr: sheet.getRange(i + 1, timeCol).getDisplayValue(),
        classTeacherEmail: row[classTeacherCol - 1]
      });
    }
  }

  if (matchingRows.length === 0) return;

  const classTeacherEmail = matchingRows[0].classTeacherEmail;
  let tableRows = "";

  for (const row of matchingRows) {
    const dates = row.dateStr.split(",").map(d => d.trim());
    const times = row.timeStr.split(",").map(t => t.trim());

    const pairCount = Math.min(dates.length, times.length);

    for (let i = 0; i < pairCount; i++) {
      tableRows += `
        <tr>
          <td style="border: 1px solid #ccc; padding: 8px;">${dates[i]}</td>
          <td style="border: 1px solid #ccc; padding: 8px;">${times[i]}</td>
        </tr>`;
    }
  }

  const emailBody = `
    <p>Dear Sir/Madam,</p>
    <p>This is to inform you that <b>${memberName}</b> has contributed time toward club activities as part of the MLSC team.</p>
    <p>Below are the attendance details:</p>
    <table style="border-collapse: collapse; width: 100%;">
      <thead>
        <tr>
          <th style="border: 1px solid #ccc; padding: 8px;">Date</th>
          <th style="border: 1px solid #ccc; padding: 8px;">Time Worked</th>
        </tr>
      </thead>
      <tbody>
        ${tableRows}
      </tbody>
    </table>
    <p>These entries have been reviewed and approved by the respective domain head
    We kindly request you to consider this as a formal attendance request</p>
    <p>Thank you,<br>MLSC Team</p>
  `;

  const subject = `Attendance Approved for ${memberName}`;
  GmailApp.sendEmail(classTeacherEmail, subject, '', { htmlBody: emailBody });

  for (const row of matchingRows) {
    sheet.getRange(row.rowIndex, emailSentCol).setValue(true);
  }
}
