📌 Automated Attendance Mailer using Google Sheets
This project automates attendance approval and email notification workflows using Google Sheets and Google Apps Script. It's designed for clubs, student groups, or teams that track member contributions and want to notify class teachers or supervisors efficiently.

✨ Features
✅ Approval-based attendance logging

📬 Automated email notifications to designated recipients after approval

📊 Combines multiple attendance entries into a single summarized email

📅 Supports multiple dates and corresponding time logs per entry

🔐 Access-controlled approval (only specific users can mark entries as approved)

🧾 Simple Google Sheets-based interface (no third-party tools required)

💡 How It Works
Members fill out attendance with their Name, Date(s), and Time Worked.

Approvers (e.g., domain heads, team leads) mark entries as YES in the "Approved (YES/NO)" column.

The script:

Detects newly approved entries.

Groups multiple records per member.

Sends a single formatted email with all date/time logs.

Marks those entries as "Email Sent" to avoid duplication.

🧾 Sheet Structure
Column Name	Purpose
Member Name	Name of the student or team member
Date	Date(s) of contribution (comma-separated)
Time Worked	Corresponding time durations (comma-separated)
Approved (YES/NO)	Approval from authorized person
Class Teacher Email	Email of recipient (e.g., class teacher)
Email Sent	Marked Yes once email is sent

📂 Files Included
Code.gs – Google Apps Script logic (copy-paste into Script Editor)
README.md – Project documentation


📧 Email Format Example

Dear Sir/Madam,

This is to inform you that <b>John Doe</b> has contributed time toward club activities.

Below are the attendance details:

<table>
  <tr><th>Date</th><th>Time Worked</th></tr>
  <tr><td>25 July</td><td>2 hrs</td></tr>
  <tr><td>26 July</td><td>1.5 hrs</td></tr>
</table>

We kindly request you to consider this for attendance purposes.

Thank you,  
Team
🛡 Access Control Tips
Use Google Sheets' Protected Range to restrict editing access to the "Approved" column.


