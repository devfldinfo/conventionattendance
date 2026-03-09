
const SHEET_ID = '1OukBzLWlDurzU5uvrEmaNXpwBoIc9jbR6RmeIg2M2TI';
const SHEET_NAME = 'Feedback';

function stressTest() {
  for (let i = 0; i < 50; i++) {
    submitData({
      Surname: "Test",
      Name: "User " + i,
      Age: "30",
      Meeting: "Test Meeting",
      Email: "Mariusmarais2008@gmai.com"
    });
  }
}


function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Convention Registration').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function submitData(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // wait up to 30 seconds

  try {

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  sheet.appendRow([
    new Date(),
    data['Surname'] || '',
    data['Name'] || '',
    data['Age'] || '',
    data['Sleeping on grounds'] || '',
    data['Using a caravan'] || '',
    data['Attending Thursday'] || '',
    data['Attending Friday'] || '',
    data['Attending Saturday'] || '',
    data['Attending Sunday'] || '',
    data['Name on working list'] || '',
    data['Job preferance'] || '',
    data['Saturday preps help'] || '',
    data['Gender'] || '',
    data['Staying for evening meals'] || '',
    data['Convention'] || '',
    data['Special requests'] || '',
    data['Meeting'] || '',
    data['Email'] || ''
  ]);

  } finally {
    lock.releaseLock();
  }
}

function buildEmail(d) {

  function esc(v) {
    return String(v || '').replace(/[&<>"']/g, s => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[s]));
  }

  const row = (label, value) =>
    `<tr>
       <td style="padding:8px;border-bottom:1px solid #ddd;"><strong>${esc(label)}</strong></td>
       <td style="padding:8px;border-bottom:1px solid #ddd;">${esc(value) || '-'}</td>
     </tr>`;

  const rowHTML = (label, value) =>
    `<tr>
     <td style="padding:8px;border-bottom:1px solid #ddd;"><strong>${esc(label)}</strong></td>
     <td style="padding:8px;border-bottom:1px solid #ddd;">${value}</td>
    </tr>`;
   
  return `
  <div style="text-align:center; font-family: Arial, sans-serif;">
    <table align="center" style="border-collapse:collapse;width:100%;max-width:600px;margin:auto;">
      <tr>
        <td style="padding:20px;">
          <h2 style="margin-top:0;">
            Thank you for completing your convention attendance form.
          </h2>

          <p>
            Here is a copy of what you have submitted. If you find that you have made a mistake while filling in your attendance form or if you would like to change any of the information or if you will not be attending anymore, please notify the responsible brother (See contact details below).
          </p>

          <table style="border-collapse:collapse;width:100%;margin-top:20px;">
            ${row('Surname', d['Surname'])}
            ${row('Name', d['Name'])}
            ${row('Age', d['Age'])}
            ${row('Sleeping on grounds', d['Sleeping on grounds'])}
            ${row('Using a caravan', d['Using a caravan'])}
            ${row('Attending Thursday', d['Attending Thursday'])}
            ${row('Attending Friday', d['Attending Friday'])}
            ${row('Attending Saturday', d['Attending Saturday'])}
            ${row('Attending Sunday', d['Attending Sunday'])}
            ${row('Name on working list', d['Name on working list'])}
            ${row('Job preferance', d['Job preferance'])}
            ${row('Saturday preps help', d['Saturday preps help'])}
            ${row('Gender', d['Gender'])}
            ${row('Staying for evening meals', d['Staying for evening meals'])}
            ${row('Convention', d['Convention'])}
            ${row('Special requests', d['Special requests'])}
            ${row('Meeting name', d['Meeting'])}
            ${row('Email', d['Email'])}
          </table>

<!-- Responsible Brothers Section -->
<h3 style="margin-top:40px;">Responsible Brothers</h3>

<table style="border-collapse:collapse;width:100%;margin-top:10px;">

  ${rowHTML('Gqeberha',
    `Julian van Wyk<br>
     📞 <a href="tel:+27723627604">+27 72 362 7604</a><br>
     ✉️ <a href="mailto:julianvanwyk@gmail.com">julianvanwyk@gmail.com</a>`)}

  ${rowHTML('Namibia',
    `Deon van Heerden<br>
     📞 <a href="tel:+27769458969">+27 76 945 8969</a><br>
     ✉️ <a href="mailto:deonvanheerden2@gmail.com">deonvanheerden2@gmail.com</a>`)}

  ${rowHTML('Cape #1 & #2',
    `Marius Marais<br>
     📞 <a href="tel:+27834454457">+27 83 445 4457</a><br>
     ✉️ <a href="mailto:mariusmarais2008@gmail.com">mariusmarais2008@gmail.com</a>`)}

  ${rowHTML('Durban',
    `Barry Vercueil<br>
     📞 <a href="tel:+27825583544">+27 82 558 3544</a><br>
     ✉️ <a href="mailto:bvercueil@gmail.com">bvercueil@gmail.com</a>`)}

  ${rowHTML('Bloemfontein',
    `Hannes Marais<br>
     📞 <a href="tel:+27832257703">+27 83 225 7703</a><br>
     ✉️ <a href="mailto:hannes.marais2007@gmail.com">hannes.marais2007@gmail.com</a>`)}

  ${rowHTML('Pretoria #1 & #2',
    `Andre de Bruyn<br>
     📞 <a href="tel:+27828280238">+27 82 828 0238</a><br>
     ✉️ <a href="mailto:andredebruyn@gmail.com">andredebruyn@gmail.com</a>`)}

  ${rowHTML('Pretoria Zulu',
    `Dickson Zivambiso<br>
     📞 <a href="tel:+263773290895">+263 77 329 0895</a><br>
     ✉️ <a href="mailto:dzivambiso@gmail.com">dzivambiso@gmail.com</a>`)}

</table>

          <!-- Submit Another Form Button -->
          <div style="margin-top:40px;">
            <a href="https://script.google.com/macros/s/AKfycbw8XOH7bvDSMqtlHwzZNNhcOsCCIl31RISU51-frTWG9ohzqC5Qerf1F0hBWXOQSVQJGg/exec"
               style="display:inline-block;
                      padding:12px 24px;
                      background-color:#1a73e8;
                      color:#ffffff;
                      text-decoration:none;
                      border-radius:4px;
                      font-weight:bold;">
              Submit Another Form
            </a>
          </div>

        </td>
      </tr>
    </table>
  </div>
  `;
}
/*
function sendMailjetEmail(toEmail, subject, htmlContent) {
  const mailjetApiKey = "e83b1e12a836c127f69040c3f812d929";
  const mailjetSecret = "ee307931eb89ed99261faae711d9b386";

  const url = "https://api.mailjet.com/v3.1/send";

  const payload = {
    Messages: [
      {
        From: {
          Email: "devfldinfo@gmail.com",
          Name: "Convention Attendance"
        },
        To: [
          {
            Email: toEmail,
            Name: toEmail.split("@")[0]
          }
        ],
        Subject: subject,
        HTMLPart: htmlContent
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(mailjetApiKey + ":" + mailjetSecret)
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}



*/

function sendPendingConfirmationEmails() {

  const LOCK_WAIT_TIME = 30000;       // 30 seconds
  const RATE_LIMIT_DELAY = 150;       // ms between emails
  const MAX_EMAILS_PER_RUN = 50;      // safety cap
  const MIN_QUOTA_BUFFER = 2;        // stop if quota drops to this

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(LOCK_WAIT_TIME);

    const remainingQuotaStart = MailApp.getRemainingDailyQuota();
    if (remainingQuotaStart <= MIN_QUOTA_BUFFER) {
      Logger.log("Quota too low. Auto-pausing.");
      return;
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const dataRange = sheet.getRange(2, 1, lastRow - 1, 21);
    const values = dataRange.getValues();

    let emailsSent = 0;
    let remainingQuota = remainingQuotaStart;

    for (let i = 0; i < values.length; i++) {

      if (emailsSent >= MAX_EMAILS_PER_RUN) break;

      if (remainingQuota <= MIN_QUOTA_BUFFER) {
        Logger.log("Quota exhausted during run. Auto-pausing.");
        break;
      }

      const row = values[i];

      const emailSentStamp = row[19]; // Column T
      const emailAddress   = row[18]; // Column S

      if (!emailSentStamp && emailAddress) {

        const data = {
          'Surname': row[1],
          'Name': row[2],
          'Age': row[3],
          'Sleeping on grounds': row[4],
          'Using a caravan': row[5],
          'Attending Thursday': row[6],
          'Attending Friday': row[7],
          'Attending Saturday': row[8],
          'Attending Sunday': row[9],
          'Name on working list': row[10],
          'Job preferance': row[11],
          'Saturday preps help': row[12],
          'Gender': row[13],
          'Staying for evening meals': row[14],
          'Convention': row[15],
          'Special requests': row[16],
          'Meeting': row[17],
          'Email': row[18]
        };

        try {

          MailApp.sendEmail({
            to: emailAddress,
            subject: 'RSA Convention Attendance',
            htmlBody: buildEmail(data),
            name: 'RSA Convention Attendance'
          });

          // Success → timestamp column T
          sheet.getRange(i + 2, 20).setValue(new Date());

          emailsSent++;
          remainingQuota--;

          Utilities.sleep(RATE_LIMIT_DELAY);

        } catch (err) {

          // Write error message in column U
          sheet.getRange(i + 2, 21).setValue(String(err));

          // Abort entire function immediately
          throw new Error("Aborted after failure sending to: " + emailAddress);
        }
      }
    }

    Logger.log("Run completed. Emails sent: " + emailsSent);

  } catch (outerErr) {
    Logger.log(outerErr);
    throw outerErr;

  } finally {
    lock.releaseLock();
  }
}
