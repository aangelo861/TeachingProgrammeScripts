async function main(workbook: ExcelScript.Workbook): Promise<void> {
    console.log("=== LETTER GENERATOR ===");
  
    const COLS_TOPIC_RELEVANT = [13, 24];
    const COLS_SPEAKER_CLEAR = [17];
    const COLS_CAREER_INTEREST = [18];
    const COLS_SKILLS_KNOWLEDGE = [19];
    const COLS_CONFIDENCE_MH = [20];
    const COLS_PROGRAM_SATISFACTION = [21];
    const COLS_COMMENTS = [8, 22];
    const COLS_CLARITY_RATING = [12];
  
    // ============================================
    // EMAIL DELIVERY CONFIG — fill in before use
    // --------------------------------------------
    // EMAIL_WORKER_URL is your Cloudflare Worker that proxies to the Resend API
    // (the Worker holds the Resend API key as a secret, so it never appears here).
    // EMAIL_TO is the inbox a Google Apps Script watches to build the PDFs.
    // See README.md → "Certificate / letter pipeline".
    // ============================================
    const EMAIL_WORKER_URL = "https://YOUR-WORKER-SUBDOMAIN.workers.dev/";
    const EMAIL_FROM = "onboarding@resend.dev";        // Resend sender (verified domain or sandbox)
    const EMAIL_TO = "your-inbox@example.com";         // inbox watched by the Apps Script
    const EMAIL_REPLY_TO = "your-reply-to@example.com";

    const letterCreationWS = workbook.getWorksheet("Letter creation");
    const sampleDataWS = workbook.getWorksheet("Sample Data");
  
    if (!letterCreationWS) { console.log("ERROR: Could not find 'Letter creation'"); return; }
    if (!sampleDataWS) { console.log("ERROR: Could not find 'Sample Data'"); return; }
  
    // --- Rebuild the user-friendly input form every run (preserves typed values) ---
    setupInputSheet(letterCreationWS);
  
    // --- Read inputs from the form ---
    const presenterName = letterCreationWS.getRange("C4").getValue();    // Presenter Name
    const teachingDate = letterCreationWS.getRange("C5").getValue();     // Date of Presentation
    const presentationName = letterCreationWS.getRange("C6").getValue(); // Presentation Title
  
    if (!teachingDate || !presenterName || !presentationName) {
      const msg = "Please fill in Presenter Name, Date of Presentation, and Presentation Title.";
      console.log("ERROR: " + msg);
      writeStatus(letterCreationWS, "⚠️ " + msg, "#e67700");
      return;
    }
  
    // --- Default signatures (the live signatures shown on the PDF come from the template) ---
    const signaturesStr = `
  FY Teaching Programme Coordinator`;
  
    console.log(`Teaching Date: ${teachingDate}`);
    console.log(`Presenter: ${presenterName}`);
    console.log(`Presentation: ${presentationName}`);
    console.log(`Signatures source: default block (live version comes from template)`);
    console.log(`Signatures preview: ${signaturesStr.substring(0, 80)}...`);
  
    // --- Parse date ---
    let targetDate: Date;
    try {
      if (typeof teachingDate === "number") {
        targetDate = new Date((teachingDate - 25569) * 86400 * 1000);
      } else {
        targetDate = new Date(String(teachingDate));
      }
      if (isNaN(targetDate.getTime())) throw new Error("Invalid date");
      console.log(`Parsed date: ${targetDate.toDateString()}`);
    } catch (error) {
      const msg = `Invalid date in the Date of Presentation field: ${teachingDate}`;
      console.log(`ERROR: ${msg}`);
      writeStatus(letterCreationWS, "⚠️ " + msg, "#e67700");
      return;
    }
  
    // --- Get sample data ---
    const sampleDataRange = sampleDataWS.getUsedRange();
    if (!sampleDataRange) { console.log("ERROR: No data in Sample Data"); return; }
  
    const sampleData = sampleDataRange.getValues();
    console.log(`Sample Data rows: ${sampleData.length}, columns: ${sampleData[0]?.length || 0}`);
  
    // --- Header verification ---
    if (sampleData.length > 0) {
      const headers = sampleData[0];
      console.log("\n--- Metric column mapping verification ---");
      const allMetricCols = [
        { name: "Topic Relevant", indices: COLS_TOPIC_RELEVANT },
        { name: "Speaker Clear", indices: COLS_SPEAKER_CLEAR },
        { name: "Career Interest", indices: COLS_CAREER_INTEREST },
        { name: "Skills/Knowledge", indices: COLS_SKILLS_KNOWLEDGE },
        { name: "Confidence MH", indices: COLS_CONFIDENCE_MH },
        { name: "Program Satisfaction", indices: COLS_PROGRAM_SATISFACTION },
        { name: "Comments", indices: COLS_COMMENTS },
        { name: "Clarity Rating", indices: COLS_CLARITY_RATING },
      ];
      for (const metric of allMetricCols) {
        const colHeaders = metric.indices
          .filter(idx => idx < headers.length)
          .map(idx => `${getColumnLetter(idx)}(${idx}): "${String(headers[idx]).replace(/\n/g, " ").substring(0, 60)}"`)
          .join(" | ");
        console.log(`  ${metric.name}: ${colHeaders}`);
      }
      console.log("---\n");
    }
  
    // --- Filter rows by date ---
    const filteredData: (string | number | boolean)[][] = [];
    for (let i = 1; i < sampleData.length; i++) {
      const startTimeStr = String(sampleData[i][1]);
      if (!startTimeStr || startTimeStr === "undefined" || startTimeStr === "") continue;
      try {
        let startDate: Date;
        const numericValue = parseFloat(startTimeStr);
        if (!isNaN(numericValue) && numericValue > 1000) {
          startDate = new Date((numericValue - 25569) * 86400 * 1000);
        } else {
          startDate = new Date(startTimeStr);
          if (isNaN(startDate.getTime())) {
            const m: RegExpMatchArray | null = startTimeStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)?/i);
            if (m) {
              const [, month, day, year, hour, minute, second, ampm = ""] = m;
              let h = parseInt(hour);
              if (ampm.toUpperCase() === "PM" && h !== 12) h += 12;
              else if (ampm.toUpperCase() === "AM" && h === 12) h = 0;
              startDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day), h, parseInt(minute), parseInt(second));
            }
          }
        }
        if (!isNaN(startDate.getTime()) && isSameDate(startDate, targetDate)) {
          filteredData.push(sampleData[i]);
        }
      } catch (error) { continue; }
    }
  
    console.log(`Found ${filteredData.length} responses for ${targetDate.toDateString()}`);
    if (filteredData.length === 0) {
      const msg = `No survey responses found for ${formatDateForLetter(targetDate)}.`;
      console.log(msg);
      writeStatus(letterCreationWS, "⚠️ " + msg, "#e67700");
      return;
    }
  
    // --- Debug dump ---
    console.log("\n=== DEBUG: First 3 rows — raw cell values ===");
    const debugCount = Math.min(3, filteredData.length);
    for (let d = 0; d < debugCount; d++) {
      const row = filteredData[d];
      const name = String(row[6]).substring(0, 30);
      console.log(`\n  Row ${d + 1} (${name}):`);
      for (let col = 7; col < Math.min(row.length, 27); col++) {
        const val = String(row[col]).trim();
        if (val && val !== "" && val !== "undefined") {
          console.log(`    ${getColumnLetter(col)}(${col}): "${val.substring(0, 60)}"`);
        } else {
          console.log(`    ${getColumnLetter(col)}(${col}): [empty]`);
        }
      }
    }
    console.log("=== END DEBUG DUMP ===\n");
  
    // --- Calculate metrics ---
    const percentages = calculatePercentages(filteredData, {
      topicRelevant: COLS_TOPIC_RELEVANT,
      speakerClear: COLS_SPEAKER_CLEAR,
      careerInterest: COLS_CAREER_INTEREST,
      skillsKnowledge: COLS_SKILLS_KNOWLEDGE,
    });
  
    const commentsArray = extractCommentsArray(filteredData, COLS_COMMENTS);
    const formattedDate = formatDateForLetter(targetDate);
  
    // Letter is still generated in memory for the email body + JSON payload,
    // but is NO LONGER written onto the worksheet.
    const letter = generateLetter(
      String(presenterName),
      String(presentationName),
      formattedDate,
      filteredData.length,
      percentages,
      commentsArray.map(c => `• ${c}`).join('\n'),
      signaturesStr
    );
  
    console.log("Letter generated in memory (not written to sheet).");
  
    // ============================================
    // BUILD STRUCTURED JSON PAYLOAD
    // ============================================
    interface StructuredPayload {
      schema_version: string;
      generated_at: string;
      source: string;
      session: {
        presenter_name: string;
        presentation_title: string;
        teaching_date: string;
        teaching_date_iso: string;
        total_responses: number;
      };
      metrics: {
        topic_relevant_and_speaker_clear_pct: number;
        skills_and_knowledge_gained_pct: number;
        career_interest_in_psychiatry_pct: number;
      };
      comments: string[];
      signatures: string;
      signatures_source: string;
      letter_plain_text: string;
    }
  
    const structuredPayload: StructuredPayload = {
      schema_version: "1.0",
      generated_at: new Date().toISOString(),
      source: "FY_Psychiatry_Teaching_Programme_Excel",
      session: {
        presenter_name: String(presenterName),
        presentation_title: String(presentationName),
        teaching_date: formattedDate,
        teaching_date_iso: targetDate.toISOString().split("T")[0],
        total_responses: filteredData.length,
      },
      metrics: {
        topic_relevant_and_speaker_clear_pct: percentages.relevantClear,
        skills_and_knowledge_gained_pct: percentages.skillsKnowledge,
        career_interest_in_psychiatry_pct: percentages.careerInterest,
      },
      comments: commentsArray,
      signatures: signaturesStr,
      signatures_source: "template",
      letter_plain_text: letter,
    };
  
    const structuredJsonString = JSON.stringify(structuredPayload, null, 2);
    console.log("Structured JSON built. Preview: " + structuredJsonString.substring(0, 200) + "...");
  
    // ============================================
    // BUILD EMAIL HTML
    // ============================================
    const letterAsHtml = letter
      .split('\n')
      .map(line => line.trim() === "" ? "<br/>" : `<p style="margin:0 0 4px 0;">${line}</p>`)
      .join('\n');
  
    const commentsAsHtml = commentsArray
      .map(c => `<li style="margin-bottom:6px;">${c}</li>`)
      .join('\n');
  
    const emailHtml = `
          <div style="font-family: Arial, sans-serif; font-size: 14px; color: #222; max-width: 800px;">
  
              <!-- SUMMARY PANEL -->
              <div style="background:#f0f4ff; border-left:4px solid #3b5bdb; padding:16px 20px; margin-bottom:24px; border-radius:4px;">
                  <h2 style="margin:0 0 12px 0; color:#3b5bdb;">📋 Teaching Contribution Letter</h2>
                  <table style="border-collapse:collapse; width:100%;">
                      <tr>
                          <td style="padding:4px 12px 4px 0; font-weight:bold; white-space:nowrap;">Presenter:</td>
                          <td style="padding:4px 0;">${presenterName}</td>
                      </tr>
                      <tr>
                          <td style="padding:4px 12px 4px 0; font-weight:bold; white-space:nowrap;">Presentation:</td>
                          <td style="padding:4px 0;">${presentationName}</td>
                      </tr>
                      <tr>
                          <td style="padding:4px 12px 4px 0; font-weight:bold; white-space:nowrap;">Teaching Date:</td>
                          <td style="padding:4px 0;">${formattedDate}</td>
                      </tr>
                      <tr>
                          <td style="padding:4px 12px 4px 0; font-weight:bold; white-space:nowrap;">Responses:</td>
                          <td style="padding:4px 0;">${filteredData.length} attendees</td>
                      </tr>
                  </table>
              </div>
  
              <!-- METRICS PANEL -->
              <div style="background:#f9f9f9; border-left:4px solid #868e96; padding:16px 20px; margin-bottom:24px; border-radius:4px;">
                  <h3 style="margin:0 0 12px 0; color:#555;">📊 Key Metrics</h3>
                  <table style="border-collapse:collapse; width:100%;">
                      <tr>
                          <td style="padding:4px 12px 4px 0;">Topic relevant &amp; speaker clear:</td>
                          <td style="padding:4px 0; font-weight:bold; color:#2f9e44;">${percentages.relevantClear}%</td>
                      </tr>
                      <tr>
                          <td style="padding:4px 12px 4px 0;">Skills &amp; knowledge gained:</td>
                          <td style="padding:4px 0; font-weight:bold; color:#2f9e44;">${percentages.skillsKnowledge}%</td>
                      </tr>
                      <tr>
                          <td style="padding:4px 12px 4px 0;">Career interest in psychiatry:</td>
                          <td style="padding:4px 0; font-weight:bold; color:#2f9e44;">${percentages.careerInterest}%</td>
                      </tr>
                  </table>
              </div>
  
              <!-- COMMENTS PANEL -->
              <div style="background:#fffdf0; border-left:4px solid #f08c00; padding:16px 20px; margin-bottom:24px; border-radius:4px;">
                  <h3 style="margin:0 0 12px 0; color:#e67700;">💬 Attendee Comments (${commentsArray.length})</h3>
                  <ul style="margin:0; padding-left:20px;">
                      ${commentsAsHtml}
                  </ul>
              </div>
  
              <!-- FULL LETTER -->
              <div style="background:#fff; border:1px solid #dee2e6; padding:24px 28px; border-radius:4px; line-height:1.7; margin-bottom:24px;">
                  <h3 style="margin:0 0 16px 0; color:#333; border-bottom:1px solid #dee2e6; padding-bottom:8px;">📄 Full Letter</h3>
                  ${letterAsHtml}
              </div>
  
              <!-- STRUCTURED JSON BLOCK — for Python extraction -->
              <div style="background:#1e1e1e; border-radius:4px; padding:20px; margin-bottom:24px;">
                  <p style="margin:0 0 10px 0; color:#888; font-size:12px; font-family:monospace;">
                      ⚙️ STRUCTURED_JSON_START — Python: import json, re; data = json.loads(re.search(r'STRUCTURED_JSON_START(.*?)STRUCTURED_JSON_END', body, re.DOTALL).group(1))
                  </p>
                  <pre style="margin:0; color:#d4d4d4; font-family:monospace; font-size:12px; white-space:pre-wrap; word-break:break-all;">${structuredJsonString}</pre>
                  <p style="margin:10px 0 0 0; color:#888; font-size:12px; font-family:monospace;">⚙️ STRUCTURED_JSON_END</p>
              </div>
  
              <p style="margin-top:20px; font-size:12px; color:#aaa;">
                  Auto-generated by the FY Psychiatry Teaching Programme Excel tool on ${formattedDate}.
              </p>
          </div>
      `;
  
    const emailPayload = {
      from: EMAIL_FROM,
      to: EMAIL_TO,
      replyTo: EMAIL_REPLY_TO,
      subject: `Teaching Letter: ${presenterName} — ${presentationName} (${formattedDate})`,
      html: emailHtml
    };
  
    console.log("\n=== EMAIL SENDING ===");
    console.log(`To: ${emailPayload.to}`);
    console.log(`Subject: ${emailPayload.subject}`);
  
    interface ResendResponse {
      id: string;
      object?: string;
      message?: string;
      name?: string;
      statusCode?: number;
    }
  
    let emailResponse: Response;
  
    try {
      emailResponse = await fetch(EMAIL_WORKER_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(emailPayload)
      });
      console.log("Fetch completed. HTTP status = " + emailResponse.status);
      console.log("Response OK = " + emailResponse.ok);
    } catch (fetchError) {
      console.log("❌ FETCH EXCEPTION");
      if (fetchError instanceof Error) {
        console.log("Error name = " + fetchError.name);
        console.log("Error message = " + fetchError.message);
      } else {
        console.log("Error dump = " + JSON.stringify(fetchError));
      }
      writeStatus(letterCreationWS, "❌ Email failed: fetch exception", "#e03131");
      console.log("=== EMAIL FAILED ===");
      return;
    }
  
    let emailResult: ResendResponse;
    try {
      emailResult = await emailResponse.json() as ResendResponse;
      console.log("Parsed response = " + JSON.stringify(emailResult));
    } catch (jsonError) {
      console.log("❌ Failed to parse response JSON");
      writeStatus(letterCreationWS, "❌ Email failed: could not parse response", "#e03131");
      console.log("=== EMAIL FAILED ===");
      return;
    }
  
    if (emailResponse.ok) {
      console.log("✅ EMAIL SENT SUCCESSFULLY to " + EMAIL_TO);
      console.log("Resend email ID = " + emailResult.id);
      writeStatus(letterCreationWS, "✅ Email sent to " + EMAIL_TO + " — " + new Date().toLocaleTimeString(), "#2f9e44");
    } else {
      console.log("❌ API error — Status: " + emailResponse.status);
      console.log("Error message: " + emailResult.message);
      if (emailResponse.status === 401) console.log("  → API key invalid or revoked");
      if (emailResponse.status === 403) console.log("  → Domain/recipient not verified in Resend");
      if (emailResponse.status === 422) console.log("  → Payload validation failed");
      if (emailResponse.status === 429) console.log("  → Rate limit hit");
      if (emailResponse.status === 500) console.log("  → Resend server error");
      writeStatus(letterCreationWS, "❌ Email failed: " + emailResult.message, "#e03131");
    }
  
    console.log("=== SCRIPT COMPLETE ===");
  }
  
  // ============================================
  // INPUT FORM SETUP
  // ============================================
  
  // Template documents shown on the sheet.
  // Edit these documents to change how the generated PDFs look — any change
  // appears on the next certificate / letter that is generated.
  const CERTIFICATE_DOC_URL = "https://docs.google.com/document/d/YOUR_CERTIFICATE_DOC_ID/edit?tab=t.0";
  const LETTER_DOC_URL = "https://docs.google.com/document/d/YOUR_LETTER_DOC_ID/edit?tab=t.0";
  
  /**
   * Rebuilds the "Letter creation" sheet as a clean input form on every run.
   * Preserves anything already typed into the input cells so the values
   * survive repeated runs, and clears all old layout / letter output.
   *
   *   C4 = Presenter Name
   *   C5 = Date of Presentation
   *   C6 = Presentation Title
   *   C8 = Status (written by the script, not the user)
   *   Rows 10+ = instructions + template links
   */
  function setupInputSheet(ws: ExcelScript.Worksheet): void {
    // 1. Preserve any values the user has already entered
    const existingPresenter = ws.getRange("C4").getValue();
    const existingDate = ws.getRange("C5").getValue();
    const existingTitle = ws.getRange("C6").getValue();
  
    // 2. Wipe the whole form area. This removes the old scattered inputs
    //    (G2 / O2 / W2 / B9), the old on-sheet letter output (B6 / B8) and
    //    any old status cell.
    ws.getRange("A1:X40").clear(ExcelScript.ClearApplyTo.all);
  
    // 3. Column widths: A = labels, B/D = spacers, C = input fields
    ws.getRange("A1").getFormat().setColumnWidth(170);
    ws.getRange("B1").getFormat().setColumnWidth(14);
    ws.getRange("C1").getFormat().setColumnWidth(380);
    ws.getRange("D1").getFormat().setColumnWidth(14);
  
    // 4. Title bar
    const title = ws.getRange("A1:D1");
    title.merge(false);
    title.setValue("Foundation Year Psychiatry Teaching Programme");
    const titleFmt = title.getFormat();
    titleFmt.getFont().setBold(true);
    titleFmt.getFont().setSize(16);
    titleFmt.getFont().setColor("#FFFFFF");
    titleFmt.getFill().setColor("#3B5BDB");
    titleFmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    titleFmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    titleFmt.setRowHeight(36);
  
    // 5. Subtitle
    const subtitle = ws.getRange("A2:D2");
    subtitle.merge(false);
    subtitle.setValue("Enter the session details below, then run the script to email the contribution letter.");
    subtitle.getFormat().getFont().setItalic(true);
    subtitle.getFormat().getFont().setSize(11);
    subtitle.getFormat().getFont().setColor("#666666");
  
    // 6. Labels + input cells
    setupLabel(ws, "A4", "Presenter Name");
    setupLabel(ws, "A5", "Date of Presentation");
    setupLabel(ws, "A6", "Presentation Title");
    setupLabel(ws, "A8", "Status");
  
    styleInput(ws, "C4");
    styleInput(ws, "C5");
    styleInput(ws, "C6");
  
    // Date cell: readable date number format
    ws.getRange("C5").setNumberFormat([["dd mmm yyyy"]]);
  
    // 7. Instructions panel
    const instrHeader = ws.getRange("A10:D10");
    instrHeader.merge(false);
    instrHeader.setValue("How it works & how to use this script");
    const ihFmt = instrHeader.getFormat();
    ihFmt.getFont().setBold(true);
    ihFmt.getFont().setSize(12);
    ihFmt.getFont().setColor("#FFFFFF");
    ihFmt.getFill().setColor("#868E96");
    ihFmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    ihFmt.setRowHeight(26);
  
    const instrBody = ws.getRange("A11:D11");
    instrBody.merge(false);
    instrBody.setValue(
      "WHAT THIS DOES\n" +
      "Generates the teaching contribution letter and certificate of appreciation as PDFs from Google Docs templates, then emails them automatically.\n\n" +
      "HOW TO USE\n" +
      "1. Fill in Presenter Name, Date of Presentation and Presentation Title above.\n" +
      "2. Run this script (Automate tab). Generation can take up to 5 minutes — please wait and do not re-run it repeatedly.\n" +
      "3. The Status cell above confirms when the email has been sent.\n\n" +
      "CHANGING THE LETTER OR CERTIFICATE\n" +
      "The wording, signatures, layout and logos live in the two template documents linked below. Edit a template and any change automatically appears on the next PDF generated — you do not need to change this script. In the certificate template, keep the {{Name}} and {{Date}} placeholders so they fill in automatically."
    );
    const ibFmt = instrBody.getFormat();
    ibFmt.setWrapText(true);
    ibFmt.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
    ibFmt.getFont().setSize(11);
    ibFmt.getFont().setColor("#333333");
    ibFmt.getFill().setColor("#F9FAFB");
    ibFmt.setRowHeight(200);
  
    // 8. Clickable template links
    addLink(ws, "A13:D13", LETTER_DOC_URL, "📄  Open the letter template  (edit to change the letter PDF)");
    addLink(ws, "A14:D14", CERTIFICATE_DOC_URL, "🏅  Open the certificate template  (edit to change the certificate PDF)");
  
    // 9. Restore preserved input values
    if (existingPresenter !== "") ws.getRange("C4").setValue(existingPresenter);
    if (existingDate !== "") ws.getRange("C5").setValue(existingDate);
    if (existingTitle !== "") ws.getRange("C6").setValue(existingTitle);
  }
  
  function addLink(ws: ExcelScript.Worksheet, address: string, url: string, text: string): void {
    const range = ws.getRange(address);
    range.merge(false);
    range.setHyperlink({ address: url, textToDisplay: text, screenTip: url });
    const fmt = range.getFormat();
    fmt.getFont().setBold(true);
    fmt.getFont().setSize(12);
    fmt.getFont().setColor("#1A56DB");
    fmt.getFill().setColor("#EEF2FF");
    fmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    fmt.setRowHeight(30);
    const edges = [
      ExcelScript.BorderIndex.edgeTop,
      ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.edgeLeft,
      ExcelScript.BorderIndex.edgeRight
    ];
    for (const edge of edges) {
      const b = fmt.getRangeBorder(edge);
      b.setStyle(ExcelScript.BorderLineStyle.continuous);
      b.setColor("#B5C2F0");
      b.setWeight(ExcelScript.BorderWeight.thin);
    }
  }
  
  function setupLabel(ws: ExcelScript.Worksheet, address: string, text: string): void {
    const cell = ws.getRange(address);
    cell.setValue(text);
    const fmt = cell.getFormat();
    fmt.getFont().setBold(true);
    fmt.getFont().setSize(11);
    fmt.getFont().setColor("#333333");
    fmt.setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  }
  
  function styleInput(ws: ExcelScript.Worksheet, address: string): void {
    const cell = ws.getRange(address);
    const fmt = cell.getFormat();
    fmt.getFill().setColor("#F4F6FB");
    fmt.getFont().setSize(11);
    const edges = [
      ExcelScript.BorderIndex.edgeTop,
      ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.edgeLeft,
      ExcelScript.BorderIndex.edgeRight
    ];
    for (const edge of edges) {
      const b = fmt.getRangeBorder(edge);
      b.setStyle(ExcelScript.BorderLineStyle.continuous);
      b.setColor("#C0C7D6");
      b.setWeight(ExcelScript.BorderWeight.thin);
    }
  }
  
  function writeStatus(ws: ExcelScript.Worksheet, message: string, color: string): void {
    const cell = ws.getRange("C8");
    cell.setValue(message);
    const fmt = cell.getFormat();
    fmt.getFont().setBold(true);
    fmt.getFont().setColor(color);
    fmt.setWrapText(true);
  }
  
  // ============================================
  // HELPERS
  // ============================================
  
  function getFirstNonEmpty(row: (string | number | boolean)[], colIndices: number[]): string {
    for (const idx of colIndices) {
      if (idx < row.length) {
        const val = String(row[idx]).trim();
        if (val && val !== "" && val !== "undefined") return val;
      }
    }
    return "";
  }
  
  function isAgreeOrStronglyAgree(response: string): boolean {
    if (!response) return false;
    const lower = response.toLowerCase().trim();
    if (lower === "strongly agree" || lower === "agree") return true;
    if (lower.includes("strongly agree")) return true;
    if (lower.endsWith("agree") && !lower.includes("disagree")) return true;
    const numVal = parseFloat(response);
    if (!isNaN(numVal) && (numVal === 4 || numVal === 5)) return true;
    return false;
  }
  
  function isSameDate(date1: Date, date2: Date): boolean {
    return date1.getFullYear() === date2.getFullYear() &&
      date1.getMonth() === date2.getMonth() &&
      date1.getDate() === date2.getDate();
  }
  
  function getColumnLetter(colIndex: number): string {
    let letter = "";
    let temp = colIndex;
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  }
  
  function calculatePercentages(
    data: (string | number | boolean)[][],
    colGroups: {
      topicRelevant: number[],
      speakerClear: number[],
      careerInterest: number[],
      skillsKnowledge: number[],
    }
  ): { relevantClear: number, skillsKnowledge: number, careerInterest: number } {
    const totalResponses = data.length;
    let relevantCount = 0, clearCount = 0, skillsKnowledgeCount = 0, careerInterestCount = 0;
  
    const relevantValues = new Set<string>();
    const clearValues = new Set<string>();
    const careerValues = new Set<string>();
    const skillsValues = new Set<string>();
  
    for (let row of data) {
      const r = getFirstNonEmpty(row, colGroups.topicRelevant);
      relevantValues.add(r || "[empty]");
      if (isAgreeOrStronglyAgree(r)) relevantCount++;
  
      const c = getFirstNonEmpty(row, colGroups.speakerClear);
      clearValues.add(c || "[empty]");
      if (isAgreeOrStronglyAgree(c)) clearCount++;
  
      const ca = getFirstNonEmpty(row, colGroups.careerInterest);
      careerValues.add(ca || "[empty]");
      if (isAgreeOrStronglyAgree(ca)) careerInterestCount++;
  
      const sk = getFirstNonEmpty(row, colGroups.skillsKnowledge);
      skillsValues.add(sk || "[empty]");
      if (isAgreeOrStronglyAgree(sk)) skillsKnowledgeCount++;
    }
  
    console.log("\n--- Unique values found per metric ---");
    console.log(`  Topic Relevant:   ${JSON.stringify([...relevantValues])}`);
    console.log(`  Speaker Clear:    ${JSON.stringify([...clearValues])}`);
    console.log(`  Career Interest:  ${JSON.stringify([...careerValues])}`);
    console.log(`  Skills/Knowledge: ${JSON.stringify([...skillsValues])}`);
    console.log("---");
  
    console.log(`\nTopic Relevant: ${relevantCount}/${totalResponses}`);
    console.log(`Speaker Clear: ${clearCount}/${totalResponses}`);
    console.log(`Skills/Knowledge: ${skillsKnowledgeCount}/${totalResponses}`);
    console.log(`Career Interest: ${careerInterestCount}/${totalResponses}`);
  
    return {
      relevantClear: Math.round(((relevantCount + clearCount) / (totalResponses * 2)) * 100),
      skillsKnowledge: Math.round((skillsKnowledgeCount / totalResponses) * 100),
      careerInterest: Math.round((careerInterestCount / totalResponses) * 100),
    };
  }
  
  function extractCommentsArray(data: (string | number | boolean)[][], commentColIndices: number[]): string[] {
    const commentSet = new Set<string>();
    const likertValues = new Set([
      "strongly agree", "agree", "neither agree nor disagree",
      "disagree", "strongly disagree"
    ]);
    for (let row of data) {
      for (const colIdx of commentColIndices) {
        if (colIdx < row.length) {
          const comment = String(row[colIdx]).trim();
          if (comment && comment !== "undefined" && comment !== "" && comment.length > 3) {
            if (likertValues.has(comment.toLowerCase())) continue;
            commentSet.add(comment);
          }
        }
      }
    }
    const comments = Array.from(commentSet);
    console.log(`\nFound ${comments.length} unique comments from columns ${commentColIndices.map(c => getColumnLetter(c)).join(", ")}`);
    if (comments.length > 0) {
      console.log("Comments preview:");
      for (let i = 0; i < Math.min(5, comments.length); i++) {
        console.log(`  ${i + 1}. "${comments[i].substring(0, 80)}"`);
      }
    }
    return comments;
  }
  
  function formatDateForLetter(date: Date): string {
    const months = ["January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"];
    return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
  }
  
  function generateLetter(
    presenterName: string,
    presentationName: string,
    formattedDate: string,
    attendanceCount: number,
    percentages: { relevantClear: number, skillsKnowledge: number, careerInterest: number },
    comments: string,
    signatures: string
  ): string {
    return `Foundation Year Psychiatry Teaching Programme Teaching Contribution Letter
  
  Dear ${presenterName},
  
  This letter serves to confirm your valuable contribution to the award-winning Foundation Year Psychiatry Teaching Programme for Central and North West London. Your presentation titled "${presentationName}" took place on ${formattedDate}.
  
  Your session formed part of a regional hybrid series attended by ${attendanceCount} Foundation Year doctors across all CNWL boroughs. You delivered an educational session that provided our trainees with essential knowledge and practical insights into ${presentationName}. Your expertise and contribution helped to develop our trainees' understanding and confidence in psychiatry. We appreciate your time, preparation, and dedication to medical education.
  
  ${percentages.relevantClear}% of Foundation doctors agreed/strongly agreed your presentation teaching topic was both relevant and clear. ${percentages.skillsKnowledge}% agreed/strongly agreed your teaching session provided them with skills and knowledge to use in their practice. ${percentages.careerInterest}% agreed/strongly agreed your teaching session increased their interest in choosing psychiatry as a career specialty. Specific audience feedback included the following:
  
  ${comments}
  
  On behalf of the organisers of the Foundation Year Psychiatry Teaching Programme, we are very grateful for your time and we would be delighted to welcome you back in the future.
  
  Yours sincerely,
  
  ${signatures}`;
  }