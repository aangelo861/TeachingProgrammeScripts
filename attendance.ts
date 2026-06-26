// Office Scripts TypeScript cannot parse union types directly inside new Map<>() angle
// brackets or inline arrow-function parameter lists — it confuses | with a comparison
// operator and then expects =>. Using a named alias avoids the ambiguity.
type ExcelValue = string | number | boolean;

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  console.log("=== Attendance cross-reference started ===");

  // ============================================
  // EMAIL DELIVERY CONFIG — fill in before use
  // --------------------------------------------
  // EMAIL_WORKER_URL is your Cloudflare Worker that proxies to the Resend API
  // (the Worker holds the Resend API key as a secret, so it never appears here).
  // EMAIL_TO is the inbox that receives the digest. See README.md.
  // ============================================
  const EMAIL_WORKER_URL = "https://YOUR-WORKER-SUBDOMAIN.workers.dev/";
  const EMAIL_FROM = "onboarding@resend.dev";        // Resend sender (verified domain or sandbox)
  const EMAIL_TO = "your-inbox@example.com";         // inbox that receives the digest email
  const EMAIL_REPLY_TO = "your-reply-to@example.com";

  // Worksheets
  const attendanceRecordWS = workbook.getWorksheet("Attendance Record");
  const sampleDataWS = workbook.getWorksheet("Sample Data");

  if (!attendanceRecordWS || !sampleDataWS) {
    console.log("ERROR: Required worksheets not found: 'Attendance Record' and/or 'Sample Data'");
    return;
  }

  // ============================================
  // COLUMN CONFIGURATION - Attendance Record
  // ============================================
  const ATTENDANCE_FIRSTNAME_COL = 5;   // F
  const ATTENDANCE_SURNAME_COL = 6;     // G
  const ATTENDANCE_EMAIL_COL = 7;       // H
  const DATE_COLUMNS_START = 8;         // I onwards
  // Col A (0): Site | Col B (1): Specialty | Col C (2): CS Trainer
  // Col D (3): CS Trainer email | Col E (4): OUTER/INNER

  // ============================================
  // COLUMN CONFIGURATION - Sample Data
  // ============================================
  const SAMPLE_COMPLETION_TIME_COL = 2; // C
  const SAMPLE_EMAIL_COL = 5;           // F

  // ============================================================
  // SESSION FEEDBACK column indices (0-based)
  // ------------------------------------------------------------
  // ⚠️ CORRECTED 28 May 2026 — verified against the generated digest email.
  // The OLD indices were off by TWO columns: the real sheet has a facilitator
  // name column (1) and a session title column (2) that the previous comment
  // block omitted, so every field from AggComments onward was shifted +2.
  // Symptoms that confirmed this: "Responses" rendered the session TITLE,
  // the comments box rendered a NAME, and Topic Relevant / Speaker Clear
  // rendered 1300% / 400% (they were actually the 13 responses / 4 comments
  // counts being multiplied by 100).
  //
  // 👉 If your sheet differs, this single block is the only thing to adjust.
  //    Open "Session Feedback" and confirm each constant points at the right
  //    column before relying on the per-session metrics.
  // ============================================================
  const SF_DATE = 0;             // A  session date
  const SF_FACILITATOR = 1;      // B  facilitator / speaker name
  const SF_TITLE = 2;            // C  session title
  const SF_AGG_COMMENTS = 3;     // D  aggregated free-text comments
  const SF_RESPONSES = 4;        // E  number of survey responses (count)
  const SF_UNIQUE_COMMENTS = 5;  // F  number of unique comments (count)
  const SF_TOPIC_RELEVANT = 6;   // G  topic relevant %        (stored 0-1)
  const SF_SPEAKER_CLEAR = 7;    // H  speaker clear %          (stored 0-1)
  const SF_COMBINED = 8;         // I  combined %               (stored 0-1)  [not displayed]
  const SF_CAREER_INTEREST = 9;  // J  career interest %        (stored 0-1)
  const SF_SKILLS = 10;          // K  skills / knowledge %     (stored 0-1)
  const SF_CONFIDENCE = 11;      // L  confidence in MH mgmt %  (stored 0-1)
  const SF_SATISFACTION = 12;    // M  programme satisfaction % (stored 0-1)
  // (further columns such as AvgClarity / ClarityResponses are not displayed)

  // Get data
  const attendanceRange = attendanceRecordWS.getUsedRange();
  const sampleDataRange = sampleDataWS.getUsedRange();

  if (!attendanceRange || !sampleDataRange) {
    console.log("ERROR: No data found in one or both worksheets");
    return;
  }

  const attendanceData = attendanceRange.getValues();
  const sampleData = sampleDataRange.getValues();

  if (attendanceData.length < 2 || sampleData.length < 2) {
    console.log("ERROR: One or both worksheets contain no usable data rows");
    return;
  }

  console.log(`Attendance rows: ${attendanceData.length - 1}`);
  console.log(`Sample rows: ${sampleData.length - 1}`);

  // ============================================
  // BUILD DATE COLUMN MAP
  // ============================================
  const dateToColumnsMap = buildDateColumnMap(attendanceData[0], DATE_COLUMNS_START);

  if (dateToColumnsMap.size === 0) {
    console.log("ERROR: No valid date columns found in Attendance Record");
    return;
  }

  let duplicateDateGroups = 0;
  for (const [, cols] of dateToColumnsMap) {
    if (cols.length > 1) duplicateDateGroups++;
  }

  console.log(`Date groups found: ${dateToColumnsMap.size}`);
  console.log(`Duplicate date groups: ${duplicateDateGroups}`);

  // ============================================
  // BUILD LOOKUP MAPS
  // ============================================
  const emailToRowMap = new Map<string, number>();
  const surnameToRowsMap = new Map<string, number[]>();
  const fullNameToRowsMap = new Map<string, number[]>();

  for (let row = 1; row < attendanceData.length; row++) {
    const firstName = normalizeNamePart(attendanceData[row][ATTENDANCE_FIRSTNAME_COL]);
    const surname = normalizeNamePart(attendanceData[row][ATTENDANCE_SURNAME_COL]);
    const email = normalizeEmail(attendanceData[row][ATTENDANCE_EMAIL_COL]);

    if (email) {
      emailToRowMap.set(email, row);
    }

    if (surname) {
      if (!surnameToRowsMap.has(surname)) {
        surnameToRowsMap.set(surname, []);
      }
      surnameToRowsMap.get(surname)!.push(row);
    }

    if (firstName && surname) {
      const fullNameKey = `${firstName}|${surname}`;
      if (!fullNameToRowsMap.has(fullNameKey)) {
        fullNameToRowsMap.set(fullNameKey, []);
      }
      fullNameToRowsMap.get(fullNameKey)!.push(row);
    }
  }

  // ============================================
  // PROCESS SAMPLE DATA
  // ============================================
  let totalConsidered = 0;
  let totalMatched = 0;
  let emailMatches = 0;
  let backupMatches = 0;
  let ambiguousBackupSkipped = 0;
  let dateNotFoundSkipped = 0;
  let invalidDateSkipped = 0;
  let noMatchSkipped = 0;
  let alreadyMarkedCount = 0;
  let attendanceMarked = 0;

  for (let i = 1; i < sampleData.length; i++) {
    const rawEmail = sampleData[i][SAMPLE_EMAIL_COL];
    const rawCompletion = sampleData[i][SAMPLE_COMPLETION_TIME_COL];

    const sampleEmail = normalizeEmail(rawEmail);
    if (!sampleEmail) {
      continue;
    }

    const completionDate = parseExcelDateValue(rawCompletion);
    if (!completionDate) {
      invalidDateSkipped++;
      continue;
    }

    totalConsidered++;

    // --------------------------------------------
    // Match row: exact email first, NHS email parse backup
    // --------------------------------------------
    let matchedRowIndex = -1;
    let matchType = "";

    // Primary: exact email
    const exactRow = emailToRowMap.get(sampleEmail);
    if (exactRow !== undefined) {
      matchedRowIndex = exactRow;
      matchType = "email";
      emailMatches++;
    } else {
      // Backup: parse firstname.lastname[number]@nhs.net
      const parsed = parseNhsEmail(sampleEmail);

      if (parsed) {
        const fullNameKey = `${parsed.firstName}|${parsed.lastName}`;
        const fullNameRows = fullNameToRowsMap.get(fullNameKey) || [];

        if (fullNameRows.length === 1) {
          matchedRowIndex = fullNameRows[0];
          matchType = "backup-fullname";
          backupMatches++;
        } else {
          const surnameRows = surnameToRowsMap.get(parsed.lastName) || [];

          if (surnameRows.length === 1) {
            matchedRowIndex = surnameRows[0];
            matchType = "backup-surname";
            backupMatches++;
          } else if (surnameRows.length > 1) {
            const narrowedRows = surnameRows.filter(rowIndex => {
              const attendanceFirstName = normalizeNamePart(attendanceData[rowIndex][ATTENDANCE_FIRSTNAME_COL]);
              return attendanceFirstName === parsed.firstName;
            });

            if (narrowedRows.length === 1) {
              matchedRowIndex = narrowedRows[0];
              matchType = "backup-surname+firstname";
              backupMatches++;
            } else {
              ambiguousBackupSkipped++;
            }
          }
        }
      }
    }

    if (matchedRowIndex === -1) {
      noMatchSkipped++;
      continue;
    }

    totalMatched++;

    // --------------------------------------------
    // Find date column
    // --------------------------------------------
    const dateKey = toDateKey(completionDate);
    const candidateColumns = dateToColumnsMap.get(dateKey);

    if (!candidateColumns || candidateColumns.length === 0) {
      dateNotFoundSkipped++;
      continue;
    }

    const dateColumnIndex = chooseDateColumn(candidateColumns, attendanceData, matchedRowIndex);

    // --------------------------------------------
    // Mark attendance
    // --------------------------------------------
    const existingValue = String(attendanceData[matchedRowIndex][dateColumnIndex] ?? "").trim().toUpperCase();
    if (existingValue === "Y") {
      alreadyMarkedCount++;
      continue;
    }

    const cell = attendanceRecordWS.getCell(matchedRowIndex, dateColumnIndex);
    cell.setValue("Y");
    cell.getFormat().getFill().setColor("#90EE90");

    attendanceData[matchedRowIndex][dateColumnIndex] = "Y";
    attendanceMarked++;

    if (attendanceMarked % 50 === 0) {
      console.log(`Progress: ${attendanceMarked} attendance marks written`);
    }
  }

  // ============================================
  // CROSS-REFERENCE SUMMARY
  // ============================================
  console.log("=== Attendance cross-reference completed ===");
  console.log(`Rows considered: ${totalConsidered}`);
  console.log(`Matched rows: ${totalMatched}`);
  console.log(`  Exact email matches: ${emailMatches}`);
  console.log(`  Backup NHS email matches: ${backupMatches}`);
  console.log(`Attendance marks written: ${attendanceMarked}`);
  console.log(`Already marked as Y: ${alreadyMarkedCount}`);
  console.log(`Skipped - ambiguous backup match: ${ambiguousBackupSkipped}`);
  console.log(`Skipped - no match found: ${noMatchSkipped}`);
  console.log(`Skipped - invalid completion date: ${invalidDateSkipped}`);
  console.log(`Skipped - matching date header not found: ${dateNotFoundSkipped}`);
  console.log(`Duplicate date groups detected: ${duplicateDateGroups}`);

  // ============================================================
  // EMAIL DIGEST
  // Attendance marking above always runs.
  // The digest email is gated: skip if < 3 calendar days since last send.
  // Last send date is persisted in Z100 of "Attendance Record" as mm/dd/yyyy.
  // Z100 is written ONLY on a successful HTTP 2xx response from the Worker.
  // ============================================================
  console.log("\n=== EMAIL DIGEST ===");

  // --- Z100: last email send date ---
  const LAST_EMAIL_ROW = 99;  // row 100, 0-based
  const LAST_EMAIL_COL = 25;  // col Z, 0-based
  const lastEmailCell = attendanceRecordWS.getCell(LAST_EMAIL_ROW, LAST_EMAIL_COL);
  const lastEmailRaw = lastEmailCell.getValue();

  let lastEmailDate: Date | null = null;
  if (lastEmailRaw && String(lastEmailRaw).trim() !== "") {
    lastEmailDate = parseExcelDateValue(lastEmailRaw);
    console.log(`Last email date (Z100): ${lastEmailDate ? lastEmailDate.toDateString() : "could not parse — treating as first run"}`);
  } else {
    console.log("Z100 is empty — first run, digest will cover all past sessions");
  }

  // --- 7-day gate (attendance marking already done above) ---
  const runDate = new Date();
  // Do NOT call setHours() here — we use the local date COMPONENTS below via Date.UTC()
  // to build timezone-safe calendar-day values.  Direct ms subtraction is broken in BST/DST:
  // runDate midnight (local) sits at 23:00 UTC the night before, so Math.floor gives 6 days
  // on the 7th calendar day.  Using Date.UTC(y, m, d) from LOCAL components avoids this.

  if (lastEmailDate) {
    const runDayUTC = Date.UTC(runDate.getFullYear(), runDate.getMonth(), runDate.getDate());
    const lastDayUTC = Date.UTC(lastEmailDate.getFullYear(), lastEmailDate.getMonth(), lastEmailDate.getDate());
    const daysDiff = Math.round((runDayUTC - lastDayUTC) / 86400000);
    console.log(`Days since last digest (calendar days): ${daysDiff}`);
    if (daysDiff < 3) {
      console.log(`EMAIL SKIPPED: ${daysDiff} calendar day(s) since last digest — need ≥3. Attendance marks were still applied.`);
      return;
    }
  }

  // --- Letter sign-off ---
  // B9 of "Letter creation" is no longer used; every non-attendance letter signs
  // off with the fixed line below.
  const nonAttendanceSignatures = `Foundation Teaching Programme Coordinator`;
  console.log("Non-attendance letter signatures: fixed sign-off ('Foundation Teaching Programme Coordinator')");

  // --- Load Session Feedback sheet ---
  const sessionFeedbackWS = workbook.getWorksheet("Session Feedback");
  const sessionFeedbackMap = new Map<string, ExcelValue[]>();

  if (sessionFeedbackWS) {
    const sfRange = sessionFeedbackWS.getUsedRange();
    if (sfRange) {
      const sfData = sfRange.getValues();
      for (let i = 1; i < sfData.length; i++) {
        const d = parseExcelDateValue(sfData[i][SF_DATE]);
        if (d) sessionFeedbackMap.set(toDateKey(d), sfData[i]);
      }
    }
    console.log(`Session Feedback rows loaded: ${sessionFeedbackMap.size}`);
  } else {
    console.log("WARNING: 'Session Feedback' sheet not found — per-session metrics will be omitted");
  }

  // --- Detect active trainees dynamically ---
  // Active = has a real first name and surname (not "Vacant", not blank)
  const activeTraineeRows: number[] = [];
  for (let row = 1; row < attendanceData.length; row++) {
    const fn = String(attendanceData[row][ATTENDANCE_FIRSTNAME_COL] ?? "").trim();
    const sn = String(attendanceData[row][ATTENDANCE_SURNAME_COL] ?? "").trim();
    if (!fn || !sn) continue;
    if (fn.toLowerCase() === "vacant") continue;
    activeTraineeRows.push(row);
  }
  const totalActive = activeTraineeRows.length;
  console.log(`Active trainees detected: ${totalActive}`);

  // --- Build ordered session list from header row ---
  interface SessionInfo {
    dateKey: string;
    date: Date;
    colIndices: number[];
    label: string;
  }

  const sessionMap = new Map<string, SessionInfo>();
  const hdr = attendanceData[0];

  for (let col = DATE_COLUMNS_START; col < hdr.length; col++) {
    const d = parseExcelDateValue(hdr[col]);
    if (!d) continue;
    const key = toDateKey(d);
    if (!sessionMap.has(key)) {
      sessionMap.set(key, { dateKey: key, date: d, colIndices: [], label: formatDateLabel(d) });
    }
    sessionMap.get(key)!.colIndices.push(col);
  }

  // Sort sessions chronologically
  const allSessions: SessionInfo[] = [...sessionMap.values()].sort(
    (a, b) => a.date.getTime() - b.date.getTime()
  );

  // Pre-compute today as a UTC calendar-day value (same approach as the gate above)
  const runDayUTC = Date.UTC(runDate.getFullYear(), runDate.getMonth(), runDate.getDate());

  // Past sessions = session calendar day <= today's calendar day
  const pastSessions = allSessions.filter(s => {
    const sDayUTC = Date.UTC(s.date.getFullYear(), s.date.getMonth(), s.date.getDate());
    return sDayUTC <= runDayUTC;
  });

  // Window sessions = past AND session calendar day > last-email calendar day
  const windowSessions = lastEmailDate
    ? pastSessions.filter(s => {
      const sDayUTC = Date.UTC(s.date.getFullYear(), s.date.getMonth(), s.date.getDate());
      const lastDayUTC = Date.UTC(lastEmailDate!.getFullYear(), lastEmailDate!.getMonth(), lastEmailDate!.getDate());
      return sDayUTC > lastDayUTC;
    })
    : pastSessions;

  console.log(`Past sessions total: ${pastSessions.length}`);
  console.log(`Sessions in digest window: ${windowSessions.length}`);

  if (windowSessions.length === 0) {
    console.log("No new sessions since last email — skipping digest");
    return;
  }

  // ============================================
  // PER-SESSION STATS
  // ============================================
  interface SessionStats {
    session: SessionInfo;
    attended: number;
    absent: string[];  // "Firstname Surname (Site)"
    rate: number;
    feedback: ExcelValue[] | null;
  }

  const sessionStats: SessionStats[] = [];

  for (const s of windowSessions) {
    // A trainee counts as present if ANY duplicate column for that date has "Y"
    const attended = activeTraineeRows.filter(r =>
      s.colIndices.some(col =>
        String(attendanceData[r][col] ?? "").trim().toUpperCase() === "Y"
      )
    ).length;

    const absent = activeTraineeRows
      .filter(r => !s.colIndices.some(col =>
        String(attendanceData[r][col] ?? "").trim().toUpperCase() === "Y"
      ))
      .map(r => {
        const fn = String(attendanceData[r][ATTENDANCE_FIRSTNAME_COL]).trim();
        const sn = String(attendanceData[r][ATTENDANCE_SURNAME_COL]).trim();
        const site = String(attendanceData[r][0] ?? "").trim();
        return `${fn} ${sn} (${site})`;
      });

    const rate = totalActive > 0 ? Math.round((attended / totalActive) * 100) : 0;
    sessionStats.push({
      session: s,
      attended,
      absent,
      rate,
      feedback: sessionFeedbackMap.get(s.dateKey) || null
    });
  }

  // ============================================
  // CURRENT CONSECUTIVE MISSED SESSIONS (trailing run since last attendance)
  // ------------------------------------------------------------
  // Per request: letters and concern flags must reflect ONLY the consecutive
  // sessions missed since the trainee last attended. As soon as a trainee
  // attends, the streak resets and all earlier non-attendance is discounted.
  // Example: misses #1,#2, attends #3, misses #4,#5  ->  current streak = 2,
  // and the letter lists only #4 and #5.
  // ============================================
  interface TraineeConcern {
    firstName: string;
    surname: string;
    email: string;        // col H — empty string if not populated; used to gate letter generation
    site: string;
    specialty: string;
    csTrainer: string;
    innerOuter: string;
    consecutiveMissed: number;      // CURRENT trailing streak (since last attended)
    totalMissed: number;            // total across all past sessions (context only)
    trailingMissedLabels: string[]; // ONLY the current trailing run — used in the letter
    lastAttendedLabel: string;
  }

  const concerns: TraineeConcern[] = [];

  for (const row of activeTraineeRows) {
    const firstName = String(attendanceData[row][ATTENDANCE_FIRSTNAME_COL]).trim();
    const surname = String(attendanceData[row][ATTENDANCE_SURNAME_COL]).trim();
    const rowEmail = normalizeEmail(attendanceData[row][ATTENDANCE_EMAIL_COL]);  // col H
    const site = String(attendanceData[row][0] ?? "").trim();
    const specialty = String(attendanceData[row][1] ?? "").trim();
    const csTrainer = String(attendanceData[row][2] ?? "").trim();
    const innerOuter = String(attendanceData[row][4] ?? "").trim();

    let consecutiveMissed = 0;            // resets on every attendance
    let trailingMissedLabels: string[] = []; // resets on every attendance
    let totalMissed = 0;
    let lastAttendedLabel = "Never attended";

    for (const session of pastSessions) {
      const present = session.colIndices.some(col =>
        String(attendanceData[row][col] ?? "").trim().toUpperCase() === "Y"
      );
      if (present) {
        // Attended — discount everything before this point.
        lastAttendedLabel = session.label;
        consecutiveMissed = 0;
        trailingMissedLabels = [];
      } else {
        consecutiveMissed++;
        totalMissed++;
        trailingMissedLabels.push(session.label);
      }
    }
    // After the loop, consecutiveMissed / trailingMissedLabels describe ONLY the
    // run of misses since the trainee last attended (or since the start if never).

    if (consecutiveMissed >= 2) {
      concerns.push({
        firstName, surname, email: rowEmail, site, specialty, csTrainer, innerOuter,
        consecutiveMissed, totalMissed,
        trailingMissedLabels,
        lastAttendedLabel
      });
    }
  }

  // Sort: worst current run first, then most total missed
  concerns.sort((a, b) => b.consecutiveMissed - a.consecutiveMissed || b.totalMissed - a.totalMissed);
  console.log(`Trainees currently on a 2+ consecutive-miss run: ${concerns.length}`);

  // Scan col H — only generate non-attendance letters for trainees who have an email address.
  // Trainees without an email still appear in the concerns table but are excluded from letters.
  const concernsWithEmail = concerns.filter(c => c.email && c.email.includes("@"));
  console.log(`  Of which have a valid email in col H (letters generated): ${concernsWithEmail.length}`);

  // ============================================
  // PROGRAMME-LEVEL STATS (all past sessions)
  // ============================================
  const avgRateWindow = sessionStats.length > 0
    ? Math.round(sessionStats.reduce((acc, s) => acc + s.rate, 0) / sessionStats.length)
    : 0;

  // Overall attendance rate across all past sessions
  let overallAttended = 0;
  let overallPossible = 0;
  for (const s of pastSessions) {
    for (const r of activeTraineeRows) {
      overallPossible++;
      if (s.colIndices.some(col =>
        String(attendanceData[r][col] ?? "").trim().toUpperCase() === "Y"
      )) overallAttended++;
    }
  }
  const overallRate = overallPossible > 0 ? Math.round((overallAttended / overallPossible) * 100) : 0;

  // Site breakdown (attendance across all past sessions, this window)
  const siteMap = new Map<string, { attended: number; possible: number }>();
  for (const r of activeTraineeRows) {
    const site = String(attendanceData[r][0] ?? "").trim() || "Unknown";
    for (const s of windowSessions) {
      if (!siteMap.has(site)) siteMap.set(site, { attended: 0, possible: 0 });
      const entry = siteMap.get(site)!;
      entry.possible++;
      if (s.colIndices.some(col =>
        String(attendanceData[r][col] ?? "").trim().toUpperCase() === "Y"
      )) entry.attended++;
    }
  }

  // Inner vs Outer breakdown (this window)
  const innerOuter = { INNER: { attended: 0, possible: 0 }, OUTER: { attended: 0, possible: 0 } };
  for (const r of activeTraineeRows) {
    const track = String(attendanceData[r][4] ?? "").trim().toUpperCase();
    const bucket = track === "INNER" ? innerOuter.INNER : innerOuter.OUTER;
    for (const s of windowSessions) {
      bucket.possible++;
      if (s.colIndices.some(col =>
        String(attendanceData[r][col] ?? "").trim().toUpperCase() === "Y"
      )) bucket.attended++;
    }
  }

  // ============================================
  // BUILD EMAIL HTML
  // ============================================
  const periodStart = lastEmailDate ? formatDateLabel(lastEmailDate) : "Programme start";
  const periodEnd = formatDateLabel(runDate);

  // --- Session cards ---
  let sessionCardsHtml = "";
  for (const stats of sessionStats) {
    const { session, attended, absent, rate, feedback } = stats;
    const rateColor = rate >= 80 ? "#2f9e44" : rate >= 60 ? "#e67700" : "#e03131";

    let metricsHtml = "";
    if (feedback) {
      const fb = feedback;

      // Counts are integers, not percentages. Coerce safely for display + pluralisation.
      const respCount = toIntOrZero(fb[SF_RESPONSES]);
      const uniqCount = toIntOrZero(fb[SF_UNIQUE_COMMENTS]);

      metricsHtml = `
            <div style="margin-top:14px;">
                <p style="margin:0 0 6px 0; font-size:11px; color:#888; font-weight:bold;
                          text-transform:uppercase; letter-spacing:0.5px;">
                    Survey — ${respCount} response${respCount !== 1 ? "s" : ""},
                    ${uniqCount} unique comment${uniqCount !== 1 ? "s" : ""}
                </p>
                <table style="border-collapse:collapse; width:100%;">
                    ${metricBarRow("Topic Relevant", normalisePercent(fb[SF_TOPIC_RELEVANT]))}
                    ${metricBarRow("Speaker Clear", normalisePercent(fb[SF_SPEAKER_CLEAR]))}
                    ${metricBarRow("Skills / Knowledge", normalisePercent(fb[SF_SKILLS]))}
                    ${metricBarRow("Career Interest in Psychiatry", normalisePercent(fb[SF_CAREER_INTEREST]))}
                    ${metricBarRow("Confidence in MH Management", normalisePercent(fb[SF_CONFIDENCE]))}
                    ${metricBarRow("Programme Satisfaction", normalisePercent(fb[SF_SATISFACTION]))}
                </table>`;

      const rawComments = String(fb[SF_AGG_COMMENTS] ?? "").trim();
      if (rawComments.length > 5) {
        // Truncate FIRST, then escape, so we never split an HTML entity like &amp;
        const truncated = rawComments.length > 600 ? rawComments.substring(0, 600) + "…" : rawComments;
        const safe = truncated
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
        metricsHtml += `
                <div style="margin-top:10px; padding:10px 12px; background:#fffdf0;
                            border-left:3px solid #f08c00; border-radius:3px;
                            font-size:12px; color:#555; white-space:pre-line; line-height:1.6;">
                    <strong style="color:#e67700;">💬 Attendee comments:</strong><br/>
                    ${safe}
                </div>`;
      }
      metricsHtml += `</div>`;
    } else {
      metricsHtml = `<p style="margin:8px 0 0 0; font-size:12px; color:#aaa;
                                     font-style:italic;">No survey feedback recorded.</p>`;
    }

    const absentSection = absent.length === 0
      ? `<p style="margin:6px 0 0 0; font-size:12px; color:#2f9e44; font-weight:bold;">
                   ✅ Full attendance!
               </p>`
      : `<p style="margin:6px 0 0 0; font-size:12px; color:#666; line-height:1.6;">
                   <strong>Absent (${absent.length}):</strong><br/>
                   ${absent.join(" &nbsp;·&nbsp; ")}
               </p>`;

    sessionCardsHtml += `
        <div style="background:#fff; border:1px solid #dee2e6; border-radius:4px;
                    padding:16px 20px; margin-bottom:16px;">
            <table style="border-collapse:collapse; width:100%; margin-bottom:10px;">
                <tr>
                    <td style="padding:0;">
                        <h3 style="margin:0; color:#333; font-size:15px;">📅 ${session.label}</h3>
                    </td>
                    <td style="padding:0; text-align:right; white-space:nowrap;">
                        <span style="font-size:19px; font-weight:bold; color:${rateColor};">
                            ${attended}&thinsp;/&thinsp;${totalActive}
                        </span>
                        <span style="font-size:12px; color:${rateColor}; margin-left:4px;">(${rate}%)</span>
                    </td>
                </tr>
            </table>
            <hr style="margin:0 0 10px 0; border:none; border-top:1px solid #eee;"/>
            ${absentSection}
            ${metricsHtml}
        </div>`;
  }

  // --- Site breakdown table ---
  let siteRowsHtml = "";
  for (const [site, data] of [...siteMap.entries()].sort((a, b) => b[1].possible - a[1].possible)) {
    const pct = data.possible > 0 ? Math.round((data.attended / data.possible) * 100) : 0;
    const col = pct >= 80 ? "#2f9e44" : pct >= 60 ? "#e67700" : "#e03131";
    siteRowsHtml += `<tr>
            <td style="padding:4px 12px 4px 0; font-size:13px;">${site}</td>
            <td style="padding:4px 12px; font-size:13px;">${data.attended} / ${data.possible}</td>
            <td style="padding:4px 0; font-size:13px; font-weight:bold; color:${col};">${pct}%</td>
        </tr>`;
  }

  const innerPct = innerOuter.INNER.possible > 0
    ? Math.round((innerOuter.INNER.attended / innerOuter.INNER.possible) * 100) : 0;
  const outerPct = innerOuter.OUTER.possible > 0
    ? Math.round((innerOuter.OUTER.attended / innerOuter.OUTER.possible) * 100) : 0;

  // --- Concerns panel ---
  let concernsHtml = "";
  if (concerns.length > 0) {
    const trows = concerns.map(c => {
      const missColor = c.consecutiveMissed >= 3 ? "#c92a2a" : "#e03131";
      return `<tr style="border-bottom:1px solid #fee2e2;">
                <td style="padding:6px 10px 6px 0; font-weight:bold; white-space:nowrap; font-size:13px;">
                    Dr ${c.firstName} ${c.surname}
                </td>
                <td style="padding:6px 10px; font-size:12px;">${c.site}</td>
                <td style="padding:6px 10px; font-size:12px;">${c.specialty}</td>
                <td style="padding:6px 10px; font-size:12px; color:#888;">${c.csTrainer}</td>
                <td style="padding:6px 10px; font-size:12px;">${c.innerOuter}</td>
                <td style="padding:6px 10px; text-align:center; font-weight:bold;
                            color:${missColor}; font-size:14px;">${c.consecutiveMissed}</td>
                <td style="padding:6px 10px; text-align:center; font-size:12px; color:#555;">
                    ${c.totalMissed} / ${pastSessions.length}
                </td>
                <td style="padding:6px 10px; font-size:12px; color:#888;">${c.lastAttendedLabel}</td>
            </tr>`;
    }).join("");

    concernsHtml = `
        <div style="background:#fff5f5; border-left:4px solid #e03131;
                    padding:16px 20px; margin-bottom:24px; border-radius:4px;">
            <h3 style="margin:0 0 14px 0; color:#c92a2a; font-size:15px;">
                ⚠️ Attendance Concerns — ${concerns.length} trainee${concerns.length !== 1 ? "s" : ""} currently on a 2+ consecutive-miss run
            </h3>
            <div style="overflow-x:auto;">
                <table style="border-collapse:collapse; width:100%; min-width:640px;">
                    <thead>
                        <tr style="background:#fee2e2;">
                            <th style="padding:6px 10px 6px 0; text-align:left; font-size:11px;
                                       color:#666; text-transform:uppercase; white-space:nowrap;">Name</th>
                            <th style="padding:6px 10px; text-align:left; font-size:11px;
                                       color:#666; text-transform:uppercase;">Site</th>
                            <th style="padding:6px 10px; text-align:left; font-size:11px;
                                       color:#666; text-transform:uppercase;">Specialty</th>
                            <th style="padding:6px 10px; text-align:left; font-size:11px;
                                       color:#666; text-transform:uppercase;">CS Trainer</th>
                            <th style="padding:6px 10px; text-align:left; font-size:11px;
                                       color:#666; text-transform:uppercase;">Track</th>
                            <th style="padding:6px 10px; text-align:center; font-size:11px;
                                       color:#666; text-transform:uppercase; white-space:nowrap;">Current Run</th>
                            <th style="padding:6px 10px; text-align:center; font-size:11px;
                                       color:#666; text-transform:uppercase; white-space:nowrap;">Total Missed</th>
                            <th style="padding:6px 10px; text-align:left; font-size:11px;
                                       color:#666; text-transform:uppercase; white-space:nowrap;">Last Attended</th>
                        </tr>
                    </thead>
                    <tbody>${trows}</tbody>
                </table>
            </div>
        </div>`;
  } else {
    concernsHtml = `
        <div style="background:#f0fff4; border-left:4px solid #2f9e44;
                    padding:12px 20px; margin-bottom:24px; border-radius:4px;">
            <p style="margin:0; color:#2f9e44; font-weight:bold; font-size:13px;">
                ✅ No trainees are currently on a run of 2 or more consecutive missed sessions.
            </p>
        </div>`;
  }

  // --- Non-attendance letters ---
  // Only generated for trainees whose col H email is populated (concernsWithEmail).
  // Each letter lists ONLY the current trailing run of misses (trailingMissedLabels).
  let lettersHtml = "";
  if (concernsWithEmail.length > 0) {
    let blocks = "";
    for (const c of concernsWithEmail) {
      const letterText = generateNonAttendanceLetter(c.surname, c.trailingMissedLabels, nonAttendanceSignatures);
      const letterHtml = letterText
        .split("\n")
        .map(line =>
          line.trim() === ""
            ? `<div style="height:8px;"></div>`
            : `<p style="margin:0 0 2px 0; font-size:13px;">${line
              .replace(/&/g, "&amp;")
              .replace(/</g, "&lt;")
              .replace(/>/g, "&gt;")}</p>`
        )
        .join("\n");

      // Show email address under the trainee name so recipients know who it's for
      const emailLine = `<p style="margin:2px 0 0 0; font-size:12px; color:#888;">
                                ${c.site} · ${c.specialty} · ${c.csTrainer}
                                &nbsp;|&nbsp; <span style="color:#3b5bdb;">${c.email}</span>
                               </p>`;

      blocks += `
            <div style="background:#fff; border:1px solid #dee2e6; border-radius:4px;
                        padding:20px 24px; margin-bottom:14px;">
                <table style="border-collapse:collapse; width:100%; margin-bottom:12px;">
                    <tr>
                        <td style="padding:0;">
                            <h4 style="margin:0; color:#333; font-size:14px;">
                                📄 Dr ${c.firstName} ${c.surname}
                            </h4>
                            ${emailLine}
                        </td>
                        <td style="padding:0; text-align:right; white-space:nowrap;">
                            <span style="background:#fee2e2; color:#c92a2a; font-size:11px;
                                         font-weight:bold; padding:3px 8px; border-radius:12px;">
                                ${c.consecutiveMissed} consecutive misses
                            </span>
                        </td>
                    </tr>
                </table>
                <hr style="margin:0 0 14px 0; border:none; border-top:1px solid #eee;"/>
                ${letterHtml}
            </div>`;
    }

    const noEmailCount = concerns.length - concernsWithEmail.length;
    const noEmailNote = noEmailCount > 0
      ? `<p style="margin:0 0 10px 0; font-size:12px; color:#e67700;">
                   ⚠️ ${noEmailCount} trainee${noEmailCount !== 1 ? "s" : ""} in the concerns table
                   above have no email address in column H — no letter generated for them.
               </p>`
      : "";

    lettersHtml = `
        <div style="margin-bottom:24px;">
            <h3 style="margin:0 0 10px 0; color:#333; font-size:15px;
                        border-bottom:2px solid #eee; padding-bottom:6px;">
                📝 Non-Attendance Letters (${concernsWithEmail.length})
            </h3>
            ${noEmailNote}
            ${blocks}
        </div>`;
  } else if (concerns.length > 0) {
    // There are concerns but none have an email address
    lettersHtml = `
        <div style="background:#fffbf0; border-left:4px solid #e67700;
                    padding:12px 20px; margin-bottom:24px; border-radius:4px;">
            <p style="margin:0; color:#e67700; font-size:13px;">
                ⚠️ ${concerns.length} trainee${concerns.length !== 1 ? "s" : ""} flagged above
                but none have an email address in column H — no letters generated.
                Populate column H to enable letter generation.
            </p>
        </div>`;
  }

  // --- Full email HTML ---
  const avgRateColor = avgRateWindow >= 80 ? "#2f9e44" : avgRateWindow >= 60 ? "#e67700" : "#e03131";
  const overallColor = overallRate >= 80 ? "#2f9e44" : overallRate >= 60 ? "#e67700" : "#e03131";

  const emailHtml = `
    <div style="font-family: Arial, sans-serif; font-size:14px; color:#222;
                max-width:800px; line-height:1.5;">

        <!-- HEADER PANEL -->
        <div style="background:#f0f4ff; border-left:4px solid #3b5bdb;
                    padding:16px 20px; margin-bottom:24px; border-radius:4px;">
            <h2 style="margin:0 0 14px 0; color:#3b5bdb; font-size:18px;">
                📊 FY Psychiatry Teaching Programme — Attendance &amp; Feedback Digest
            </h2>
            <table style="border-collapse:collapse; width:100%; font-size:13px;">
                <tr>
                    <td style="padding:3px 20px 3px 0; font-weight:bold; color:#555;
                                white-space:nowrap;">Period covered:</td>
                    <td style="padding:3px 0;">${periodStart} → ${periodEnd}</td>
                </tr>
                <tr>
                    <td style="padding:3px 20px 3px 0; font-weight:bold; color:#555;
                                white-space:nowrap;">Sessions in this digest:</td>
                    <td style="padding:3px 0;">${windowSessions.length}</td>
                </tr>
                <tr>
                    <td style="padding:3px 20px 3px 0; font-weight:bold; color:#555;
                                white-space:nowrap;">Active cohort size:</td>
                    <td style="padding:3px 0;">${totalActive} trainees</td>
                </tr>
                <tr>
                    <td style="padding:3px 20px 3px 0; font-weight:bold; color:#555;
                                white-space:nowrap;">Avg attendance (this digest):</td>
                    <td style="padding:3px 0; font-weight:bold; color:${avgRateColor};">${avgRateWindow}%</td>
                </tr>
                <tr>
                    <td style="padding:3px 20px 3px 0; font-weight:bold; color:#555;
                                white-space:nowrap;">Overall programme attendance:</td>
                    <td style="padding:3px 0; font-weight:bold; color:${overallColor};">
                        ${overallRate}% (${overallAttended}/${overallPossible} across ${pastSessions.length} sessions)
                    </td>
                </tr>
                <tr>
                    <td style="padding:3px 20px 3px 0; font-weight:bold; color:#555;
                                white-space:nowrap;">Inner track attendance:</td>
                    <td style="padding:3px 0; color:${innerPct >= 80 ? "#2f9e44" : "#e67700"};">
                        ${innerPct}% (${innerOuter.INNER.attended}/${innerOuter.INNER.possible})
                    </td>
                </tr>
                <tr>
                    <td style="padding:3px 20px 3px 0; font-weight:bold; color:#555;
                                white-space:nowrap;">Outer track attendance:</td>
                    <td style="padding:3px 0; color:${outerPct >= 80 ? "#2f9e44" : "#e67700"};">
                        ${outerPct}% (${innerOuter.OUTER.attended}/${innerOuter.OUTER.possible})
                    </td>
                </tr>
                <tr>
                    <td style="padding:3px 20px 3px 0; font-weight:bold; color:#555;
                                white-space:nowrap;">Trainees on a 2+ consecutive-miss run:</td>
                    <td style="padding:3px 0; font-weight:bold;
                                color:${concerns.length > 0 ? "#e03131" : "#2f9e44"};">
                        ${concerns.length}
                    </td>
                </tr>
            </table>
        </div>

        <!-- SITE BREAKDOWN -->
        <div style="background:#f9f9f9; border-left:4px solid #868e96;
                    padding:14px 20px; margin-bottom:24px; border-radius:4px;">
            <h3 style="margin:0 0 10px 0; color:#555; font-size:13px;
                        text-transform:uppercase; letter-spacing:0.5px;">
                🏥 Attendance by Site (this digest)
            </h3>
            <table style="border-collapse:collapse; font-size:13px;">
                ${siteRowsHtml}
            </table>
        </div>

        <!-- SESSION CARDS -->
        <h3 style="margin:0 0 14px 0; color:#333; font-size:15px;
                    border-bottom:2px solid #eee; padding-bottom:6px;">
            📋 Session Breakdown (${windowSessions.length} session${windowSessions.length !== 1 ? "s" : ""})
        </h3>
        ${sessionCardsHtml}

        <!-- ATTENDANCE CONCERNS -->
        ${concernsHtml}

        <!-- NON-ATTENDANCE LETTERS -->
        ${lettersHtml}

        <p style="margin-top:24px; font-size:11px; color:#aaa;
                   border-top:1px solid #eee; padding-top:12px;">
            Auto-generated by the FY Psychiatry Teaching Programme Excel attendance tracker
            on ${periodEnd}. Next digest will be sent ≥3 days from this date.
        </p>
    </div>`;

  // ============================================
  // SEND EMAIL
  // ============================================
  const emailSubject = `FY Teaching Digest: ${windowSessions.length} session${windowSessions.length !== 1 ? "s" : ""} · ${periodStart} → ${periodEnd} · ${avgRateWindow}% avg attendance`;

  const emailPayload = {
    from: EMAIL_FROM,
    to: EMAIL_TO,
    replyTo: EMAIL_REPLY_TO,
    subject: emailSubject,
    html: emailHtml
  };

  console.log("\n=== SENDING EMAIL ===");
  console.log(`To: ${emailPayload.to}`);
  console.log(`Subject: ${emailSubject}`);

  interface ResendResponse { id: string; message?: string; statusCode?: number; }

  let emailResponse: Response;
  try {
    emailResponse = await fetch(EMAIL_WORKER_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(emailPayload)
    });
    console.log(`HTTP status: ${emailResponse.status} | OK: ${emailResponse.ok}`);
  } catch (fetchError) {
    console.log("❌ FETCH EXCEPTION — email not sent, Z100 not updated");
    if (fetchError instanceof Error) {
      console.log(`${fetchError.name}: ${fetchError.message}`);
    }
    return;
  }

  let emailResult: ResendResponse;
  try {
    emailResult = await emailResponse.json() as ResendResponse;
    console.log(`Response: ${JSON.stringify(emailResult).substring(0, 200)}`);
  } catch (parseError) {
    console.log("❌ Could not parse response JSON — Z100 not updated");
    return;
  }

  if (emailResponse.ok) {
    console.log(`✅ EMAIL SENT to ${EMAIL_TO} | Resend ID: ${emailResult.id}`);

    // Write today's date to Z100 only on success.
    // Stored as a text string "mm/dd/yyyy" — parseExcelDateValue() handles this format
    // on the next run.  We do NOT call setNumberFormat() here because that method does
    // not exist on ExcelScript.RangeFormat in the Office Scripts runtime.
    const mm = String(runDate.getMonth() + 1).padStart(2, "0");
    const dd = String(runDate.getDate()).padStart(2, "0");
    const yyyy = String(runDate.getFullYear());
    const todayStr = `${mm}/${dd}/${yyyy}`;
    lastEmailCell.setValue(todayStr);
    lastEmailCell.getFormat().getFont().setSize(9);
    lastEmailCell.getFormat().getFont().setColor("#888888");
    console.log(`Z100 updated to: ${todayStr}`);
  } else {
    console.log(`❌ API error — Status ${emailResponse.status}: ${emailResult.message ?? "unknown"}`);
    if (emailResponse.status === 401) console.log("  → API key invalid or revoked");
    if (emailResponse.status === 403) console.log("  → Domain or recipient not verified in Resend");
    if (emailResponse.status === 422) console.log("  → Payload validation failed");
    if (emailResponse.status === 429) console.log("  → Rate limit exceeded");
    if (emailResponse.status === 500) console.log("  → Resend server error");
    console.log("Z100 not updated");
  }

  console.log("=== SCRIPT COMPLETE ===");
}

// ============================================
// HELPERS — EXISTING (unchanged)
// ============================================

/**
 * Builds a map of YYYY-M-D -> list of attendance date column indexes
 */
function buildDateColumnMap(
  headerRow: (string | number | boolean)[],
  startCol: number
): Map<string, number[]> {
  const map = new Map<string, number[]>();

  for (let col = startCol; col < headerRow.length; col++) {
    const parsedDate = parseExcelDateValue(headerRow[col]);
    if (!parsedDate) {
      continue;
    }

    const key = toDateKey(parsedDate);
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key)!.push(col);
  }

  return map;
}

/**
 * Chooses a date column from candidate columns.
 * If duplicates exist for the same date:
 * - prefer a column already containing data in that row
 * - otherwise use the first column
 */
function chooseDateColumn(
  candidateColumns: number[],
  attendanceData: (string | number | boolean)[][],
  matchedRowIndex: number
): number {
  if (candidateColumns.length === 1) {
    return candidateColumns[0];
  }

  for (const colIdx of candidateColumns) {
    const existingValue = String(attendanceData[matchedRowIndex][colIdx] ?? "").trim();
    if (existingValue && existingValue.toLowerCase() !== "undefined") {
      return colIdx;
    }
  }

  return candidateColumns[0];
}

/**
 * Parses Excel serial dates and standard date strings.
 * Returns null if parsing fails.
 */
function parseExcelDateValue(value: string | number | boolean): Date | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  // Excel serial number
  if (typeof value === "number" && value > 1000) {
    const date = new Date((value - 25569) * 86400 * 1000);
    return isValidDate(date) ? date : null;
  }

  const text = String(value).trim();
  if (!text || text.toLowerCase() === "undefined") {
    return null;
  }

  // Numeric-looking string serial date
  const numericValue = parseFloat(text);
  if (!isNaN(numericValue) && numericValue > 1000 && /^\d+(\.\d+)?$/.test(text)) {
    const date = new Date((numericValue - 25569) * 86400 * 1000);
    return isValidDate(date) ? date : null;
  }

  // mm/dd/yyyy [hh:mm[:ss] [AM/PM]] — checked BEFORE native parsing so the
  // day/month order is deterministic (native new Date("01/02/2026") is engine
  // and locale dependent and can silently swap day/month).
  const match = text.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?$/i
  );

  if (match) {
    const [, month, day, year, hour = "0", minute = "0", second = "0", ampm = ""] = match;

    let hour24 = parseInt(hour, 10);
    if (ampm.toUpperCase() === "PM" && hour24 !== 12) {
      hour24 += 12;
    } else if (ampm.toUpperCase() === "AM" && hour24 === 12) {
      hour24 = 0;
    }

    const date = new Date(
      parseInt(year, 10),
      parseInt(month, 10) - 1,
      parseInt(day, 10),
      hour24,
      parseInt(minute, 10),
      parseInt(second, 10)
    );

    return isValidDate(date) ? date : null;
  }

  // Native parsing fallback (ISO strings etc.)
  const native = new Date(text);
  if (isValidDate(native)) {
    return native;
  }

  return null;
}

/**
 * Normalizes names for reliable matching
 * e.g. "O'Neil-Smith" -> "oneilsmith"
 */
function normalizeNamePart(value: string | number | boolean): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

/**
 * Normalizes email
 */
function normalizeEmail(value: string | number | boolean): string {
  return String(value ?? "")
    .trim()
    .toLowerCase();
}

/**
 * Backup parser for NHS-style emails:
 * firstname.lastname[number]@nhs.net
 *
 * Examples:
 * - john.smith@nhs.net
 * - john.smith2@nhs.net
 * - anne-marie.o'neil12@nhs.net
 */
function parseNhsEmail(email: string): { firstName: string; lastName: string } | null {
  const cleaned = normalizeEmail(email);

  const match = cleaned.match(/^([a-z][a-z'-]*)\.([a-z][a-z'-]*?)(\d*)@nhs\.net$/i);
  if (!match) {
    return null;
  }

  const firstName = normalizeNamePart(match[1]);
  const lastName = normalizeNamePart(match[2]);

  if (!firstName || !lastName) {
    return null;
  }

  return { firstName, lastName };
}

/**
 * Date validation
 */
function isValidDate(date: Date): boolean {
  return !isNaN(date.getTime()) && date.getFullYear() > 1900;
}

/**
 * Consistent day-level key (month is 0-based, matching Date.getMonth())
 */
function toDateKey(date: Date): string {
  return `${date.getFullYear()}-${date.getMonth()}-${date.getDate()}`;
}

/**
 * Utility retained for future debugging if needed
 */
function getColumnLetter(colIndex: number): string {
  let letter = "";
  let temp = colIndex;

  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }

  return letter;
}

// ============================================
// HELPERS — NEW / UPDATED (email digest)
// ============================================

/**
 * Formats a Date as "22 April 2026"
 */
function formatDateLabel(date: Date): string {
  const months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
}

/**
 * Coerce a cell value to a non-negative integer, or 0 if it isn't a number.
 * Used for the survey "responses" / "unique comments" COUNTS (not percentages).
 */
function toIntOrZero(value: ExcelValue): number {
  const n = parseFloat(String(value ?? "").replace(/[^0-9.\-]/g, ""));
  if (isNaN(n) || n < 0) return 0;
  return Math.round(n);
}

/**
 * Normalises a survey-metric cell into an integer percentage 0-100, or null
 * if the cell is blank / non-numeric.
 *
 * Robust against the two failure modes that produced "1300%" / "400%":
 *  - values stored as a fraction (0-1)        -> scaled up by 100
 *  - values stored as a whole percent (0-100) -> left as-is
 *  - anything outside 0-100 is clamped, so a stray count or double-scaled
 *    value can never blow the bar chart up again
 *  - returns null for blank/garbage so the row shows "—" instead of "NaN%"
 */
function normalisePercent(value: ExcelValue): number | null {
  if (value === null || value === undefined || String(value).trim() === "") {
    return null;
  }
  let n = parseFloat(String(value).replace("%", "").trim());
  if (isNaN(n)) {
    return null;
  }
  // Stored as a fraction (e.g. 0.88) -> convert to percent.
  if (n > 0 && n <= 1) {
    n = n * 100;
  }
  // Clamp into the valid range so the bar can never overflow.
  if (n < 0) n = 0;
  if (n > 100) n = 100;
  return Math.round(n);
}

/**
 * Generates a non-attendance letter for a trainee.
 * missedDates: ONLY the trailing run of consecutive misses since the trainee
 *              last attended (earlier non-attendance is intentionally excluded).
 * signatures:  fixed sign-off line ("Foundation Teaching Programme Coordinator")
 */
function generateNonAttendanceLetter(surname: string, missedDates: string[], signatures: string): string {
  const datesList = missedDates.length > 0
    ? missedDates.map(d => `  • ${d}`).join("\n")
    : "  • (no sessions recorded)";

  const sessionWord = missedDates.length === 1 ? "session" : "consecutive sessions";

  return `Dear Dr ${surname},

I hope you are well.

We have noted that you have been unable to attend the following ${sessionWord} of the Foundation Teaching programme:

${datesList}

As you know, attendance at these teaching sessions is a mandatory part of the Foundation Programme. Dedicated time is allocated away from clinical duties to allow you to attend. If you are unable to join a session, we expect notification in advance, either to confirm annual leave or to explain the reason for non-attendance.

Could you please let me know the reason you have been unable to attend these recent sessions? This will help us understand if there are any barriers to attendance and ensure you are supported to meet the programme requirements going forward.

Best Wishes,

${signatures}`;
}

/**
 * Returns an HTML table row for a survey metric with a text progress bar.
 * pct: already-normalised integer 0-100, or null for "no data".
 *
 * The bar fill is hard-clamped to 0-10 segments so String.repeat() can never
 * receive a negative argument (which throws RangeError and would abort the
 * whole script) — this was a latent crash if any metric ever exceeded 100%
 * or went negative.
 */
function metricBarRow(label: string, pct: number | null): string {
  if (pct === null) {
    return `<tr>
        <td style="padding:3px 12px 3px 0; color:#555; font-size:13px; white-space:nowrap;">${label}</td>
        <td style="padding:3px 8px; font-weight:bold; color:#aaa; font-size:13px;
                   white-space:nowrap;">—</td>
        <td style="padding:3px 0; font-family:monospace; color:#ccc; font-size:11px;
                   white-space:nowrap;">no data</td>
    </tr>`;
  }

  const clamped = Math.max(0, Math.min(100, pct));
  const color = clamped >= 90 ? "#2f9e44" : clamped >= 75 ? "#e67700" : "#e03131";
  const filled = Math.max(0, Math.min(10, Math.round(clamped / 10)));
  const bar = "█".repeat(filled) + "░".repeat(10 - filled);
  return `<tr>
        <td style="padding:3px 12px 3px 0; color:#555; font-size:13px; white-space:nowrap;">${label}</td>
        <td style="padding:3px 8px; font-weight:bold; color:${color}; font-size:13px;
                   white-space:nowrap;">${clamped}%</td>
        <td style="padding:3px 0; font-family:monospace; color:${color}; font-size:11px;
                   white-space:nowrap;">${bar}</td>
    </tr>`;
}