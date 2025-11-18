// @ts-nocheck
import { readFile, readdir, stat } from "fs/promises";
import { google } from "googleapis";
import { join, basename } from "path";
import logger from "./logger.js";
import dotenv from "dotenv";
dotenv.config();

// ==================== KONFIGURASI ====================
const CONFIG = {
  // Google Sheets Configuration
  CREDENTIALS_FILE:
    process.env.GOOGLE_CREDENTIALS_FILE?.replace(/['"]/g, "") ||
    "./credentials.json",
  SPREADSHEET_ID: process.env.SPREADSHEET_ID?.replace(/['"]/g, ""),
  WORKSHEET_NAME:
    process.env.WORKSHEET_NAME?.replace(/['"]/g, "") ||
    "SIBER PD LABU - TEKNIS",

  // File Processing Configuration
  INPUT_MODE: process.env.INPUT_MODE?.replace(/['"]/g, "") || "directory",
  DIRECTORY_PATH:
    process.env.DIRECTORY_PATH?.replace(/['"]/g, "") || "./sources",
  FILE_PATTERN: new RegExp(
    process.env.FILE_PATTERN?.replace(/['"]/g, "") || "\\.txt$",
    "i"
  ),

  // Batch Processing Configuration
  BATCH_SIZE: parseInt(process.env.BATCH_SIZE) || 5,
  DELAY_BETWEEN_BATCHES: parseInt(process.env.DELAY_BETWEEN_BATCHES) || 1000,
  CONTINUE_ON_ERROR: process.env.CONTINUE_ON_ERROR !== "false",

  // Highlight Color Configuration (RGB values 0-1)
  HIGHLIGHT_COLOR: { red: 1, green: 0.8, blue: 0.6 },
};

// ==================== RECAP SYSTEM ====================
class ProcessingRecap {
  constructor() {
    this.results = {
      highlighted: [],
      notFound: [],
      failed: [],
      total: 0,
    };
    this.startTime = Date.now();
  }

  addResult(filename, status, details = {}) {
    this.results.total++;

    const result = {
      filename: basename(filename),
      status,
      timestamp: new Date().toLocaleString("id-ID"),
      ...details,
    };

    switch (status) {
      case "highlighted":
        this.results.highlighted.push(result);
        break;
      case "notFound":
        this.results.notFound.push(result);
        break;
      case "failed":
        this.results.failed.push(result);
        break;
    }
  }

  printRecap() {
    const duration = Date.now() - this.startTime;

    console.log("\n" + "═".repeat(80));
    console.log("🎨 HIGHLIGHTING RECAP - DETAILED RESULTS");
    console.log("═".repeat(80));

    // HIGHLIGHTED FILES
    if (this.results.highlighted.length > 0) {
      console.log(
        `\n✅ HIGHLIGHTED: ${this.results.highlighted.length} file(s)`
      );
      console.log("─".repeat(40));
      this.results.highlighted.forEach((result, index) => {
        console.log(`   ${index + 1}. ${result.filename}`);
        console.log(
          `      ↳ Row: ${result.rowIndex} | Match: ${result.matchType} | Value: ${result.matchValue}`
        );
      });
    } else {
      console.log(`\n✅ HIGHLIGHTED: 0 files`);
    }

    // NOT FOUND FILES
    if (this.results.notFound.length > 0) {
      console.log(`\n⚠️  NOT FOUND: ${this.results.notFound.length} file(s)`);
      console.log("─".repeat(40));
      this.results.notFound.forEach((result, index) => {
        console.log(`   ${index + 1}. ${result.filename}`);
        console.log(`      ↳ Reason: ${result.reason}`);
      });
    } else {
      console.log(`\n⚠️  NOT FOUND: 0 files`);
    }

    // FAILED FILES
    if (this.results.failed.length > 0) {
      console.log(`\n❌ FAILED: ${this.results.failed.length} file(s)`);
      console.log("─".repeat(40));
      this.results.failed.forEach((result, index) => {
        console.log(`   ${index + 1}. ${result.filename}`);
        console.log(`      ↳ Error: ${result.error}`);
      });
    } else {
      console.log(`\n❌ FAILED: 0 files`);
    }

    // SUMMARY
    console.log("\n" + "═".repeat(80));
    console.log("📈 SUMMARY");
    console.log("─".repeat(40));
    console.log(`   Total Files: ${this.results.total}`);
    console.log(
      `   ✅ Highlighted: ${this.results.highlighted.length} (${(
        (this.results.highlighted.length / this.results.total) *
        100
      ).toFixed(1)}%)`
    );
    console.log(
      `   ⚠️  Not Found: ${this.results.notFound.length} (${(
        (this.results.notFound.length / this.results.total) *
        100
      ).toFixed(1)}%)`
    );
    console.log(
      `   ❌ Failed: ${this.results.failed.length} (${(
        (this.results.failed.length / this.results.total) *
        100
      ).toFixed(1)}%)`
    );
    console.log(`   ⏱️  Duration: ${(duration / 1000).toFixed(2)} seconds`);
    console.log("═".repeat(80));

    // Log recap
    logger.info("Highlighting recap generated", {
      totalFiles: this.results.total,
      highlighted: this.results.highlighted.length,
      notFound: this.results.notFound.length,
      failed: this.results.failed.length,
      successRate: `${(
        (this.results.highlighted.length / this.results.total) *
        100
      ).toFixed(1)}%`,
      duration: `${duration}ms`,
      highlightedFiles: this.results.highlighted.map((r) => r.filename),
      notFoundFiles: this.results.notFound.map((r) => ({
        file: r.filename,
        reason: r.reason,
      })),
      failedFiles: this.results.failed.map((r) => ({
        file: r.filename,
        error: r.error,
      })),
    });
  }
}

// ==================== EKSTRAKSI DATA DARI FILENAME ====================
function extractIdentifiersFromFilename(filename) {
  const data = {};

  console.log("📁 Parsing filename:", filename);

  // Extract Serial Number
  const serialMatch = filename.match(
    /sn[._-]([A-Za-z0-9-]+)(?=_{0,2}[^A-Za-z0-9-]|_hn|$)/i
  );
  data.serialNumber = serialMatch?.[1] || null;

  // Extract Rack Number
  const rackMatch = filename.match(/r[._-]?(\d+(?:-\d+)?)/i);
  data.rackNumber = rackMatch?.[1] || null;

  // Extract U Slot Number (ambil nilai pertama jika range)
  const unitMatch = filename.match(/u[._-]?(\d+(?:-\d+)?)/i);
  const unitValue = unitMatch?.[1] || null;
  data.uSlotNumber = unitValue ? unitValue.split("-")[0] : null;

  console.log("   Extracted identifiers:", {
    serialNumber: data.serialNumber,
    rackNumber: data.rackNumber,
    uSlotNumber: data.uSlotNumber,
  });

  return data;
}

// ==================== AUTENTIKASI ====================
async function authenticate() {
  const credentials = JSON.parse(
    await readFile(CONFIG.CREDENTIALS_FILE, "utf-8")
  );
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return await auth.getClient();
}

// ==================== FIND ROW BY SN OR RACK+UNIT ====================
async function findTargetRow(sheets, identifiers) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: `${CONFIG.WORKSHEET_NAME}!A:Q`,
  });
  const rows = res.data.values || [];

  // PRIORITAS 1: Match SEMPURNA (SN + Rack + U Slot semuanya cocok)
  if (
    identifiers.serialNumber &&
    identifiers.rackNumber &&
    identifiers.uSlotNumber
  ) {
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const sheetSerialNumber = row[11]; // Column L (index 11)
      const sheetRackNumber = row[4]; // Column E (index 4)
      const sheetUnitNumber = row[5]; // Column F (index 5)

      if (
        sheetSerialNumber === identifiers.serialNumber &&
        sheetRackNumber === identifiers.rackNumber &&
        sheetUnitNumber === identifiers.uSlotNumber
      ) {
        logger.debug("Found PERFECT match (SN + Rack + Unit)", {
          serialNumber: sheetSerialNumber,
          rack: sheetRackNumber,
          unit: sheetUnitNumber,
          rowIndex: i + 1,
        });
        return {
          rowIndex: i + 1,
          matchType: "Perfect Match (SN + Rack + Unit)",
          matchValue: `SN:${sheetSerialNumber} R${sheetRackNumber} U${sheetUnitNumber}`,
        };
      }
    }
  }

  // PRIORITAS 2: Match by Rack + U Slot (lebih reliable karena physical location)
  if (identifiers.rackNumber && identifiers.uSlotNumber) {
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const sheetRackNumber = row[4]; // Column E (index 4)
      const sheetUnitNumber = row[5]; // Column F (index 5)

      if (
        sheetRackNumber === identifiers.rackNumber &&
        sheetUnitNumber === identifiers.uSlotNumber
      ) {
        const sheetSerialNumber = row[11]; // Column L (index 11)

        // Warning jika SN tidak cocok (kemungkinan data lama)
        if (
          identifiers.serialNumber &&
          sheetSerialNumber &&
          sheetSerialNumber !== identifiers.serialNumber
        ) {
          logger.warn("Rack+Unit matched but Serial Number differs", {
            expectedSN: identifiers.serialNumber,
            foundSN: sheetSerialNumber,
            rack: sheetRackNumber,
            unit: sheetUnitNumber,
            rowIndex: i + 1,
          });
        }

        logger.debug("Found row by Rack & U Slot", {
          rack: sheetRackNumber,
          unit: sheetUnitNumber,
          serialNumber: sheetSerialNumber,
          rowIndex: i + 1,
        });
        return {
          rowIndex: i + 1,
          matchType: "Rack & U Slot",
          matchValue: `R${sheetRackNumber} U${sheetUnitNumber}`,
          warning:
            identifiers.serialNumber &&
            sheetSerialNumber &&
            sheetSerialNumber !== identifiers.serialNumber
              ? `SN mismatch: expected ${identifiers.serialNumber}, found ${sheetSerialNumber}`
              : null,
        };
      }
    }
  }

  // PRIORITAS 3: Match by Serial Number saja (paling tidak reliable)
  if (identifiers.serialNumber) {
    const matches = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const sheetSerialNumber = row[11]; // Column L (index 11)

      if (sheetSerialNumber && sheetSerialNumber === identifiers.serialNumber) {
        matches.push({
          rowIndex: i + 1,
          rack: row[4],
          unit: row[5],
        });
      }
    }

    if (matches.length > 1) {
      logger.warn("Multiple rows found with same Serial Number", {
        serialNumber: identifiers.serialNumber,
        matches: matches,
        note: "This indicates duplicate data in sheet. Using first match.",
      });
      console.log(
        `\n⚠️  WARNING: Serial Number "${identifiers.serialNumber}" found in ${matches.length} rows:`
      );
      matches.forEach((m, idx) => {
        console.log(
          `   ${idx + 1}. Row ${m.rowIndex} - Rack: ${m.rack}, Unit: ${m.unit}`
        );
      });
      console.log(`   → Using first match (Row ${matches[0].rowIndex})\n`);
    }

    if (matches.length > 0) {
      const match = matches[0];
      logger.debug("Found row by Serial Number", {
        serialNumber: identifiers.serialNumber,
        rowIndex: match.rowIndex,
        duplicateCount: matches.length,
      });
      return {
        rowIndex: match.rowIndex,
        matchType: "Serial Number Only",
        matchValue: identifiers.serialNumber,
        warning:
          matches.length > 1
            ? `Found ${matches.length} duplicate entries`
            : null,
      };
    }
  }

  return null;
}

// ==================== HIGHLIGHT ROW ====================
async function highlightRow(sheets, rowIndex) {
  // Get sheet ID
  const sheetMetadata = await sheets.spreadsheets.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
  });

  const targetSheet = sheetMetadata.data.sheets.find(
    (sheet) => sheet.properties.title === CONFIG.WORKSHEET_NAME
  );

  const sheetId = targetSheet?.properties?.sheetId || 0;

  // Apply background color to entire row
  const requests = [
    {
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: rowIndex - 1, // 0-based index
          endRowIndex: rowIndex,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: CONFIG.HIGHLIGHT_COLOR,
          },
        },
        fields: "userEnteredFormat.backgroundColor",
      },
    },
  ];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    requestBody: { requests },
  });

  logger.info("Row highlighted", {
    rowIndex: rowIndex,
    color: CONFIG.HIGHLIGHT_COLOR,
  });
}

// ==================== PROCESS SINGLE FILE ====================
async function processSingleFile(authClient, filePath, recap) {
  const startTime = Date.now();

  try {
    const stats = await stat(filePath);
    logger.fileProcessing(filePath, "started", {
      size: stats.size,
    });

    const identifiers = extractIdentifiersFromFilename(basename(filePath));

    // Validate identifiers
    if (
      !identifiers.serialNumber &&
      !(identifiers.rackNumber && identifiers.uSlotNumber)
    ) {
      const reason =
        "No valid identifiers found (need Serial Number OR Rack+Unit)";
      logger.fileProcessing(filePath, "skipped", { reason });
      recap.addResult(filePath, "notFound", { reason });
      return { success: false, reason };
    }

    const sheets = google.sheets({ version: "v4", auth: authClient });
    const target = await findTargetRow(sheets, identifiers);

    if (!target) {
      const reason = "No matching row found in sheet";
      logger.fileProcessing(filePath, "not_found", {
        reason,
        identifiers,
      });
      recap.addResult(filePath, "notFound", {
        reason,
        searchedBy: identifiers.serialNumber
          ? `SN: ${identifiers.serialNumber}`
          : `Rack: ${identifiers.rackNumber}, U: ${identifiers.uSlotNumber}`,
      });
      return { success: false, reason };
    }

    // Highlight the row
    await highlightRow(sheets, target.rowIndex);

    const duration = Date.now() - startTime;

    logger.fileProcessing(filePath, "highlighted", {
      rowIndex: target.rowIndex,
      matchType: target.matchType,
      duration: `${duration}ms`,
    });

    recap.addResult(filePath, "highlighted", {
      rowIndex: target.rowIndex,
      matchType: target.matchType,
      matchValue: target.matchValue,
      duration: duration,
    });

    return {
      success: true,
      rowIndex: target.rowIndex,
      matchType: target.matchType,
    };
  } catch (error) {
    const duration = Date.now() - startTime;

    logger.error("File processing failed", {
      file: filePath,
      error: error.message,
      duration: `${duration}ms`,
      operation: "file_processing",
    });

    recap.addResult(filePath, "failed", {
      error: error.message,
      duration: duration,
    });

    if (!CONFIG.CONTINUE_ON_ERROR) throw error;
    return { success: false, error: error.message };
  }
}

// ==================== MAIN ====================
async function main() {
  const startTime = Date.now();

  try {
    logger.info("Starting Google Sheets Row Highlighting Process");

    const files = (await readdir(CONFIG.DIRECTORY_PATH))
      .filter((f) => CONFIG.FILE_PATTERN.test(f))
      .map((f) => join(CONFIG.DIRECTORY_PATH, f));

    logger.info("Files discovered", {
      totalFiles: files.length,
      directory: CONFIG.DIRECTORY_PATH,
    });

    if (files.length === 0) {
      console.log("⚠️  No files found to process");
      return;
    }

    const auth = await authenticate();
    logger.info("Authenticated with Google Sheets API");

    const recap = new ProcessingRecap();

    // Process all files
    for (const file of files) {
      await processSingleFile(auth, file, recap);
    }

    const totalDuration = Date.now() - startTime;

    // Print final recap
    recap.printRecap();

    // Log batch operation summary
    logger.batchOperation(
      files.length,
      recap.results.highlighted.length,
      recap.results.failed.length,
      totalDuration
    );

    logger.info("All files processed", {
      totalFiles: files.length,
      highlighted: recap.results.highlighted.length,
      notFound: recap.results.notFound.length,
      failed: recap.results.failed.length,
      totalDuration: `${totalDuration}ms`,
    });
  } catch (error) {
    const duration = Date.now() - startTime;
    logger.error("Main processing failed", {
      error: error.message,
      duration: `${duration}ms`,
      operation: "main_processing",
    });
    process.exit(1);
  }
}

// Run the application
main().catch((err) => {
  logger.error("Fatal error in main process", {
    error: err.message,
    stack: err.stack,
  });
  process.exit(1);
});

export { extractIdentifiersFromFilename, highlightRow, findTargetRow };
