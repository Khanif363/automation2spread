// @ts-nocheck
import { readFile, readdir, appendFile, stat } from "fs/promises";
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

  // Spreadsheet Column Configuration
  START_COLUMN_INDEX: parseInt(process.env.START_COLUMN_INDEX) || 18,
};

// ==================== FIELD MAPPING ====================
const FIELD_MAPPING = {
  "Rack Number": {
    custom: (_, filename) => filename.match(/r[._-]?(\d+)/i)?.[1] ?? "N/A",
    required: false,
  },

  "U Slot Number": {
    custom: (_, filename) =>
      filename.match(/u[._-]?(\d+(?:-\d+)?)/i)?.[1] ?? "N/A",
    required: false,
  },
  "Prosesor (CPU)": {
    pattern: /^\s*model\s*name\s*[:\-]?\s*(.+)$/im,
    required: false,
  },
  "Penggunaan CPU (%)": {
    pattern: /CPU\s+Usage:\s*([\d.]+)%/,
    required: false,
  },
  Memori: {
    custom: (content) => {
      const m = content.match(/Mem:\s+([\d.]+[GMK]i?)\s+([\d.]+[GMK]i?)/);
      return m ? `TOTAL ${m[1]}, USED ${m[2]}` : "N/A";
    },
    required: false,
  },
  "Penggunaan Memori (%)": { pattern: /Used:\s*([\d.]+)%/, required: false },
  Tipe: {
    custom: (content) => {
      const sectionMatch = content.match(
        /Disk Model and Capacity:[\s\S]*?(?=\n\s*\n|====================|$)/
      );
      if (!sectionMatch) return "N/A";

      const section = sectionMatch[0];
      const types = [
        ...section.matchAll(/\s(\b(?:disk|rom|lvm|part)\b)\s*/gi),
      ].map((m) => m[1].toLowerCase());

      const unique = [...new Set(types)];
      return unique.join(", ") || "N/A";
    },
    required: false,
  },
  Model: {
    custom: (content) => {
      const sectionMatch = content.match(
        /Disk Model and Capacity:[\s\S]*?(?=\n\s*\n|====================|$)/
      );
      if (!sectionMatch) return "N/A";

      const section = sectionMatch[0];
      const models = [
        ...section.matchAll(/^\s*\S+\s+(.+?)\s+\d+[GMKTB]\s+(disk|rom)/gm),
      ]
        .map((m) => m[1].trim())
        .filter((x) => x !== "");

      return models.join(", ") || "N/A";
    },
    required: false,
  },
  "Kapasitas Disk": {
    pattern: /^\s*total\s+([\d.]+[KMGT])/im,
    required: false,
  },
  "Konfigurasi RAID": {
    pattern:
      /RAID Configuration:\s*([\s\S]*?)(?=\n[A-Z][^\n]*:|====================|$)/,
    required: false,
  },
  Partisi: {
    custom: (content) => {
      const sectionMatch = content.match(
        /Disk Model and Capacity:[\s\S]*?(?=\n\s*\n|====================|$)/
      );
      if (!sectionMatch) return "N/A";

      const section = sectionMatch[0];
      const lines = section
        .split("\n")
        .map((l) => l.trim())
        .filter(
          (l) =>
            /^(?:[├└─]*\s*)?(sd|nvme|vd|sr)\w*/i.test(l) &&
            /(disk|part|rom|lvm)/i.test(l)
        );

      const partitions = lines.map((line) => {
        const clean = line.replace(/[├└─│]+/g, "").trim();
        const parts = clean.split(/\s+/);

        const name = parts[0] ?? "";
        const model = parts.slice(1, -3).join(", ") || "";
        const size = parts.at(-3) ?? "";
        const type = parts.at(-2) ?? "";
        const mount = parts.at(-1)?.startsWith("/") ? parts.at(-1) : "";

        const details = [model, size, type, mount].filter(Boolean).join(", ");
        return `${name} (${details})`;
      });

      return partitions.join("; ") || "N/A";
    },
    required: false,
  },
  "Penggunaan (%)": {
    pattern: /CPU\s*Usage:\s*([\d.]+)%/i,
    required: false,
  },
  "Nama dan Versi OS": {
    pattern: /OS Name & Version\s*:\s*(.+)/,
    required: false,
  },
  "Versi Kernel/Build (Potensi Kerentanan)": {
    pattern: /Kernel Version\s*:\s*(.+)/,
    required: false,
  },
  Architecture: { pattern: /Architecture\s*:\s*(.+)/, required: false },
  "Versi Firmware": { pattern: /Firmware Version\s*:\s*(.+)/, required: false },
  "Tgl Firmware": { pattern: /Firmware Release\s*:\s*(.+)/, required: false },
  Hostname: {
    pattern: /Hostname\s*:\s*([^\n\r]+)/,
    required: false,
  },
  "Nama Interface": {
    custom: (content) => {
      const interfaces = [
        ...content.matchAll(/[\n\r]\s*\d+:\s*([\w-]+):\s*<[^>]+>/g),
      ].map((m) => m[1]);
      return interfaces.join(", ") || "N/A";
    },
    required: false,
  },
  "Alamat IP dan Subnet Mask": {
    custom: (content) => {
      const ips = [
        ...content.matchAll(/\w+:\s*(\d+\.\d+\.\d+\.\d+\/\d+)/g),
      ].map((m) => m[1]);
      return [...new Set(ips)].join(", ") || "N/A";
    },
    required: false,
  },
  "Interface Count": {
    pattern: /Interface Count:\s*(\d+)/i,
    required: false,
  },
  "IP Gateway": { pattern: /default via\s*([\d.]+)/, required: false },
  "IP DNS": {
    custom: (content) => {
      const dnsList = new Set();

      // Method 1: Pattern original Anda - "DNS Servers (from ...)"
      const dnsMatch = content.match(
        /DNS Servers \(from [^)]+\):\n([\s\S]*?)(?=\n\n|DNS Configuration|$)/
      );
      if (dnsMatch) {
        const dnsLines = dnsMatch[1]
          .trim()
          .split("\n")
          .filter((line) => line.trim());
        dnsLines.forEach((line) => {
          const ip = line.trim().match(/[0-9a-fA-F:.]+/);
          if (ip) dnsList.add(ip[0]);
        });
      }

      // Method 2: Dari /etc/resolv.conf
      const resolvMatch = content.match(
        /--- \/etc\/resolv\.conf ---\n([\s\S]*?)(?=\n---|$)/
      );
      if (resolvMatch) {
        const nameservers = resolvMatch[1].matchAll(
          /nameserver\s+([0-9a-fA-F:.]+)/g
        );
        for (const match of nameservers) {
          dnsList.add(match[1]);
        }
      }

      // Method 3: Dari systemd-resolved
      const systemdMatch = content.match(
        /--- systemd-resolved.*?---\n([\s\S]*?)(?=\n---|$)/
      );
      if (systemdMatch) {
        const dnsServers = systemdMatch[1].matchAll(
          /DNS Servers?:\s*([0-9a-fA-F:.\s]+)/gi
        );
        for (const match of dnsServers) {
          const ips = match[1].trim().split(/\s+/);
          ips.forEach((ip) => {
            if (/^[0-9a-fA-F:.]+$/.test(ip)) dnsList.add(ip);
          });
        }
      }

      // Method 4: Dari NetworkManager
      const nmMatch = content.match(
        /--- NetworkManager.*?---\n([\s\S]*?)(?=\n---|$)/
      );
      if (nmMatch) {
        const dnsEntries = nmMatch[1].matchAll(
          /IP[46]\.DNS\[\d+\]:\s*([0-9a-fA-F:.]+)/g
        );
        for (const match of dnsEntries) {
          dnsList.add(match[1]);
        }
      }

      // Method 5: Dari DHCP leases
      const dhcpMatch = content.match(
        /--- DHCP leases.*?---\n([\s\S]*?)(?=\n---|$)/g
      );
      if (dhcpMatch) {
        dhcpMatch.forEach((block) => {
          const dnsEntries = block.matchAll(
            /domain-name-servers\s+([0-9a-fA-F:.,\s]+);/g
          );
          for (const match of dnsEntries) {
            const ips = match[1].split(/[,\s]+/);
            ips.forEach((ip) => {
              if (/^[0-9a-fA-F:.]+$/.test(ip.trim())) dnsList.add(ip.trim());
            });
          }
        });
      }

      // Method 6: Dari /etc/netplan/
      const netplanMatch = content.match(
        /--- \/etc\/netplan\/ ---\n([\s\S]*?)(?=\n---|$)/
      );
      if (netplanMatch) {
        // Format: addresses: [ip1, ip2]
        const bracketAddresses = netplanMatch[1].matchAll(
          /addresses:\s*\[([\s\S]*?)\]/g
        );
        for (const match of bracketAddresses) {
          const ips = match[1].match(/[0-9a-fA-F:.]+/g);
          if (ips) ips.forEach((ip) => dnsList.add(ip));
        }
        // Format list dengan dash: - ip
        const listAddresses = netplanMatch[1].matchAll(
          /^\s*-\s*([0-9a-fA-F:.]+)/gm
        );
        for (const match of listAddresses) {
          dnsList.add(match[1]);
        }
      }

      // Method 7: Dari /etc/network/interfaces
      const interfacesMatch = content.match(
        /--- \/etc\/network\/interfaces ---\n([\s\S]*?)(?=\n---|$)/
      );
      if (interfacesMatch) {
        const dnsEntries = interfacesMatch[1].matchAll(
          /dns-nameservers\s+([0-9a-fA-F:.\s]+)/g
        );
        for (const match of dnsEntries) {
          const ips = match[1].trim().split(/\s+/);
          ips.forEach((ip) => {
            if (/^[0-9a-fA-F:.]+$/.test(ip)) dnsList.add(ip);
          });
        }
      }

      // Method 8: Dari /etc/sysconfig/network-scripts/
      const ifcfgMatch = content.match(
        /--- \/etc\/sysconfig\/network-scripts\/ ---\n([\s\S]*?)(?=\n---|$)/
      );
      if (ifcfgMatch) {
        const dnsEntries = ifcfgMatch[1].matchAll(
          /DNS\d*=["']?([0-9a-fA-F:.]+)["']?/g
        );
        for (const match of dnsEntries) {
          dnsList.add(match[1]);
        }
      }

      // Filter hanya IP valid (IPv4 & IPv6)
      const validDNS = Array.from(dnsList).filter(
        (ip) =>
          /^([0-9]{1,3}\.){3}[0-9]{1,3}$/.test(ip) || // IPv4
          /^([0-9a-fA-F]{0,4}:){2,7}[0-9a-fA-F]{0,4}$/.test(ip) // IPv6
      );

      return validDNS.length > 0 ? validDNS.join(", ") : "N/A";
    },
    required: false,
  },
  "MAC Address": {
    custom: (content) => {
      const macs = [...content.matchAll(/link\/ether\s+([a-f0-9:]+)/gi)].map(
        (m) => m[1]
      );
      return [...new Set(macs)].slice(0, 3).join(", ") || "N/A";
    },
    required: false,
  },
  "Daftar Aplikasi Mayor & Versinya": {
    pattern:
      /Top Installed Applications:([\s\S]*?)(?=\n====================|$)/,
    custom: (content) => {
      if (!content) return "N/A";
      const packages = [];
      const lines = content.split("\n");

      for (const line of lines) {
        const cleanLine = line.trim();
        if (!cleanLine) continue;

        const match = cleanLine.match(/^\d+\s+([a-zA-Z0-9\-_.]+)\s+/);
        if (match) {
          packages.push(match[1]);
        }
      }

      return [...new Set(packages)].slice(0, 8).join(", ") || "N/A";
    },
    required: false,
  },
  "Running Service": {
    custom: (content) => {
      const services = [
        ...content.matchAll(/(\S+\.service)\s+loaded\s+active\s+running/g),
      ].map((m) => m[1]);
      return services.slice(0, 10).join(", ") || "N/A";
    },
    required: false,
  },
  "Port yang Terbuka": {
    custom: (content) => {
      const ports = [
        ...new Set(
          [...content.matchAll(/LISTEN\s+\d+\s+\d+\s+[\d.]+:(\d+)/g)].map(
            (m) => m[1]
          )
        ),
      ];
      return ports.slice(0, 15).join(", ") || "N/A";
    },
    required: false,
  },
};

// ==================== RECAP SYSTEM ====================
class ProcessingRecap {
  constructor() {
    this.results = {
      success: [],
      skipped: [],
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
      case "success":
        this.results.success.push(result);
        break;
      case "skipped":
        this.results.skipped.push(result);
        break;
      case "failed":
        this.results.failed.push(result);
        break;
    }
  }

  printRecap() {
    const duration = Date.now() - this.startTime;

    console.log("\n" + "═".repeat(80));
    console.log("📊 PROCESSING RECAP - DETAILED RESULTS");
    console.log("═".repeat(80));

    // SUCCESS FILES
    if (this.results.success.length > 0) {
      console.log(`\n✅ SUCCESS: ${this.results.success.length} file(s)`);
      console.log("─".repeat(40));
      this.results.success.forEach((result, index) => {
        console.log(`   ${index + 1}. ${result.filename}`);
        console.log(
          `      ↳ Row: ${result.rowIndex} | Type: ${result.serverType} | Hostname: ${result.hostname}`
        );
      });
    } else {
      console.log(`\n✅ SUCCESS: 0 files`);
    }

    // SKIPPED FILES
    if (this.results.skipped.length > 0) {
      console.log(`\n⏭️  SKIPPED: ${this.results.skipped.length} file(s)`);
      console.log("─".repeat(40));
      this.results.skipped.forEach((result, index) => {
        console.log(`   ${index + 1}. ${result.filename}`);
        console.log(`      ↳ Reason: ${result.reason}`);
      });
    } else {
      console.log(`\n⏭️  SKIPPED: 0 files`);
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
      `   ✅ Success: ${this.results.success.length} (${(
        (this.results.success.length / this.results.total) *
        100
      ).toFixed(1)}%)`
    );
    console.log(
      `   ⏭️  Skipped: ${this.results.skipped.length} (${(
        (this.results.skipped.length / this.results.total) *
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

    // Physical vs Virtual breakdown
    const physicalSuccess = this.results.success.filter(
      (r) => r.serverType === "Fisik"
    ).length;
    const virtualSuccess = this.results.success.filter(
      (r) => r.serverType === "VM"
    ).length;

    console.log(`\n   🖥️  Physical Servers: ${physicalSuccess}`);
    console.log(`   💿 Virtual Machines: ${virtualSuccess}`);
    console.log("═".repeat(80));

    // Log recap ke file juga
    logger.info("Processing recap generated", {
      totalFiles: this.results.total,
      success: this.results.success.length,
      skipped: this.results.skipped.length,
      failed: this.results.failed.length,
      successRate: `${(
        (this.results.success.length / this.results.total) *
        100
      ).toFixed(1)}%`,
      duration: `${duration}ms`,
      physicalServers: physicalSuccess,
      virtualMachines: virtualSuccess,
      successFiles: this.results.success.map((r) => r.filename),
      skippedFiles: this.results.skipped.map((r) => ({
        file: r.filename,
        reason: r.reason,
      })),
      failedFiles: this.results.failed.map((r) => ({
        file: r.filename,
        error: r.error,
      })),
    });
  }

  getRecapData() {
    return {
      summary: {
        total: this.results.total,
        success: this.results.success.length,
        skipped: this.results.skipped.length,
        failed: this.results.failed.length,
        successRate: `${(
          (this.results.success.length / this.results.total) *
          100
        ).toFixed(1)}%`,
      },
      details: {
        success: this.results.success,
        skipped: this.results.skipped,
        failed: this.results.failed,
      },
    };
  }
}

// ==================== HELPER ====================
function getColumnLetter(index) {
  let letter = "";
  while (index >= 0) {
    letter = String.fromCharCode((index % 26) + 65) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

// ==================== EKSTRAKSI DATA ====================
function extractDataFromTxt(content, filename = "") {
  const data = {};

  console.log("📁 Parsing filename:", filename);

  // IMPROVED REGEX PATTERNS
  const serialMatch = filename.match(
    /sn[._-]([A-Za-z0-9-]+)(?=_{0,2}[^A-Za-z0-9-]|_hn|$)/i
  );
  const rackMatch = filename.match(/r[._-]?(\d+(?:-\d+)?)/i);
  const unitMatch = filename.match(/u[._-]?(\d+(?:-\d+)?)/i);
  const typeMatch = filename.match(/ty[._-](svr|vm)(?:[_-].*?)?(?=_|$)/i);
  const hostnameMatch = filename.match(/hn[._-]([^_]+)/i);
  const parentIPMatch = filename.match(
    /(?:svr|vm)[_-]s[_-](\d+\.\d+\.\d+\.\d+)/i
  );

  data["Serial Num Device"] = serialMatch?.[1] || "N/A";
  data["Rack Number"] = rackMatch?.[1] || "N/A";

  // Handle U Slot ranges (e.g., "22-23" -> ambil pertama "22")
  const unitValue = unitMatch?.[1] || "N/A";
  data["U Slot Number"] = unitValue.split("-")[0];

  data["Hostname"] =
    hostnameMatch?.[1] ||
    content.match(/Hostname\s*:\s*(.+)/)?.[1]?.trim() ||
    "N/A";

  // IMPROVED Server Type detection
  if (typeMatch) {
    const typeValue = typeMatch[1].toLowerCase();
    if (typeValue === "svr") {
      data["Server Type"] = "Fisik";
    } else if (typeValue === "vm") {
      data["Server Type"] = "VM";
    } else {
      data["Server Type"] = "N/A";
    }
  } else {
    // Fallback: cari ty.svr atau ty.vm di mana saja dalam filename
    if (/ty\.svr/i.test(filename)) {
      data["Server Type"] = "Fisik";
    } else if (/ty\.vm/i.test(filename)) {
      data["Server Type"] = "VM";
    } else {
      data["Server Type"] = "N/A";
    }
  }

  // Parent IP untuk VM
  data["Parent IP"] = parentIPMatch?.[1] || null;

  console.log("   Extracted:", {
    serial: data["Serial Num Device"],
    rack: data["Rack Number"],
    unit: data["U Slot Number"],
    type: data["Server Type"],
    parentIP: data["Parent IP"],
  });

  // Extract field mapping data
  for (const [field, cfg] of Object.entries(FIELD_MAPPING)) {
    let val = "N/A";
    try {
      if (cfg.custom) {
        val = cfg.custom(content, filename);
      } else if (cfg.pattern) {
        val = content.match(cfg.pattern)?.[1]?.trim() ?? "N/A";
      }
      if (cfg.required && (val === "N/A" || val === null)) {
        throw new Error();
      }
    } catch {
      if (cfg.required) {
        throw new Error(`Missing required field: ${field}`);
      }
    }
    data[field] = val;
  }

  data["Source File"] = filename;
  data["Processed At"] = new Date().toLocaleString("id-ID", {
    timeZone: "Asia/Jakarta",
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
async function findTargetRow(sheets, data) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: `${CONFIG.WORKSHEET_NAME}!A:Q`,
  });
  const rows = res.data.values || [];

  // console.log(rows);
  

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const sn = row[11];
    const rackNum = row[4];
    const unitNum = row[5];

    if (
      sn &&
      data["Serial Num Device"] !== "N/A" &&
      sn === data["Serial Num Device"]
    ) {
      logger.debug("Found server by Serial Num Device", {
        serialNumber: sn,
        rowIndex: i + 1,
      });
      return { rowIndex: i + 1, matchType: "sn" };
    }

    console.log(data["U Slot Number"]);
    console.log('data["Rack Number"]', data["Rack Number"]);
    console.log('rackNum', rackNum);
    console.log('unitNum', unitNum);
    
    
    
    if (data["Rack Number"] !== "N/A" && data["U Slot Number"] !== "N/A") {
      if (
        rackNum === data["Rack Number"] &&
        unitNum === data["U Slot Number"]
      ) {
        logger.debug("Found server by rack and unit", {
          rack: rackNum,
          unit: unitNum,
          rowIndex: i + 1,
        });
        return { rowIndex: i + 1, matchType: "rack_unit" };
      }
    }
  }

  return null;
}

// ==================== FIND PARENT SERVER ROW FOR VM (Optional - for positioning only) ====================
async function findParentServerRow(sheets, data) {
  const parentIP = data["Parent IP"];
  const rackNumber = data["Rack Number"];
  const uSlotNumber = data["U Slot Number"];

  if (!parentIP && !rackNumber && !uSlotNumber) {
    logger.warn("No parent identifiers provided", {
      parentIP: parentIP,
      rackNumber: rackNumber,
      uSlotNumber: uSlotNumber,
    });
    return null;
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: `${CONFIG.WORKSHEET_NAME}!A:AL`,
  });
  const rows = res.data.values || [];

  // Priority 1: Search by Rack Number and U Slot Number
  // Berdasarkan screenshot: Rack Position = Column E (index 4), U Slot = Column F (index 5)
  if (
    rackNumber &&
    rackNumber !== "N/A" &&
    uSlotNumber &&
    uSlotNumber !== "N/A"
  ) {
    // console.log(`Searching for parent: Rack=${rackNumber}, U Slot=${uSlotNumber}`);

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const rowRackNumber = (row[4] || "").toString().trim(); // Column E (index 4)
      const rowUSlotNumber = (row[5] || "").toString().trim(); // Column F (index 5)

      //   console.log(`Row ${i + 1}: Rack=${rowRackNumber}, U Slot=${rowUSlotNumber}`);

      // Check if rack matches
      if (rowRackNumber === rackNumber) {
        // Check if U slot matches (handle ranges like "22-23")
        const rowUSlotFirst = rowUSlotNumber.split("-")[0];
        const searchUSlotFirst = uSlotNumber.split("-")[0];

        if (
          rowUSlotFirst === searchUSlotFirst ||
          rowUSlotNumber === uSlotNumber
        ) {
          logger.debug("Found parent server by rack & unit", {
            rackNumber: rackNumber,
            uSlotNumber: uSlotNumber,
            rowIndex: i + 1,
            matchType: "rack_u_slot",
            matchedRack: rowRackNumber,
            matchedUSlot: rowUSlotNumber,
          });
          return { rowIndex: i + 1, row };
        }
      }
    }

    console.log("No parent found by Rack+USlot, trying IP...");
  }

  // Priority 2: Search by Parent IP
  if (parentIP && parentIP !== "N/A") {
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const ipColumn = row[37] || ""; // IP address column (index 37)

      if (ipColumn.includes(parentIP)) {
        logger.debug("Found parent server by IP", {
          parentIP: parentIP,
          rowIndex: i + 1,
          matchType: "ip",
        });
        return { rowIndex: i + 1, row };
      }
    }
  }

  logger.warn("Parent server not found", {
    parentIP: parentIP,
    rackNumber: rackNumber,
    uSlotNumber: uSlotNumber,
  });
  return null;
}

// ==================== APPEND TO GOOGLE SHEET ====================
async function appendToGoogleSheet(sheets, data) {
  const TARGET_COLUMN_BLOCKS = [
    { start: 18, end: 41 },
    { start: 48, end: 50 },
  ];

  const fieldOrder = Object.keys(FIELD_MAPPING);
  const values = fieldOrder.map((f) => data[f] ?? "N/A");

  console.log("data", data);

  async function writeBlocks(rowIndex) {
    let idx = 0;
    for (const b of TARGET_COLUMN_BLOCKS) {
      const start = getColumnLetter(b.start);
      const end = getColumnLetter(b.end);
      const len = b.end - b.start + 1;
      const range = `${CONFIG.WORKSHEET_NAME}!${start}${rowIndex}:${end}${rowIndex}`;
      const blockValues = values.slice(idx, idx + len);
      idx += len;

      const cellsUpdated = blockValues.length;

      logger.spreadsheetOperation("update", CONFIG.SPREADSHEET_ID, {
        range: range,
        cellsUpdated: cellsUpdated,
        rowIndex: rowIndex,
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        range,
        valueInputOption: "RAW",
        requestBody: { values: [blockValues] },
      });
    }
  }

  // CASE 1: Server Fisik
  if (data["Server Type"] === "Fisik") {
    const target = await findTargetRow(sheets, data);
    if (!target) {
      return {
        success: true,
        filtered: true,
        reason: "Server tidak ditemukan di sheet (no match by SN or Rack+Unit)",
      };
    }

    await writeBlocks(target.rowIndex);
    return {
      success: true,
      processed: true,
      rowIndex: target.rowIndex,
    };
  }

  // CASE 2: VM - ALWAYS CREATE NEW ROW
  if (data["Server Type"] === "VM") {
    console.log("Processing VM - will create new row");

    // Cari parent server (optional - hanya untuk positioning)
    const parentRow = await findParentServerRow(sheets, data);

    let insertPosition;
    let newRow;

    if (parentRow) {
      // Insert row tepat setelah parent server
      insertPosition = parentRow.rowIndex; // Position to insert AFTER
      newRow = parentRow.rowIndex + 1; // The new row number after insert

      logger.info("VM will be inserted after parent server", {
        parentIP: data["Parent IP"],
        rackNumber: data["Rack Number"],
        uSlotNumber: data["U Slot Number"],
        parentRowIndex: parentRow.rowIndex,
        insertPosition: insertPosition,
        newRowIndex: newRow,
      });
    } else {
      // Jika parent tidak ditemukan, tambahkan di akhir sheet
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        range: `${CONFIG.WORKSHEET_NAME}!A:A`,
      });
      const lastRow = (res.data.values || []).length;
      insertPosition = lastRow; // Insert at the end
      newRow = lastRow + 1;

      logger.warn("Parent server not found, VM will be added at end of sheet", {
        parentIP: data["Parent IP"],
        rackNumber: data["Rack Number"],
        uSlotNumber: data["U Slot Number"],
        lastRow: lastRow,
        insertPosition: insertPosition,
        newRowIndex: newRow,
      });
    }

    // Get sheet ID (might not be 0)
    const sheetMetadata = await sheets.spreadsheets.get({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
    });

    const targetSheet = sheetMetadata.data.sheets.find(
      (sheet) => sheet.properties.title === CONFIG.WORKSHEET_NAME
    );

    const sheetId = targetSheet?.properties?.sheetId || 0;

    console.log("Sheet metadata:", {
      sheetId: sheetId,
      worksheetName: CONFIG.WORKSHEET_NAME,
      insertPosition: insertPosition,
    });

    // Insert new row using insertDimension
    const batchUpdateResponse = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            insertDimension: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: insertPosition, // 0-based index
                endIndex: insertPosition + 1, // Insert 1 row
              },
              inheritFromBefore: false, // Don't inherit formatting from row above
            },
          },
        ],
      },
    });

    logger.info("New row inserted for VM", {
      rowIndex: newRow,
      hostname: data["Hostname"],
      insertResponse: batchUpdateResponse.data,
    });

    // Write data to new row
    await writeBlocks(newRow);

    return {
      success: true,
      processed: true,
      isVM: true,
      rowIndex: newRow,
      parentFound: !!parentRow,
      insertPosition: insertPosition,
    };
  }

  return {
    success: true,
    filtered: true,
    reason: "Server type tidak dikenali",
  };
}

// ==================== PROCESS SINGLE FILE ====================
async function processSingleFile(authClient, filePath, recap) {
  const startTime = Date.now();

  try {
    const stats = await stat(filePath);
    logger.fileProcessing(filePath, "started", {
      size: stats.size,
    });

    const content = await readFile(filePath, "utf-8");
    const data = extractDataFromTxt(content, filePath);

    const sheets = google.sheets({ version: "v4", auth: authClient });
    const result = await appendToGoogleSheet(sheets, data);

    const duration = Date.now() - startTime;

    if (result.processed) {
      logger.fileProcessing(filePath, "completed", {
        rowIndex: result.rowIndex,
        serverType: data["Server Type"],
        duration: `${duration}ms`,
      });

      recap.addResult(filePath, "success", {
        rowIndex: result.rowIndex,
        serverType: data["Server Type"],
        hostname: data["Hostname"],
        duration: duration,
      });
    } else {
      logger.fileProcessing(filePath, "skipped", {
        reason: result.reason,
        duration: `${duration}ms`,
      });

      recap.addResult(filePath, "skipped", {
        reason: result.reason,
        duration: duration,
      });
    }

    return result;
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
    logger.info("Starting Google Sheets Update Process");

    const files = (await readdir(CONFIG.DIRECTORY_PATH))
      .filter((f) => CONFIG.FILE_PATTERN.test(f))
      .map((f) => join(CONFIG.DIRECTORY_PATH, f));

    logger.info("Files discovered", {
      totalFiles: files.length,
      directory: CONFIG.DIRECTORY_PATH,
    });

    const auth = await authenticate();
    logger.info("Authenticated with Google Sheets API");

    const recap = new ProcessingRecap();

    // Separate files by type
    const physicalServers = files.filter((f) => /ty\.svr/i.test(f));
    const virtualMachines = files.filter((f) => /ty\.vm/i.test(f));

    logger.info("File distribution", {
      physicalServers: physicalServers.length,
      virtualMachines: virtualMachines.length,
    });

    // Process physical servers first
    if (physicalServers.length > 0) {
      logger.info("Processing physical servers", {
        count: physicalServers.length,
      });

      for (const file of physicalServers) {
        await processSingleFile(auth, file, recap);
      }

      logger.info("Physical servers processing completed", {
        processed: physicalServers.length,
      });
    }

    // Process virtual machines after physical servers
    if (virtualMachines.length > 0) {
      logger.info("Processing virtual machines", {
        count: virtualMachines.length,
      });

      for (const file of virtualMachines) {
        await processSingleFile(auth, file, recap);
      }

      logger.info("Virtual machines processing completed", {
        processed: virtualMachines.length,
      });
    }

    const totalDuration = Date.now() - startTime;

    // Print final recap
    recap.printRecap();

    // Log batch operation summary
    logger.batchOperation(
      files.length,
      recap.results.success.length,
      recap.results.failed.length,
      totalDuration
    );

    logger.info("All files processed successfully", {
      totalFiles: files.length,
      success: recap.results.success.length,
      skipped: recap.results.skipped.length,
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

export { extractDataFromTxt, appendToGoogleSheet };
