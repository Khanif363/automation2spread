// @ts-nocheck
import { readFile, readdir, stat } from "fs/promises";
import { google } from "googleapis";
import { join, basename } from "path";
import logger from "./logger.js";
import dotenv from "dotenv";
dotenv.config();

// ==================== KONFIGURASI ====================
const CONFIG = {
  CREDENTIALS_FILE:
    process.env.GOOGLE_CREDENTIALS_FILE?.replace(/['"]/g, "") ||
    "./credentials.json",
  SPREADSHEET_ID: process.env.SPREADSHEET_ID?.replace(/['"]/g, ""),
  WORKSHEET_NAME:
    process.env.WORKSHEET_NAME?.replace(/['"]/g, "") ||
    "SIBER PD LABU - TEKNIS",

  DIRECTORY_PATH:
    process.env.DIRECTORY_PATH?.replace(/['"]/g, "") || "./sources",
  FILE_PATTERN: new RegExp(
    process.env.FILE_PATTERN?.replace(/['"]/g, "") || "\\.txt$",
    "i"
  ),

  BATCH_SIZE: parseInt(process.env.BATCH_SIZE) || 5,
  DELAY_BETWEEN_BATCHES: parseInt(process.env.DELAY_BETWEEN_BATCHES) || 1000,
  CONTINUE_ON_ERROR: process.env.CONTINUE_ON_ERROR !== "false",

  START_COLUMN_INDEX: parseInt(process.env.START_COLUMN_INDEX) || 16,
};

// ==================== FIELD MAPPING UNTUK WINDOWS ====================
const FIELD_MAPPING_WINDOWS = {
  "Rack Number": {
    custom: (_, filename) => {
      const match = filename.match(/[Rr][._-]?(\d+)/);
      return match?.[1] ?? "N/A";
    },
    required: false,
  },
  "U Slot Number": {
    custom: (_, filename) => {
      const match = filename.match(/[Uu][._-]?(\d+)/);
      return match?.[1] ?? "N/A";
    },
    required: false,
  },
  "Prosesor (CPU)": {
    custom: (content) => {
      const match = content.match(/Name\s*:\s*([^\r\n]+)/);
      return match?.[1]?.trim() || "N/A";
    },
    required: false,
  },
  "Penggunaan CPU (%)": {
    custom: (content) => {
      const matches = [...content.matchAll(/LoadPercentage\s*:\s*(\d+)/g)];
      if (matches.length > 0) {
        const sum = matches.reduce((acc, m) => acc + parseInt(m[1]), 0);
        const avg = (sum / matches.length).toFixed(2);
        return `${avg}%`;
      }
      return "N/A";
    },
    required: false,
  },
  Memori: {
    custom: (content) => {
      const totalMatch = content.match(/TotalPhysicalMemory\s*:\s*(\d+)/);
      const freeMatch = content.match(/FreePhysicalMemory\s*:\s*(\d+)/);
      
      if (totalMatch && freeMatch) {
        const totalGB = (parseInt(totalMatch[1]) / 1024 / 1024).toFixed(2);
        const usedGB = ((parseInt(totalMatch[1]) - parseInt(freeMatch[1])) / 1024 / 1024).toFixed(2);
        return `TOTAL ${totalGB}GB, USED ${usedGB}GB`;
      }
      return "N/A";
    },
    required: false,
  },
  "Penggunaan Memori (%)": {
    custom: (content) => {
      const totalMatch = content.match(/TotalPhysicalMemory\s*:\s*(\d+)/);
      const freeMatch = content.match(/FreePhysicalMemory\s*:\s*(\d+)/);
      
      if (totalMatch && freeMatch) {
        const total = parseInt(totalMatch[1]);
        const free = parseInt(freeMatch[1]);
        const usedPercent = (((total - free) / total) * 100).toFixed(2);
        return `${usedPercent}%`;
      }
      return "N/A";
    },
    required: false,
  },
  Tipe: {
    custom: (content) => {
      const types = new Set();
      
      // Check for SCSI, USB, etc
      if (content.includes("SCSI")) types.add("SCSI");
      if (content.includes("USB")) types.add("USB");
      if (content.includes("Fixed hard disk")) types.add("Fixed Disk");
      if (content.includes("Removable Media")) types.add("Removable");
      
      return types.size > 0 ? Array.from(types).join(", ") : "N/A";
    },
    required: false,
  },
  Model: {
    custom: (content) => {
      const models = new Set();
      const diskSection = content.match(/6\. PHYSICAL DISKS[\s\S]*?(?=7\.|$)/);
      
      if (diskSection) {
        const modelMatches = [...diskSection[0].matchAll(/Model\s*:\s*([^\r\n]+)/g)];
        modelMatches.forEach(m => {
          const model = m[1].trim();
          if (model) models.add(model);
        });
      }
      
      return models.size > 0 ? Array.from(models).join(", ") : "N/A";
    },
    required: false,
  },
  "Kapasitas Disk": {
    custom: (content) => {
      const sizeMatches = [...content.matchAll(/Size\s*:\s*(\d+)/g)];
      
      if (sizeMatches.length > 0) {
        const totalBytes = sizeMatches.reduce((acc, m) => acc + parseInt(m[1]), 0);
        const totalGB = (totalBytes / 1024 / 1024 / 1024).toFixed(2);
        return `${totalGB}GB`;
      }
      return "N/A";
    },
    required: false,
  },
  "Konfigurasi RAID": {
    custom: (content) => {
      if (content.includes("PERC") || content.includes("RAID")) {
        const raidMatch = content.match(/(PERC|RAID)[^\r\n]*/);
        return raidMatch?.[0]?.trim() || "RAID Detected";
      }
      return "N/A";
    },
    required: false,
  },
  Partisi: {
    custom: (content) => {
      const partitions = [];
      const logicalSection = content.match(/7\. LOGICAL DISKS[\s\S]*?(?=8\.|$)/);
      
      if (logicalSection) {
        const diskMatches = [...logicalSection[0].matchAll(/DeviceID\s*:\s*([A-Z]:)[\s\S]*?Size\s*:\s*(\d+)[\s\S]*?FreeSpace\s*:\s*(\d+)[\s\S]*?FileSystem\s*:\s*([^\r\n]+)/g)];
        
        diskMatches.forEach(m => {
          const drive = m[1];
          const sizeGB = (parseInt(m[2]) / 1024 / 1024 / 1024).toFixed(2);
          const freeGB = (parseInt(m[3]) / 1024 / 1024 / 1024).toFixed(2);
          const fs = m[4].trim();
          partitions.push(`${drive} (${sizeGB}GB, Free: ${freeGB}GB, ${fs})`);
        });
      }
      
      return partitions.length > 0 ? partitions.join("; ") : "N/A";
    },
    required: false,
  },
  "Penggunaan (%)": {
    custom: (content) => {
      const logicalSection = content.match(/7\. LOGICAL DISKS[\s\S]*?(?=8\.|$)/);
      
      if (logicalSection) {
        const diskMatches = [...logicalSection[0].matchAll(/Size\s*:\s*(\d+)[\s\S]*?FreeSpace\s*:\s*(\d+)/g)];
        
        if (diskMatches.length > 0) {
          let totalSize = 0;
          let totalUsed = 0;
          
          diskMatches.forEach(m => {
            const size = parseInt(m[1]);
            const free = parseInt(m[2]);
            totalSize += size;
            totalUsed += (size - free);
          });
          
          const usagePercent = ((totalUsed / totalSize) * 100).toFixed(2);
          return `${usagePercent}%`;
        }
      }
      return "N/A";
    },
    required: false,
  },
  "Nama dan Versi OS": {
    custom: (content) => {
      const nameMatch = content.match(/Caption\s*:\s*(Microsoft Windows[^\r\n]+)/);
      const versionMatch = content.match(/Version\s*:\s*([^\r\n]+)/);
      
      if (nameMatch) {
        const osName = nameMatch[1].trim();
        const version = versionMatch?.[1]?.trim() || "";
        return version ? `${osName} (${version})` : osName;
      }
      return "N/A";
    },
    required: false,
  },
  "Versi Kernel/Build (Potensi Kerentanan)": {
    custom: (content) => {
      const buildMatch = content.match(/BuildNumber\s*:\s*([^\r\n]+)/);
      return buildMatch?.[1]?.trim() || "N/A";
    },
    required: false,
  },
  Architecture: {
    custom: (content) => {
      const archMatch = content.match(/OSArchitecture\s*:\s*([^\r\n]+)/);
      return archMatch?.[1]?.trim() || "N/A";
    },
    required: false,
  },
  "Versi Firmware": {
    custom: (content) => {
      const biosMatch = content.match(/SMBIOSBIOSVersion\s*:\s*([^\r\n]+)/);
      return biosMatch?.[1]?.trim() || "N/A";
    },
    required: false,
  },
  "Tgl Firmware": {
    custom: (content) => {
      const dateMatch = content.match(/ReleaseDate\s*:\s*([^\r\n]+)/);
      return dateMatch?.[1]?.trim() || "N/A";
    },
    required: false,
  },
  Hostname: {
    custom: (content) => {
      const hostnameMatch = content.match(/(?:CSName|DNSHostName)\s*:\s*([^\r\n]+)/);
      return hostnameMatch?.[1]?.trim() || "N/A";
    },
    required: false,
  },
  "Nama Interface": {
    custom: (content) => {
      const interfaces = new Set();
      const adapterSection = content.match(/9\. NETWORK ADAPTERS[\s\S]*?(?=10\.|$)/);
      
      if (adapterSection) {
        const nameMatches = [...adapterSection[0].matchAll(/Name\s*:\s*([^\r\n]+)/g)];
        nameMatches.forEach(m => {
          const name = m[1].trim();
          if (name && !name.includes("Miniport") && !name.includes("Kernel")) {
            interfaces.add(name);
          }
        });
      }
      
      return interfaces.size > 0 ? Array.from(interfaces).join(", ") : "N/A";
    },
    required: false,
  },
  "Alamat IP dan Subnet Mask": {
    custom: (content) => {
      const ips = new Set();
      const ipMatches = [...content.matchAll(/IPAddress\s*:\s*\{([^}]+)\}/g)];
      
      ipMatches.forEach(m => {
        const ipList = m[1].split(",").map(ip => ip.trim());
        ipList.forEach(ip => {
          if (ip && !ip.startsWith("fe80") && ip !== "::1") {
            ips.add(ip);
          }
        });
      });
      
      const subnetMatches = [...content.matchAll(/IPSubnet\s*:\s*\{([^}]+)\}/g)];
      const subnets = [];
      subnetMatches.forEach(m => {
        const subnetList = m[1].split(",").map(s => s.trim());
        subnets.push(...subnetList);
      });
      
      const result = Array.from(ips).map((ip, i) => {
        const subnet = subnets[i] || "";
        return subnet ? `${ip}/${subnet}` : ip;
      });
      
      return result.length > 0 ? result.join(", ") : "N/A";
    },
    required: false,
  },
  "Interface Count": {
    custom: (content) => {
      const adapterSection = content.match(/9\. NETWORK ADAPTERS[\s\S]*?(?=10\.|$)/);
      
      if (adapterSection) {
        const count = (adapterSection[0].match(/NetConnectionStatus\s*:/g) || []).length;
        return count.toString();
      }
      return "N/A";
    },
    required: false,
  },
  "IP Gateway": {
    custom: (content) => {
      const gatewayMatch = content.match(/DefaultIPGateway\s*:\s*\{([^}]+)\}/);
      if (gatewayMatch) {
        const gateway = gatewayMatch[1].trim();
        return gateway || "N/A";
      }
      return "N/A";
    },
    required: false,
  },
  "IP DNS": {
    custom: (content) => {
      const dnsMatches = [...content.matchAll(/DNSServerSearchOrder\s*:\s*\{([^}]+)\}/g)];
      const dnsServers = new Set();
      
      dnsMatches.forEach(m => {
        const dnsList = m[1].split(",").map(dns => dns.trim());
        dnsList.forEach(dns => {
          if (dns) dnsServers.add(dns);
        });
      });
      
      return dnsServers.size > 0 ? Array.from(dnsServers).join(", ") : "N/A";
    },
    required: false,
  },
  "MAC Address": {
    custom: (content) => {
      const macs = new Set();
      const macMatches = [...content.matchAll(/MACAddress\s*:\s*([A-F0-9:-]+)/gi)];
      
      macMatches.forEach(m => {
        const mac = m[1].trim();
        if (mac && mac !== "00:00:00:00:00:00") {
          macs.add(mac);
        }
      });
      
      return macs.size > 0 ? Array.from(macs).slice(0, 3).join(", ") : "N/A";
    },
    required: false,
  },
  "Daftar Aplikasi Mayor & Versinya": {
    custom: (content) => {
      const productSection = content.match(/18\. INSTALLED PRODUCTS[\s\S]*?(?=19\.|$)/);
      const apps = new Set();
      
      if (productSection) {
        const productMatches = [...productSection[0].matchAll(/Name\s*:\s*([^\r\n]+)[\s\S]*?Version\s*:\s*([^\r\n]+)/g)];
        
        productMatches.forEach(m => {
          const name = m[1].trim();
          const version = m[2].trim();
          if (name && !name.includes("Minimum Runtime") && !name.includes("Additional Runtime")) {
            apps.add(`${name} (${version})`);
          }
        });
      }
      
      return apps.size > 0 ? Array.from(apps).slice(0, 8).join(", ") : "N/A";
    },
    required: false,
  },
  "Running Service": {
    custom: (content) => {
      // Windows services bisa diekstrak dari section lain jika ada
      // Untuk sementara return N/A atau bisa ditambahkan command Get-Service di script PowerShell
      return "N/A";
    },
    required: false,
  },
  "Port yang Terbuka": {
    custom: (content) => {
      // Port information bisa ditambahkan dengan netstat command di script PowerShell
      return "N/A";
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
    console.log("📊 WINDOWS DATA PROCESSING RECAP");
    console.log("═".repeat(80));

    if (this.results.success.length > 0) {
      console.log(`\n✅ SUCCESS: ${this.results.success.length} file(s)`);
      console.log("─".repeat(40));
      this.results.success.forEach((result, index) => {
        console.log(`   ${index + 1}. ${result.filename}`);
        console.log(
          `      ↳ Row: ${result.rowIndex} | Type: ${result.serverType} | Hostname: ${result.hostname}`
        );
      });
    }

    if (this.results.skipped.length > 0) {
      console.log(`\n⏭️  SKIPPED: ${this.results.skipped.length} file(s)`);
      console.log("─".repeat(40));
      this.results.skipped.forEach((result, index) => {
        console.log(`   ${index + 1}. ${result.filename}`);
        console.log(`      ↳ Reason: ${result.reason}`);
      });
    }

    if (this.results.failed.length > 0) {
      console.log(`\n❌ FAILED: ${this.results.failed.length} file(s)`);
      console.log("─".repeat(40));
      this.results.failed.forEach((result, index) => {
        console.log(`   ${index + 1}. ${result.filename}`);
        console.log(`      ↳ Error: ${result.error}`);
      });
    }

    console.log("\n" + "═".repeat(80));
    console.log("📈 SUMMARY");
    console.log("─".repeat(40));
    console.log(`   Total Files: ${this.results.total}`);
    console.log(`   ✅ Success: ${this.results.success.length}`);
    console.log(`   ⏭️  Skipped: ${this.results.skipped.length}`);
    console.log(`   ❌ Failed: ${this.results.failed.length}`);
    console.log(`   ⏱️  Duration: ${(duration / 1000).toFixed(2)} seconds`);
    console.log("═".repeat(80));
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

// ==================== EKSTRAKSI DATA WINDOWS ====================
function extractDataFromWindowsTxt(content, filename = "") {
  const data = {};

  // Extract from content first, then fallback to filename
  const serialMatch = content.match(/SerialNumber\s*:\s*([^\r\n]+)/) || 
                      content.match(/Serial Number\s*:\s*([^\r\n]+)/) ||
                      filename.match(/sn[._-]?([^_]+)/i);
  
  // More flexible rack and unit matching
  const rackMatch = filename.match(/[Rr][._-]?(\d+)/) || 
                    content.match(/Rack\s*(?:Number)?\s*:\s*(\d+)/i);
  
  const unitMatch = filename.match(/[Uu][._-]?(\d+)/) || 
                    content.match(/U\s*(?:Slot)?\s*(?:Number)?\s*:\s*(\d+)/i);
  
  const typeMatch = filename.match(/ty[._-]?(\w+)/i);
  const parentIPMatch = filename.match(/vm-s-([\d.]+)-i-([\d.]+)/);
  
  const hostnameMatch = content.match(/(?:CSName|DNSHostName|Computer(?:Name)?)\s*:\s*([^\r\n]+)/) ||
                        filename.match(/hn[._-]?([^_]+)/i);

  data["Serial Number"] = serialMatch?.[1]?.trim() || "N/A";
  data["Hostname"] = hostnameMatch?.[1]?.trim() || "N/A";

  // Determine server type
  if (typeMatch) {
    const typeValue = typeMatch[1].toLowerCase();
    if (typeValue === "svr" || typeValue === "server" || typeValue === "physical") {
      data["Server Type"] = "Fisik";
    } else if (typeValue === "vm" || typeValue === "virtual") {
      data["Server Type"] = "VM";
    } else {
      data["Server Type"] = "N/A";
    }
  } else {
    // Deteksi hypervisor untuk VM
    if (content.includes("VMware") || content.includes("Hyper-V") || 
        content.includes("VirtualBox") || content.includes("HypervisorPresent") ||
        content.match(/HypervisorPresent\s*:\s*True/i)) {
      data["Server Type"] = "VM";
    } else {
      data["Server Type"] = "Fisik";
    }
  }

  data["Parent IP"] = parentIPMatch?.[1] || null;

  console.log("🔍 DEBUG - Extracted from filename:", {
    filename: basename(filename),
    serialNumber: data["Serial Number"],
    rackNumber: rackMatch?.[1] || "N/A",
    unitNumber: unitMatch?.[1] || "N/A",
    serverType: data["Server Type"],
    hostname: data["Hostname"]
  });

  for (const [field, cfg] of Object.entries(FIELD_MAPPING_WINDOWS)) {
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

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const sn = row[9];
    const rackNum = row[16];
    const unitNum = row[17];

    if (sn && data["Serial Number"] !== "N/A" && sn === data["Serial Number"]) {
      logger.debug("Found server by serial number", {
        serialNumber: sn,
        rowIndex: i + 1,
      });
      return { rowIndex: i + 1, matchType: "sn" };
    }

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

// ==================== FIND PARENT SERVER ROW FOR VM ====================
async function findParentServerRow(sheets, data) {
  const parentIP = data["Parent IP"];
  const rackNumber = data["Rack Number"];
  const uSlotNumber = data["U Slot Number"];

  if (!parentIP) {
    logger.warn("Missing Parent IP for parent search", { parentIP: parentIP });
    return null;
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: `${CONFIG.WORKSHEET_NAME}!A:AJ`,
  });
  const rows = res.data.values || [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const rowRackNumber = row[16] || "";
    const rowUSlotNumber = row[17] || "";
    const ipColumn = row[35] || "";

    const isRackAndUSlotMatch =
      rowRackNumber === rackNumber && rowUSlotNumber === uSlotNumber;
    const isParentIPMatch = ipColumn.includes(parentIP);

    if (isRackAndUSlotMatch || isParentIPMatch) {
      logger.debug("Found parent server for VM", {
        parentIP: parentIP,
        rackNumber: rackNumber,
        uSlotNumber: uSlotNumber,
        rowIndex: i + 1,
        matchType: isRackAndUSlotMatch ? "rack_u_slot" : "ip",
      });
      return { rowIndex: i + 1, row };
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
    { start: 16, end: 39 },
    { start: 46, end: 48 },
  ];

  const fieldOrder = Object.keys(FIELD_MAPPING_WINDOWS);
  const values = fieldOrder.map((f) => data[f] ?? "N/A");

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

  // CASE 2: VM
  if (data["Server Type"] === "VM") {
    const parentRow = await findParentServerRow(sheets, data);
    if (!parentRow) {
      return {
        success: true,
        filtered: true,
        reason: `Parent server dengan IP ${data["Parent IP"]} atau Rack ${data["Rack Number"]} Unit ${data["U Slot Number"]} tidak ditemukan`,
      };
    }

    const newRow = parentRow.rowIndex + 1;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            insertDimension: {
              range: {
                sheetId: 0,
                dimension: "ROWS",
                startIndex: newRow - 1,
                endIndex: newRow,
              },
            },
          },
        ],
      },
    });

    await writeBlocks(newRow);
    return {
      success: true,
      processed: true,
      isVM: true,
      rowIndex: newRow,
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
    const data = extractDataFromWindowsTxt(content, filePath);

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
    logger.info("Starting Windows Data Processing for Google Sheets");

    const allFiles = await readdir(CONFIG.DIRECTORY_PATH);
    
    console.log("\n📁 Files in directory:");
    allFiles.forEach(f => console.log(`   - ${f}`));

    const files = allFiles
      .filter((f) => CONFIG.FILE_PATTERN.test(f))
      .map((f) => join(CONFIG.DIRECTORY_PATH, f));

    console.log("\n✅ Files matching pattern:");
    files.forEach(f => console.log(`   - ${basename(f)}`));

    logger.info("Files discovered", {
      totalFiles: files.length,
      directory: CONFIG.DIRECTORY_PATH,
      pattern: CONFIG.FILE_PATTERN.toString()
    });

    if (files.length === 0) {
      console.log("\n⚠️  WARNING: No files match the pattern!");
      console.log(`   Pattern: ${CONFIG.FILE_PATTERN}`);
      console.log(`   Directory: ${CONFIG.DIRECTORY_PATH}`);
      return;
    }

    const auth = await authenticate();
    logger.info("Authenticated with Google Sheets API");

    const recap = new ProcessingRecap();

    // Separate files by type (Windows specific)
    const physicalServers = files.filter((f) => /ty\.svr/i.test(f));
    const virtualMachines = files.filter((f) => /ty\.vm/i.test(f));

    logger.info("File distribution", {
      physicalServers: physicalServers.length,
      virtualMachines: virtualMachines.length,
    });

    // Process physical servers first
    if (physicalServers.length > 0) {
      logger.info("Processing Windows physical servers", {
        count: physicalServers.length,
      });

      for (const file of physicalServers) {
        await processSingleFile(auth, file, recap);
      }

      logger.info("Physical servers processing completed", {
        processed: physicalServers.length,
      });
    }

    // Process virtual machines
    if (virtualMachines.length > 0) {
      logger.info("Processing Windows virtual machines", {
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

    logger.info("All Windows files processed successfully", {
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

export { extractDataFromWindowsTxt, appendToGoogleSheet };