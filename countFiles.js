// @ts-nocheck
import fs from "fs";
import path from "path";

const parentFolders = ["draft", "done", "notes"]; // ubah sesuai kebutuhan
const baseDir = "."; // root direktori

function countFilesAndTypes(dirPath) {
  let count = 0;
  const typeCounts = { vm: 0, svr: 0, none: 0 };

  const items = fs.readdirSync(dirPath, { withFileTypes: true });

  for (const item of items) {
    const fullPath = path.join(dirPath, item.name);

    if (item.isDirectory()) {
      const nested = countFilesAndTypes(fullPath);
      count += nested.count;
      typeCounts.vm += nested.typeCounts.vm;
      typeCounts.svr += nested.typeCounts.svr;
      typeCounts.none += nested.typeCounts.none;
    } else if (item.isFile()) {
      count++;
      // cari pola "ty.xxx" dalam nama file
      const match = item.name.match(/ty\.([a-zA-Z0-9-]+)/);
      if (match) {
        const ty = match[1].toLowerCase();
        if (ty.includes("vm")) typeCounts.vm++;
        else if (ty.includes("svr")) typeCounts.svr++;
        else typeCounts.none++;
      } else {
        typeCounts.none++;
      }
    }
  }

  return { count, typeCounts };
}

function printTypeSummary(typeCounts) {
  return `(vm: ${typeCounts.vm}, svr: ${typeCounts.svr}, none: ${typeCounts.none})`;
}

function main() {
  let grandTotal = 0;
  const grandTypeTotals = { vm: 0, svr: 0, none: 0 };

  for (const parent of parentFolders) {
    const parentPath = path.join(baseDir, parent);
    if (!fs.existsSync(parentPath)) {
      console.warn(`⚠️ Folder '${parent}' tidak ditemukan, dilewati.\n`);
      continue;
    }

    console.log(`${parent}/`);
    const subs = fs.readdirSync(parentPath, { withFileTypes: true }).filter(d => d.isDirectory());

    let parentTotal = 0;
    const parentTypeTotals = { vm: 0, svr: 0, none: 0 };

    for (const sub of subs) {
      const subPath = path.join(parentPath, sub.name);
      const { count, typeCounts } = countFilesAndTypes(subPath);

      parentTotal += count;
      parentTypeTotals.vm += typeCounts.vm;
      parentTypeTotals.svr += typeCounts.svr;
      parentTypeTotals.none += typeCounts.none;

      console.log(`  ${sub.name}: ${count} files ${printTypeSummary(typeCounts)}`);
    }

    console.log(`  ➤ Total ${parent}: ${parentTotal} files ${printTypeSummary(parentTypeTotals)}\n`);

    grandTotal += parentTotal;
    grandTypeTotals.vm += parentTypeTotals.vm;
    grandTypeTotals.svr += parentTypeTotals.svr;
    grandTypeTotals.none += parentTypeTotals.none;
  }

  console.log(`📊 Grand Total: ${grandTotal} files ${printTypeSummary(grandTypeTotals)}`);
}

main();
