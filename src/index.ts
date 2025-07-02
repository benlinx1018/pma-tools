#!/usr/bin/env node

import { input } from "@inquirer/prompts";
import ExcelJS from "exceljs";
import * as fs from "fs";

async function main() {
  let config = await loadConfig();

  const targetWb = new ExcelJS.Workbook();
  await targetWb.xlsx.readFile(config.targetFileName);
  const targetSheet = targetWb.getWorksheet(config.targetSheetName);

  if (!targetSheet) {
    console.log(
      `找不到工作表 "${config.targetSheetName}"，請檢查 config.json 中的設定`
    );
    process.exit(1);
  }
  const targetHeaderRow = targetSheet.getRow(2);

  const targetHeaderColIdMap = new Map<string, number>();
  targetHeaderRow.eachCell((cell, colNumber) => {
    const cellVal = getCellVal(cell);
    if (cellVal) {
      targetHeaderColIdMap.set(cellVal, colNumber);
    }
  });

  const targetIdColIndex = targetHeaderColIdMap.get(
    config.targetIdentifierColumnName
  );

  console.log(
    `目標識別欄位名稱:${config.targetIdentifierColumnName} ,索引位置:${targetIdColIndex}`
  );
  const targetUpdateColIndex = targetHeaderColIdMap.get(
    config.targetUpdateColumnName
  );
  console.log(
    `目標更新欄位名稱:${config.targetUpdateColumnName} ,索引位置:${targetUpdateColIndex}`
  );
  if (!targetIdColIndex || !targetUpdateColIndex) {
    throw new Error("目標欄位設定異常，請檢查 config.json");
  }

  const sourceMap = new Map<string, Map<string, ExcelJS.CellValue>>();
  for (const source of config.sourceFiles) {
    const srcWb = new ExcelJS.Workbook();
    await srcWb.xlsx.readFile(source.fileName);
    const srcSheet = srcWb.worksheets[0];
    const srcHeaderRow = srcSheet.getRow(2);
    const srcHeaderColIdMap = new Map<string, number>();
    srcHeaderRow.eachCell((cell, colNumber) => {
      const cellVal = getCellVal(cell);
      if (cellVal) {
        srcHeaderColIdMap.set(cellVal, colNumber);
      }
    });

    const srcIdColIdx = srcHeaderColIdMap.get(
      config.targetIdentifierColumnName
    );
    const srcValColIdx = srcHeaderColIdMap.get(config.targetUpdateColumnName);
    if (!srcIdColIdx || !srcValColIdx) {
      // console.log(
      //   `❌ srcIdColIdx 或 srcValColIdx 未找到，跳過來源檔案「${source.fileName}」`
      // );
      continue;
    }

    const sourceIdMap = new Map<string, ExcelJS.CellValue>();
    for (let r = 3; r <= srcSheet.actualRowCount; r++) {
      const srcRow = srcSheet.getRow(r);

      const matched = source.criteria.every((c) => {
        const srcColHeaderIndex = srcHeaderColIdMap.get(c.headerName);
        if (!srcColHeaderIndex) return false;
        const value = getCellVal(srcRow.getCell(srcColHeaderIndex));

        return c.targetValues.includes(value);
      });

      if (!matched) {
        // console.log(
        //   `❌ 第${rowNum}筆資料不符合來源檔案「${source.fileName}」的條件，跳過`
        // );
        continue;
      }

      const key = getCellVal(srcRow.getCell(srcIdColIdx));
      if (key) {
        const rolColValue = getCellVal(srcRow.getCell(srcValColIdx));
        sourceIdMap.set(key, rolColValue);
      }
    }
    sourceMap.set(source.fileName, sourceIdMap);
  }

  for (let rowNum = 3; rowNum <= targetSheet.actualRowCount; rowNum++) {
    // console.log(`開始檢查第${rowNum}筆資料...`);
    const targetRow = targetSheet.getRow(rowNum);
    const targetIdColValue = getCellVal(targetRow.getCell(targetIdColIndex));

    if (targetIdColValue == "[object Object]") {
      console.log(
        `targetIdColValue異常`,
        JSON.stringify(targetRow.getCell(targetIdColIndex).value)
      );
    }

    if (!targetIdColValue) {
      console.log(`❌ 目標識別欄位值為空，跳過第${rowNum}筆資料`);
      continue;
    }

    for (const source of config.sourceFiles) {
      // 檢查 criteria 是否符合
      const matched = source.criteria.every((c) => {
        const targetColHeaderIndex = targetHeaderColIdMap.get(c.headerName);
        if (!targetColHeaderIndex) {
          return false;
        }
        const value = getCellVal(targetRow.getCell(targetColHeaderIndex));

        return c.targetValues.includes(value);
      });

      if (!matched) {
        // console.log(
        //   `❌ 第${rowNum}筆資料不符合來源檔案「${source.fileName}」的條件，跳過`
        // );
        continue;
      }

      const sourceCache = sourceMap.get(source.fileName);
      if (sourceCache) {
        const srcRowVal = sourceCache.get(targetIdColValue);
        if (srcRowVal) {
          const updateVal = srcRowVal;
          const targetCell = targetRow.getCell(targetUpdateColIndex);
          const targetCellValue = getCellVal(targetCell);

          if (targetCellValue != updateVal) {
            targetCell.value = updateVal;
            safelySetCellFill(targetCell, config.highlightColor);
            console.log(
              `✅ 更新第${rowNum}筆資料${config.targetIdentifierColumnName}:${targetIdColValue} ${config.targetColumnName} = ${updateVal} (source: ${source.fileName})`
            );
          }
        }
      }
    }
  }

  await targetWb.xlsx.writeFile(config.targetFileName);
  console.log("📄 寫入完成：" + config.targetFileName);
}

main().catch(console.error);

interface Criteria {
  headerName: string;
  targetValues: string[];
}

interface SourceFile {
  fileName: string;
  criteria: Criteria[];
}

interface Config {
  highlightColor: string;
  targetFileName: string;
  targetSheetName: string;
  targetUpdateColumnName: string;
  targetIdentifierColumnName: string;
  sourceFiles: SourceFile[];
}

async function readFileName(
  message: string,
  defaultPath: string
): Promise<string> {
  while (true) {
    const targetFileName = await input({
      message: message,
      default: defaultPath,
    });

    if (fs.existsSync(targetFileName)) {
      return targetFileName;
    }

    console.error("❌ 找不到檔案，請確認路徑是否正確\n");
  }
}

async function loadConfig(): Promise<Config> {
  let configPath = "./config.json";

  if (!fs.existsSync(configPath)) {
    console.warn(`⚠️ : ${configPath}`);
    configPath = await readFileName(
      `⚠️預設設定檔(${configPath})載入失敗，請輸入 config.json 的路徑`,
      configPath
    );
  }

  const config = JSON.parse(fs.readFileSync(configPath, "utf-8"));

  return config;
}

function getCellVal(cell: ExcelJS.Cell): string {
  const val = cell?.value;

  if (val === null || val === undefined) return "";

  if (typeof val === "object" && "result" in val) {
    return val.result?.toString?.() ?? ""; // 如果公式有 result 就用它
  }

  return val.toString?.() ?? "";
}

function safelySetCellFill(cell: ExcelJS.Cell, highlightColor: string) {
  const originalStyle = JSON.parse(JSON.stringify(cell.style || {}));

  // 強制斷開原樣式的連結，再建立新樣式物件
  cell.style = {
    ...originalStyle,
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: highlightColor },
    },
  };
}
