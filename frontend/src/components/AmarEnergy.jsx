// AmarEnergy.jsx
import React, { useState, useRef } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import * as jalaali from "jalaali-js";
import { toPng } from "html-to-image";

// تعریف مستقیم مسیر عکس در پوشه public (نیازی به import نیست)
const logoImage = "/image.webp"; 
const BACKEND_URL = "http://localhost:5000";

const AmarEnergy = () => {
  const [data, setData] = useState([]);
  const [domesticRows, setDomesticRows] = useState([]);
  const [internationalRows, setInternationalRows] = useState([]);
  const [otherRows, setOtherRows] = useState([]);

  const [titleText] = useState("آمار انرژی");
  const [SuplyText] = useState("آمار معاملات ");
  const [loading, setLoading] = useState(false);
  const [statusMessage, setStatusMessage] = useState("");

  // استفاده مستقیم از مسیر تعریف شده
  const [logoBase64] = useState(logoImage);

  // رفرنس‌ها برای خروجی تصویر
  const domesticTableRef = useRef(null);
  const internationalTableRef = useRef(null);

  // تاریخ جلالی برای دریافت با تاریخ دلخواه
  const todayJ = jalaali.toJalaali(new Date());
  const [selectedYear, setSelectedYear] = useState(todayJ.jy);
  const [selectedMonth, setSelectedMonth] = useState(todayJ.jm);
  const [selectedDay, setSelectedDay] = useState(todayJ.jd);
  const daysInMonth = jalaali.jalaaliMonthLength(selectedYear, selectedMonth);

  const MONTH_NAMES = [
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
  ];

  const getSelectedDateString = () => {
    const mm = String(selectedMonth).padStart(2, "0");
    const dd = String(selectedDay).padStart(2, "0");
    return `${selectedYear}/${mm}/${dd}`;
  };

  const columns = [
    "نام تجاری کالا",
    "عرضه کننده",
    "واحد عرضه",
    "حجم عرضه",
    "حجم تقاضا",
    "حجم کل معامله شده",
    "بیشترین قیمت معامله",
    "کمترین قیمت معامله",
    "متوسط قيمت معامله",
    "قیمت پایه",
    "قیمت مبنا",
    "مقصد",
    "نوع معامله",
  ];

  // ------------------------- نرمال‌سازی متن ----------------------------
  const normalizeText = (str) => {
    if (!str) return "";
    return str
      .toString()
      .replace(/ك/g, "ک")
      .replace(/ي/g, "ی")
      .replace(/\u200c/g, " ")
      .replace(/[()]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  };

  // ------------------------- واترمارک اکسل ----------------------------
  const createWatermark = (text1, text2) => {
    const canvas = document.createElement("canvas");
    canvas.width = 3600;
    canvas.height = 1000;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate((-8 * Math.PI) / 180);
    ctx.textAlign = "center";
    ctx.fillStyle = "rgba(140,140,140,0.25)";
    ctx.shadowColor = "rgba(0,0,0,0.25)";
    ctx.shadowBlur = 15;
    ctx.shadowOffsetX = 6;
    ctx.shadowOffsetY = 6;
    ctx.font = "bold 250px Arial";
    ctx.fillText(text1, 0, -180);
    ctx.font = "bold 230px Arial";
    ctx.fillText(text2, 0, 260);
    return canvas.toDataURL("image/png");
  };

  const fixProductName = (name) => {
    if (!name) return "";

    let normalized = name.replace(/ك/g, "ک").replace(/ي/g, "ی");
    normalized = normalized.replace(/حاصل از.*/i, "").trim();
    normalized = normalized
      .replace(/\(([^)]+)\)/g, (match, p1) => {
        if (p1.includes("حاصل")) return "";
        return match;
      })
      .trim();

    const stopWords = ["شرکت", "مجتمع"];
    const parts = normalized.split(" ");
    let filtered = [];

    for (let i = 0; i < parts.length; i++) {
      if (stopWords.includes(parts[i])) break;
      filtered.push(parts[i]);
    }

    const trimmed = name.trim();
    if (
      trimmed.endsWith("عمده") &&
      filtered.length &&
      filtered[filtered.length - 1] !== "(عمده)"
    ) {
      filtered.push("(عمده)");
    }

    return filtered.join(" ");
  };

  const formatPriceBase = (text) => {
    if (!text) return "";
    let str = text.toString();
    str = str.replace(/\([^)]*\)/g, "").trim();
    const match = str.match(/(-?\d+)(.*)/);
    if (match) {
      const isNegative = match[1].startsWith("-");
      let numberPart = match[1].replace("-", "");
      let rest = match[2].trim();
      return `${isNegative ? "-" : ""}${numberPart}${rest}`;
    }
    return str;
  };

  const extractCellValue = (v) => {
    if (v === null || v === undefined) return "";
    if (typeof v === "object" && v.richText) {
      return v.richText.map((t) => t.text).join("");
    }
    if (typeof v === "object" && v.formula !== undefined) {
      if (v.result !== undefined && v.result !== null) return v.result;
      return "";
    }
    if (v === 0 || v === "0") return 0;
    return v;
  };

  const faDigitsToEn = (s) =>
    s.replace(/[۰-۹]/g, (d) => "۰۱۲۳۴۵۶۷۸۹".indexOf(d));

  const getCleanNumberFromRow = (row, headerMap, colKey) => {
    const cIndex = headerMap[colKey];
    if (!cIndex) return null;

    const raw = extractCellValue(row.getCell(cIndex).value);
    if (raw === null || raw === undefined || raw === "") return 0;
    if (typeof raw === "number") return raw;

    const s = faDigitsToEn(raw.toString())
      .replace(/,/g, "")
      .replace(/\u200c/g, "")
      .replace(/\s+/g, "")
      .trim();

    if (!s) return 0;
    const cleaned = s.replace(/[^\d.-]/g, "");
    if (!cleaned) return 0;

    const num = Number(cleaned);
    return Number.isFinite(num) ? num : 0;
  };

  // ------------------------------------- پردازش بافر اکسل -------------------------------------
  const processExcelBuffer = async (buffer) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const worksheet = workbook.worksheets[0];
    const headerRow = worksheet.getRow(1);

    const realHeaders = {};
    headerRow.eachCell((cell, colNumber) => {
      const norm = normalizeText(cell.value);
      realHeaders[norm] = colNumber;
    });

    const columnAliases = {
      "نام تجاری کالا": ["نام تجاری کالا"],
      "عرضه کننده": ["عرضه کننده"],
      "واحد عرضه": ["واحد عرضه"],
      "حجم عرضه": ["حجم عرضه", "حجم عرضه بر حسب واحد عرضه"],
      "حجم تقاضا": ["حجم تقاضا", "تقاضا(بر حسب واحد عرضه)", "تقاضا بر حسب واحد عرضه"],
      "حجم کل معامله شده": [
        "حجم کل معامله شده بر حسب واحد عرضه",
        "حجم كل معامله شده بر حسب واحد عرضه",
      ],
      "کمترین قیمت معامله": ["کمترین قیمت معامله"],
      "متوسط قيمت معامله": ["متوسط قيمت معامله", "متوسط قیمت معامله"],
      "بیشترین قیمت معامله": ["بیشترین قیمت معامله"],
      "قیمت مبنا": [
        "قیمت مبنا",
        "قيمت مبنا",
        "قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم",
        "قیمت مبنا جزئیات محاسبه",
      ],
      "قیمت پایه": ["قیمت پایه", "قیمت پایه (ریال)"],
      "مقصد": ["مقصد"],
      "نوع معامله": ["نوع معامله", "نوع معاملـه"],
    };

    const headerMap = {};
    Object.keys(columnAliases).forEach((mainKey) => {
      const aliases = columnAliases[mainKey];
      for (let alias of aliases) {
        const normAlias = normalizeText(alias);
        const found = Object.keys(realHeaders).find((h) => h === normAlias);
        if (found) {
          headerMap[mainKey] = realHeaders[found];
          break;
        }
      }
    });

    const jsonData = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      if (headerMap["حجم عرضه"]) {
        const supplyNum = getCleanNumberFromRow(row, headerMap, "حجم عرضه");
        if (supplyNum === 0) return;
      }

      const rowData = {};
      Object.keys(columnAliases).forEach((col) => {
        const colIndex = headerMap[col];
        let value = colIndex ? extractCellValue(row.getCell(colIndex).value) : "";
        if (col === "نام تجاری کالا") value = fixProductName(value);
        rowData[col] = value;
      });

      if (!rowData["نام تجاری کالا"]) return;
      jsonData.push(rowData);
    });

    const domestic = [],
      international = [],
      other = [];

    jsonData.forEach((row) => {
      const dest = normalizeText(row["مقصد"]);
      if (dest.includes("داخلی")) domestic.push(row);
      else if (dest.includes("بین")) international.push(row);
      else other.push(row);
    });

    domestic.sort(sortByProductName);
    international.sort(sortByProductName);
    other.sort(sortByProductName);

    setDomesticRows(domestic);
    setInternationalRows(international);
    setOtherRows(other);

    const emptyRow = Object.fromEntries(columns.map((c) => [c, ""]));
    const createHeaderRow = () => {
      const r = {};
      columns.forEach((col) => (r[col] = col));
      r.__isHeader = true;
      return r;
    };

    const finalData = [];
    finalData.push(createHeaderRow());
    if (domestic.length > 0) finalData.push(...domestic);
    finalData.push(emptyRow);
    finalData.push(createHeaderRow());
    if (international.length > 0) finalData.push(...international);
    if (other.length > 0) finalData.push(...other);

    setData(finalData);
    return { domestic, international, other };
  };

  // دریافت آمار امروز
  const handleFetchToday = async () => {
    try {
      setLoading(true);
      setStatusMessage("در حال دریافت اکسل آمار امروز از بورس انرژی...");
      const res = await fetch(`${BACKEND_URL}/api/amar/excel`);
      if (!res.ok) {
        const errJson = await res.json().catch(() => ({}));
        throw new Error(errJson.message || `خطا در دریافت: ${res.status}`);
      }
      const arrayBuffer = await res.arrayBuffer();
      const result = await processExcelBuffer(arrayBuffer);
      if (result.domestic.length + result.international.length + result.other.length === 0) {
        setStatusMessage("فایل دریافت شد اما ردیفی با حجم عرضه غیر صفر یافت نشد.");
      } else {
        setStatusMessage("فایل آمار امروز با موفقیت دریافت و پردازش شد.");
      }
    } catch (err) {
      console.error(err);
      setStatusMessage(`خطا: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  // دریافت با تاریخ انتخابی
  const handleFetchByDate = async () => {
    const dateStr = getSelectedDateString();
    try {
      setLoading(true);
      setStatusMessage(`در حال دریافت اکسل آمار تاریخ ${dateStr}...`);
      const res = await fetch(
        `${BACKEND_URL}/api/amar/excel?date=${encodeURIComponent(dateStr)}`
      );
      if (!res.ok) {
        const errJson = await res.json().catch(() => ({}));
        throw new Error(errJson.message || `خطا در دریافت تاریخ ${dateStr}: ${res.status}`);
      }
      const arrayBuffer = await res.arrayBuffer();
      const result = await processExcelBuffer(arrayBuffer);
      if (result.domestic.length + result.international.length + result.other.length === 0) {
        setStatusMessage("فایل دریافت شد اما ردیفی با حجم عرضه غیر صفر یافت نشد.");
      } else {
        setStatusMessage(`آمار تاریخ ${dateStr} دریافت و پردازش شد.`);
      }
    } catch (err) {
      console.error(err);
      setStatusMessage(`خطا: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  // آپلود دستی
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      setLoading(true);
      const arrayBuffer = await file.arrayBuffer();
      await processExcelBuffer(arrayBuffer);
      setStatusMessage("فایل آپلودی پردازش شد.");
    } catch (err) {
      setStatusMessage(`خطا در پردازش فایل: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const numericColumns = [
    "حجم عرضه",
    "حجم تقاضا",
    "کمترین قیمت معامله",
    "متوسط قيمت معامله",
    "بیشترین قیمت معامله",
    "حجم کل معامله شده",
    "قیمت پایه",
    "قیمت مبنا",
  ];

  const toNumber = (val) => {
    if (val === null || val === undefined || val === "") return null;
    if (typeof val === "number") return val;
    const cleaned = faDigitsToEn(val.toString())
      .replace(/,/g, "")
      .replace(/[^\d.-]/g, "");
    const num = Number(cleaned);
    return isNaN(num) ? null : num;
  };

  const sortByProductName = (a, b) => {
    const nameA = (a["نام تجاری کالا"] || "")
      .toString()
      .replace(/ي/g, "ی")
      .replace(/ك/g, "ک")
      .trim();
    const nameB = (b["نام تجاری کالا"] || "")
      .toString()
      .replace(/ي/g, "ی")
      .replace(/ك/g, "ک")
      .trim();
    return nameA.localeCompare(nameB, "fa");
  };

  const getTodayShamsi = () => {
    const today = new Date();
    const { jy, jm, jd } = jalaali.toJalaali(today);
    const dayNames = ["یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه"];
    const dayOfWeek = dayNames[today.getDay()];
    return {
      day: dayOfWeek,
      date: `${jy}/${jm.toString().padStart(2, "0")}/${jd.toString().padStart(2, "0")}`,
    };
  };

  // محاسبه درصد رقابت برای جدول تصویر
  const calculateCompetition = (avgPrice, basePrice) => {
    const avg = toNumber(avgPrice);
    const base = toNumber(basePrice);
    if (avg && base && base > 0 && avg > base) {
      return (((avg - base) / base) * 100).toFixed(1) + "%";
    }
    return "0%";
  };

  const isCellEmpty = (v) =>
    v === null || v === undefined || v === "" || v === 0 || v === "0" || v === "-";

  // -------------------------------------------------- تابع خروجی اکسل --------------------------------------------------
  const exportExcel = async () => {
    if (data.length === 0) return;

    const workbook = new ExcelJS.Workbook();

    const autoFitColumns = (worksheet, cols) => {
      const headerRowNumber = 3;

      cols.forEach((col, index) => {
        const colNumber = index + 1;
        let maxLength = 0;

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber < headerRowNumber) return;

          const cell = row.getCell(colNumber);
          let cellText = "";

          if (cell.value !== null && cell.value !== undefined) {
            if (typeof cell.value === "object") {
              if (cell.value.result !== undefined && cell.value.result !== null) {
                cellText = String(cell.value.result);
              } else if (cell.value.richText) {
                cellText = cell.value.richText.map((t) => t.text).join("");
              } else if (cell.value.formula) {
                cellText = "10.0%";
              } else {
                cellText = String(cell.value);
              }
            } else if (typeof cell.value === "number") {
              cellText = cell.value.toLocaleString("en-US");
            } else {
              cellText = String(cell.value);
            }
          }

          if (cellText) {
            const length = cellText.trim().length;
            if (length > maxLength) {
              maxLength = length;
            }
          }
        });

        const calculatedWidth = maxLength + 5;
        worksheet.getColumn(colNumber).width = Math.min(Math.max(calculatedWidth, 12), 60);
      });
    };

    const addMergedHeader = (worksheet, cols, topText, secondText) => {
      const lastColLetter = String.fromCharCode(64 + cols.length);
      worksheet.mergeCells(`A1:${lastColLetter}1`);
      const topCell = worksheet.getCell("A1");
      topCell.value = topText;
      topCell.font = { name: "Arial", bold: true, size: 14 };
      topCell.alignment = { horizontal: "center", vertical: "middle" };
      topCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0FAF8E" } };
      worksheet.getRow(1).height = 30;
      worksheet.getRow(1).eachCell((cell) => {
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });

      worksheet.mergeCells(`A2:${lastColLetter}2`);
      const secondCell = worksheet.getCell("A2");
      secondCell.value = secondText;
      secondCell.font = { name: "Arial", bold: true, size: 18 };
      secondCell.alignment = { horizontal: "center", vertical: "middle" };
      secondCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3FD9B8" } };
      worksheet.getRow(2).height = 40;
      worksheet.getRow(2).eachCell((cell) => {
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });

      const headerRow = worksheet.addRow(cols.map((col) => col));
      headerRow.eachCell((cell) => {
        cell.font = { name: "Arial", bold: true, size: 12 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFCAC8C9" } };
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });
      worksheet.getRow(headerRow.number).height = 28;
      return headerRow.number;
    };

    const addMergedHeaderWithColors = (worksheet, cols, topText, secondText, colors) => {
      const lastColLetter = String.fromCharCode(64 + cols.length);

      worksheet.mergeCells(`A1:${lastColLetter}1`);
      const topCell = worksheet.getCell("A1");
      topCell.value = topText;
      topCell.font = { name: "Arial", bold: true, size: 14 };
      topCell.alignment = { horizontal: "center", vertical: "middle" };
      topCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.header1 } };
      worksheet.getRow(1).height = 33;
      worksheet.getRow(1).eachCell((cell) => {
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });

      worksheet.mergeCells(`A2:${lastColLetter}2`);
      const secondCell = worksheet.getCell("A2");
      secondCell.value = secondText;
      secondCell.font = { name: "Arial", bold: true, size: 18 };
      secondCell.alignment = { horizontal: "center", vertical: "middle" };
      secondCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.header2 } };
      worksheet.getRow(2).height = 40;
      worksheet.getRow(2).eachCell((cell) => {
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });

      const headerRow = worksheet.addRow(cols.map((col) => col));
      headerRow.eachCell((cell) => {
        cell.font = { name: "Arial", bold: true, size: 12 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });
      worksheet.getRow(headerRow.number).height = 28;
      return headerRow.number;
    };

    const { day, date } = getTodayShamsi();

    // =================== شیت داخلی ===================
    if (domesticRows.length > 0) {
      const internalColumns = [
        "نام تجاری کالا",
        "عرضه کننده",
        "واحد عرضه",
        "حجم عرضه",
        "حجم تقاضا",
        "حجم کل معامله شده",
        "کمترین قیمت معامله",
        "متوسط قيمت معامله",
        "بیشترین قیمت معامله",
        "قیمت پایه",
        "قیمت مبنا",
        "مقصد",
        "نوع معامله",
        "درصد رقابت",
      ];

      const domesticSheet = workbook.addWorksheet("داخلی", {
        properties: { defaultRowHeight: 25 },
        views: [{ rightToLeft: true }],
      });

      const headerRowNumber = addMergedHeader(
        domesticSheet,
        internalColumns,
        "کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی  09123011311 و 61914000-021 (داخلی 401)",
        `آمار معاملات محصولات بورس انرژی در رینگ داخلی روز ${day} مورخ ${date}`
      );

      const wmBase64 = createWatermark("کارگزاری آینده نگر خوارزمی", "09123011311");
      const wmId = workbook.addImage({ base64: wmBase64, extension: "png" });

      domesticSheet.addImage(wmId, {
        tl: { col: 2, row: 2 },
        ext: { width: 900, height: 300 },
        editAs: "oneCell",
      });

      domesticRows.forEach((rowData, i) => {
        const cleanedRow = internalColumns.map((col) => {
          if (col === "درصد رقابت") return "";
          const val = rowData[col];
          if (val === "" || val === null || val === undefined) return "";
          if (val === 0 || val === "0") return 0;
          return val;
        });

        const row = domesticSheet.addRow(cleanedRow);
        const rowNum = row.number;

        const avgPriceColLetter = String.fromCharCode(64 + internalColumns.indexOf("متوسط قيمت معامله") + 1);
        const basePriceColLetter = String.fromCharCode(64 + internalColumns.indexOf("قیمت پایه") + 1);
        const compColIndex = internalColumns.indexOf("درصد رقابت") + 1;

        row.getCell(compColIndex).value = {
          formula: `IFERROR(IF(AND(ISNUMBER(${avgPriceColLetter}${rowNum}),ISNUMBER(${basePriceColLetter}${rowNum}),${basePriceColLetter}${rowNum}<>0),(${avgPriceColLetter}${rowNum}-${basePriceColLetter}${rowNum})/${basePriceColLetter}${rowNum},0),0)`,
        };
        row.getCell(compColIndex).numFmt = "[=0]0%;0.0%";

        const fillColor = i % 2 === 0 ? "FFFFFFFF" : "FFF5F5F5";

        row.eachCell((cell, colNumber) => {
          const colName = internalColumns[colNumber - 1];

          if (numericColumns.includes(colName)) {
            const numValue = toNumber(cell.value);
            if (numValue !== null) {
              cell.value = numValue;
              cell.numFmt = "#,##0";
            }
          }

          if (colName === "درصد رقابت") {
            cell.numFmt = "[=0]0%;0.0%";
          }

          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } };
          cell.font = { name: "Arial", bold: true };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        });
      });

      const internalCompColLetter = String.fromCharCode(64 + internalColumns.indexOf("درصد رقابت") + 1);
      domesticSheet.addConditionalFormatting({
        ref: `${internalCompColLetter}${headerRowNumber + 1}:${internalCompColLetter}${domesticSheet.rowCount}`,
        rules: [
          {
            type: "cellIs",
            operator: "greaterThan",
            formulae: ["0"],
            style: { font: { color: { argb: "FF008000" }, bold: true } },
          },
        ],
      });

      domesticSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber > headerRowNumber) row.height = 25;
      });

      autoFitColumns(domesticSheet, internalColumns);

      const domesticPriceMabnaIndex = internalColumns.indexOf("قیمت مبنا") + 1;
      if (domesticPriceMabnaIndex > 0) {
        let isDomesticMabnaEmpty = true;
        domesticSheet.eachRow((row, rowNumber) => {
          if (rowNumber <= headerRowNumber) return;
          const val = row.getCell(domesticPriceMabnaIndex).value;
          if (!isCellEmpty(val)) isDomesticMabnaEmpty = false;
        });

        if (isDomesticMabnaEmpty) {
          domesticSheet.getColumn(domesticPriceMabnaIndex).hidden = true;
        }
      }
    }

    // =================== شیت بین الملل ===================
    const kashfPremiumRows = [...internationalRows, ...otherRows].filter((row) =>
      row["نوع معامله"]?.includes("کشف پریمیوم")
    );
    const otherNonPremiumRows = [...internationalRows, ...otherRows].filter(
      (row) => !row["نوع معامله"]?.includes("کشف پریمیوم")
    );
    const mainDataRows = [...kashfPremiumRows, ...otherNonPremiumRows];

    if (mainDataRows.length > 0) {
      const mainColumns = [...columns, "درصد رقابت"];

      const mainSheet = workbook.addWorksheet("بین الملل", {
        properties: { defaultRowHeight: 25 },
        views: [{ rightToLeft: true }],
      });

      const headerRowNumber = addMergedHeaderWithColors(
        mainSheet,
        mainColumns,
        "کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311 و 61914000-021 (داخلی 401)",
        `آمار معاملات محصولات بورس انرژی در رینگ صادراتی روز ${day} مورخ ${date}`,
        { header1: "FFFF6A00", header2: "FFFFB873", header3: "FFD9D9D9" }
      );

      const wmBase64 = createWatermark("کارگزاری آینده نگر خوارزمی", "09123011311");
      const wmId = workbook.addImage({ base64: wmBase64, extension: "png" });

      mainSheet.addImage(wmId, {
        tl: { col: 2, row: 2 },
        ext: { width: 900, height: 300 },
        editAs: "oneCell",
      });

      mainDataRows.forEach((rowData, i) => {
        const rowDataCopy = { ...rowData };
        if (rowDataCopy["قیمت پایه"]) rowDataCopy["قیمت پایه"] = formatPriceBase(rowDataCopy["قیمت پایه"]);

        const cleanedRow = mainColumns.map((col) => {
          if (col === "درصد رقابت") return "";
          const val = rowDataCopy[col];
          if (val === "" || val === null || val === undefined) return "";
          if (val === 0 || val === "0") return 0;
          return val;
        });

        const row = mainSheet.addRow(cleanedRow);
        const rowNum = row.number;

        const avgPriceColLetter = String.fromCharCode(64 + mainColumns.indexOf("متوسط قيمت معامله") + 1);
        const basePriceColLetter = String.fromCharCode(64 + mainColumns.indexOf("قیمت پایه") + 1);
        const compColIndex = mainColumns.indexOf("درصد رقابت") + 1;

        row.getCell(compColIndex).value = {
          formula: `IFERROR(IF(AND(ISNUMBER(${avgPriceColLetter}${rowNum}),ISNUMBER(${basePriceColLetter}${rowNum}),${basePriceColLetter}${rowNum}<>0),(${avgPriceColLetter}${rowNum}-${basePriceColLetter}${rowNum})/${basePriceColLetter}${rowNum},0),0)`,
        };
        row.getCell(compColIndex).numFmt = "[=0]0%;0.0%";

        const fillColor = i % 2 === 0 ? "FFFFFFFF" : "FFFAF7F3";

        row.eachCell((cell, colNumber) => {
          const colName = mainColumns[colNumber - 1];

          if (numericColumns.includes(colName)) {
            const numValue = toNumber(cell.value);
            if (numValue !== null) {
              cell.value = numValue;
              cell.numFmt = "#,##0";
            }
          }

          if (colName === "درصد رقابت") {
            cell.numFmt = "[=0]0%;0.0%";
          }

          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } };
          cell.font = { name: "Arial", bold: true };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        });
      });

      const mainCompColLetter = String.fromCharCode(64 + mainColumns.indexOf("درصد رقابت") + 1);
      mainSheet.addConditionalFormatting({
        ref: `${mainCompColLetter}${headerRowNumber + 1}:${mainCompColLetter}${mainSheet.rowCount}`,
        rules: [
          {
            type: "cellIs",
            operator: "greaterThan",
            formulae: ["0"],
            style: { font: { color: { argb: "FF008000" }, bold: true } },
          },
        ],
      });

      mainSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber > headerRowNumber) row.height = 25;
      });

      autoFitColumns(mainSheet, mainColumns);

      const intlPriceMabnaIndex = mainColumns.indexOf("قیمت مبنا") + 1;
      if (intlPriceMabnaIndex > 0) {
        let isIntlMabnaEmpty = true;
        mainSheet.eachRow((row, rowNumber) => {
          if (rowNumber <= headerRowNumber) return;
          const val = row.getCell(intlPriceMabnaIndex).value;
          if (!isCellEmpty(val)) isIntlMabnaEmpty = false;
        });

        if (isIntlMabnaEmpty) {
          mainSheet.getColumn(intlPriceMabnaIndex).hidden = true;
        }
      }
    }

    const buf = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buf]), "Enerzhi_Output_Separated.xlsx");
  };

  // ------------------------- توابع خروجی عکس ----------------------------
  const handleExportDomesticImage = async () => {
    if (!domesticTableRef.current || domesticRows.length === 0) {
      alert("داده‌ای برای رینگ داخلی وجود ندارد.");
      return;
    }
    try {
      setStatusMessage("در حال تولید عکس رینگ داخلی...");
      const dataUrl = await toPng(domesticTableRef.current, {
        quality: 0.95,
        pixelRatio: 2,
        backgroundColor: "#ffffff",
      });
      saveAs(dataUrl, `amar-dakheli-${Date.now()}.png`);
      setStatusMessage("عکس رینگ داخلی با موفقیت دانلود شد.");
    } catch (err) {
      console.error(err);
      setStatusMessage("خطا در تولید عکس رینگ داخلی");
    }
  };

  const handleExportInternationalImage = async () => {
    const intlData = [...internationalRows, ...otherRows];
    if (!internationalTableRef.current || intlData.length === 0) {
      alert("داده‌ای برای رینگ بین‌الملل وجود ندارد.");
      return;
    }
    try {
      setStatusMessage("در حال تولید عکس رینگ بین‌الملل...");
      const dataUrl = await toPng(internationalTableRef.current, {
        quality: 0.95,
        pixelRatio: 2,
        backgroundColor: "#ffffff",
      });
      saveAs(dataUrl, `amar-intl-${Date.now()}.png`);
      setStatusMessage("عکس رینگ بین‌الملل با موفقیت دانلود شد.");
    } catch (err) {
      console.error(err);
      setStatusMessage("خطا در تولید عکس رینگ بین‌الملل");
    }
  };

  const { day, date } = getTodayShamsi();
  const allIntlRows = [...internationalRows, ...otherRows];

  const isDomesticMabnaAllEmpty =
    domesticRows.length === 0 || domesticRows.every((r) => isCellEmpty(r["قیمت مبنا"]));
  const isIntlMabnaAllEmpty =
    allIntlRows.length === 0 || allIntlRows.every((r) => isCellEmpty(r["قیمت مبنا"]));
  const isGeneralMabnaAllEmpty = isDomesticMabnaAllEmpty && isIntlMabnaAllEmpty;

  const displayColumnsGeneral = isGeneralMabnaAllEmpty
    ? columns.filter((c) => c !== "قیمت مبنا")
    : columns;
  const domesticColumns = isDomesticMabnaAllEmpty
    ? columns.filter((c) => c !== "قیمت مبنا")
    : columns;
  const intlColumns = isIntlMabnaAllEmpty
    ? columns.filter((c) => c !== "قیمت مبنا")
    : columns;

  const getDynamicCellStyle = (isHeader = false) => {
    return {
      padding: isHeader ? "8px 12px" : "6px 12px",
      fontSize: "11px",
      border: "1px solid #cbd5e1",
      verticalAlign: "middle",
      textAlign: "center",
      whiteSpace: "nowrap",
      height: "26px",
      backgroundColor: isHeader ? undefined : "rgba(255, 255, 255, 0.9)",
    };
  };

  const watermarkContainerStyle = {
    position: "relative",
    display: "inline-block",
    width: "max-content",
    backgroundColor: "#ffffff",
    padding: "20px",
    direction: "rtl",
    fontFamily: "Tahoma, Arial",
    boxSizing: "border-box",
  };

  const watermarkImgStyle = {
    position: "absolute",
    top: "50%",
    left: "50%",
    transform: "translate(-50%, -50%)",
    width: "45%",
    opacity: 0.08,
    zIndex: 0,
    pointerEvents: "none",
  };

  return (
    <div style={{ padding: "20px", fontFamily: "Tahoma, Arial, sans-serif", direction: "rtl" }}>
      <h2 style={{ color: "#0f766e" }}>{titleText}</h2>
      <h4 style={{ color: "#475569" }}>{SuplyText}</h4>

      <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginTop: "15px" }}>
        <button
          onClick={handleFetchToday}
          disabled={loading}
          style={{
            backgroundColor: "#0f766e",
            color: "#fff",
            border: "none",
            padding: "10px 18px",
            borderRadius: "6px",
            cursor: loading ? "not-allowed" : "pointer",
            fontWeight: "bold",
          }}
        >
          {loading ? "در حال دریافت..." : "📥 دریافت آمار امروز"}
        </button>

        <div style={{ display: "flex", gap: "5px", alignItems: "center", backgroundColor: "#f1f5f9", padding: "6px 12px", borderRadius: "6px" }}>
          <span>تاریخ:</span>
          <select value={selectedYear} onChange={(e) => setSelectedYear(Number(e.target.value))}>
            {[1402, 1403, 1404, 1405].map((y) => (
              <option key={y} value={y}>{y}</option>
            ))}
          </select>
          <select value={selectedMonth} onChange={(e) => setSelectedMonth(Number(e.target.value))}>
            {MONTH_NAMES.map((name, i) => (
              <option key={i + 1} value={i + 1}>{name}</option>
            ))}
          </select>
          <select value={selectedDay} onChange={(e) => setSelectedDay(Number(e.target.value))}>
            {Array.from({ length: daysInMonth }, (_, i) => i + 1).map((d) => (
              <option key={d} value={d}>{d}</option>
            ))}
          </select>
          <button
            onClick={handleFetchByDate}
            disabled={loading}
            style={{ backgroundColor: "#2563eb", color: "#fff", border: "none", padding: "6px 12px", borderRadius: "4px", cursor: "pointer" }}
          >
            دریافت این تاریخ
          </button>
        </div>

        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} disabled={loading} style={{ fontSize: "12px" }} />
      </div>

      <div style={{ display: "flex", gap: "10px", marginTop: "15px", flexWrap: "wrap" }}>
        <button
          onClick={exportExcel}
          disabled={data.length === 0}
          style={{
            backgroundColor: "#16a34a",
            color: "#fff",
            border: "none",
            padding: "10px 18px",
            borderRadius: "6px",
            cursor: data.length === 0 ? "not-allowed" : "pointer",
            fontWeight: "bold",
          }}
        >
          📊 خروجی اکسل
        </button>

        <button
          onClick={handleExportDomesticImage}
          disabled={domesticRows.length === 0}
          style={{
            backgroundColor: "#0891b2",
            color: "#fff",
            border: "none",
            padding: "10px 18px",
            borderRadius: "6px",
            cursor: domesticRows.length === 0 ? "not-allowed" : "pointer",
            fontWeight: "bold",
          }}
        >
          📸 خروجی عکس رینگ داخلی ({domesticRows.length})
        </button>

        <button
          onClick={handleExportInternationalImage}
          disabled={allIntlRows.length === 0}
          style={{
            backgroundColor: "#ea580c",
            color: "#fff",
            border: "none",
            padding: "10px 18px",
            borderRadius: "6px",
            cursor: allIntlRows.length === 0 ? "not-allowed" : "pointer",
            fontWeight: "bold",
          }}
        >
          📸 خروجی عکس رینگ بین‌الملل ({allIntlRows.length})
        </button>
      </div>

      {statusMessage && (
        <div style={{ marginTop: "15px", padding: "10px", backgroundColor: "#e0f2fe", color: "#0369a1", borderRadius: "6px", fontSize: "14px" }}>
          {statusMessage}
        </div>
      )}

      <div style={{ overflowX: "auto", marginTop: "20px" }}>
        <table style={{ width: "max-content", minWidth: "100%", borderCollapse: "collapse", tableLayout: "auto" }}>
          <thead>
            <tr>
              {displayColumnsGeneral.map((col, idx) => (
                <th
                  key={idx}
                  style={{
                    ...getDynamicCellStyle(true),
                    backgroundColor: "#e2e8f0",
                    fontWeight: "bold",
                    color: "#1e293b",
                  }}
                >
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, idx) => (
              <tr key={idx} style={{ backgroundColor: row.__isHeader ? "#f8fafc" : "#fff" }}>
                {displayColumnsGeneral.map((col, i) => (
                  <td
                    key={i}
                    style={{
                      ...getDynamicCellStyle(false),
                      color: "#334155",
                    }}
                  >
                    {numericColumns.includes(col) && row[col] && typeof row[col] === "number"
                      ? row[col].toLocaleString()
                      : row[col]}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* کانتینرهای پنهان تولید عکس */}
      <div style={{ position: "absolute", left: "-9999px", top: "-9999px" }}>
        
        <div ref={domesticTableRef} style={watermarkContainerStyle}>
          {logoBase64 && <img src={logoBase64} alt="watermark" style={watermarkImgStyle} />}
          
          <div style={{ position: "relative", zIndex: 1 }}>
            <div style={{ backgroundColor: "#0FAF8E", color: "#ffffff", padding: "10px 16px", textAlign: "center", fontWeight: "bold", fontSize: "14px" }}>
              کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311 و 61914000-021 (داخلی 401)
            </div>
            <div style={{ backgroundColor: "#3FD9B8", color: "#064e3b", padding: "12px 16px", textAlign: "center", fontWeight: "bold", fontSize: "16px", borderBottom: "2px solid #0f766e" }}>
              آمار معاملات محصولات بورس انرژی در رینگ داخلی روز {day} مورخ {date}
            </div>
            <table style={{ width: "max-content", minWidth: "100%", borderCollapse: "collapse", tableLayout: "auto", marginTop: "2px" }}>
              <thead>
                <tr style={{ backgroundColor: "#CAC8C9", color: "#000" }}>
                  {domesticColumns.map((c, i) => (
                    <th key={i} style={{ ...getDynamicCellStyle(true), border: "1px solid #94a3b8", fontWeight: "bold", backgroundColor: "rgba(202, 200, 201, 0.9)" }}>
                      {c}
                    </th>
                  ))}
                  <th style={{ ...getDynamicCellStyle(true), border: "1px solid #94a3b8", fontWeight: "bold", backgroundColor: "rgba(202, 200, 201, 0.9)" }}>
                    درصد رقابت
                  </th>
                </tr>
              </thead>
              <tbody>
                {domesticRows.map((r, i) => {
                  const comp = calculateCompetition(r["متوسط قيمت معامله"], r["قیمت پایه"]);
                  return (
                    <tr key={i}>
                      {domesticColumns.map((c, j) => (
                        <td key={j} style={getDynamicCellStyle(false)}>
                          {numericColumns.includes(c) && r[c] ? Number(r[c]).toLocaleString() : r[c]}
                        </td>
                      ))}
                      <td style={{ ...getDynamicCellStyle(false), color: comp !== "0%" ? "#16a34a" : "#000", fontWeight: "bold" }}>
                        {comp}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        <div ref={internationalTableRef} style={watermarkContainerStyle}>
          {logoBase64 && <img src={logoBase64} alt="watermark" style={watermarkImgStyle} />}

          <div style={{ position: "relative", zIndex: 1 }}>
            <div style={{ backgroundColor: "#FF6A00", color: "#ffffff", padding: "10px 16px", textAlign: "center", fontWeight: "bold", fontSize: "14px" }}>
              کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311 و 61914000-021 (داخلی 401)
            </div>
            <div style={{ backgroundColor: "#FFB873", color: "#7c2d12", padding: "12px 16px", textAlign: "center", fontWeight: "bold", fontSize: "16px", borderBottom: "2px solid #ea580c" }}>
              آمار معاملات محصولات بورس انرژی در رینگ صادراتی روز {day} مورخ {date}
            </div>
            <table style={{ width: "max-content", minWidth: "100%", borderCollapse: "collapse", tableLayout: "auto", marginTop: "2px" }}>
              <thead>
                <tr style={{ backgroundColor: "#D9D9D9", color: "#000" }}>
                  {intlColumns.map((c, i) => (
                    <th key={i} style={{ ...getDynamicCellStyle(true), border: "1px solid #94a3b8", fontWeight: "bold", backgroundColor: "rgba(217, 217, 217, 0.9)" }}>
                      {c}
                    </th>
                  ))}
                  <th style={{ ...getDynamicCellStyle(true), border: "1px solid #94a3b8", fontWeight: "bold", backgroundColor: "rgba(217, 217, 217, 0.9)" }}>
                    درصد رقابت
                  </th>
                </tr>
              </thead>
              <tbody>
                {allIntlRows.map((r, i) => {
                  const comp = calculateCompetition(r["متوسط قيمت معامله"], r["قیمت پایه"]);
                  return (
                    <tr key={i}>
                      {intlColumns.map((c, j) => (
                        <td key={j} style={getDynamicCellStyle(false)}>
                          {numericColumns.includes(c) && r[c] ? Number(r[c]).toLocaleString() : r[c]}
                        </td>
                      ))}
                      <td style={{ ...getDynamicCellStyle(false), color: comp !== "0%" ? "#16a34a" : "#000", fontWeight: "bold" }}>
                        {comp}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

      </div>
    </div>
  );
};

export default AmarEnergy;
