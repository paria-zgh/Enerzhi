import React, { useRef, useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { toPng } from 'html-to-image';
import * as jalaali from 'jalaali-js';

const FONT_NAME = 'Shabnam';
// ارتفاع ثابت و هماهنگ تمام ردیف‌های داده (غیر از ۳ ردیف اول هدر)
const DATA_ROW_HEIGHT_PX = 56;
const EXCEL_DATA_ROW_HEIGHT = 40;

const EnerzhiExcelTransformer = () => {
  const [data, setData] = useState([]);
  const [domesticRows, setDomesticRows] = useState([]);
  const [internationalRows, setInternationalRows] = useState([]);
  const [otherRows, setOtherRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');

  // رِف‌های اختصاصی جداول برای خروجی عکس PNG
  const domesticTableRef = useRef(null);
  const internationalTableRef = useRef(null);

  const titleText = 'لیست انرژی';
  const supplyText = 'عرضه‌ها';

  const columns = [
    'نام کالا',
    'تولید کننده',
    'حجم عرضه',
    'حداکثر افزایش حجم عرضه',
    'زمان عرضه',
    'قیمت پایه',
    'قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم',
    'نوع معامله',
    'حداقل خرید',
    'سررسید کالا',
    'نوع بسته بندی',
    'محل تحویل',
    'مقصد',
  ];

  const basePriceKey = 'قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم';

  const internalColumns = [
    'نام کالا',
    'تولید کننده',
    'حجم عرضه',
    'حداکثر افزایش حجم عرضه',
    'زمان عرضه',
    'قیمت پایه',
    'نوع معامله',
    'حداقل خرید',
    'سررسید کالا',
    'نوع بسته بندی',
    'محل تحویل',
  ];

  const displayCol = (col) => {
    if (col === 'حداکثر افزایش حجم عرضه') return 'افزایش عرضه';
    if (col === basePriceKey) return 'قيمت مبنا';
    return col;
  };

  const getCellValue = (value) => {
    if (value === null || value === undefined) return '';

    if (typeof value === 'object') {
      if (value.text !== undefined) return value.text;
      if (value.result !== undefined) return value.result;
      if (value.richText) {
        return value.richText.map((x) => x.text || '').join('');
      }
      if (value.formula !== undefined) return value.result || '';
      if (value instanceof Date) return value.toLocaleString('fa-IR');
      return value.toString();
    }

    return value;
  };

  const createWatermark = (text1, text2) => {
    const canvas = document.createElement('canvas');
    canvas.width = 3600;
    canvas.height = 1000;

    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate((-8 * Math.PI) / 180);

    ctx.textAlign = 'center';
    ctx.fillStyle = 'rgba(0,70,180,0.25)';

    ctx.shadowColor = 'rgba(0,0,0,0.25)';
    ctx.shadowBlur = 15;
    ctx.shadowOffsetX = 6;
    ctx.shadowOffsetY = 6;

    ctx.font = `bold 250px ${FONT_NAME}, Arial, sans-serif`;
    ctx.fillText(text1, 0, -180);

    ctx.font = `bold 230px ${FONT_NAME}, Arial, sans-serif`;
    ctx.fillText(text2, 0, 260);

    return canvas.toDataURL('image/png');
  };

  const fixProductName = (name) => {
    if (!name) return '';

    let normalized = name
      .toString()
      .replace(/ك/g, 'ک')
      .replace(/ي/g, 'ی');

    normalized = normalized
      .replace(/\s*\(.*?حاصل از.*?\)/gi, '')
      .trim();

    normalized = normalized.replace(/حاصل از.*/i, '').trim();

    const stopWords = ['شرکت', 'مجتمع'];
    const parts = normalized.split(/\s+/);
    const filtered = [];

    for (let i = 0; i < parts.length; i++) {
      if (stopWords.includes(parts[i])) break;
      filtered.push(parts[i]);
    }

    const originalName = name.toString().trim();

    if (
      originalName.endsWith('عمده') &&
      filtered.length > 0 &&
      filtered[filtered.length - 1] !== '(عمده)'
    ) {
      filtered.push('(عمده)');
    }

    return filtered.join(' ').replace(/\s+/g, ' ').trim();
  };

  const formatPriceBase = (text) => {
    if (!text) return '';

    let str = text.toString();
    str = str.replace(/\([^)]*\)/g, '').trim();

    const match = str.match(/(-?\d+)(.*)/);
    if (match) {
      const isNegative = match[1].startsWith('-');
      const numberPart = match[1].replace('-', '');
      const rest = match[2].trim();

      return `${isNegative ? '-' : ''}${numberPart}${rest}`;
    }

    return str;
  };

  const sortByProductName = (a, b) => {
    const nameA = (a['نام کالا'] || '')
      .toString()
      .replace(/ي/g, 'ی')
      .replace(/ك/g, 'ک')
      .trim();

    const nameB = (b['نام کالا'] || '')
      .toString()
      .replace(/ي/g, 'ی')
      .replace(/ك/g, 'ک')
      .trim();

    return nameA.localeCompare(nameB, 'fa');
  };

  const getTomorrowShamsi = () => {
    const today = new Date();
    const targetDate = new Date(today);
    const dayOfWeek = today.getDay(); // 0: یکشنبه, 1: دوشنبه, ..., 3: چهارشنبه, 4: پنجشنبه, 5: جمعه, 6: شنبه
  
    // اگر چهارشنبه باشد: 3 روز اضافه شود تا به شنبه برسد
    if (dayOfWeek === 3) {
      targetDate.setDate(today.getDate() + 3);
    } 
    // اگر پنجشنبه باشد: 2 روز اضافه شود
    else if (dayOfWeek === 4) {
      targetDate.setDate(today.getDate() + 2);
    } 
    // اگر جمعه باشد: 1 روز اضافه شود
    else if (dayOfWeek === 5) {
      targetDate.setDate(today.getDate() + 1);
    } 
    // سایر روزها (شنبه تا سه‌شنبه): همان روز بعد
    else {
      targetDate.setDate(today.getDate() + 1);
    }
  
    const { jy, jm, jd } = jalaali.toJalaali(targetDate);
  
    const dayNames = [
      'یکشنبه',
      'دوشنبه',
      'سه‌شنبه',
      'چهارشنبه',
      'پنجشنبه',
      'جمعه',
      'شنبه',
    ];
  
    return {
      day: dayNames[targetDate.getDay()],
      date: `${jy}/${String(jm).padStart(2, '0')}/${String(jd).padStart(2, '0')}`,
    };
  };
  
  const processExcelFile = async (fileOrBlob) => {
    const buffer = await fileOrBlob.arrayBuffer();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const worksheet = workbook.worksheets[0];

    if (!worksheet) {
      throw new Error('هیچ شیتی در فایل اکسل پیدا نشد.');
    }

    const headerRow = worksheet.getRow(1);
    const headerMap = {};

    headerRow.eachCell((cell, colNumber) => {
      const header = getCellValue(cell.value).toString().trim();
      if (header) headerMap[header] = colNumber;
    });

    const jsonData = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      const rowData = {};

      columns.forEach((column) => {
        const columnIndex = headerMap[column];

        let value = columnIndex
          ? getCellValue(row.getCell(columnIndex).value)
          : '';

        if (column === 'نام کالا') {
          value = fixProductName(value);
        }

        rowData[column] = value;
      });

      const hasValue = columns.some((column) => {
        const value = rowData[column];

        return (
          value !== null &&
          value !== undefined &&
          value.toString().trim() !== ''
        );
      });

      if (hasValue) jsonData.push(rowData);
    });

    const domestic = [];
    const international = [];
    const other = [];

    jsonData.forEach((row) => {
      const destination = (row['مقصد'] || '')
        .toString()
        .trim()
        .replace(/ي/g, 'ی')
        .replace(/ك/g, 'ک');

      if (destination.includes('داخلی')) {
        domestic.push(row);
      } else if (destination.includes('بین')) {
        international.push(row);
      } else {
        other.push(row);
      }
    });

    domestic.sort(sortByProductName);
    international.sort(sortByProductName);
    other.sort(sortByProductName);

    const finalData = [...domestic, ...international, ...other];

    setData(finalData);
    setDomesticRows(domestic);
    setInternationalRows(international);
    setOtherRows(other);

    return { domestic, international, other };
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setLoading(true);
      setMessage('در حال پردازش فایل اکسل...');

      await processExcelFile(file);

      setMessage('فایل با موفقیت پردازش شد. حالا خروجی موردنظر را بگیر.');
    } catch (error) {
      console.error('Manual file processing error:', error);
      setMessage(`خطا در پردازش فایل: ${error.message}`);
    } finally {
      setLoading(false);
    }
  };

  const getInternationalRowsOrdered = () => {
    const allMainRows = [...internationalRows, ...otherRows];

    const kashfPremiumRows = allMainRows.filter((row) =>
      (row['نوع معامله'] || '').toString().includes('کشف پریمیوم')
    );

    const otherNonPremiumRows = allMainRows.filter(
      (row) => !(row['نوع معامله'] || '').toString().includes('کشف پریمیوم')
    );

    return [...kashfPremiumRows, ...otherNonPremiumRows].map((r) => {
      const rowDataCopy = { ...r };
      if (rowDataCopy['قیمت پایه']) {
        let base = rowDataCopy['قیمت پایه'].toString();
        base = base.replace(
          '(مرکز مبادله- مطابق مکانیزم مصوب هیئت مدیره بورس انرژی ایران)',
          ''
        ).trim();
        rowDataCopy['قیمت پایه'] = formatPriceBase(base);
      }
      return rowDataCopy;
    });
  };

  const getInternationalColumns = () => {
    const rows = getInternationalRowsOrdered();
    const cols = [...columns];

    const idx = cols.findIndex((c) => c === basePriceKey);
    if (idx !== -1) {
      const allEmpty = rows.every((row) => {
        const v = row[basePriceKey];
        return !v || v.toString().trim() === '';
      });
      if (allEmpty) cols.splice(idx, 1);
    }
    return cols;
  };

  const exportDomesticImage = async () => {
    if (domesticRows.length === 0) {
      setMessage('هیچ داده‌ای در رینگ داخلی وجود ندارد.');
      return;
    }
    if (!domesticTableRef.current) return;

    try {
      setLoading(true);
      setMessage('در حال ساخت عکس از جدول داخلی...');
      const dataUrl = await toPng(domesticTableRef.current, {
        cacheBust: true,
        backgroundColor: '#ffffff',
        pixelRatio: 2,
      });
      saveAs(dataUrl, 'Enerzhi_Domestic_Table.png');
      setMessage('عکس جدول داخلی با موفقیت دانلود شد.');
    } catch (error) {
      console.error('Domestic image error:', error);
      setMessage(`خطا در ساخت عکس: ${error.message}`);
    } finally {
      setLoading(false);
    }
  };

  const exportInternationalImage = async () => {
    const intlRows = getInternationalRowsOrdered();
    if (intlRows.length === 0) {
      setMessage('هیچ داده‌ای در رینگ بین‌الملل وجود ندارد.');
      return;
    }
    if (!internationalTableRef.current) return;

    try {
      setLoading(true);
      setMessage('در حال ساخت عکس از جدول بین‌الملل...');
      const dataUrl = await toPng(internationalTableRef.current, {
        cacheBust: true,
        backgroundColor: '#ffffff',
        pixelRatio: 2,
      });
      saveAs(dataUrl, 'Enerzhi_International_Table.png');
      setMessage('عکس جدول بین‌الملل با موفقیت دانلود شد.');
    } catch (error) {
      console.error('International image error:', error);
      setMessage(`خطا در ساخت عکس: ${error.message}`);
    } finally {
      setLoading(false);
    }
  };

  const autoFitColumns = (worksheet, worksheetColumns) => {
    const skipRows = 2;

    worksheetColumns.forEach((column, index) => {
      let maxLength = 1;

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber <= skipRows) return;

        const cell = row.getCell(index + 1);
        let value = cell.value;

        if (value && typeof value === 'object' && value.richText) {
          value = value.richText.map((item) => item.text || '').join('');
        } else if (
          value !== null &&
          value !== undefined &&
          typeof value !== 'string'
        ) {
          value = value.toString();
        }

        if (value) {
          maxLength = Math.max(maxLength, value.toString().length);
        }
      });

      worksheet.getColumn(index + 1).width = Math.min(
        Math.max(maxLength + 2, 12),
        45
      );
    });
  };

  const addMergedHeader = (worksheet, worksheetColumns, topText, secondText) => {
    const lastColLetter = worksheet.getColumn(worksheetColumns.length).letter;

    worksheet.mergeCells(`A1:${lastColLetter}1`);

    const topCell = worksheet.getCell('A1');
    topCell.value = topText;
    topCell.font = { name: FONT_NAME, bold: true, size: 14 };
    topCell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      readingOrder: 2,
    };
    topCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF6E8BC5' },
    };

    worksheet.getRow(1).height = 33;
    worksheet.getRow(1).eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    worksheet.mergeCells(`A2:${lastColLetter}2`);

    const secondCell = worksheet.getCell('A2');
    secondCell.value = secondText;
    secondCell.font = { name: FONT_NAME, bold: true, size: 18 };
    secondCell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      readingOrder: 2,
    };
    secondCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFB3C6E7' },
    };

    worksheet.getRow(2).height = 40;
    worksheet.getRow(2).eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    const headerRow = worksheet.addRow(worksheetColumns);

    headerRow.eachCell((cell) => {
      cell.font = { name: FONT_NAME, bold: true, size: 12 };
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        readingOrder: 2,
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFCAC8C9' },
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    worksheet.getRow(headerRow.number).height = 28;
    return headerRow.number;
  };

  const addMergedHeaderWithColors = (
    worksheet,
    worksheetColumns,
    topText,
    secondText,
    colors
  ) => {
    const lastColLetter = worksheet.getColumn(worksheetColumns.length).letter;

    worksheet.mergeCells(`A1:${lastColLetter}1`);

    const topCell = worksheet.getCell('A1');
    topCell.value = topText;
    topCell.font = { name: FONT_NAME, bold: true, size: 14 };
    topCell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      readingOrder: 2,
    };
    topCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: colors.header1 },
    };

    worksheet.getRow(1).height = 33;
    worksheet.getRow(1).eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    worksheet.mergeCells(`A2:${lastColLetter}2`);

    const secondCell = worksheet.getCell('A2');
    secondCell.value = secondText;
    secondCell.font = { name: FONT_NAME, bold: true, size: 18 };
    secondCell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      readingOrder: 2,
    };
    secondCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: colors.header2 },
    };

    worksheet.getRow(2).height = 40;
    worksheet.getRow(2).eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    const headerRow = worksheet.addRow(worksheetColumns);

    headerRow.eachCell((cell) => {
      cell.font = { name: FONT_NAME, bold: true, size: 12 };
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        readingOrder: 2,
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: colors.header3 },
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    worksheet.getRow(headerRow.number).height = 28;
    return headerRow.number;
  };

  const exportExcel = async (processedData = null) => {
    const domestic = processedData?.domestic || domesticRows;
    const international = processedData?.international || internationalRows;
    const other = processedData?.other || otherRows;

    if (
      domestic.length === 0 &&
      international.length === 0 &&
      other.length === 0
    ) {
      setMessage('هیچ داده‌ای برای خروجی وجود ندارد.');
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const { day, date } = getTomorrowShamsi();

    // ================= شیت داخلی =================
    if (domestic.length > 0) {
      const displayInternalColumns = internalColumns.map((col) =>
        col === 'حداکثر افزایش حجم عرضه' ? 'افزایش عرضه' : col
      );

      const domesticSheet = workbook.addWorksheet('داخلی', {
        properties: { defaultRowHeight: EXCEL_DATA_ROW_HEIGHT },
        views: [{ rightToLeft: true }],
      });

      const headerRowNumber = addMergedHeader(
        domesticSheet,
        displayInternalColumns,
        'کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی  09123011311 و 61914000-021 (داخلی 401)',
        `عرضه محصولات بورس انرژی در رینگ داخلی روز ${day} مورخ ${date}`
      );

      const wmBase64 = createWatermark(
        'کارگزاری آینده نگر خوارزمی',
        '09123011311'
      );

      const wmId = workbook.addImage({
        base64: wmBase64,
        extension: 'png',
      });

      domesticSheet.addImage(wmId, {
        tl: { col: 2, row: 2 },
        ext: { width: 900, height: 300 },
        editAs: 'oneCell',
      });

      domestic.forEach((rowData, i) => {
        const row = domesticSheet.addRow(
          internalColumns.map((col) => {
            let value = rowData[col] ?? '';
            value = value.toString();
            return '\u200F' + value;
          })
        );

        const fillColor = i % 2 === 0 ? 'FFFFFFFF' : 'FFF5F5F5';

        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: fillColor },
          };
          cell.font = { name: FONT_NAME, bold: true };
          cell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            readingOrder: 2,
            wrapText: true,
          };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        });
      });

      domesticSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber > headerRowNumber) row.height = EXCEL_DATA_ROW_HEIGHT;
      });

      autoFitColumns(domesticSheet, internalColumns);
    }

    // ================= شیت بین‌الملل =================
    const mainDataRows = getInternationalRowsOrdered();
    const mainColumns = getInternationalColumns();

    if (mainDataRows.length > 0) {
      const mainSheet = workbook.addWorksheet('بین الملل', {
        properties: { defaultRowHeight: EXCEL_DATA_ROW_HEIGHT },
        views: [{ rightToLeft: true }],
      });

      const headerRowNumber = addMergedHeaderWithColors(
        mainSheet,
        mainColumns.map(displayCol),
        'کارگزاری آینده نگر خوارزمی 09123011311 و 61914000-021 (داخلی 401)',
        `عرضه محصولات بورس انرژی در رینگ صادراتی روز ${day} مورخ ${date}`,
        {
          header1: 'FFCC2B52',
          header2: 'FFD9A299',
          header3: 'FFCAC8C9',
        }
      );

      const wmBase64 = createWatermark(
        'کارگزاری آینده نگر خوارزمی',
        '09123011311'
      );

      const wmId = workbook.addImage({
        base64: wmBase64,
        extension: 'png',
      });

      mainSheet.addImage(wmId, {
        tl: { col: 2, row: 2 },
        ext: { width: 900, height: 300 },
        editAs: 'oneCell',
      });

      mainDataRows.forEach((rowDataCopy, i) => {
        const row = mainSheet.addRow(
          mainColumns.map((col) => {
            let value = rowDataCopy[col] ?? '';
            value = value.toString();
            return '\u200F' + value;
          })
        );

        const fillColor = i % 2 === 0 ? 'FFFFFFFF' : 'FFFAF7F3';

        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: fillColor },
          };
          cell.font = { name: FONT_NAME, bold: true };
          cell.alignment = {
            horizontal: 'center',
            vertical: 'middle',
            readingOrder: 2,
            wrapText: true,
          };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        });
      });

      mainSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber > headerRowNumber) row.height = EXCEL_DATA_ROW_HEIGHT;
      });

      autoFitColumns(mainSheet, mainColumns);
    }

    const buffer = await workbook.xlsx.writeBuffer();

    saveAs(
      new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
      'Enerzhi_Output_Separated.xlsx'
    );
  };

  const downloadAndTransformExcel = async () => {
    try {
      setLoading(true);
      setMessage(
        'در حال دریافت فایل اکسل بورس انرژی (ممکن است تا ۱ دقیقه زمان ببرد)...'
      );

      const response = await fetch('/api/offers/excel');

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(errorText || 'دریافت فایل از بک‌اند ناموفق بود.');
      }

      const rawExcelBlob = await response.blob();

      if (
        rawExcelBlob.type.includes('text') ||
        rawExcelBlob.type.includes('html')
      ) {
        const errorHtml = await rawExcelBlob.text();
        throw new Error(
          'پاسخ سرور فایل اکسل معتبر نیست: ' + errorHtml.slice(0, 100)
        );
      }

      const processedData = await processExcelFile(rawExcelBlob);
      await exportExcel(processedData);

      setMessage('فایل دریافت، پردازش و با موفقیت دانلود شد.');
    } catch (error) {
      console.error('Download and transform error:', error);
      setMessage(`خطا: ${error.message}`);
    } finally {
      setLoading(false);
    }
  };

  const intlTableRows = getInternationalRowsOrdered();
  const intlTableCols = getInternationalColumns();
  const { day: dDay, date: dDate } = getTomorrowShamsi();

  const watermarkUrl = createWatermark('کارگزاری آینده نگر خوارزمی', '09123011311');

  const htmlFont = `${FONT_NAME}, Arial, sans-serif`;

  return (
    <div style={{ padding: 20, fontFamily: htmlFont, direction: 'rtl' }}>
      <h2>{titleText}</h2>
      <h4>{supplyText}</h4>

      <div
        style={{
          display: 'flex',
          gap: 10,
          flexWrap: 'wrap',
          alignItems: 'center',
        }}
      >
        <button
          onClick={downloadAndTransformExcel}
          disabled={loading}
          style={{
            padding: '10px 20px',
            fontFamily: htmlFont,
            cursor: loading ? 'not-allowed' : 'pointer',
          }}
        >
          {loading ? 'در حال آماده‌سازی...' : 'دریافت و تبدیل اکسل بورس انرژی'}
        </button>

        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileUpload}
          disabled={loading}
        />

        <button
          onClick={() => exportExcel()}
          disabled={
            loading ||
            (domesticRows.length === 0 &&
              internationalRows.length === 0 &&
              otherRows.length === 0)
          }
          style={{ padding: '10px 20px', fontFamily: htmlFont, cursor: 'pointer' }}
        >
          خروجی اکسل کامل
        </button>

        <button
          onClick={exportDomesticImage}
          disabled={loading || domesticRows.length === 0}
          style={{
            padding: '10px 20px',
            fontFamily: htmlFont,
            backgroundColor: '#2e7d32',
            color: '#fff',
            border: 'none',
            borderRadius: '4px',
            cursor: domesticRows.length === 0 ? 'not-allowed' : 'pointer',
          }}
        >
          خروجی عکس شیت داخلی (PNG)
        </button>

        <button
          onClick={exportInternationalImage}
          disabled={loading || intlTableRows.length === 0}
          style={{
            padding: '10px 20px',
            fontFamily: htmlFont,
            backgroundColor: '#c2185b',
            color: '#fff',
            border: 'none',
            borderRadius: '4px',
            cursor: intlTableRows.length === 0 ? 'not-allowed' : 'pointer',
          }}
        >
          خروجی عکس شیت بین‌الملل (PNG)
        </button>
      </div>

      {message && (
        <p
          style={{
            marginTop: 15,
            color: message.startsWith('خطا') ? 'red' : 'green',
          }}
        >
          {message}
        </p>
      )}

      {/* پیش‌نمایش جدول داخلی */}
      {domesticRows.length > 0 && (
        <div style={{ marginTop: 30 }}>
          <h3 style={{ fontFamily: htmlFont }}>پیش‌نمایش شیت داخلی (مشابه فایل اکسل)</h3>
          <div
            ref={domesticTableRef}
            style={{
              position: 'relative',
              overflow: 'hidden',
              border: '1px solid #ccc',
              backgroundColor: '#ffffff',
              minWidth: 1200,
            }}
          >
            {/* ردیف اول هدر */}
            <div
              style={{
                position: 'relative',
                zIndex: 1,
                background: '#6E8BC5',
                color: '#000',
                fontFamily: htmlFont,
                fontWeight: 700,
                fontSize: 14,
                textAlign: 'center',
                padding: '10px 8px',
                borderBottom: '1px solid #999',
              }}
            >
              کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311 و 61914000-021 (داخلی 401)
            </div>

            {/* ردیف دوم هدر */}
            <div
              style={{
                position: 'relative',
                zIndex: 1,
                background: '#B3C6E7',
                color: '#000',
                fontFamily: htmlFont,
                fontWeight: 700,
                fontSize: 18,
                textAlign: 'center',
                padding: '12px 8px',
                borderBottom: '1px solid #999',
              }}
            >
              {`عرضه محصولات بورس انرژی در رینگ داخلی روز ${dDay} مورخ ${dDate}`}
            </div>

            {/* بدنه جدول همراه واترمارک */}
            <div style={{ position: 'relative' }}>
              <div
                style={{
                  position: 'absolute',
                  inset: 0,
                  backgroundImage: `url(${watermarkUrl})`,
                  backgroundRepeat: 'no-repeat',
                  backgroundPosition: 'center center',
                  backgroundSize: '800px auto',
                  pointerEvents: 'none',
                  zIndex: 0,
                }}
              />

              <table
                style={{
                  width: '100%',
                  borderCollapse: 'collapse',
                  tableLayout: 'fixed',
                  fontFamily: htmlFont,
                  fontWeight: 700,
                  position: 'relative',
                  zIndex: 1,
                }}
              >
                {/* ردیف سوم هدر (سرستون‌ها) */}
                <thead>
                  <tr style={{ height: 40 }}>
                    {internalColumns.map((col) => (
                      <th
                        key={col}
                        style={{
                          background: '#CAC8C9',
                          border: '1px solid #999',
                          padding: '6px 8px',
                          textAlign: 'center',
                          verticalAlign: 'middle',
                          whiteSpace: 'nowrap',
                          fontSize: 12,
                        }}
                      >
                        {displayCol(col)}
                      </th>
                    ))}
                  </tr>
                </thead>
                {/* ردیف‌های داده با ارتفاع یکدست و ثابت */}
                <tbody>
                  {domesticRows.map((row, i) => (
                    <tr
                      key={i}
                      style={{
                        height: DATA_ROW_HEIGHT_PX,
                        background: i % 2 === 0 ? '#fff' : '#F5F5F5',
                      }}
                    >
                      {internalColumns.map((col) => (
                        <td
                          key={col}
                          style={{
                            border: '1px solid #999',
                            padding: '4px 6px',
                            textAlign: 'center',
                            verticalAlign: 'middle',
                            fontSize: 12,
                            height: DATA_ROW_HEIGHT_PX,
                            boxSizing: 'border-box',
                          }}
                        >
                          <div
                            style={{
                              height: '100%',
                              maxHeight: DATA_ROW_HEIGHT_PX - 8,
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              overflow: 'hidden',
                              wordBreak: 'break-word',
                              lineHeight: '1.3',
                            }}
                          >
                            {'\u200F' + (row[col] ?? '').toString()}
                          </div>
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* پیش‌نمایش جدول بین‌الملل */}
      {intlTableRows.length > 0 && (
        <div style={{ marginTop: 30 }}>
          <h3 style={{ fontFamily: htmlFont }}>پیش‌نمایش شیت بین‌الملل (مشابه فایل اکسل)</h3>
          <div
            ref={internationalTableRef}
            style={{
              position: 'relative',
              overflow: 'hidden',
              border: '1px solid #ccc',
              backgroundColor: '#ffffff',
              minWidth: 1200,
            }}
          >
            {/* ردیف اول هدر */}
            <div
              style={{
                position: 'relative',
                zIndex: 1,
                background: '#CC2B52',
                color: '#000',
                fontFamily: htmlFont,
                fontWeight: 700,
                fontSize: 14,
                textAlign: 'center',
                padding: '10px 8px',
                borderBottom: '1px solid #999',
              }}
            >
              کارگزاری آینده نگر خوارزمی 09123011311 و 61914000-021 (داخلی 401)
            </div>

            {/* ردیف دوم هدر */}
            <div
              style={{
                position: 'relative',
                zIndex: 1,
                background: '#D9A299',
                color: '#000',
                fontFamily: htmlFont,
                fontWeight: 700,
                fontSize: 18,
                textAlign: 'center',
                padding: '12px 8px',
                borderBottom: '1px solid #999',
              }}
            >
              {`عرضه محصولات بورس انرژی در رینگ صادراتی روز ${dDay} مورخ ${dDate}`}
            </div>

            {/* بدنه جدول همراه واترمارک */}
            <div style={{ position: 'relative' }}>
              <div
                style={{
                  position: 'absolute',
                  inset: 0,
                  backgroundImage: `url(${watermarkUrl})`,
                  backgroundRepeat: 'no-repeat',
                  backgroundPosition: 'center center',
                  backgroundSize: '800px auto',
                  pointerEvents: 'none',
                  zIndex: 0,
                }}
              />

              <table
                style={{
                  width: '100%',
                  borderCollapse: 'collapse',
                  tableLayout: 'fixed',
                  fontFamily: htmlFont,
                  fontWeight: 700,
                  position: 'relative',
                  zIndex: 1,
                }}
              >
                {/* ردیف سوم هدر (سرستون‌ها) */}
                <thead>
                  <tr style={{ height: 40 }}>
                    {intlTableCols.map((col) => (
                      <th
                        key={col}
                        style={{
                          background: '#CAC8C9',
                          border: '1px solid #999',
                          padding: '6px 8px',
                          textAlign: 'center',
                          verticalAlign: 'middle',
                          whiteSpace: 'nowrap',
                          fontSize: 12,
                        }}
                      >
                        {displayCol(col)}
                      </th>
                    ))}
                  </tr>
                </thead>
                {/* ردیف‌های داده با ارتفاع یکدست و ثابت */}
                <tbody>
                  {intlTableRows.map((row, i) => (
                    <tr
                      key={i}
                      style={{
                        height: DATA_ROW_HEIGHT_PX,
                        background: i % 2 === 0 ? '#fff' : '#FAF7F3',
                      }}
                    >
                      {intlTableCols.map((col) => (
                        <td
                          key={col}
                          style={{
                            border: '1px solid #999',
                            padding: '4px 6px',
                            textAlign: 'center',
                            verticalAlign: 'middle',
                            fontSize: 12,
                            height: DATA_ROW_HEIGHT_PX,
                            boxSizing: 'border-box',
                          }}
                        >
                          <div
                            style={{
                              height: '100%',
                              maxHeight: DATA_ROW_HEIGHT_PX - 8,
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              overflow: 'hidden',
                              wordBreak: 'break-word',
                              lineHeight: '1.3',
                            }}
                          >
                            {'\u200F' + (row[col] ?? '').toString()}
                          </div>
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default EnerzhiExcelTransformer;
