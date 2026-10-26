import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import * as jalaali from 'jalaali-js';


const EnerzhiExcelTransformer = () => {
  const [data, setData] = useState([]);
  const [domesticRows, setDomesticRows] = useState([]);
  const [internationalRows, setInternationalRows] = useState([]);
  const [otherRows, setOtherRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');

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

  const getCellValue = (value) => {
    if (value === null || value === undefined) return '';

    if (typeof value === 'object') {
      if (value.text !== undefined) return value.text;
      if (value.result !== undefined) return value.result;
      if (value.richText) return value.richText.map((x) => x.text || '').join('');
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

    ctx.font = 'bold 250px B Nazanin';
    ctx.fillText(text1, 0, -180);

    ctx.font = 'bold 230px B Nazanin';
    ctx.fillText(text2, 0, 260);

    return canvas.toDataURL('image/png');
  };

  const fixProductName = (name) => {
    if (!name) return '';

    let normalized = name.toString()
      .replace(/ك/g, 'ک')
      .replace(/ي/g, 'ی');

    normalized = normalized.replace(/\s*\(.*?حاصل از.*?\)/gi, '').trim();
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
    str = str.replace(/\\([^)]*\\)/g, '').trim();

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

      if (hasValue) {
        jsonData.push(rowData);
      }
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



 
    const finalData = [];

    if (domestic.length > 0) {
        finalData.push(...domestic);
    }
    
    if (international.length > 0) {
        finalData.push(...international);
    }
    
    if (other.length > 0) {
        finalData.push(...other);
    }
    
    setData(finalData);
    setDomesticRows(domestic);
    setInternationalRows(international);
    setOtherRows(other);
    
    return {
        domestic,
        international,
        other
    };
    

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

      setMessage('فایل با موفقیت پردازش شد. حالا خروجی اکسل را بگیر.');
    } catch (error) {
      console.error('Manual file processing error:', error);
      setMessage(`خطا در پردازش فایل: ${error.message}`);
    } finally {
      setLoading(false);
    }
  };

  const getTomorrowShamsi = () => {
    const today = new Date();

    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);

    const { jy, jm, jd } = jalaali.toJalaali(tomorrow);

    const dayNames = [
        'یکشنبه',
        'دوشنبه',
        'سه‌شنبه',
        'چهارشنبه',
        'پنجشنبه',
        'جمعه',
        'شنبه'
    ];

    return {
        day: dayNames[tomorrow.getDay()],
        date: `${jy}/${String(jm).padStart(2, '0')}/${String(jd).padStart(2, '0')}`
    };
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
        } else if (value !== null && value !== undefined && typeof value !== 'string') {
          value = value.toString();
        }

        if (value) {
          const length = value.toString().length;
          if (length > maxLength) maxLength = length;
        }
      });

      worksheet.getColumn(index + 1).width = Math.min(Math.max(maxLength + 2, 12), 45);
    });
  };

  const addMergedHeader = (worksheet, worksheetColumns, topText, secondText) => {
    const lastColLetter = worksheet.getColumn(worksheetColumns.length).letter;

    worksheet.mergeCells(`A1:${lastColLetter}1`);
    const topCell = worksheet.getCell('A1');
    topCell.value = topText;
    topCell.font = { name: 'B Nazanin', bold: true, size: 14 };
    topCell.alignment = { horizontal: 'center', vertical: 'middle', readingOrder: 2 };
    topCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6E8BC5' } };

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
    secondCell.font = { name: 'B Nazanin', bold: true, size: 18 };
    secondCell.alignment = { horizontal: 'center', vertical: 'middle', readingOrder: 2 };
    secondCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB3C6E7' } };

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
      cell.font = { name: 'B Nazanin', bold: true, size: 12 };
      cell.alignment = { horizontal: 'center', vertical: 'middle', readingOrder: 2 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCAC8C9' } };
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
    topCell.font = { name: 'B Nazanin', bold: true, size: 14 };
    topCell.alignment = { horizontal: 'center', vertical: 'middle', readingOrder: 2 };
    topCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.header1 } };

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
    secondCell.font = { name: 'B Nazanin', bold: true, size: 18 };
    secondCell.alignment = { horizontal: 'center', vertical: 'middle', readingOrder: 2 };
    secondCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.header2 } };

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
      cell.font = { name: 'B Nazanin', bold: true, size: 12 };
      cell.alignment = { horizontal: 'center', vertical: 'middle', readingOrder: 2 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.header3 } };
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

      const displayInternalColumns = internalColumns.map((col) =>
        col === 'حداکثر افزایش حجم عرضه' ? 'افزایش عرضه' : col
      );

      const domesticSheet = workbook.addWorksheet('داخلی', {
        properties: { defaultRowHeight: 25 },
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
            value = '\u200F' + value;
            return value;
          })
        );

        const fillColor = i % 2 === 0 ? 'FFFFFFFF' : 'FFF5F5F5';

        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: fillColor },
          };
          cell.font = { name: 'B Nazanin', bold: true };
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
        if (rowNumber > headerRowNumber) row.height = 25;
      });

      autoFitColumns(domesticSheet, internalColumns);
    }

    // ================= شیت بین الملل =================
    const basePriceKey = 'قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم';

    const allMainRows = [...international, ...other];

    const kashfPremiumRows = allMainRows.filter((row) =>
      (row['نوع معامله'] || '').toString().includes('کشف پریمیوم')
    );

    const otherNonPremiumRows = allMainRows.filter(
      (row) => !(row['نوع معامله'] || '').toString().includes('کشف پریمیوم')
    );

    const mainDataRows = [...kashfPremiumRows, ...otherNonPremiumRows];
    let mainColumns = [...columns];

    const priceColumnIndex = mainColumns.findIndex((col) => col === basePriceKey);

    if (priceColumnIndex !== -1) {
      const allEmpty = mainDataRows.every(
        (row) => !row[basePriceKey] || row[basePriceKey].toString().trim() === ''
      );

      if (allEmpty) {
        mainColumns.splice(priceColumnIndex, 1);
      }
    }

    if (mainDataRows.length > 0) {
      const mainSheet = workbook.addWorksheet('بین الملل', {
        properties: { defaultRowHeight: 25 },
        views: [{ rightToLeft: true }],
      });

      const displayColumnsMap = (col) => {
        if (col === 'حداکثر افزایش حجم عرضه') return 'افزایش عرضه';
        if (col === basePriceKey) return 'قيمت مبنا';
        return col;
      };

      const headerRowNumber = addMergedHeaderWithColors(
        mainSheet,
        mainColumns.map(displayColumnsMap),
        'کارگزاری آینده نگر00-021 (داخلی  ذوقی 09123011311 و 61914000-021 (داخلی 401)',
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

      mainDataRows.forEach((rowData, i) => {
        const rowDataCopy = { ...rowData };

        if (rowDataCopy['قیمت پایه']) {
          let base = rowDataCopy['قیمت پایه'].toString();

          base = base.replace(
            '(مرکز مبادله- مطابق مکانیزم مصوب هیئت مدیره بورس انرژی ایران)',
            ''
          ).trim();

          rowDataCopy['قیمت پایه'] = formatPriceBase(base);
        }

        const row = mainSheet.addRow(
          mainColumns.map((col) => {
            let value = rowDataCopy[col] ?? '';
            value = value.toString();
            value = '\u200F' + value;
            return value;
          })
        );

        const fillColor = i % 2 === 0 ? 'FFFFFFFF' : 'FFFAF7F3';

        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: fillColor },
          };
          cell.font = { name: 'B Nazanin', bold: true };
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
        if (rowNumber > headerRowNumber) row.height = 25;
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
      setMessage('در حال دریافت فایل اکسل بورس انرژی (ممکن است تا ۱ دقیقه زمان ببرد)...');

      // آدرس کامل بک‌اند روی Render:
      const response = await fetch('https://enerzhi-3.onrender.com/api/offers/excel');

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(errorText || 'دریافت فایل از بک‌اند ناموفق بود.');
      }

      const rawExcelBlob = await response.blob();

      // چک کردن اینکه آیا واقعاً اکسل دانلود شده یا متن خطا برگشته:
      if (rawExcelBlob.type.includes('text') || rawExcelBlob.type.includes('html')) {
        const errorHtml = await rawExcelBlob.text();
        throw new Error('پاسخ سرور فایل اکسل معتبر نیست: ' + errorHtml.slice(0, 100));
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


  return (
    <div style={{ padding: 20, fontFamily: 'Arial', direction: 'rtl' }}>
      <h2>{titleText}</h2>
      <h4>{supplyText}</h4>

      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
        <button
          onClick={downloadAndTransformExcel}
          disabled={loading}
          style={{
            padding: '10px 20px',
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
          style={{ padding: '10px 20px', cursor: 'pointer' }}
        >
          خروجی اکسل فایل انتخاب‌شده
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

      <div style={{ overflowX: 'auto', marginTop: 20 }}>
        <table
          style={{
            width: '100%',
            borderCollapse: 'collapse',
            minWidth: 1200,
          }}
        >
          <thead>
            <tr>
              {columns.map((column, index) => (
                <th
                  key={index}
                  style={{
                    backgroundColor: '#D3D3D3',
                    padding: '8px',
                    border: '1px solid #ccc',
                    whiteSpace: 'nowrap',
                  }}
                >
                  {column}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {columns.map((column, columnIndex) => (
                  <td
                    key={columnIndex}
                    style={{
                      padding: '8px',
                      border: '1px solid #ccc',
                      textAlign: 'center',
                    }}
                  >
                    {row[column]}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default EnerzhiExcelTransformer;
