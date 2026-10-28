// backend/excelProcessor.js

const ExcelJS = require('exceljs');
const jalaali = require('jalaali-js');
const path = require('path');

// --- توابع کمکی ---

function fixProductName(name) {
    if (!name) return '';
    return name.replace(/_/g, ' ');
}

function getPriceAsNumber(price) {
    if (typeof price === 'undefined' || price === null) return 0;
    const priceNum = Number(String(price).replace(/,/g, ''));
    return isNaN(priceNum) ? 0 : priceNum;
}

function getNextTradingDayShamsi() {
    const today = new Date();
    const targetDate = new Date(today);
    const dayOfWeek = today.getDay(); // 0: یکشنبه, 3: چهارشنبه, 4: پنجشنبه, 5: جمعه, 6: شنبه

    let daysToAdd = 1;
    if (dayOfWeek === 3) {
        daysToAdd = 3; // چهارشنبه -> شنبه
    } else if (dayOfWeek === 4) {
        daysToAdd = 2; // پنجشنبه -> شنبه
    } else if (dayOfWeek === 5) {
        daysToAdd = 1; // جمعه -> شنبه
    } else {
        daysToAdd = 1;
    }

    targetDate.setDate(today.getDate() + daysToAdd);

    const jalaaliDate = jalaali.g2j(
        targetDate.getFullYear(),
        targetDate.getMonth() + 1,
        targetDate.getDate()
    );

    const formattedMonth = jalaaliDate.gm < 10 ? `0${jalaaliDate.gm}` : jalaaliDate.gm;
    const formattedDay = jalaaliDate.gd < 10 ? `0${jalaaliDate.gd}` : jalaaliDate.gd;
    
    return `${jalaaliDate.gy}/${formattedMonth}/${formattedDay}`;
}

function sortByProductName(a, b) {
    const nameA = (a['نام کالا'] || '').toLowerCase();
    const nameB = (b['نام کالا'] || '').toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
}

// --- تابع اصلی پردازش ---

async function processExcelData(scrapedData) {
    // ✅ گارد امنیتی: اگر هیچ دیتایی نبود، پردازش نکن و null برگردان
    if (!scrapedData || !Array.isArray(scrapedData) || scrapedData.length === 0) {
        console.warn('⚠️ داده‌ای برای پردازش یافت نشد (لیست خالی است).');
        return null;
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Petki';
    workbook.lastModifiedBy = 'Petki';

    // --- پردازش و دسته‌بندی داده‌ها ---
    let processedData = scrapedData.map(item => ({
        ...item,
        'قیمت پایه': getPriceAsNumber(item['قیمت پایه']),
        'نام کالا': fixProductName(item['نام کالا']),
        'تاریخ فردا': getNextTradingDayShamsi()
    }));

    // فیلتر کردن و مرتب‌سازی
    const internalOffers = processedData.filter(item => item.destination === 'داخلی');
    const internationalOffers = processedData.filter(item => item.destination === 'بین‌الملل');
    const otherOffers = processedData.filter(item => item.destination !== 'داخلی' && item.destination !== 'بین‌الملل');

    internalOffers.sort(sortByProductName);
    internationalOffers.sort(sortByProductName);
    otherOffers.sort(sortByProductName);

    // استایل هدر
    const headerStyle = {
        font: { bold: true, name: 'Calibri', size: 12 },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    };
    const currencyFormat = '#,##0';

    // --- ساخت شیت داخلی ---
    if (internalOffers.length > 0) {
        const internalSheet = workbook.addWorksheet('داخلی', {
            views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }]
        });
        internalSheet.addRows([
            ['تاریخ فردا', 'مقصد', 'نام کالا', 'قیمت پایه (ریال)', 'توضیحات']
        ]);
        internalSheet.getRow(1).style = headerStyle;

        internalOffers.forEach((offer, index) => {
            const rowNum = index + 2;
            internalSheet.addRow([
                offer['تاریخ فردا'],
                offer.destination,
                offer['نام کالا'],
                offer['قیمت پایه'],
                offer.description
            ]);
            const row = internalSheet.getRow(rowNum);
            row.font = { name: 'Calibri', size: 11 };
            internalSheet.getCell(rowNum, 1).alignment = { horizontal: 'center' };
            internalSheet.getCell(rowNum, 2).alignment = { horizontal: 'center' };
            internalSheet.getCell(rowNum, 4).numFmt = currencyFormat;
            internalSheet.getCell(rowNum, 4).alignment = { horizontal: 'right' };
        });

        internalSheet.getColumn(1).width = 15;
        internalSheet.getColumn(2).width = 12;
        internalSheet.getColumn(3).width = 40;
        internalSheet.getColumn(4).width = 18;
        internalSheet.getColumn(5).width = 30;
    }

    // --- ساخت شیت بین‌الملل ---
    if (internationalOffers.length > 0) {
        const internationalSheet = workbook.addWorksheet('بین الملل', {
            views: [{state: 'frozen', xSplit: 1, ySplit: 1}]
        });
        internationalSheet.addRows([
            ['تاریخ فردا', 'مقصد', 'نام کالا', 'قیمت پایه (ریال)', 'توضیحات', 'جزئیات قیمت مبنا']
        ]);
        internationalSheet.getRow(1).style = headerStyle;

        internationalOffers.forEach((offer, index) => {
            const rowNum = index + 2;
            internationalSheet.addRow([
                offer['تاریخ فردا'],
                offer.destination,
                offer['نام کالا'],
                offer['قیمت پایه'],
                offer.description,
                offer['قیمت مبنا جزئیات محاسبه قیمت نهایی معاملات کشف پریمیوم']
            ]);
            const row = internationalSheet.getRow(rowNum);
            row.font = { name: 'Calibri', size: 11 };
            internationalSheet.getCell(rowNum, 1).alignment = { horizontal: 'center' };
            internationalSheet.getCell(rowNum, 2).alignment = { horizontal: 'center' };
            internationalSheet.getCell(rowNum, 4).numFmt = currencyFormat;
            internationalSheet.getCell(rowNum, 4).alignment = { horizontal: 'right' };
        });

        internationalSheet.getColumn(1).width = 15;
        internationalSheet.getColumn(2).width = 12;
        internationalSheet.getColumn(3).width = 40;
        internationalSheet.getColumn(4).width = 18;
        internationalSheet.getColumn(5).width = 30;
        internationalSheet.getColumn(6).width = 50;
    }

    // --- ساخت شیت سایر ---
    if (otherOffers.length > 0) {
        const otherSheet = workbook.addWorksheet('سایر', {
            views: [{state: 'frozen', xSplit: 1, ySplit: 1}]
        });
        otherSheet.addRows([
            ['تاریخ فردا', 'مقصد', 'نام کالا', 'قیمت پایه (ریال)', 'توضیحات']
        ]);
        otherSheet.getRow(1).style = headerStyle;

        otherOffers.forEach((offer, index) => {
            const rowNum = index + 2;
            otherSheet.addRow([
                offer['تاریخ فردا'],
                offer.destination,
                offer['نام کالا'],
                offer['قیمت پایه'],
                offer.description
            ]);
            const row = otherSheet.getRow(rowNum);
            row.font = { name: 'Calibri', size: 11 };
            otherSheet.getCell(rowNum, 1).alignment = { horizontal: 'center' };
            otherSheet.getCell(rowNum, 2).alignment = { horizontal: 'center' };
            otherSheet.getCell(rowNum, 4).numFmt = currencyFormat;
            otherSheet.getCell(rowNum, 4).alignment = { horizontal: 'right' };
        });

        otherSheet.getColumn(1).width = 15;
        otherSheet.getColumn(2).width = 12;
        otherSheet.getColumn(3).width = 40;
        otherSheet.getColumn(4).width = 18;
        otherSheet.getColumn(5).width = 30;
    }

    // --- تولید بافر ---
    try {
        const buffer = await workbook.xlsx.writeBuffer();
        console.log('Excel buffer generated successfully.');
        return buffer;
    } catch (error) {
        console.error('Error writing Excel buffer:', error);
        return null;
    }
}

module.exports = { processExcelData };
