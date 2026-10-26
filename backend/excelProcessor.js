// backend/excelProcessor.js

const ExcelJS = require('exceljs');
const jalaali = require('jalaali-js');
// const { createCanvas, registerFont } = require('@napi-rs/canvas'); // واترمارک فعلا غیرفعال است
const path = require('path');

// --- تنظیم فونت (فعلا غیرفعال) ---
// const fontPath = '/usr/share/fonts/truetype/msttcorefonts/Calibri.ttf';
// let isFontRegistered = false;
// function registerCalibriFont() { /* ... */ }


// --- توابع کمکی ---

function fixProductName(name) {
    if (!name) return '';
    // منطق نرمال‌سازی نام کالا
    return name.replace(/_/g, ' '); // مثال: جایگزینی آندرلاین با فاصله
}

// تابع فرمت‌دهی قیمت که فقط عدد خالص را برمی‌گرداند
function getPriceAsNumber(price) {
    if (typeof price === 'undefined' || price === null) return 0;
    const priceNum = Number(price.replace(/,/g, '')); // حذف کاما و تبدیل به عدد
    return isNaN(priceNum) ? 0 : priceNum;
}

function getTomorrowShamsi() {
    const today = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const jalaaliTomorrow = jalaali.g2j(tomorrow.getFullYear(), tomorrow.getMonth() + 1, tomorrow.getDate());
    // فرمت تاریخ به صورت YYYY/MM/DD
    const formattedMonth = jalaaliTomorrow.gm < 10 ? `0${jalaaliTomorrow.gm}` : jalaaliTomorrow.gm;
    const formattedDay = jalaaliTomorrow.gd < 10 ? `0${jalaaliTomorrow.gd}` : jalaaliTomorrow.gd;
    return `${jalaaliTomorrow.gy}/${formattedMonth}/${formattedDay}`;
}

// تابع مرتب‌سازی بر اساس نام کالا
function sortByProductName(a, b) {
    // کلید 'نام کالا' را مستقیماً استفاده می‌کنیم
    const nameA = (a['نام کالا'] || '').toLowerCase();
    const nameB = (b['نام کالا'] || '').toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
}

// --- تابع اصلی پردازش ---

async function processExcelData(scrapedData) {
    // registerCalibriFont(); // غیرفعال شده

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Petki';
    workbook.lastModifiedBy = 'Petki';

    // --- ساخت واترمارک (فعلا غیرفعال) ---
    // ... (کدهای مربوط به canvas کامنت شده‌اند)

    // --- پردازش و دسته‌بندی داده‌ها ---
    let processedData = scrapedData.map(item => ({
        ...item,
        // قیمت پایه را فقط به عدد تبدیل می‌کنیم
        'قیمت پایه': getPriceAsNumber(item['قیمت پایه']),
        'نام کالا': fixProductName(item['نام کالا']),
        'تاریخ فردا': getTomorrowShamsi()
    }));

    // فیلتر کردن و مرتب‌سازی
    const internalOffers = processedData.filter(item => item.destination === 'داخلی');
    const internationalOffers = processedData.filter(item => item.destination === 'بین‌الملل');
    const otherOffers = processedData.filter(item => item.destination !== 'داخلی' && item.destination !== 'بین‌الملل');

    internalOffers.sort(sortByProductName);
    internationalOffers.sort(sortByProductName);
    otherOffers.sort(sortByProductName);


    // --- استایل هدر ---
    const headerStyle = {
        font: { bold: true, name: 'Calibri', size: 12 }, // نام فونت Calibri
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } }, // خاکستری روشن
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    };
    const currencyFormat = '#,##0'; // فرمت عددی برای واحد پولی (ریال)

    // --- ساخت شیت داخلی ---
    const internalSheet = workbook.addWorksheet('داخلی', {
        views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }] // فریز کردن هدر
    });
    internalSheet.addRows([
        ['تاریخ فردا', 'مقصد', 'نام کالا', 'قیمت پایه (ریال)', 'توضیحات']
    ]);
    internalSheet.getRow(1).style = headerStyle; // اعمال استایل به هدر

    internalOffers.forEach((offer, index) => {
        const rowNum = index + 2;
        internalSheet.addRow([
            offer['تاریخ فردا'],
            offer.destination,
            offer['نام کالا'],
            offer['قیمت پایه'], // حالا عدد خام است
            offer.description
        ]);
        // تنظیم استایل و فرمت سلول‌ها
        const row = internalSheet.getRow(rowNum);
        row.font = { name: 'Calibri', size: 11 };
        internalSheet.getCell(rowNum, 1).alignment = { horizontal: 'center' }; // تاریخ
        internalSheet.getCell(rowNum, 2).alignment = { horizontal: 'center' }; // مقصد
        internalSheet.getCell(rowNum, 4).numFmt = currencyFormat; // فرمت عدد برای قیمت
        internalSheet.getCell(rowNum, 4).alignment = { horizontal: 'right' }; // راست‌چین برای قیمت
    });

    // تنظیم عرض ستون‌ها
    internalSheet.getColumn(1).width = 15;
    internalSheet.getColumn(2).width = 12;
    internalSheet.getColumn(3).width = 40;
    internalSheet.getColumn(4).width = 18;
    internalSheet.getColumn(5).width = 30;


    // --- ساخت شیت بین‌الملل ---
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
            offer['قیمت پایه'], // عدد خام
            offer.description,
            offer['قیمت مبنا جزئیات محاسبه قیمت نهایی معاملات کشف پریمیوم'] // این ستون را هم بررسی کنید
        ]);
        const row = internationalSheet.getRow(rowNum);
        row.font = { name: 'Calibri', size: 11 };
        internationalSheet.getCell(rowNum, 1).alignment = { horizontal: 'center' };
        internationalSheet.getCell(rowNum, 2).alignment = { horizontal: 'center' };
        internationalSheet.getCell(rowNum, 4).numFmt = currencyFormat; // فرمت عدد
        internationalSheet.getCell(rowNum, 4).alignment = { horizontal: 'right' };
    });

    internationalSheet.getColumn(1).width = 15;
    internationalSheet.getColumn(2).width = 12;
    internationalSheet.getColumn(3).width = 40;
    internationalSheet.getColumn(4).width = 18;
    internationalSheet.getColumn(5).width = 30;
    internationalSheet.getColumn(6).width = 50;


    // --- شیت "سایر" ---
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
                offer['قیمت پایه'], // عدد خام
                offer.description
            ]);
            const row = otherSheet.getRow(rowNum);
            row.font = { name: 'Calibri', size: 11 };
            otherSheet.getCell(rowNum, 1).alignment = { horizontal: 'center' };
            otherSheet.getCell(rowNum, 2).alignment = { horizontal: 'center' };
            otherSheet.getCell(rowNum, 4).numFmt = currencyFormat; // فرمت عدد
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
        // در صورت بروز خطا هنگام نوشتن بافر، null برمی‌گردانیم تا server.js متوجه شود.
        return null;
    }
}

module.exports = { processExcelData };
