const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const jalaali = require('jalaali-js');

// پیدا کردن مسیر Chrome یا Edge در ویندوز
function getBrowserPath() {
    if (process.platform !== 'win32') return null;
  
    const possiblePaths = [
      'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
      'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
      'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
      'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
    ];
  
    for (const browserPath of possiblePaths) {
      if (fs.existsSync(browserPath)) return browserPath;
    }
  
    return null;
}

// تبدیل اعداد فارسی و عربی به انگلیسی
function toEnglishDigits(value) {
    if (!value) return '';
    return String(value)
        .replace(/[۰-۹]/g, digit => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(digit)))
        .replace(/[٠-٩]/g, digit => String('٠١٢٣٤٥٦٧٨٩'.indexOf(digit)));
}

// تبدیل اعداد انگلیسی به فارسی
function toPersianDigits(value) {
    if (!value) return '';
    return String(value).replace(/\d/g, digit => '۰۱۲۳۴۵۶۷۸۹'[digit]);
}

// ✅ اصلاح‌شده: تاریخ شمسی روز کاری بعدی (چهارشنبه، پنجشنبه و جمعه -> شنبه)
function getNextTradingDayJalali() {
    const today = new Date();
    const targetDate = new Date(today);
    const dayOfWeek = today.getDay(); // 0: یکشنبه, 1: دوشنبه, ..., 3: چهارشنبه, 4: پنجشنبه, 5: جمعه, 6: شنبه

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

    const jalaliDate = jalaali.toJalaali(
        targetDate.getFullYear(),
        targetDate.getMonth() + 1,
        targetDate.getDate()
    );

    return (
        `${jalaliDate.jy}/` +
        `${String(jalaliDate.jm).padStart(2, '0')}/` +
        `${String(jalaliDate.jd).padStart(2, '0')}`
    );
}

// واردکردن مقدار در فیلتر تاریخ
async function setDateValue(page, inputLocator, dateValue) {
    await inputLocator.waitFor({
        state: 'visible',
        timeout: 10000
    });

    await inputLocator.click();

    // پاک‌کردن مقدار قبلی
    await inputLocator.press('Control+A');
    await inputLocator.press('Backspace');

    // واردکردن تاریخ
    await inputLocator.type(dateValue, { delay: 80 });

    // تأیید DatePicker
    await inputLocator.press('Enter');
    await inputLocator.press('Tab');

    await page.waitForTimeout(700);

    const actualValue = await inputLocator.inputValue();
    const normalizedActualValue = toEnglishDigits(actualValue);
    const normalizedExpectedValue = toEnglishDigits(dateValue);

    console.log(`Date input value: ${actualValue} | Expected: ${dateValue}`);

    if (normalizedActualValue !== normalizedExpectedValue) {
        throw new Error(
            `تاریخ فیلتر درست وارد نشد. مقدار واقعی: ${actualValue}، مقدار مورد انتظار: ${dateValue}`
        );
    }
}

// اسکرپ اطلاعات و در صورت نیاز دانلود اکسل اصلی سایت
async function scrapeOfferAnnouncements(options = {}) {
    const { downloadExcel = false } = options;

    let downloadedExcelPath = null;
    let downloadedExcelName = null;

    const executablePath = getBrowserPath();

    const browser = await chromium.launch({
      headless: true,
      executablePath: executablePath || undefined,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--ignore-certificate-errors'
      ]
    });

    const context = await browser.newContext({
        ignoreHTTPSErrors: true,
        acceptDownloads: true,
        userAgent:
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' +
            'AppleWebKit/537.36 (KHTML, like Gecko) ' +
            'Chrome/120.0.0.0 Safari/537.36'
    });
    
    const page = await context.newPage();

    try {
        console.log('Opening IEE page...');

        await page.goto(
            'https://oldsite.iee.ir/PhysicalMarket/OffersAnnouncements',
            {
                waitUntil: 'domcontentloaded',
                timeout: 60000
            }
        );

        await page.waitForTimeout(3000);

        // محاسبه تاریخ هدف (شنبه در چهارشنبه‌ها)
        const targetTradingDayJalali = getNextTradingDayJalali();
        console.log(`Target trading date: ${targetTradingDayJalali}`);

        // پیدا کردن دو فیلتر تاریخ
        const dateInputs = page.locator('input.rmdp-input');
        if (await dateInputs.count() < 2) {
            throw new Error('دو فیلتر تاریخ در صفحه پیدا نشد.');
        }

        const fromDateInput = dateInputs.nth(0);
        const toDateInput = dateInputs.nth(1);

        await fromDateInput.waitFor({ state: 'visible', timeout: 20000 });
        await toDateInput.waitFor({ state: 'visible', timeout: 20000 });

        // وارد کردن تاریخ
        await setDateValue(page, fromDateInput, targetTradingDayJalali);
        await page.waitForTimeout(500);
        await setDateValue(page, toDateInput, targetTradingDayJalali);

        // بررسی مقدار اعمال شده
        const actualFromDate = await fromDateInput.inputValue();
        const actualToDate = await toDateInput.inputValue();

        if (
            toEnglishDigits(actualFromDate) !== targetTradingDayJalali ||
            toEnglishDigits(actualToDate) !== targetTradingDayJalali
        ) {
            throw new Error(
                `تاریخ فیلترها اشتباه است. From: ${actualFromDate}, To: ${actualToDate}, Expected: ${targetTradingDayJalali}`
            );
        }

        // کلیک روی دکمه جستجو
        const searchButton = page.locator('button.btn-success:has-text("جستجو")').first();
        await searchButton.waitFor({ state: 'visible', timeout: 15000 });

        console.log('Clicking search button...');
        await searchButton.click();

        await page.waitForTimeout(3000);
        await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});

        /*
         * ✅ مرحله ۱: اول استخراج جدول را بررسی می‌کنیم تا ببینیم اصلاً داده‌ای هست یا نه
         */
        await page.waitForSelector('table', { timeout: 15000 }).catch(() => {});

        const data = await page.evaluate(() => {
            const tables = Array.from(document.querySelectorAll('table'));
            const targetTable = tables.find((table) => {
                const tableText = table.innerText || '';
                return (
                    tableText.includes('بازار هدف') ||
                    tableText.includes('نام کالا') ||
                    tableText.includes('تاریخ و زمان عرضه')
                );
            });

            if (!targetTable) return null;

            return Array.from(targetTable.querySelectorAll('tr'))
                .map((row) => Array.from(row.querySelectorAll('th, td')).map(cell => cell.innerText.trim()))
                .filter((row) => row.length > 0);
        });

        // اگر داده‌ای برای این تاریخ نبود:
        const hasRows = data && data.length > 1;

        if (!hasRows) {
            console.log(`⚠️ هیچ عرضه‌ای برای تاریخ ${targetTradingDayJalali} ثبت نشده است.`);
            return {
                isEmpty: true,
                headers: [],
                rows: [],
                excelPath: null,
                excelName: null,
                message: `هیچ عرضه‌ای برای تاریخ ${targetTradingDayJalali} ثبت نشده است.`,
                meta: {
                    totalRows: 0,
                    targetDateJalali: targetTradingDayJalali,
                    targetDatePersian: toPersianDigits(targetTradingDayJalali),
                    scrapedAt: new Date().toISOString()
                }
            };
        }

        const headers = data[0];
        const rows = data.slice(1);

        /*
         * ✅ مرحله ۲: دانلود اکسل اصلی فقط زمانی که داده واقعاً وجود داشته باشد
         */
        if (downloadExcel) {
            const excelButtonSelector = 'button.btn.btn-outline-success.ms-2';
            const excelButton = page.locator(excelButtonSelector).first();

            const isBtnVisible = await excelButton.isVisible().catch(() => false);

            if (isBtnVisible) {
                console.log('Excel button found. Attempting to download original Excel...');

                try {
                    // با تایم‌اوت ۱۰ ثانیه تا اگر دکمه کار نکرد، برنامه تا ابد معطل نشود
                    const [download] = await Promise.all([
                        page.waitForEvent('download', { timeout: 10000 }),
                        excelButton.click()
                    ]);

                    downloadedExcelName = download.suggestedFilename() || `offers-${Date.now()}.xlsx`;
                    const tempFileName = `${Date.now()}-${downloadedExcelName}`;
                    downloadedExcelPath = path.join(__dirname, tempFileName);

                    await download.saveAs(downloadedExcelPath);
                    console.log(`Original Excel saved temporarily at: ${downloadedExcelPath}`);
                } catch (dlErr) {
                    console.warn('⚠️ دکمه اکسل کلیک شد اما فایلی دانلود نشد (شاید سایت دانلودی ارائه نداده):', dlErr.message);
                }
            } else {
                console.warn('⚠️ دکمه دانلود اکسل در صفحه فعال یا موجود نبود.');
            }
        }

        console.log(`Successfully scraped ${rows.length} rows for ${targetTradingDayJalali}.`);

        return {
            isEmpty: false,
            headers,
            rows,
            excelPath: downloadedExcelPath,
            excelName: downloadedExcelName,
            meta: {
                totalRows: rows.length,
                targetDateJalali: targetTradingDayJalali,
                targetDatePersian: toPersianDigits(targetTradingDayJalali),
                scrapedAt: new Date().toISOString()
            }
        };

    } catch (error) {
        console.error('Scraper Error:', error.message);

        // حذف فایل موقت در صورت بروز خطا
        if (downloadedExcelPath && fs.existsSync(downloadedExcelPath)) {
            try {
                fs.unlinkSync(downloadedExcelPath);
                console.log('Temporary Excel deleted after scraper error.');
            } catch (deleteError) {
                console.error('Could not delete temporary Excel:', deleteError.message);
            }
        }

        throw error;
    } finally {
        await browser.close();
    }
}

module.exports = {
    scrapeOfferAnnouncements
};
