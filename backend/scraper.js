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
    if (!value) {
        return '';
    }

    return String(value)
        .replace(/[۰-۹]/g, (digit) => {
            return String('۰۱۲۳۴۵۶۷۸۹'.indexOf(digit));
        })
        .replace(/[٠-٩]/g, (digit) => {
            return String('٠١٢٣٤٥٦٧٨٩'.indexOf(digit));
        });
}

// تبدیل اعداد انگلیسی به فارسی
function toPersianDigits(value) {
    if (!value) {
        return '';
    }

    return String(value).replace(
        /\d/g,
        (digit) => '۰۱۲۳۴۵۶۷۸۹'[digit]
    );
}

// تاریخ شمسی فردا با اعداد انگلیسی
function getTomorrowJalali() {
    const today = new Date();

    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);

    const jalaliDate = jalaali.toJalaali(
        tomorrow.getFullYear(),
        tomorrow.getMonth() + 1,
        tomorrow.getDate()
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
    await inputLocator.type(dateValue, {
        delay: 80
    });

    // تأیید DatePicker
    await inputLocator.press('Enter');

    // خروج از فیلد؛ از Escape استفاده نمی‌کنیم
    await inputLocator.press('Tab');

    await page.waitForTimeout(700);

    const actualValue = await inputLocator.inputValue();

    const normalizedActualValue = toEnglishDigits(actualValue);
    const normalizedExpectedValue = toEnglishDigits(dateValue);

    console.log(
        `Date input value: ${actualValue} | Expected: ${dateValue}`
    );

    if (normalizedActualValue !== normalizedExpectedValue) {
        throw new Error(
            `تاریخ فیلتر درست وارد نشد. مقدار واقعی: ${actualValue}، مقدار مورد انتظار: ${dateValue}`
        );
    }
}

// اسکرپ اطلاعات و در صورت نیاز دانلود اکسل اصلی سایت
async function scrapeOfferAnnouncements(options = {}) {
    const {
        downloadExcel = false
    } = options;

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

        /*
         * محاسبه تاریخ فردا
         *
         * مثال:
         * 1405/06/01
         */
        const tomorrowJalali = getTomorrowJalali();

        console.log(
            `Target date: ${tomorrowJalali}`
        );

        /*
         * پیدا کردن دو فیلتر تاریخ
         */
        const dateInputs = page.locator('input.rmdp-input');

        if (await dateInputs.count() < 2) {
            throw new Error(
                'دو فیلتر تاریخ در صفحه پیدا نشد.'
            );
        }

        const fromDateInput = dateInputs.nth(0);
        const toDateInput = dateInputs.nth(1);

        await fromDateInput.waitFor({
            state: 'visible',
            timeout: 20000
        });

        await toDateInput.waitFor({
            state: 'visible',
            timeout: 20000
        });

        /*
         * واردکردن تاریخ فردا در فیلتر تاریخ شروع
         */
        await setDateValue(
            page,
            fromDateInput,
            tomorrowJalali
        );

        await page.waitForTimeout(700);

        /*
         * واردکردن تاریخ فردا در فیلتر تاریخ پایان
         */
        await setDateValue(
            page,
            toDateInput,
            tomorrowJalali
        );

        /*
         * بررسی مقدار واقعی دو فیلتر
         */
        const actualFromDate = await fromDateInput.inputValue();
        const actualToDate = await toDateInput.inputValue();

        console.log(
            `From date: ${actualFromDate}`
        );

        console.log(
            `To date: ${actualToDate}`
        );

        if (
            toEnglishDigits(actualFromDate) !== tomorrowJalali ||
            toEnglishDigits(actualToDate) !== tomorrowJalali
        ) {
            throw new Error(
                `تاریخ فیلترها اشتباه است. ` +
                `From: ${actualFromDate}, ` +
                `To: ${actualToDate}, ` +
                `Expected: ${tomorrowJalali}`
            );
        }

        /*
         * پیدا کردن دکمه جستجو
         */
        const searchButton = page
            .locator('button.btn-success:has-text("جستجو")')
            .first();

        await searchButton.waitFor({
            state: 'visible',
            timeout: 15000
        });

        console.log(
            'Clicking search button...'
        );

        /*
         * کلیک روی دکمه جستجو
         */
        await searchButton.click();

        await page.waitForTimeout(3000);

        await page
            .waitForLoadState('networkidle', {
                timeout: 30000
            })
            .catch(() => {});

        /*
         * دانلود اکسل اصلی سایت، در صورت درخواست
         */
        if (downloadExcel) {
            try {
                const excelButtonSelector =
                    'button.btn.btn-outline-success.ms-2';

                await page.waitForSelector(
                    excelButtonSelector,
                    {
                        state: 'visible',
                        timeout: 15000
                    }
                );

                console.log(
                    'Excel button found. Downloading original Excel...'
                );

                const [download] = await Promise.all([
                    page.waitForEvent('download'),
                    page.click(excelButtonSelector)
                ]);

                downloadedExcelName =
                    download.suggestedFilename();

                if (!downloadedExcelName) {
                    downloadedExcelName =
                        `offers-${Date.now()}.xlsx`;
                }

                const tempFileName =
                    `${Date.now()}-${downloadedExcelName}`;

                downloadedExcelPath = path.join(
                    __dirname,
                    tempFileName
                );

                await download.saveAs(
                    downloadedExcelPath
                );

                console.log(
                    `Original Excel saved temporarily at: ${downloadedExcelPath}`
                );
            } catch (error) {
                console.error(
                    'Could not download original Excel:',
                    error.message
                );

                throw error;
            }
        }

        /*
         * صبر برای آماده‌شدن جدول نتایج
         */
        await page.waitForSelector('table', {
            timeout: 20000
        });

        /*
         * استخراج جدول هدف
         */
        const data = await page.evaluate(() => {
            const tables = Array.from(
                document.querySelectorAll('table')
            );

            const targetTable = tables.find((table) => {
                const tableText =
                    table.innerText || '';

                return (
                    tableText.includes('بازار هدف') ||
                    tableText.includes('نام کالا') ||
                    tableText.includes('تاریخ و زمان عرضه')
                );
            });

            if (!targetTable) {
                return null;
            }

            return Array.from(
                targetTable.querySelectorAll('tr')
            )
                .map((row) => {
                    return Array.from(
                        row.querySelectorAll('th, td')
                    ).map((cell) =>
                        cell.innerText.trim()
                    );
                })
                .filter((row) => row.length > 0);
        });

        /*
         * بررسی وجود جدول و ردیف‌های واقعی
         */
        if (!data || data.length <= 1) {
            throw new Error(
                `اطلاعات قابل‌استفاده‌ای برای تاریخ ${tomorrowJalali} استخراج نشد.`
            );
        }

        const headers = data[0];
        const rows = data.slice(1);

        if (rows.length === 0) {
            throw new Error(
                `هیچ ردیفی برای تاریخ ${tomorrowJalali} پیدا نشد.`
            );
        }

        console.log(
            `Successfully scraped ${rows.length} rows for ${tomorrowJalali}.`
        );

        return {
            headers,
            rows,

            // مسیر فایل موقت اکسل
            excelPath: downloadedExcelPath,

            // نام اصلی فایل اکسل سایت
            excelName: downloadedExcelName,

            meta: {
                totalRows: rows.length,
                targetDateJalali: tomorrowJalali,
                targetDatePersian:
                    toPersianDigits(tomorrowJalali),
                scrapedAt: new Date().toISOString()
            }
        };
    } catch (error) {
        console.error(
            'Scraper Error:',
            error.message
        );

        /*
         * حذف فایل موقت در صورت خطا
         */
        if (
            downloadedExcelPath &&
            fs.existsSync(downloadedExcelPath)
        ) {
            try {
                fs.unlinkSync(
                    downloadedExcelPath
                );

                console.log(
                    'Temporary Excel deleted after scraper error.'
                );
            } catch (deleteError) {
                console.error(
                    'Could not delete temporary Excel:',
                    deleteError.message
                );
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
