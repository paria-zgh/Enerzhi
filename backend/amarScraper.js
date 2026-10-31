const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const jalaali = require('jalaali-js');

// پیدا کردن مسیر مرورگر نصب‌شده روی سیستم (دقیقاً مشابه scraper.js)
function getBrowserPath() {
    const isWindows = process.platform === 'win32';
    if (!isWindows) return null;

    const possiblePaths = [
        'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
        'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
    ];

    for (const p of possiblePaths) {
        if (fs.existsSync(p)) {
            return p;
        }
    }
    return null;
}

// تابع کمکی تبدیل اعداد انگلیسی به فارسی (برای فیلد تاریخ سایت)
function toPersianDigits(str) {
    const persianDigits = ['۰', '۱', '۲', '۳', '۴', '۵', '۶', '۷', '۸', '۹'];
    return String(str).replace(/[0-9]/g, (w) => persianDigits[+w]);
}

// تابع دریافت تاریخ امروز شمسی به فرمت ۱۴۰۳/۰۸/۰۶
function getTodayJalali() {
    const now = new Date();
    const j = jalaali.toJalaali(now.getFullYear(), now.getMonth() + 1, now.getDate());
    const mm = String(j.jm).padStart(2, '0');
    const dd = String(j.jd).padStart(2, '0');
    return `${j.jy}/${mm}/${dd}`;
}

async function scrapeAmarStatistics({ dateJalali = null, downloadExcel = false } = {}) {
    const targetDateEnglish = dateJalali || getTodayJalali();
    const targetDatePersian = toPersianDigits(targetDateEnglish);

    const downloadsDir = path.join(__dirname, 'downloads');
    if (!fs.existsSync(downloadsDir)) {
        fs.mkdirSync(downloadsDir, { recursive: true });
    }

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

    let downloadedFilePath = null;
    let excelFileName = null;

    try {
        console.log(`[amarScraper] Navigating to Statistics page...`);
        await page.goto('https://oldsite.iee.ir/PhysicalMarket/Statistics', {
            waitUntil: 'networkidle',
            timeout: 60000
        });

        // ۱. پر کردن فیلترهای تاریخ (هر دو اینپوت از/تا)
        const dateInputs = page.locator('input.rmdp-input');
        const inputCount = await dateInputs.count();

        if (inputCount >= 2) {
            console.log(`[amarScraper] Setting date to: ${targetDatePersian} (${targetDateEnglish})`);
            // اینپوت اول: از تاریخ
            await dateInputs.nth(0).click({ clickCount: 3 });
            await dateInputs.nth(0).fill(targetDatePersian);
            await dateInputs.nth(0).press('Enter');

            // اینپوت دوم: تا تاریخ
            await dateInputs.nth(1).click({ clickCount: 3 });
            await dateInputs.nth(1).fill(targetDatePersian);
            await dateInputs.nth(1).press('Enter');
        } else {
            console.warn('[amarScraper] Date inputs not matched expected count:', inputCount);
        }

        // ۲. کلیک دکمه جستجو
        console.log('[amarScraper] Clicking Search button...');
        const searchButton = page.locator('button:has-text("جستجو")');
        await searchButton.first().click();

        // صبر برای لود اطلاعات
        await page.waitForTimeout(3000);

        // ۳. اگر فقط اکسل خواسته شده بود
        if (downloadExcel) {
            console.log('[amarScraper] Triggering Excel download...');

            const excelButton = page.locator('button.btn-outline-success, button:has(img[src*="excel"])').first();

            const [download] = await Promise.all([
                page.waitForEvent('download', { timeout: 30000 }),
                excelButton.click()
            ]);

            excelFileName = download.suggestedFilename() || `amar-${Date.now()}.xlsx`;
            downloadedFilePath = path.join(downloadsDir, `${Date.now()}-${excelFileName}`);
            await download.saveAs(downloadedFilePath);

            console.log(`[amarScraper] Excel saved to: ${downloadedFilePath}`);

            return {
                isEmpty: false,
                excelPath: downloadedFilePath,
                excelName: excelFileName,
                meta: {
                    targetDateJalali: targetDateEnglish
                }
            };
        }

        // ۴. اگر داده‌های جدول به صورت JSON خواسته شده بود
        const table = page.locator('table').first();
        const hasTable = await table.isVisible().catch(() => false);

        if (!hasTable) {
            return {
                isEmpty: true,
                message: 'هیچ داده‌ای برای این تاریخ یافت نشد.',
                headers: [],
                rows: [],
                meta: { targetDateJalali: targetDateEnglish }
            };
        }

        const headers = await table.locator('thead th').allInnerTexts();
        const cleanedHeaders = headers.map(h => h.trim()).filter(Boolean);

        const rowsData = [];
        const rows = table.locator('tbody tr');
        const count = await rows.count();

        for (let i = 0; i < count; i++) {
            const cells = await rows.nth(i).locator('td').allInnerTexts();
            const cleanedCells = cells.map(c => c.trim());
            if (cleanedCells.length > 0) {
                rowsData.push(cleanedCells);
            }
        }

        return {
            isEmpty: rowsData.length === 0,
            headers: cleanedHeaders,
            rows: rowsData,
            meta: {
                totalRows: rowsData.length,
                targetDateJalali: targetDateEnglish
            }
        };

    } finally {
        await browser.close();
    }
}

module.exports = {
    scrapeAmarStatistics
};
