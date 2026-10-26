const express = require('express');
const cors = require('cors');
const fs = require('fs');

const {
    scrapeOfferAnnouncements
} = require('./scraper');

const app = express();

app.use(cors());
app.use(express.json());


// تست سلامت سرور
app.get('/', (req, res) => {
    res.send('Backend is running');
});

// دریافت اطلاعات جدول به‌صورت JSON
app.get('/api/offers', async (req, res) => {
    try {
        /*
         * در این مسیر اکسل دانلود نمی‌شود.
         * فقط اطلاعات جدول برگردانده می‌شود.
         */
        const result = await scrapeOfferAnnouncements({
            downloadExcel: false
        });

        res.json({
            success: true,
            headers: result.headers,
            rows: result.rows,
            meta: result.meta
        });
    } catch (error) {
        console.error('Offers API Error:', error);

        res.status(500).json({
            success: false,
            message: 'خطا در دریافت اطلاعات عرضه‌ها',
            error: error.message
        });
    }
});

// دانلود همان اکسل اصلی سایت
app.get('/api/offers/excel', async (req, res) => {
    let excelPath = null;

    try {
        /*
         * در این مسیر:
         * 1. صفحه سایت باز می‌شود
         * 2. اکسل اصلی سایت دانلود می‌شود
         * 3. همان اکسل به کاربر ارسال می‌شود
         */
        const result = await scrapeOfferAnnouncements({
            downloadExcel: true
        });

        excelPath = result.excelPath;

        if (!excelPath || !fs.existsSync(excelPath)) {
            throw new Error(
                'فایل اکسل اصلی سایت دریافت نشد.'
            );
        }

        /*
         * نامی که کاربر در کامپیوتر خود می‌بیند.
         * اگر بخواهید نام اصلی سایت حفظ شود، می‌توانید
         * به‌جای این مقدار از result.excelName استفاده کنید.
         */
        const fileName =
            result.excelName || `offers-${Date.now()}.xlsx`;

        console.log(`Sending file to frontend: ${excelPath}`);

        res.download(
            excelPath,
            fileName,
            (downloadError) => {
                /*
                 * بعد از ارسال فایل، نسخه موقت بک‌اند حذف می‌شود.
                 */
                if (
                    excelPath &&
                    fs.existsSync(excelPath)
                ) {
                    fs.unlink(
                        excelPath,
                        (deleteError) => {
                            if (deleteError) {
                                console.error(
                                    'Error deleting temporary Excel:',
                                    deleteError.message
                                );
                            } else {
                                console.log(
                                    'Temporary Excel deleted.'
                                );
                            }
                        }
                    );
                }

                if (downloadError) {
                    console.error(
                        'Download Error:',
                        downloadError.message
                    );
                }
            }
        );
    } catch (error) {
        console.error('Excel API Error:', error);

        if (
            excelPath &&
            fs.existsSync(excelPath)
        ) {
            try {
                fs.unlinkSync(excelPath);
            } catch (deleteError) {
                console.error(
                    'Could not delete Excel after error:',
                    deleteError.message
                );
            }
        }

        if (!res.headersSent) {
            res.status(500).json({
                success: false,
                message: 'خطا در دانلود فایل اکسل',
                error: error.message
            });
        }
    }
});

const PORT = process.env.PORT || 5000;

app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});

