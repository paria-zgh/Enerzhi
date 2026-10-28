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
        const result = await scrapeOfferAnnouncements({
            downloadExcel: false
        });

        // ✅ بررسی اینکه آیا داده‌ای پیدا شد یا خالی است
        if (result.isEmpty) {
            return res.status(200).json({
                success: true,
                empty: true,
                message: result.message || 'هیچ عرضه‌ای برای این تاریخ ثبت نشده است.',
                headers: [],
                rows: [],
                meta: result.meta
            });
        }

        res.json({
            success: true,
            empty: false,
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
        const result = await scrapeOfferAnnouncements({
            downloadExcel: true
        });

        // ✅ اگر داده‌ای برای عرضه نبود و اکسلی تولید نشده بود
        if (result.isEmpty || !result.excelPath) {
            return res.status(404).json({
                success: false,
                empty: true,
                message: result.message || 'برای این تاریخ هیچ عرضه‌ای ثبت نشده و فایل اکسلی وجود ندارد.'
            });
        }

        excelPath = result.excelPath;

        if (!fs.existsSync(excelPath)) {
            throw new Error('فایل اکسل اصلی سایت یافت نشد.');
        }

        const fileName = result.excelName || `offers-${Date.now()}.xlsx`;

        console.log(`Sending file to frontend: ${excelPath}`);

        res.download(
            excelPath,
            fileName,
            (downloadError) => {
                // حذف فایل موقت بعد از دانلود
                if (excelPath && fs.existsSync(excelPath)) {
                    fs.unlink(excelPath, (deleteError) => {
                        if (deleteError) {
                            console.error('Error deleting temporary Excel:', deleteError.message);
                        } else {
                            console.log('Temporary Excel deleted.');
                        }
                    });
                }

                if (downloadError) {
                    console.error('Download Error:', downloadError.message);
                }
            }
        );
    } catch (error) {
        console.error('Excel API Error:', error);

        if (excelPath && fs.existsSync(excelPath)) {
            try {
                fs.unlinkSync(excelPath);
            } catch (deleteError) {
                console.error('Could not delete Excel after error:', deleteError.message);
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
