import React, { useState } from 'react';
import './Enerzhi.css';

function Enerzhi() {
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState(null);

    const handleDownload = async () => {
        setIsLoading(true);
        setError(null);

        try {
            const response = await fetch('https://enerzhi-1.onrender.com/api/offers/excel');

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Server error: ${response.status} - ${errorText}`);
            }

            const disposition = response.headers.get('content-disposition');
            let filename = `offers_${Date.now()}.xlsx`;

            if (disposition && disposition.includes('attachment')) {
                const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                const matches = filenameRegex.exec(disposition);
                if (matches && matches[1]) {
                    filename = matches[1].replace(/['"]/g, '');
                }
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);

            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = filename;

            document.body.appendChild(a);
            a.click();

            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        } catch (err) {
            console.error('Download failed:', err);
            setError(`خطا در دانلود فایل: ${err.message}`);
        } finally {
            setIsLoading(false);
        }
    };

    return (
        <div className="enerzhi-container">
            <h2>دانلود فایل معاملات</h2>
            <button onClick={handleDownload} disabled={isLoading} className="download-button">
                {isLoading ? 'در حال پردازش...' : 'دانلود فایل اکسل'}
            </button>
            {error && <p className="error-message">{error}</p>}
            {!isLoading && !error && (
                <p className="info-message">
                    با کلیک روی دکمه، همان فایل اکسل ذخیره‌شده در بک‌اند دانلود خواهد شد.
                </p>
            )}
        </div>
    );
}

export default Enerzhi;
