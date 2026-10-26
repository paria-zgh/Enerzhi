import { useState } from 'react';

export default function DownloadRawExcelButton() {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const handleDownload = async () => {
    setLoading(true);
    setError('');

    try {
      // درخواست به بک‌اند؛ با این مسیر فایل خام دانلود می‌شه
      const response = await fetch('http://localhost:5000/api/offers/raw-excel');

      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.message || 'خطا در دانلود فایل');
      }

      // گرفتن blob از پاسخ و ساخت لینک دانلود در مرورگر
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);

      // استخراج نام فایل از هدر Content-Disposition (اختیاری)
      const contentDisposition = response.headers.get('Content-Disposition');
      let filename = 'offers-raw.xlsx';
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="?([^"]+)"?/);
        if (match) filename = match[1];
      }

      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      console.log('✅ فایل خام دانلود شد');
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div>
      <button
        onClick={handleDownload}
        disabled={loading}
        className="btn btn-outline-success"
      >
        {loading ? 'در حال دانلود...' : '⬇️ دانلود فایل اکسل خام'}
      </button>

      {error && (
        <p style={{ color: 'red', marginTop: '8px' }}>
          {error}
        </p>
      )}
    </div>
  );
}
