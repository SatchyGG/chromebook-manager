const express = require('express');
const path = require('path');
const ExcelJS = require('exceljs');
const app = express();
const PORT = 3000;

app.use(express.json());

let scanHistory = [];

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/manifest.json', (req, res) => res.sendFile(path.join(__dirname, 'manifest.json')));
app.get('/sw.js', (req, res) => res.sendFile(path.join(__dirname, 'sw.js')));
app.post('/api/assign', (req, res) => {
    const { cabinet_barcode, device_barcode } = req.body;
    const now = new Date().toLocaleString('ja-JP');

    // 直前のデータと「完全に同じ組み合わせ」なら無視する
    if (scanHistory.length > 0) {
        const lastData = scanHistory[scanHistory.length - 1];
        if (lastData.cabinet === cabinet_barcode && lastData.device === device_barcode) {
            console.log(`【サーバー拒否】直前と全く同じ登録のため無視しました: 端末[${device_barcode}]`);
            return res.json({ success: false, message: "既に直前に登録済みです（重複ブロック）" });
        }
    }
    scanHistory.push({
        time: now,
        cabinet: cabinet_barcode,
        device: device_barcode
    });
    
    console.log(`【データ保存成功】日時: ${now} | キャビネット: ${cabinet_barcode} ➔ 端末: ${device_barcode}`);
    res.json({ success: true, message: "配置情報を更新しました" });
});

// 📊 Excelダウンロード用API
app.get('/api/download-excel', async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('配置登録データ');

        worksheet.views = [{ showGridLines: true }];

        worksheet.columns = [
            { header: '登録日時', key: 'time', width: 25 },
            { header: 'キャビネット（棚）コード', key: 'cabinet', width: 25 },
            { header: '端末シリアル / バーコード', key: 'device', width: 30 },
            { header: 'ステータス', key: 'status', width: 15 }
        ];

        const headerRow = worksheet.getRow(1);
        headerRow.font = { name: 'Meiryo', size: 11, bold: true, color: { argb: 'FFFFFF' } };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '1F4E78' }
        };
        headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
        scanHistory.forEach(item => {
            worksheet.addRow({
                time: item.time,
                cabinet: item.cabinet,
                device: item.device,
                status: '配置完了'
            });
        });

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                row.font = { name: 'Meiryo', size: 11 };
                row.getCell('time').alignment = { horizontal: 'center' };
                row.getCell('cabinet').alignment = { horizontal: 'center' };
                row.getCell('status').alignment = { horizontal: 'center' };
                
                row.eachCell(cell => {
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'D9D9D9' } },
                        left: { style: 'thin', color: { argb: 'D9D9D9' } },
                        bottom: { style: 'thin', color: { argb: 'D9D9D9' } },
                        right: { style: 'thin', color: { argb: 'D9D9D9' } }
                    };
                });
            }
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=chromebook_layout.xlsx');

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error(error);
        res.status(500).send('Excel生成エラーが発生しました');
    }
});

app.listen(PORT, () => {
    console.log(`サーバーが起動しました: http://localhost:${PORT}`);
});
