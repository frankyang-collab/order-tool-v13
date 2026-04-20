const msoffcrypto = require('msoffcrypto-tool');
const XLSX = require('xlsx');

exports.handler = async (event) => {
    if (event.httpMethod !== 'POST') return { statusCode: 405, body: 'Only POST allowed' };
    
    try {
        const body = JSON.parse(event.body);
        // 解析前端傳來的 Base64 檔案
        const fileBuffer = Buffer.from(body.data.split(',')[1], 'base64');
        
        // 載入加密檔案
        const input = new msoffcrypto(fileBuffer);
        
        if (!input.hasPassword()) {
            return { statusCode: 400, body: JSON.stringify({ error: "檔案未加密，無法使用此通道解析" }) };
        }
        
        // 帶入 MOMO 固定密碼
        input.setPassword('90575147');
        
        // 執行解密
        const decryptedBuffer = await new Promise((resolve, reject) => {
            const chunks = [];
            const outStream = new require('stream').Writable({
                write(chunk, enc, next) {
                    chunks.push(chunk);
                    next();
                }
            });
            try {
                input.decrypt(outStream);
                outStream.on('finish', () => resolve(Buffer.concat(chunks)));
                outStream.on('error', reject);
            } catch (err) {
                reject(err);
            }
        });

        // 讀取解密後的 Excel 並轉為 CSV 格式 (為了與前端現有邏輯對接)
        const workbook = XLSX.read(decryptedBuffer, { type: 'buffer' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const csvStr = XLSX.utils.sheet_to_csv(worksheet);
        
        return {
            statusCode: 200,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ csv: csvStr })
        };

    } catch (error) {
        return {
            statusCode: 500,
            body: JSON.stringify({ error: error.message || "密碼錯誤或解密失敗" })
        };
    }
};
