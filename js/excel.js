// ============================================================
// Excel 解析与导出工具
// 依赖 SheetJS (xlsx.full.min.js)
// ============================================================

/**
 * 解析上传的 Excel 文件（存货收发存汇总表）
 * 提取：会计期间、物料代码、物料名称、计量单位、期末结存数量
 * @param {File} file
 * @returns {Promise<{period: string, items: Array}>}
 */
async function parseInventoryExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                // 读取第一个 sheet
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (jsonData.length < 2) {
                    reject(new Error('Excel 文件为空或格式不正确'));
                    return;
                }

                // 表头行（第0行）
                const headers = jsonData[0];
                // 查找列索引
                const colMap = {
                    period: headers.indexOf('会计期间'),
                    code: headers.indexOf('物料代码'),
                    name: headers.indexOf('物料名称'),
                    unit: headers.indexOf('计量单位'),
                    endQty: headers.indexOf('期末结存数量')
                };

                // 检查必须列是否都存在
                for (const [key, idx] of Object.entries(colMap)) {
                    if (idx === -1) {
                        reject(new Error(`找不到列：${key}`));
                        return;
                    }
                }

                let period = '';
                const items = [];

                for (let i = 1; i < jsonData.length; i++) {
                    const row = jsonData[i];
                    if (!row || !row[colMap.code]) continue; // 跳过空行

                    if (!period && row[colMap.period]) {
                        period = String(row[colMap.period]);
                    }

                    items.push({
                        period: period,
                        material_code: String(row[colMap.code] || '').trim(),
                        material_name: String(row[colMap.name] || '').trim(),
                        unit: String(row[colMap.unit] || '').trim(),
                        jiecun: parseFloat(row[colMap.endQty]) || 0
                    });
                }

                resolve({ period, items });
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error('文件读取失败'));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * 导出单个档口报表为 Excel
 * @param {string} filename - 文件名（如"干料领料单"）
 * @param {Array} data - [{material_name, value}]
 * @param {string} sheetTitle - 工作表标题
 */
function exportToExcel(filename, data, sheetTitle = '报表') {
    const wb = XLSX.utils.book_new();

    const wsData = [
        ['物料名称', '数量']
    ];
    data.forEach(item => {
        wsData.push([item.material_name, item.value]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // 设置列宽
    ws['!cols'] = [
        { wch: 30 },
        { wch: 15 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, sheetTitle);
    XLSX.writeFile(wb, `${filename}.xlsx`);
}

/**
 * 批量导出所有档口报表（一个 Excel 多个 Sheet）
 * @param {Object} reports - { '干料领料单': [{material_name, value}], ... }
 * @param {string} filename
 */
function exportAllReports(reports, filename) {
    const wb = XLSX.utils.book_new();

    for (const [sheetName, data] of Object.entries(reports)) {
        if (!data || data.length === 0) continue;

        const wsData = [['物料名称', '数量']];
        data.forEach(item => {
            wsData.push([item.material_name, item.value]);
        });

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        ws['!cols'] = [{ wch: 30 }, { wch: 15 }];
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    }

    XLSX.writeFile(wb, `${filename}.xlsx`);
}
