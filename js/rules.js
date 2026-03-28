// ============================================================
// 规则引擎 - 差异处理与报表生成
// ============================================================

// 6 个档口常量
const CATEGORIES = ['ganlia', 'shiliao', 'shushi', 'xiancai', 'douyou', 'yuangongcan'];
const CATEGORY_NAMES = {
    ganlia: '干料',
    shiliao: '湿料',
    shushi: '熟食',
    xiancai: '咸菜',
    douyou: '豆油',
    yuangongcan: '员工餐'
};

// 03. 前缀的档口对应规则
const PREFIX_03_MAP = {
    '03.01.01': 'ganlia',
    '03.01.02': 'shiliao',
    '03.02': 'shushi',
    '03.03': 'xiancai',
    '03.04': 'xiancai',
    '03.07': 'xiancai',
    '03.05': 'douyou'
};

/**
 * 判断物料代码的前缀归属
 * @param {string} code
 * @returns {'01_02'|'03'|'04'|'unknown'}
 */
function getCodePrefix(code) {
    if (code.startsWith('01.') || code.startsWith('02.')) return '01_02';
    if (code.startsWith('03.')) return '03';
    if (code.startsWith('04.')) return '04';
    return 'unknown';
}

/**
 * 根据 03. 前缀匹配档口
 * @param {string} code
 * @returns {string|null} 档口 key
 */
function match03Category(code) {
    // 按前缀长度从长到短匹配，确保精确优先
    const sortedPrefixes = Object.keys(PREFIX_03_MAP).sort((a, b) => b.length - a.length);
    for (const prefix of sortedPrefixes) {
        if (code.startsWith(prefix)) {
            return PREFIX_03_MAP[prefix];
        }
    }
    return null;
}

/**
 * 核心规则引擎：处理所有物料差异，生成分类报表
 * @param {Array} inventoryData - 库存数据（含 diankv）
 * @param {Array} rulesData - 规则数据
 * @returns {Object} {
 *   reports: {干料领料单: [...], 干料入库单: [...], ...},
 *   inquiries: [...],       // 询问清单
 *   noRule: [...],          // 无规则物料
 *   skipped: [...],         // 跳过的物料（不动/04./无差异）
 *   summary: {...}          // 汇总统计
 * }
 */
function processInventory(inventoryData, rulesData) {
    // 构建规则查询表
    const rulesMap = new Map();
    rulesData.forEach(r => rulesMap.set(r.material_code, r));

    // 初始化报表容器：领料单 + 入库单
    const reports = {};
    for (const cat of CATEGORIES) {
        reports[`${CATEGORY_NAMES[cat]}领料单`] = [];
        reports[`${CATEGORY_NAMES[cat]}入库单`] = [];
    }

    const inquiries = [];  // 询问清单
    const noRule = [];     // 无规则物料
    const skipped = [];    // 跳过的物料

    for (const item of inventoryData) {
        const code = item.material_code;
        const name = item.material_name;
        const unit = item.unit;
        const jiecun = item.jiecun ?? 0;
        // 点库未录入 → 视为 0
        const diankv = item.diankv ?? 0;

        const prefix = getCodePrefix(code);
        const rule = rulesMap.get(code);

        // ---- 04. 前缀 → 不处理 ----
        if (prefix === '04') {
            // 但结存为负时，移到询问
            if (jiecun < 0) {
                inquiries.push({ ...item, reason: '04前缀但结存为负', diff: diankv - jiecun });
            } else {
                skipped.push({ ...item, reason: '04前缀不处理' });
            }
            continue;
        }

        // ---- 01./02. 前缀处理 ----
        if (prefix === '01_02') {
            // 查规则
            if (!rule) {
                noRule.push({ ...item, reason: '未找到规则' });
                continue;
            }

            const attr = rule.attribute;

            // 属性="不动"
            if (attr === '不动') {
                // 结存为负 → 改为询问
                if (jiecun < 0) {
                    inquiries.push({ ...item, reason: '不动但结存为负', diff: jiecun - diankv });
                } else {
                    skipped.push({ ...item, reason: '属性为不动' });
                }
                continue;
            }

            // 属性="询问"
            if (attr === '询问') {
                const diff = jiecun - diankv;
                if (diff !== 0) {
                    inquiries.push({ ...item, reason: '属性为询问', diff });
                }
                continue;
            }

            // 属性="正常" → 按规则处理
            const pendingData = jiecun - diankv; // 待处理数据

            // 待处理数据<=0 → 不处理（但结存为负 → 询问）
            if (pendingData <= 0) {
                if (jiecun < 0) {
                    inquiries.push({ ...item, reason: '结存为负', diff: pendingData });
                } else {
                    skipped.push({ ...item, reason: '无需处理（结存≤点库）' });
                }
                continue;
            }

            // 收集有系数的档口
            const activeCats = [];
            for (const cat of CATEGORIES) {
                const ratio = rule[cat];
                if (ratio && ratio > 0) {
                    activeCats.push({ key: cat, ratio });
                }
            }

            if (activeCats.length === 0) {
                // 有规则但没有具体系数 → 询问
                inquiries.push({ ...item, reason: '有规则但无具体系数', diff: pendingData });
                continue;
            }

            // 计算分配
            const divider = unit === '斤' ? 2 : 1;
            let remaining = pendingData;

            for (let i = 0; i < activeCats.length; i++) {
                const { key, ratio } = activeCats[i];
                let value;

                if (i < activeCats.length - 1) {
                    // 非最后一个：round(待处理*比例, 0)
                    value = Math.round(pendingData * ratio);
                    remaining -= value;
                } else {
                    // 最后一个：用剩余量
                    value = remaining;
                }

                // 单位为斤 → ÷2
                const finalValue = value / divider;

                if (finalValue !== 0) {
                    reports[`${CATEGORY_NAMES[key]}领料单`].push({
                        material_name: name,
                        value: finalValue
                    });
                }
            }

            continue;
        }

        // ---- 03. 前缀处理 ----
        if (prefix === '03') {
            const pendingData = diankv - jiecun; // 注意方向：点库 - 结存

            // 不需要规则库，按代码前缀判断档口
            const category = match03Category(code);

            if (!category) {
                // 找不到对应的档口规则
                inquiries.push({ ...item, reason: '03前缀但档口未定义', diff: pendingData });
                continue;
            }

            // 结存为负且 pendingData <= 0 → 询问
            if (jiecun < 0 && pendingData <= 0) {
                inquiries.push({ ...item, reason: '03前缀结存为负', diff: pendingData });
                continue;
            }

            // pendingData <= 0 → 不处理
            if (pendingData <= 0) {
                skipped.push({ ...item, reason: '03无需处理（点库≤结存）' });
                continue;
            }

            const divider = unit === '斤' ? 2 : 1;
            const finalValue = pendingData / divider;

            reports[`${CATEGORY_NAMES[category]}入库单`].push({
                material_name: name,
                value: finalValue
            });

            continue;
        }

        // ---- 未知前缀 ----
        noRule.push({ ...item, reason: '未知物料代码前缀' });
    }

    // 汇总统计
    const summary = {
        total: inventoryData.length,
        processed: 0,
        inquired: inquiries.length,
        noRule: noRule.length,
        skipped: skipped.length
    };

    for (const arr of Object.values(reports)) {
        summary.processed += arr.length;
    }

    return { reports, inquiries, noRule, skipped, summary };
}
