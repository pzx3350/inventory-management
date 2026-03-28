// ============================================================
// Supabase 数据库操作封装
// ============================================================

let supabaseClient = null;

/**
 * 初始化 Supabase 客户端
 */
function initSupabase() {
    if (!SUPABASE_URL || SUPABASE_URL === 'YOUR_SUPABASE_URL') {
        console.error('请在 config.js 中配置 Supabase URL 和 Key');
        return false;
    }
    supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    return true;
}

/**
 * 批量写入/更新库存数据（upsert）
 * @param {Array} items - [{period, material_code, material_name, unit, jiecun}]
 * @param {boolean} isSecondUpload - 是否是第二次上传（只更新结存）
 * @returns {Object} {success, inserted, updated, errors}
 */
async function upsertInventory(items, isSecondUpload = false) {
    const result = { success: true, inserted: 0, updated: 0, errors: [] };

    // 先查询已有数据判断新增/更新
    const { data: existing } = await supabaseClient
        .from('inventory')
        .select('material_code, diankv')
        .eq('period', items[0]?.period);

    const existingMap = new Map();
    (existing || []).forEach(e => existingMap.set(e.material_code, e));

    // 准备 upsert 数据
    const upsertData = items.map(item => {
        const exists = existingMap.has(item.material_code);
        if (exists) result.updated++;
        else result.inserted++;

        const record = {
            period: item.period,
            material_code: item.material_code,
            material_name: item.material_name,
            unit: item.unit,
            jiecun: item.jiecun,
            updated_at: new Date().toISOString()
        };

        // 第一次上传时初始化 diankv 为 null，第二次上传保留已有 diankv
        if (!isSecondUpload && !exists) {
            record.diankv = null;
        }

        return record;
    });

    // 分批 upsert（每批 100 条）
    const batchSize = 100;
    for (let i = 0; i < upsertData.length; i += batchSize) {
        const batch = upsertData.slice(i, i + batchSize);
        const { error } = await supabaseClient
            .from('inventory')
            .upsert(batch, { onConflict: 'period,material_code' });

        if (error) {
            result.success = false;
            result.errors.push(error.message);
        }
    }

    return result;
}

/**
 * 获取所有库存数据
 * @param {string} period - 会计期间
 * @returns {Array}
 */
async function getInventory(period) {
    const { data, error } = await supabaseClient
        .from('inventory')
        .select('*')
        .eq('period', period)
        .order('material_code');

    if (error) {
        console.error('获取库存数据失败:', error);
        return [];
    }
    return data || [];
}

/**
 * 更新单条物料的点库数量
 * @param {string} period
 * @param {string} materialCode
 * @param {number} diankvValue
 */
async function updateDiankv(period, materialCode, diankvValue) {
    const { error } = await supabaseClient
        .from('inventory')
        .update({
            diankv: diankvValue,
            updated_at: new Date().toISOString()
        })
        .eq('period', period)
        .eq('material_code', materialCode);

    if (error) {
        console.error('更新点库数据失败:', error);
        return false;
    }
    return true;
}

/**
 * 获取所有规则
 * @returns {Array}
 */
async function getRules() {
    const { data, error } = await supabaseClient
        .from('rules')
        .select('*')
        .order('material_code');

    if (error) {
        console.error('获取规则失败:', error);
        return [];
    }
    return data || [];
}

/**
 * 新建/修改规则
 * @param {Object} rule
 */
async function upsertRule(rule) {
    const { error } = await supabaseClient
        .from('rules')
        .upsert(rule, { onConflict: 'material_code' });

    if (error) {
        console.error('保存规则失败:', error);
        return false;
    }
    return true;
}

/**
 * 删除规则
 * @param {string} materialCode
 */
async function deleteRule(materialCode) {
    const { error } = await supabaseClient
        .from('rules')
        .delete()
        .eq('material_code', materialCode);

    if (error) {
        console.error('删除规则失败:', error);
        return false;
    }
    return true;
}

/**
 * 获取所有不重复的会计期间
 */
async function getPeriods() {
    const { data, error } = await supabaseClient
        .from('inventory')
        .select('period')
        .order('period', { ascending: false });

    if (error) return [];
    const unique = [...new Set((data || []).map(d => d.period))];
    return unique;
}
