// ============================================================
// app.js - 主应用逻辑（页面路由、交互、数据绑定）
// ============================================================

// 全局状态
let currentPeriod = '';
let inventoryCache = [];
let rulesCache = [];
let reportCache = null;
let diankvFilterMode = 'all'; // 'all' | 'empty' | 'done'
let currentEditItem = null;

// ---- 初始化 ----
document.addEventListener('DOMContentLoaded', async () => {
    // 初始化 Supabase
    if (!initSupabase()) {
        showToast('请先在 config.js 中配置 Supabase', 'error');
        return;
    }

    // 页面路由
    setupNavigation();

    // 上传区域
    setupUploadZone();

    // 搜索功能
    setupSearch();

    // 规则表单
    setupRuleForm();

    // 规则Excel导入
    setupRuleImport();

    // 加载已有期间
    await loadPeriods();

    // 加载规则
    await loadRules();
});

// ---- 页面导航 ----
function setupNavigation() {
    const tabs = document.querySelectorAll('.nav-tab');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const page = tab.dataset.page;
            switchPage(page);
        });
    });
}

function switchPage(pageName) {
    // 更新 tab 高亮
    document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
    document.querySelector(`[data-page="${pageName}"]`).classList.add('active');

    // 切换页面
    document.querySelectorAll('.page-section').forEach(p => p.classList.remove('active'));
    document.getElementById(`page-${pageName}`).classList.add('active');

    // 进入点库页面时刷新数据
    if (pageName === 'diankv' && currentPeriod) {
        loadDiankvList();
    }

    // 进入规则页面时刷新
    if (pageName === 'rules') {
        renderRulesTable();
    }
}

// ---- Toast 通知 ----
function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = 'opacity 0.3s';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// ---- 上传区域 ----
function setupUploadZone() {
    const zone = document.getElementById('uploadZone');
    const input = document.getElementById('fileInput');

    zone.addEventListener('click', () => input.click());

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('dragover');
    });

    zone.addEventListener('dragleave', () => {
        zone.classList.remove('dragover');
    });

    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) handleFileUpload(file);
    });

    input.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleFileUpload(file);
        input.value = '';
    });
}

// ---- 文件上传处理 ----
async function handleFileUpload(file) {
    const resultDiv = document.getElementById('uploadResult');

    // 显示加载中
    resultDiv.innerHTML = `
        <div class="alert alert-info">
            <span class="spinner"></span>
            <span>正在解析 ${file.name} ...</span>
        </div>
    `;

    try {
        // 解析 Excel
        const { period, items } = await parseInventoryExcel(file);

        if (!items.length) {
            resultDiv.innerHTML = '<div class="alert alert-danger">⚠️ 解析结果为空，请检查 Excel 格式</div>';
            return;
        }

        // 判断是否是第二次上传（已有该期间数据）
        const existing = await getInventory(period);
        const isSecondUpload = existing.length > 0;

        // 写入 Supabase
        const result = await upsertInventory(items, isSecondUpload);

        if (result.success) {
            currentPeriod = period;
            resultDiv.innerHTML = `
                <div class="alert alert-success">
                    ✅ 上传成功！期间：${period}，共 ${items.length} 条物料
                    ${isSecondUpload ? `（更新结存 ${result.updated} 条）` : `（新增 ${result.inserted} 条）`}
                </div>
            `;

            // 更新期间选择器
            await loadPeriods();
            // 显示概览
            await refreshOverview();

            showToast('上传成功！', 'success');
        } else {
            resultDiv.innerHTML = `
                <div class="alert alert-danger">
                    ❌ 上传失败：${result.errors.join('，')}
                </div>
            `;
        }
    } catch (err) {
        resultDiv.innerHTML = `
            <div class="alert alert-danger">
                ❌ 解析失败：${err.message}
            </div>
        `;
    }
}

// ---- 加载期间列表 ----
async function loadPeriods() {
    const periods = await getPeriods();
    const select = document.getElementById('periodSelect');
    const info = document.getElementById('periodInfo');

    if (periods.length === 0) {
        select.innerHTML = '<option value="">请先上传 Excel</option>';
        info.textContent = '';
        return;
    }

    select.innerHTML = periods.map(p => `<option value="${p}" ${p === currentPeriod ? 'selected' : ''}>${p}</option>`).join('');

    if (!currentPeriod) {
        currentPeriod = periods[0];
    }

    select.addEventListener('change', async (e) => {
        currentPeriod = e.target.value;
        await refreshOverview();
    });

    info.textContent = `共 ${periods.length} 个期间`;
}

// ---- 刷新概览 ----
async function refreshOverview() {
    if (!currentPeriod) return;

    const data = await getInventory(currentPeriod);
    inventoryCache = data;

    const total = data.length;
    const recorded = data.filter(d => d.diankv !== null && d.diankv !== undefined).length;
    const progress = total > 0 ? Math.round(recorded / total * 100) : 0;

    // 计算结存为负的数量
    const negativeCount = data.filter(d => d.jiecun < 0).length;

    const statsGrid = document.getElementById('statsGrid');
    statsGrid.innerHTML = `
        <div class="stat-card">
            <div class="stat-value">${total}</div>
            <div class="stat-label">物料总数</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${recorded}</div>
            <div class="stat-label">已点库</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${total - recorded}</div>
            <div class="stat-label">未点库</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${negativeCount}</div>
            <div class="stat-label">结存为负</div>
        </div>
    `;

    document.getElementById('progressFill').style.width = `${progress}%`;
    document.getElementById('progressText').textContent = `点库进度：${recorded} / ${total} (${progress}%)`;
    document.getElementById('overviewCard').style.display = 'block';
}

// ---- 点库列表 ----
async function loadDiankvList() {
    if (!currentPeriod) return;

    const data = await getInventory(currentPeriod);
    inventoryCache = data;
    renderDiankvList(data);
}

function renderDiankvList(data) {
    const container = document.getElementById('materialList');

    if (!data || data.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon">📦</div>
                <div class="empty-title">暂无数据</div>
                <div class="empty-desc">请先在电脑端上传 Excel 文件</div>
            </div>
        `;
        return;
    }

    // 过滤
    let filtered = data;
    if (diankvFilterMode === 'empty') {
        filtered = data.filter(d => d.diankv === null || d.diankv === undefined);
    } else if (diankvFilterMode === 'done') {
        filtered = data.filter(d => d.diankv !== null && d.diankv !== undefined);
    }

    // 搜索过滤
    const searchTerm = document.getElementById('searchInput').value.trim().toLowerCase();
    if (searchTerm) {
        filtered = filtered.filter(d =>
            d.material_name.toLowerCase().includes(searchTerm) ||
            d.material_code.toLowerCase().includes(searchTerm)
        );
    }

    // 按物料代码前缀分组
    const groups = {};
    filtered.forEach(item => {
        const prefix = getCodePrefix(item.material_code);
        let groupName;
        if (prefix === '01_02') {
            // 更细的分组
            if (item.material_code.startsWith('01.01')) groupName = '肉类';
            else if (item.material_code.startsWith('01.02')) groupName = '蔬菜';
            else if (item.material_code.startsWith('01.03')) groupName = '调料';
            else if (item.material_code.startsWith('02.')) groupName = '02类物料';
            else groupName = '01类其他';
        } else if (prefix === '03') {
            groupName = '03类物料';
        } else if (prefix === '04') {
            groupName = '04类物料';
        } else {
            groupName = '其他';
        }

        if (!groups[groupName]) groups[groupName] = [];
        groups[groupName].push(item);
    });

    // 更新进度
    const total = data.length;
    const recorded = data.filter(d => d.diankv !== null && d.diankv !== undefined).length;
    const progress = total > 0 ? Math.round(recorded / total * 100) : 0;
    document.getElementById('diankvProgress').textContent = `已录入 ${recorded} / ${total}`;
    document.getElementById('diankvProgressFill').style.width = `${progress}%`;

    // 渲染列表
    let html = '';
    for (const [groupName, items] of Object.entries(groups)) {
        const groupRecorded = items.filter(i => i.diankv !== null && i.diankv !== undefined).length;
        html += `
            <div class="material-group-header" onclick="toggleGroup(this)">
                <span>▸ ${groupName}</span>
                <span class="count">${groupRecorded}/${items.length}</span>
            </div>
            <div class="material-group-items" style="display: none;">
        `;

        items.forEach(item => {
            const hasDiankv = item.diankv !== null && item.diankv !== undefined;
            html += `
                <div class="material-item" onclick="openDiankvModal('${item.material_code}')">
                    <div class="material-info">
                        <div class="material-name">${item.material_name}</div>
                        <div class="material-meta">${item.material_code} · ${item.unit} · 结存: ${item.jiecun ?? 0}</div>
                    </div>
                    <div class="material-diankv ${hasDiankv ? 'recorded' : 'empty'}">
                        ${hasDiankv ? item.diankv : '未录入'}
                    </div>
                </div>
            `;
        });

        html += '</div>';
    }

    container.innerHTML = html;
}

function toggleGroup(header) {
    const items = header.nextElementSibling;
    const isOpen = items.style.display !== 'none';
    items.style.display = isOpen ? 'none' : 'block';
    header.querySelector('span:first-child').textContent = (isOpen ? '▸ ' : '▾ ') + header.querySelector('span:first-child').textContent.slice(2);
}

function toggleDiankvFilter() {
    const btn = document.getElementById('diankvFilter');
    if (diankvFilterMode === 'all') {
        diankvFilterMode = 'empty';
        btn.textContent = '仅未录入';
        btn.className = 'badge badge-warning';
    } else if (diankvFilterMode === 'empty') {
        diankvFilterMode = 'done';
        btn.textContent = '仅已录入';
        btn.className = 'badge badge-success';
    } else {
        diankvFilterMode = 'all';
        btn.textContent = '全部';
        btn.className = 'badge badge-info';
    }
    renderDiankvList(inventoryCache);
}

// ---- 点库弹窗 ----
function openDiankvModal(materialCode) {
    const item = inventoryCache.find(i => i.material_code === materialCode);
    if (!item) return;

    currentEditItem = item;

    document.getElementById('modalTitle').textContent = item.material_name;
    document.getElementById('modalSubtitle').textContent = `${item.material_code} · ${item.unit}`;

    const jiecun = item.jiecun ?? 0;
    const jiecunEl = document.getElementById('modalJiecun');
    jiecunEl.textContent = jiecun;
    jiecunEl.className = 'modal-field-value' + (jiecun < 0 ? ' negative' : '');

    const input = document.getElementById('modalDiankvInput');
    input.value = item.diankv ?? '';

    // 计算差异
    updateModalDiff();

    document.getElementById('diankvModal').classList.add('open');

    // 聚焦输入框
    setTimeout(() => input.focus(), 300);

    // 监听输入实时更新差异
    input.oninput = updateModalDiff;
}

function updateModalDiff() {
    const input = document.getElementById('modalDiankvInput');
    const diffEl = document.getElementById('modalDiff');
    const val = parseFloat(input.value);

    if (isNaN(val) || !currentEditItem) {
        diffEl.textContent = '—';
        diffEl.className = 'modal-field-value';
        return;
    }

    const jiecun = currentEditItem.jiecun ?? 0;
    const diff = val - jiecun;
    diffEl.textContent = diff > 0 ? `+${diff}` : diff;
    diffEl.className = 'modal-field-value' + (diff < 0 ? ' negative' : '');
}

function closeModal() {
    document.getElementById('diankvModal').classList.remove('open');
    currentEditItem = null;
}

async function saveDiankv() {
    if (!currentEditItem) return;

    const input = document.getElementById('modalDiankvInput');
    const val = input.value.trim();

    // 允许清空（设为 null）或输入数字
    let diankvValue = null;
    if (val !== '') {
        diankvValue = parseFloat(val);
        if (isNaN(diankvValue)) {
            showToast('请输入有效数字', 'error');
            return;
        }
    }

    const success = await updateDiankv(currentPeriod, currentEditItem.material_code, diankvValue);

    if (success) {
        // 更新缓存
        const idx = inventoryCache.findIndex(i => i.material_code === currentEditItem.material_code);
        if (idx !== -1) {
            inventoryCache[idx].diankv = diankvValue;
        }
        renderDiankvList(inventoryCache);
        closeModal();
        showToast('保存成功', 'success');
    } else {
        showToast('保存失败，请重试', 'error');
    }
}

// ---- 搜索功能 ----
function setupSearch() {
    // 点库搜索
    document.getElementById('searchInput').addEventListener('input', () => {
        renderDiankvList(inventoryCache);
    });

    // 规则搜索
    document.getElementById('ruleSearchInput').addEventListener('input', () => {
        renderRulesTable();
    });
}

// ---- 报表生成 ----
async function generateReports() {
    if (!currentPeriod) {
        showToast('请先选择会计期间', 'error');
        return;
    }

    // 获取最新数据
    const data = await getInventory(currentPeriod);
    inventoryCache = data;

    if (data.length === 0) {
        showToast('没有库存数据', 'error');
        return;
    }

    // 获取规则
    const rules = await getRules();

    // 运行规则引擎
    const result = processInventory(data, rules);
    reportCache = result;

    // 显示汇总
    document.getElementById('reportResult').innerHTML = `
        <div class="alert alert-success">✅ 报表生成完成</div>
    `;

    const statsDiv = document.getElementById('reportStats');
    statsDiv.innerHTML = `
        <div class="stat-card">
            <div class="stat-value">${result.summary.total}</div>
            <div class="stat-label">物料总数</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${result.summary.processed}</div>
            <div class="stat-label">已处理</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${result.summary.inquired}</div>
            <div class="stat-label">需询问</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${result.summary.noRule}</div>
            <div class="stat-label">无规则</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${result.summary.skipped}</div>
            <div class="stat-label">已跳过</div>
        </div>
    `;

    // 渲染各报表预览
    const previewsDiv = document.getElementById('reportPreviews');
    let previewHtml = '';

    for (const [name, items] of Object.entries(result.reports)) {
        if (items.length === 0) continue;

        previewHtml += `
            <div class="card">
                <div class="card-header">
                    <div class="card-title">📄 ${name}</div>
                    <div style="display: flex; gap: 8px; align-items: center;">
                        <span class="badge badge-success">${items.length} 条</span>
                        <button class="btn btn-sm btn-secondary" onclick="downloadSingleReport('${name}')">📥 下载</button>
                    </div>
                </div>
                <div class="table-container" style="max-height: 300px; overflow-y: auto;">
                    <table>
                        <thead>
                            <tr><th>物料名称</th><th>数量</th></tr>
                        </thead>
                        <tbody>
                            ${items.map(i => `<tr><td>${i.material_name}</td><td>${i.value}</td></tr>`).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        `;
    }

    previewsDiv.innerHTML = previewHtml;

    // 渲染询问清单
    if (result.inquiries.length > 0) {
        document.getElementById('inquiryCard').style.display = 'block';
        document.getElementById('inquiryBody').innerHTML = result.inquiries.map(i => `
            <tr>
                <td>${i.material_code}</td>
                <td>${i.material_name}</td>
                <td>${i.jiecun ?? 0}</td>
                <td>${i.diankv ?? '未录'}</td>
                <td>${i.diff ?? '—'}</td>
                <td><span class="badge badge-warning">${i.reason}</span></td>
            </tr>
        `).join('');
    } else {
        document.getElementById('inquiryCard').style.display = 'none';
    }

    // 渲染无规则物料
    if (result.noRule.length > 0) {
        document.getElementById('noRuleCard').style.display = 'block';
        document.getElementById('noRuleBody').innerHTML = result.noRule.map(i => `
            <tr>
                <td>${i.material_code}</td>
                <td>${i.material_name}</td>
                <td>${i.unit}</td>
                <td><button class="btn btn-sm btn-primary" onclick="openRuleDialogFromNoRule('${i.material_code}', '${i.material_name}', '${i.unit}')">➕ 新建规则</button></td>
            </tr>
        `).join('');
    } else {
        document.getElementById('noRuleCard').style.display = 'none';
    }

    document.getElementById('reportSummary').style.display = 'block';
    showToast('报表生成完成！', 'success');
}

function downloadSingleReport(reportName) {
    if (!reportCache) return;
    const data = reportCache.reports[reportName];
    if (!data || data.length === 0) {
        showToast('该报表无数据', 'error');
        return;
    }
    exportToExcel(`${currentPeriod}_${reportName}`, data, reportName);
}

function downloadAllReports() {
    if (!reportCache) return;
    exportAllReports(reportCache.reports, `${currentPeriod}_库存盘点报表`);
    showToast('下载完成！', 'success');
}

// ---- 规则管理 ----
async function loadRules() {
    rulesCache = await getRules();
    renderRulesTable();
}

function renderRulesTable() {
    const searchTerm = document.getElementById('ruleSearchInput').value.trim().toLowerCase();
    let filtered = rulesCache;
    if (searchTerm) {
        filtered = rulesCache.filter(r =>
            r.material_name.toLowerCase().includes(searchTerm) ||
            r.material_code.toLowerCase().includes(searchTerm)
        );
    }

    const tbody = document.getElementById('rulesBody');

    if (filtered.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="11" style="text-align:center; color: var(--text-muted); padding: 40px;">
                    ${searchTerm ? '无匹配结果' : '暂无规则数据，请导入或新建'}
                </td>
            </tr>
        `;
        return;
    }

    tbody.innerHTML = filtered.map(r => `
        <tr>
            <td style="font-size: 0.8rem;">${r.material_code}</td>
            <td>${r.material_name}</td>
            <td>${r.unit || ''}</td>
            <td><span class="badge ${
                r.attribute === '正常' ? 'badge-success' :
                r.attribute === '不动' ? 'badge-info' : 'badge-warning'
            }">${r.attribute}</span></td>
            <td>${r.ganlia || ''}</td>
            <td>${r.shiliao || ''}</td>
            <td>${r.shushi || ''}</td>
            <td>${r.xiancai || ''}</td>
            <td>${r.douyou || ''}</td>
            <td>${r.yuangongcan || ''}</td>
            <td>
                <button class="btn btn-sm btn-secondary" onclick="editRule('${r.material_code}')">✏️</button>
            </td>
        </tr>
    `).join('');
}

function openRuleDialog(rule = null) {
    document.getElementById('ruleDialogTitle').textContent = rule ? '编辑规则' : '新建规则';
    document.getElementById('ruleCode').value = rule?.material_code || '';
    document.getElementById('ruleCode').disabled = !!rule;
    document.getElementById('ruleName').value = rule?.material_name || '';
    document.getElementById('ruleUnit').value = rule?.unit || '';
    document.getElementById('ruleAttribute').value = rule?.attribute || '正常';
    document.getElementById('ruleGanlia').value = rule?.ganlia || '';
    document.getElementById('ruleShiliao').value = rule?.shiliao || '';
    document.getElementById('ruleShushi').value = rule?.shushi || '';
    document.getElementById('ruleXiancai').value = rule?.xiancai || '';
    document.getElementById('ruleDouyou').value = rule?.douyou || '';
    document.getElementById('ruleYuangongcan').value = rule?.yuangongcan || '';

    document.getElementById('ruleDialog').classList.add('open');
}

function openRuleDialogFromNoRule(code, name, unit) {
    openRuleDialog({ material_code: code, material_name: name, unit: unit });
    // 允许编辑代码（因为是新建）
    document.getElementById('ruleCode').disabled = false;
    // 切到规则 tab
    switchPage('rules');
}

function editRule(code) {
    const rule = rulesCache.find(r => r.material_code === code);
    if (rule) openRuleDialog(rule);
}

function closeRuleDialog() {
    document.getElementById('ruleDialog').classList.remove('open');
}

function setupRuleForm() {
    document.getElementById('ruleForm').addEventListener('submit', async (e) => {
        e.preventDefault();

        const rule = {
            material_code: document.getElementById('ruleCode').value.trim(),
            material_name: document.getElementById('ruleName').value.trim(),
            unit: document.getElementById('ruleUnit').value.trim(),
            attribute: document.getElementById('ruleAttribute').value,
            ganlia: parseFloat(document.getElementById('ruleGanlia').value) || null,
            shiliao: parseFloat(document.getElementById('ruleShiliao').value) || null,
            shushi: parseFloat(document.getElementById('ruleShushi').value) || null,
            xiancai: parseFloat(document.getElementById('ruleXiancai').value) || null,
            douyou: parseFloat(document.getElementById('ruleDouyou').value) || null,
            yuangongcan: parseFloat(document.getElementById('ruleYuangongcan').value) || null,
        };

        if (!rule.material_code || !rule.material_name) {
            showToast('物料代码和名称不能为空', 'error');
            return;
        }

        const success = await upsertRule(rule);
        if (success) {
            closeRuleDialog();
            await loadRules();
            showToast('规则保存成功', 'success');
        } else {
            showToast('规则保存失败', 'error');
        }
    });
}

// ---- 规则 Excel 导入 ----
function setupRuleImport() {
    const input = document.getElementById('ruleFileInput');
    input.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        try {
            const data = await readRuleExcel(file);
            if (data.length === 0) {
                showToast('未解析到规则数据', 'error');
                return;
            }

            // 批量 upsert
            let successCount = 0;
            const batchSize = 50;
            for (let i = 0; i < data.length; i += batchSize) {
                const batch = data.slice(i, i + batchSize);
                const { error } = await supabaseClient
                    .from('rules')
                    .upsert(batch, { onConflict: 'material_code' });

                if (!error) successCount += batch.length;
            }

            await loadRules();
            showToast(`成功导入 ${successCount} 条规则`, 'success');
        } catch (err) {
            showToast('导入失败：' + err.message, 'error');
        }

        input.value = '';
    });
}

function importRulesFromExcel() {
    document.getElementById('ruleFileInput').click();
}

/**
 * 解析规则 Excel
 */
function readRuleExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const wb = XLSX.read(data, { type: 'array' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

                if (rows.length < 2) {
                    resolve([]);
                    return;
                }

                // 查找列（根据物料处理规则.xlsx 的真实格式）
                const headers = rows[0];
                const colMap = {
                    code: findCol(headers, ['物料代码']),
                    name: findCol(headers, ['物料名称']),
                    unit: findCol(headers, ['单位']),
                    attr: findCol(headers, ['属性']),
                    ganlia: findCol(headers, ['干料']),
                    shiliao: findCol(headers, ['湿料']),
                    shushi: findCol(headers, ['熟食']),
                    xiancai: findCol(headers, ['咸菜']),
                    douyou: findCol(headers, ['豆油']),
                    yuangongcan: findCol(headers, ['员工餐']),
                };

                const rules = [];
                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row || !row[colMap.code]) continue;

                    rules.push({
                        material_code: String(row[colMap.code]).trim(),
                        material_name: String(row[colMap.name] || '').trim(),
                        unit: String(row[colMap.unit] || '').trim(),
                        attribute: String(row[colMap.attr] || '正常').trim(),
                        ganlia: parseFloat(row[colMap.ganlia]) || null,
                        shiliao: parseFloat(row[colMap.shiliao]) || null,
                        shushi: parseFloat(row[colMap.shushi]) || null,
                        xiancai: parseFloat(row[colMap.xiancai]) || null,
                        douyou: parseFloat(row[colMap.douyou]) || null,
                        yuangongcan: parseFloat(row[colMap.yuangongcan]) || null,
                    });
                }

                resolve(rules);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error('读取文件失败'));
        reader.readAsArrayBuffer(file);
    });
}

function findCol(headers, candidates) {
    for (const c of candidates) {
        const idx = headers.indexOf(c);
        if (idx !== -1) return idx;
    }
    return -1;
}

// ---- 点库弹窗：点击外部关闭 ----
document.getElementById('diankvModal').addEventListener('click', (e) => {
    if (e.target.id === 'diankvModal') closeModal();
});

document.getElementById('ruleDialog').addEventListener('click', (e) => {
    if (e.target.id === 'ruleDialog') closeRuleDialog();
});

// ---- 键盘快捷键 ----
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        closeModal();
        closeRuleDialog();
    }
});
