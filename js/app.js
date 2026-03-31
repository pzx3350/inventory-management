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
const collapsedMobileGroups = new Set(); // 记录手机端已折叠的分组名

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

    // 手机端默认进点库页
    if (window.innerWidth <= 768) {
        switchPage('diankv');
        setTimeout(() => {
            const s = document.getElementById('searchInputMobile');
            if (s) s.focus();
        }, 400);
    }
});

// ---- 页面导航 ----
function setupNavigation() {
    document.querySelectorAll('.nav-item, .mobile-nav-item').forEach(tab => {
        tab.addEventListener('click', () => {
            if (tab.dataset.page) switchPage(tab.dataset.page);
        });
    });
}

function switchPage(pageName) {
    // 更新侧边栏高亮
    document.querySelectorAll('.nav-item').forEach(t => t.classList.remove('active'));
    const sidebarItem = document.querySelector(`.nav-item[data-page="${pageName}"]`);
    if (sidebarItem) sidebarItem.classList.add('active');

    // 更新底部导航高亮
    document.querySelectorAll('.mobile-nav-item').forEach(t => t.classList.remove('active'));
    const mobileItem = document.querySelector(`.mobile-nav-item[data-page="${pageName}"]`);
    if (mobileItem) mobileItem.classList.add('active');

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

// ---- 物料分组 ----
function normCode(code) {
    // 将物料代码每段补齐到3位，总共补到5段，用于范围比较
    const segs = code.split('.');
    while (segs.length < 5) segs.push('0');
    return segs.map(s => s.padStart(3, '0')).join('.');
}

function getMaterialGroup(code) {
    const n = normCode(code);
    if (n >= normCode('01.03.01.01.001') && n <= normCode('01.03.03.009')) return '调料';
    if (n >= normCode('03.01.01.01.001') && n <= normCode('03.01.01.02.018')) return '干料';
    if ((n >= normCode('03.04.001') && n <= normCode('03.05.332')) ||
        (n >= normCode('03.07.019') && n <= normCode('03.07.021'))) return '咸菜';
    if (code.startsWith('04.')) return '库存商品';
    return '其他';
}

const GROUP_ORDER = ['调料', '干料', '咸菜', '库存商品', '其他'];

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

    // 搜索过滤（支持汉字、全拼、首字母）
    const searchTerm = document.getElementById('searchInput').value.trim();
    if (searchTerm) {
        filtered = filtered.filter(d =>
            fuzzyMatch(d.material_name, searchTerm) ||
            d.material_code.toLowerCase().includes(searchTerm.toLowerCase())
        );
    }

    // 更新进度
    const total = data.length;
    const recorded = data.filter(d => d.diankv !== null && d.diankv !== undefined).length;
    const progress = total > 0 ? Math.round(recorded / total * 100) : 0;
    document.getElementById('diankvProgress').textContent = `已录入 ${recorded} / ${total}`;
    document.getElementById('diankvProgressFill').style.width = `${progress}%`;

    // 按分组渲染
    const groups = {};
    GROUP_ORDER.forEach(g => groups[g] = []);
    filtered.forEach(item => {
        const g = getMaterialGroup(item.material_code);
        groups[g].push(item);
    });

    let html = '';
    GROUP_ORDER.forEach(groupName => {
        const items = groups[groupName];
        if (items.length === 0) return;
        const doneCount = items.filter(i => i.diankv !== null && i.diankv !== undefined).length;
        const isCollapsed = collapsedMobileGroups.has(groupName);
        html += `
            <div class="material-group-header" data-group="${groupName}" onclick="toggleGroup(this)">
                <span>${isCollapsed ? '▸' : '▾'} ${groupName}</span>
                <span class="count">${doneCount}/${items.length}</span>
            </div>
            <div class="group-items" style="${isCollapsed ? 'display:none' : ''}">
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
        html += `</div>`;
    });

    container.innerHTML = html;

    // 同步渲染桌面表格
    renderDiankvTable(data);
}

// ---- 桌面端点库表格 ----
function renderDiankvTable(data) {
    const container = document.getElementById('diankvDesktopWrap');
    if (!container) return;

    if (!data || data.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon">📦</div>
                <div class="empty-title">暂无数据</div>
                <div class="empty-desc">请先在电脑端上传 Excel 文件</div>
            </div>`;
        return;
    }

    // 应用同 renderDiankvList 相同的过滤逻辑
    let filtered = data;
    if (diankvFilterMode === 'empty') {
        filtered = data.filter(d => d.diankv === null || d.diankv === undefined);
    } else if (diankvFilterMode === 'done') {
        filtered = data.filter(d => d.diankv !== null && d.diankv !== undefined);
    }

    const searchTerm = document.getElementById('searchInput').value.trim();
    if (searchTerm) {
        filtered = filtered.filter(d =>
            fuzzyMatch(d.material_name, searchTerm) ||
            d.material_code.toLowerCase().includes(searchTerm.toLowerCase())
        );
    }

    // 按分组整理
    const groups = {};
    GROUP_ORDER.forEach(g => groups[g] = []);
    filtered.forEach(item => {
        const g = getMaterialGroup(item.material_code);
        groups[g].push(item);
    });

    let html = `<table class="diankv-desktop-table">
        <colgroup>
            <col class="col-num">
            <col class="col-name">
            <col class="col-unit">
            <col class="col-num2">
            <col class="col-num2">
            <col class="col-num2">
        </colgroup>
        <thead>
            <tr>
                <th style="text-align:center;">#</th>
                <th>物料名称</th>
                <th>计量单位</th>
                <th style="text-align:right;">结存</th>
                <th style="text-align:right;">点库</th>
                <th style="text-align:right;">差异</th>
            </tr>
        </thead>
        <tbody>`;

    let groupSeq = 0;
    GROUP_ORDER.forEach(groupName => {
        const items = groups[groupName];
        if (items.length === 0) return;

        const gid = `g${groupSeq++}`;
        const doneCount = items.filter(i => i.diankv !== null && i.diankv !== undefined).length;

        html += `<tr class="diankv-group-row" data-gid="${gid}" onclick="toggleDesktopGroup('${gid}')">
            <td colspan="6">
                <span class="group-toggle-icon">▾</span>
                ${groupName}
                <span class="group-count">${doneCount} / ${items.length}</span>
            </td>
        </tr>`;

        items.forEach((item, idx) => {
            const hasDiankv = item.diankv !== null && item.diankv !== undefined;
            const jiecun = item.jiecun ?? 0;
            // 点库未输入视为 0 计算差异，parseFloat+toFixed 消除浮点误差
            const diankvVal = hasDiankv ? item.diankv : 0;
            const diff = parseFloat((diankvVal - jiecun).toFixed(4));
            const diffStr = diff > 0 ? `+${diff}` : `${diff}`;
            const diffClass = diff !== 0 ? (diff > 0 ? 'diff-positive' : 'diff-negative') : '';

            html += `<tr class="diankv-data-row" data-gmember="${gid}">
                <td class="col-num">${idx + 1}</td>
                <td class="cell-name" title="${item.material_name}">${item.material_name}</td>
                <td>${item.unit}</td>
                <td class="col-right">${jiecun}</td>
                <td class="diankv-cell ${hasDiankv ? 'has-value' : ''}"
                    data-code="${item.material_code}"
                    data-jiecun="${jiecun}"
                    onclick="startInlineEdit(this)">
                    ${hasDiankv ? item.diankv : '<span class="placeholder">点击输入</span>'}
                </td>
                <td class="col-right ${diffClass}">${diffStr}</td>
            </tr>`;
        });
    });

    html += '</tbody></table>';
    container.innerHTML = html;
}

// ---- 桌面端内联编辑 ----
function startInlineEdit(cell) {
    if (cell.querySelector('input')) return; // 已在编辑中

    const code = cell.dataset.code;
    const jiecun = parseFloat(cell.dataset.jiecun) || 0;
    const item = inventoryCache.find(i => i.material_code === code);
    const currentVal = (item && item.diankv !== null && item.diankv !== undefined) ? item.diankv : '';

    const input = document.createElement('input');
    input.type = 'number';
    input.step = 'any';
    input.value = currentVal;
    input.className = 'inline-edit-input';
    input.onclick = e => e.stopPropagation();

    cell.innerHTML = '';
    cell.appendChild(input);
    input.focus();
    input.select();

    // 实时更新差异列
    input.addEventListener('input', () => {
        const val = parseFloat(input.value);
        const diffCell = cell.parentElement.querySelector('td:last-child');
        if (!diffCell) return;
        if (isNaN(val)) {
            diffCell.textContent = '—';
            diffCell.className = 'col-right';
        } else {
            const diff = parseFloat((val - jiecun).toFixed(4));
            diffCell.textContent = diff > 0 ? `+${diff}` : `${diff}`;
            diffCell.className = 'col-right ' + (diff > 0 ? 'diff-positive' : diff < 0 ? 'diff-negative' : '');
        }
    });

    let saved = false;

    async function save() {
        if (saved) return;
        saved = true;

        const val = input.value.trim();
        let diankvValue = null;
        if (val !== '') {
            diankvValue = parseFloat(val);
            if (isNaN(diankvValue)) {
                saved = false;
                input.focus();
                return;
            }
        }

        const success = await updateDiankv(currentPeriod, code, diankvValue);
        if (success) {
            const idx = inventoryCache.findIndex(i => i.material_code === code);
            if (idx !== -1) inventoryCache[idx].diankv = diankvValue;
            // 只刷新表格，不重渲移动端列表（避免滚动位置跳动）
            renderDiankvTable(inventoryCache);
            // 更新顶部进度
            const total = inventoryCache.length;
            const recorded = inventoryCache.filter(d => d.diankv !== null && d.diankv !== undefined).length;
            const progress = total > 0 ? Math.round(recorded / total * 100) : 0;
            document.getElementById('diankvProgress').textContent = `已录入 ${recorded} / ${total}`;
            document.getElementById('diankvProgressFill').style.width = `${progress}%`;
        } else {
            showToast('保存失败，请重试', 'error');
            saved = false;
            renderDiankvTable(inventoryCache);
        }
    }

    input.addEventListener('blur', save);
    input.addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
        if (e.key === 'Escape') {
            saved = true; // 阻止 blur 触发保存
            renderDiankvTable(inventoryCache);
        }
    });
}

// ---- 一键导出点库数据 ----
function exportDiankvExcel() {
    if (!inventoryCache || inventoryCache.length === 0) {
        showToast('暂无数据可导出', 'error');
        return;
    }

    const wsData = [['物料代码', '物料名称', '计量单位', '结存', '点库', '差异']];

    inventoryCache.forEach(item => {
        const hasDiankv = item.diankv !== null && item.diankv !== undefined;
        const diff = hasDiankv ? (item.diankv - (item.jiecun ?? 0)) : '';
        wsData.push([
            item.material_code,
            item.material_name,
            item.unit,
            item.jiecun ?? 0,
            hasDiankv ? item.diankv : '',
            diff
        ]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [
        { wch: 20 }, { wch: 30 }, { wch: 10 },
        { wch: 12 }, { wch: 12 }, { wch: 12 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, '点库数据');
    XLSX.writeFile(wb, `${currentPeriod || '点库'}_点库数据.xlsx`);
    showToast('导出成功！', 'success');
}

// ---- 桌面端分组折叠 ----
function toggleDesktopGroup(gid) {
    const rows = document.querySelectorAll(`tr[data-gmember="${gid}"]`);
    const header = document.querySelector(`tr[data-gid="${gid}"]`);
    const icon = header ? header.querySelector('.group-toggle-icon') : null;
    if (!rows.length) return;

    const isOpen = rows[0].style.display !== 'none';
    rows.forEach(r => { r.style.display = isOpen ? 'none' : ''; });
    if (icon) icon.textContent = isOpen ? '▸' : '▾';
}

function toggleGroup(header) {
    const groupName = header.dataset.group;
    const items = header.nextElementSibling;
    const isOpen = items.style.display !== 'none';
    items.style.display = isOpen ? 'none' : 'block';
    header.querySelector('span:first-child').textContent = (isOpen ? '▸ ' : '▾ ') + groupName;
    if (isOpen) collapsedMobileGroups.add(groupName);
    else collapsedMobileGroups.delete(groupName);
}

async function clearAllDiankv() {
    if (!currentPeriod || !inventoryCache || inventoryCache.length === 0) return;
    if (!confirm('确定清空本期所有点库记录？')) return;
    // 分批处理，每批50条，避免URL过长
    const codes = inventoryCache.map(i => i.material_code);
    const batchSize = 50;
    for (let i = 0; i < codes.length; i += batchSize) {
        const batch = codes.slice(i, i + batchSize);
        const { error } = await supabaseClient
            .from('inventory')
            .update({ diankv: null, updated_at: new Date().toISOString() })
            .eq('period', currentPeriod)
            .in('material_code', batch);
        if (error) { showToast('清空失败，请重试', 'error'); return; }
    }
    showToast('已清空本期所有点库记录', 'success');
    await loadDiankvList();
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
        // 保存后搜索框清空并聚焦
        setTimeout(() => {
            const search = document.getElementById('searchInput');
            const searchMobile = document.getElementById('searchInputMobile');
            if (search) { search.value = ''; search.dispatchEvent(new Event('input')); }
            if (searchMobile) { searchMobile.value = ''; searchMobile.focus(); }
            else if (search) { search.focus(); }
        }, 50);
    } else {
        showToast('保存失败，请重试', 'error');
    }
}

// 获取当前生效的搜索框（手机用底部，桌面用顶部）
function getSearchInput() {
    const mobile = document.getElementById('searchInputMobile');
    if (mobile && window.getComputedStyle(mobile).display !== 'none') return mobile;
    return document.getElementById('searchInput');
}

// ---- 搜索功能 ----
function setupSearch() {
    // 桌面搜索框
    document.getElementById('searchInput').addEventListener('input', () => {
        renderDiankvList(inventoryCache);
    });

    // 手机底部搜索框
    const mobileInput = document.getElementById('searchInputMobile');
    if (mobileInput) {
        mobileInput.addEventListener('input', () => {
            // 同步到顶部搜索框（renderDiankvList 读取顶部）
            document.getElementById('searchInput').value = mobileInput.value;
            renderDiankvList(inventoryCache);
        });
    }

    // 规则搜索
    document.getElementById('ruleSearchInput').addEventListener('input', () => {
        renderRulesTable();
    });
}

// ---- 产品入库单辅助函数 ----

// 期间末日：'2026.3' → '2026/03/31'
function getPeriodLastDay(period) {
    const [y, m] = period.split('.').map(Number);
    const lastDay = new Date(y, m, 0).getDate(); // m不减1，Date自动溢出到上月末
    return `${y}/${String(m).padStart(2,'0')}/${String(lastDay).padStart(2,'0')}`;
}

// 收货仓库
function getWarehouse(code) {
    if (code.startsWith('01.') || code.startsWith('02.')) return '原材料仓';
    if (code.startsWith('03.')) return '产成品仓';
    if (code.startsWith('04.')) return '库存商品仓';
    return '';
}

// 档口 → 交货单位
const VENDOR_MAP = {
    ganlia: '干料车间',
    shiliao: '湿料车间',
    shushi: '熟食车间',
    xiancai: '咸菜车间',
    douyou: '豆油混合',
    yuangongcan: '员工餐车间'
};

// CIN编号递增：'CIN008044' + 1 → 'CIN008045'
function incrementCin(cin, n) {
    const num = parseInt(cin.replace(/^CIN0*/,''), 10) + n;
    return 'CIN' + String(num).padStart(6, '0');
}

// 构建产品入库单行数据
function buildRukkuRows(reports, period, startCin) {
    const date = getPeriodLastDay(period);
    const rows = [];
    let groupIndex = 0;

    for (const cat of CATEGORIES) {
        const items = reports[`${CATEGORY_NAMES[cat]}入库单`];
        if (!items || items.length === 0) continue;
        const cin = incrementCin(startCin, groupIndex);
        const vendor = VENDOR_MAP[cat] || CATEGORY_NAMES[cat];
        for (const item of items) {
            rows.push({
                日期: date,
                交货单位: vendor,
                编号: cin,
                验收: '李磊',
                保管: '李磊',
                物料编码: item.material_code,
                单位: item.unit === '斤' ? 'kg' : item.unit,
                实收数量: item.value,
                收货仓库: getWarehouse(item.material_code),
                仓位: ''
            });
        }
        groupIndex++;
    }
    return rows;
}

// 渲染产品入库单 HTML 表格
function renderRukkuTable() {
    if (!reportCache) return;
    const startCin = document.getElementById('cinInput').value.trim() || 'CIN000000';
    const rows = buildRukkuRows(reportCache.reports, currentPeriod, startCin);

    if (rows.length === 0) {
        document.getElementById('rukkuTableWrap').innerHTML = '<div class="empty-state"><div class="empty-title">无入库数据</div></div>';
        return;
    }

    let html = `<table><thead><tr>
        <th>日期</th><th>交货单位</th><th>编号</th><th>验收</th><th>保管</th>
        <th>物料编码</th><th>单位</th><th>实收数量</th><th>收货仓库</th><th>仓位</th>
    </tr></thead><tbody>`;

    rows.forEach(r => {
        html += `<tr>
            <td>${r.日期}</td>
            <td>${r.交货单位}</td>
            <td>${r.编号}</td>
            <td>${r.验收}</td>
            <td>${r.保管}</td>
            <td>${r.物料编码}</td>
            <td>${r.单位}</td>
            <td>${r.实收数量}</td>
            <td>${r.收货仓库}</td>
            <td>${r.仓位}</td>
        </tr>`;
    });
    html += '</tbody></table>';
    document.getElementById('rukkuTableWrap').innerHTML = html;
}

// ---- 生产领料单辅助函数 ----

// SOUT编号递增
function incrementSout(sout, n) {
    const num = parseInt(sout.replace(/^SOUT0*/,''), 10) + n;
    return 'SOUT' + String(num).padStart(6, '0');
}

function buildLiaoliaoRows(reports, period, startSout) {
    const date = getPeriodLastDay(period);
    const rows = [];
    let groupIndex = 0;

    for (const cat of CATEGORIES) {
        const items = reports[`${CATEGORY_NAMES[cat]}领料单`];
        if (!items || items.length === 0) continue;
        const sout = incrementSout(startSout, groupIndex);
        const dept = VENDOR_MAP[cat] || CATEGORY_NAMES[cat];
        for (const item of items) {
            rows.push({
                日期: date,
                领料部门: dept,
                编号: sout,
                领料: '李磊',
                发料: '李磊',
                领料类型: '一般领料',
                物料代码: item.material_code,
                是否返工: '否',
                单位: item.unit === '斤' ? 'kg' : item.unit,
                实发数量: item.value,
                发料仓库: getWarehouse(item.material_code),
                仓位: ''
            });
        }
        groupIndex++;
    }
    return rows;
}

function renderLiaoliaoTable() {
    if (!reportCache) return;
    const startSout = document.getElementById('soutInput').value.trim() || 'SOUT000000';
    const rows = buildLiaoliaoRows(reportCache.reports, currentPeriod, startSout);

    if (rows.length === 0) {
        document.getElementById('liaoliaoTableWrap').innerHTML = '<div class="empty-state"><div class="empty-title">无领料数据</div></div>';
        return;
    }

    let html = `<table><thead><tr>
        <th>日期</th><th>领料部门</th><th>编号</th><th>领料</th><th>发料</th>
        <th>领料类型</th><th>物料代码</th><th>是否返工</th><th>单位</th><th>实发数量</th><th>发料仓库</th><th>仓位</th>
    </tr></thead><tbody>`;

    rows.forEach(r => {
        html += `<tr>
            <td>${r.日期}</td>
            <td>${r.领料部门}</td>
            <td>${r.编号}</td>
            <td>${r.领料}</td>
            <td>${r.发料}</td>
            <td>${r.领料类型}</td>
            <td>${r.物料代码}</td>
            <td>${r.是否返工}</td>
            <td>${r.单位}</td>
            <td>${r.实发数量}</td>
            <td>${r.发料仓库}</td>
            <td>${r.仓位}</td>
        </tr>`;
    });
    html += '</tbody></table>';
    document.getElementById('liaoliaoTableWrap').innerHTML = html;
}

// 将期间末日转为 Excel 日期序号
function getPeriodLastDateSerial(period) {
    const [y, m] = period.split('.').map(Number);
    const d = new Date(y, m - 1, new Date(y, m, 0).getDate());
    return Math.floor((d - new Date(1899, 11, 30)) / 86400000);
}

// 通用：拉取模板 → 清空数据行 → 填入新数据 → 下载
async function exportFromTemplate(templatePath, sheetName, dataRows, filename) {
    let templateBuf;
    try {
        const resp = await fetch(templatePath);
        if (!resp.ok) throw new Error('模板文件加载失败');
        templateBuf = await resp.arrayBuffer();
    } catch (e) {
        showToast('模板加载失败：' + e.message, 'error');
        return;
    }

    const wb = XLSX.read(new Uint8Array(templateBuf), { type: 'array', cellStyles: true });
    const ws = wb.Sheets[sheetName];
    if (!ws) { showToast('模板Sheet不存在', 'error'); return; }

    // 清除第2行起的所有数据（保留第1行表头）
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = 1; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            delete ws[addr];
        }
    }

    // 写入新数据行
    const dateSerial = getPeriodLastDateSerial(currentPeriod);
    dataRows.forEach((row, ri) => {
        row.forEach((val, ci) => {
            const addr = XLSX.utils.encode_cell({ r: ri + 1, c: ci });
            const cell = { v: val };
            if (ci === 0) { cell.t = 'n'; cell.z = 'yyyy/mm/dd'; }
            else if (typeof val === 'number') { cell.t = 'n'; }
            else { cell.t = 's'; }
            ws[addr] = cell;
        });
    });

    // 更新 range
    ws['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: dataRows.length, c: range.e.c }
    });

    XLSX.writeFile(wb, filename, { bookType: 'xls' });
    showToast('下载完成！', 'success');
}

function copyLiaoliaoData() {
    if (!reportCache) return;
    const startSout = document.getElementById('soutInput').value.trim() || 'SOUT000000';
    const rows = buildLiaoliaoRows(reportCache.reports, currentPeriod, startSout);
    if (rows.length === 0) { showToast('无领料数据', 'error'); return; }
    const date = getPeriodLastDay(currentPeriod);
    const tsv = rows.map(r => [
        date, r.领料部门, r.编号, r.领料, r.发料,
        r.领料类型, r.物料代码, r.是否返工,
        r.单位.toUpperCase(), r.实发数量, r.发料仓库, r.仓位
    ].join('\t')).join('\n');
    navigator.clipboard.writeText(tsv).then(() => {
        showToast('已复制！请在生产领料单.xls 中选中 A2 单元格，粘贴→只保留值', 'success');
    }).catch(() => showToast('复制失败，请检查浏览器权限', 'error'));
}

function copyRukkuData() {
    if (!reportCache) return;
    const startCin = document.getElementById('cinInput').value.trim() || 'CIN000000';
    const rows = buildRukkuRows(reportCache.reports, currentPeriod, startCin);
    if (rows.length === 0) { showToast('无入库数据', 'error'); return; }
    const date = getPeriodLastDay(currentPeriod);
    const tsv = rows.map(r => [
        date, r.交货单位, r.编号, r.验收, r.保管,
        r.物料编码, r.单位.toUpperCase(), r.实收数量, r.收货仓库, r.仓位
    ].join('\t')).join('\n');
    navigator.clipboard.writeText(tsv).then(() => {
        showToast('已复制！请在产品入库.xls 中选中 A2 单元格，粘贴→只保留值', 'success');
    }).catch(() => showToast('复制失败，请检查浏览器权限', 'error'));
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

    // 渲染产品入库单
    document.getElementById('rukkuCard').style.display = 'block';
    renderRukkuTable();

    // 渲染生产领料单
    document.getElementById('liaoliaoCard').style.display = 'block';
    renderLiaoliaoTable();

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
            fuzzyMatch(r.material_name, searchTerm) ||
            r.material_code.toLowerCase().includes(searchTerm.toLowerCase())
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

// ---- 侧边栏折叠 ----
window._toggleSidebar = function() {
    const sidebar = document.getElementById('sidebar');
    const isCollapsed = sidebar.classList.toggle('collapsed');
    localStorage.setItem('sidebar_collapsed', isCollapsed ? '1' : '0');
};

// ---- 主题切换 ----
window._toggleTheme = function() {
    const isDark = document.documentElement.classList.toggle('dark');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
};

// ---- 恢复侧边栏和主题状态 ----
(function() {
    if (localStorage.getItem('sidebar_collapsed') === '1') {
        document.getElementById('sidebar').classList.add('collapsed');
    }
    if (window.innerWidth <= 640) {
        document.getElementById('sidebar').classList.add('collapsed');
    }
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'dark') {
        document.documentElement.classList.add('dark');
    }
})();
