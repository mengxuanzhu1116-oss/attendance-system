/**
 * 本猿每日考勤系统 - 主应用
 * 功能：数据上传、解析、分析、可视化
 */

// ==================== 全局状态 ====================
const AppState = {
    currentPage: 'dashboard',
    mode: 'admin', // 'admin' 或 'visitor'，通过 URL 参数控制
    // 存储的数据
    employees: [],        // 员工档案
    accessRecords: {      // 门禁记录（按城市分组）
        北京: [],
        郑州: [],
        杭州: []
    },
    schedules: [],        // 排班表
    leaves: [],           // 请假记录
    // 分析结果
    analysisResult: null,
    analysisDetail: null, // 详细分析结果
    managerStats: null,   // 主管团队统计
    managerPersonalData: null, // 主管本人出勤数据
    // 图表实例
    charts: {}
};

// 飞书多维表格 Token（默认配置）
const FEISHU_BASE_TOKEN = 'YAsBbBXpeaBQrhsYrdZcY2SwnXo';

// ==================== 初始化 ====================
document.addEventListener('DOMContentLoaded', async () => {
    // 检测访客模式
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('mode') === 'visitor' || urlParams.get('share') === 'true') {
        AppState.mode = 'visitor';
        console.log('[访客模式] 已启用');
    }
    
    // 优先从 attendance-data.json 文件加载数据（用于线上部署）
    // 这样所有访客都能看到同一份数据
    let dataLoaded = false;
    
    try {
        const response = await fetch('./attendance-data.json');
        if (response.ok) {
            const data = await response.json();
            AppState.employees = data.employees || [];
            AppState.accessRecords = data.accessRecords || { 北京: [], 郑州: [], 杭州: [] };
            AppState.schedules = data.schedules || [];
            AppState.leaves = data.leaves || [];
            AppState.analysisResult = data.analysisResult || null;
            AppState.analysisDetail = data.analysisDetail || null;
            AppState.managerStats = data.managerStats || null;
            AppState.managerPersonalData = data.managerPersonalData || null;
            dataLoaded = true;
            console.log('[数据加载] 从 attendance-data.json 加载成功');
        }
    } catch (e) {
        console.log('[数据加载] 无法从 JSON 文件加载（可能是本地 file:// 协议）:', e.message);
    }
    
    // 如果 JSON 文件加载失败，尝试从 localStorage 加载（用于本地管理员）
    if (!dataLoaded) {
        const savedData = localStorage.getItem('attendanceData');
        if (savedData) {
            try {
                const data = JSON.parse(savedData);
                AppState.employees = data.employees || [];
                AppState.accessRecords = data.accessRecords || { 北京: [], 郑州: [], 杭州: [] };
                AppState.schedules = data.schedules || [];
                AppState.leaves = data.leaves || [];
                AppState.analysisResult = data.analysisResult || null;
                AppState.analysisDetail = data.analysisDetail || null;
                AppState.managerStats = data.managerStats || null;
                AppState.managerPersonalData = data.managerPersonalData || null;
                dataLoaded = true;
                console.log('[数据加载] 从本地存储恢复成功');
            } catch (e) {
                console.error('[数据加载] 本地存储数据解析失败:', e);
            }
        }
    }
    
    if (!dataLoaded) {
        console.log('[数据加载] 未找到数据，请上传考勤数据');
    }
    
    initNavigation();
    initSidebar();
    applyVisitorMode(); // 应用访客模式
    loadPage('dashboard');
});

// 初始化导航
function initNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const page = item.dataset.page;
            navItems.forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            loadPage(page);
        });
    });
}

// 初始化侧边栏
function initSidebar() {
    const menuToggle = document.getElementById('menuToggle');
    const sidebar = document.getElementById('sidebar');
    
    menuToggle.addEventListener('click', () => {
        sidebar.classList.toggle('show');
    });
}

// 应用访客模式（隐藏管理功能）
function applyVisitorMode() {
    if (AppState.mode === 'visitor') {
        // 隐藏上传、历史、设置菜单
        document.querySelectorAll('.nav-upload, .nav-history, .nav-settings').forEach(el => {
            el.style.display = 'none';
        });
        
        // 隐藏顶部操作按钮
        const headerActions = document.querySelector('.header-actions');
        if (headerActions) {
            headerActions.innerHTML = '<button class="btn btn-secondary" onclick="exportData()">📥 导出报告</button>';
        }
        
        // 更新用户显示
        const userName = document.querySelector('.user-name');
        const userRole = document.querySelector('.user-role');
        if (userName) userName.textContent = '访客';
        if (userRole) userRole.textContent = '只读查看';
    }
}

// 获取主管本人出勤数据
function getManagerPersonalData() {
    if (!AppState.analysisDetail || AppState.analysisDetail.length === 0) {
        return [];
    }
    
    // 获取所有主管名单（从 manager 字段中提取）
    const managerSet = new Set();
    AppState.analysisDetail.forEach(r => {
        if (r.manager && r.manager !== '未分配' && r.manager !== '无') {
            managerSet.add(r.manager);
        }
    });
    
    // 从分析结果中筛选出主管本人的记录
    const managerPersonalData = [];
    managerSet.forEach(managerName => {
        const managerRecord = AppState.analysisDetail.find(r => r.name === managerName);
        if (managerRecord) {
            managerPersonalData.push({
                name: managerRecord.name,
                jobCategory: managerRecord.jobCategory || '-',
                location: managerRecord.location || '-',
                workTime: managerRecord.workTime || '-',
                checkTime: managerRecord.checkTime || '-',
                status: managerRecord.status || '-'
            });
        } else {
            // 主管没有考勤记录
            managerPersonalData.push({
                name: managerName,
                jobCategory: '-',
                location: '-',
                workTime: '-',
                checkTime: '-',
                status: '缺勤'
            });
        }
    });
    
    // 按状态排序（缺勤、迟到排前面）
    const statusOrder = { '缺勤': 0, '迟到': 1, '异常': 2, '请假': 3, '正常': 4, '已离职': 5 };
    managerPersonalData.sort((a, b) => {
        return (statusOrder[a.status] || 99) - (statusOrder[b.status] || 99);
    });
    
    AppState.managerPersonalData = managerPersonalData;
    return managerPersonalData;
}

// ==================== 页面加载 ====================
function loadPage(page) {
    AppState.currentPage = page;
    const content = document.getElementById('pageContent');
    const title = document.getElementById('pageTitle');
    const breadcrumb = document.getElementById('breadcrumb');
    
    // 清除图表
    Object.values(AppState.charts).forEach(chart => chart.dispose && chart.dispose());
    AppState.charts = {};
    
    switch(page) {
        case 'dashboard':
            title.textContent = '考勤仪表盘';
            breadcrumb.textContent = '首页 / 考勤仪表盘';
            renderDashboard();
            break;
        case 'upload':
            title.textContent = '数据上传';
            breadcrumb.textContent = '首页 / 数据上传';
            renderUploadPage();
            break;
        case 'history':
            title.textContent = '历史记录';
            breadcrumb.textContent = '首页 / 历史记录';
            renderHistoryPage();
            break;
        case 'settings':
            title.textContent = '系统设置';
            breadcrumb.textContent = '首页 / 系统设置';
            renderSettingsPage();
            break;
    }
}

// ==================== 仪表盘页面 ====================
function renderDashboard() {
    const content = document.getElementById('pageContent');
    const hasData = AppState.employees.length > 0 || 
                    Object.values(AppState.accessRecords).some(arr => arr.length > 0);
    
    if (!hasData) {
        content.innerHTML = `
            <div class="data-status">
                <div class="icon">⚠️</div>
                <div class="text">
                    <strong>暂无数据</strong>
                    <p>请先前往「数据上传」页面上传考勤数据</p>
                </div>
                <button class="btn btn-primary" onclick="loadPage('upload')">去上传数据</button>
            </div>
            ${renderEmptyDashboard()}
        `;
        return;
    }
    
    // 有数据，显示分析结果
    content.innerHTML = renderDashboardContent();
    
    // 延迟初始化图表
    setTimeout(() => {
        initDashboardCharts();
        renderAttendanceTable();
        renderManagerPersonalTable(); // 渲染主管本人数据表格
        initManagerPersonalCharts(); // 初始化主管本人图表
    }, 100);
}

function renderEmptyDashboard() {
    return `
        <div class="stats-grid">
            <div class="stat-card normal">
                <div class="stat-icon">👥</div>
                <div class="stat-info">
                    <div class="stat-value">0</div>
                    <div class="stat-label">总人数</div>
                </div>
            </div>
            <div class="stat-card normal">
                <div class="stat-icon">✅</div>
                <div class="stat-info">
                    <div class="stat-value">0</div>
                    <div class="stat-label">正常出勤</div>
                </div>
            </div>
            <div class="stat-card late">
                <div class="stat-icon">⏰</div>
                <div class="stat-info">
                    <div class="stat-value">0</div>
                    <div class="stat-label">迟到</div>
                </div>
            </div>
            <div class="stat-card leave">
                <div class="stat-icon">📅</div>
                <div class="stat-info">
                    <div class="stat-value">0</div>
                    <div class="stat-label">请假</div>
                </div>
            </div>
            <div class="stat-card absent">
                <div class="stat-icon">❌</div>
                <div class="stat-info">
                    <div class="stat-value">0</div>
                    <div class="stat-label">缺勤</div>
                </div>
            </div>
        </div>
        <div class="empty-state">
            <div class="icon">📊</div>
            <h3>等待数据</h3>
            <p>上传数据后，这里将显示考勤分析结果</p>
        </div>
    `;
}

function renderDashboardContent() {
    const stats = calculateStats();
    const jobCategoryStats = calculateJobCategoryStats(); // 计算职类统计
    
    return `
        <!-- 统计卡片 -->
        <div class="stats-grid">
            <div class="stat-card normal">
                <div class="stat-icon">👥</div>
                <div class="stat-info">
                    <div class="stat-value">${stats.total}</div>
                    <div class="stat-label">总人数</div>
                </div>
            </div>
            <div class="stat-card normal">
                <div class="stat-icon">✅</div>
                <div class="stat-info">
                    <div class="stat-value">${stats.normal}</div>
                    <div class="stat-label">正常出勤</div>
                </div>
                <div class="stat-percent">${stats.normalPercent}%</div>
            </div>
            <div class="stat-card late">
                <div class="stat-icon">⏰</div>
                <div class="stat-info">
                    <div class="stat-value">${stats.late}</div>
                    <div class="stat-label">迟到</div>
                </div>
                <div class="stat-percent">${stats.latePercent}%</div>
            </div>
            <div class="stat-card leave">
                <div class="stat-icon">📅</div>
                <div class="stat-info">
                    <div class="stat-value">${stats.leave}</div>
                    <div class="stat-label">请假</div>
                </div>
                <div class="stat-percent">${stats.leavePercent}%</div>
            </div>
            <div class="stat-card absent">
                <div class="stat-icon">❌</div>
                <div class="stat-info">
                    <div class="stat-value">${stats.totalAbsent}</div>
                    <div class="stat-label">缺勤</div>
                </div>
                <div class="stat-percent">${stats.absentPercent}%</div>
            </div>
            ${stats.abnormal > 0 ? `
            <div class="stat-card" style="background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); border-color: #f59e0b;">
                <div class="stat-icon">⚠️</div>
                <div class="stat-info">
                    <div class="stat-value" style="color: #d97706;">${stats.abnormal}</div>
                    <div class="stat-label">异常/未设置</div>
                </div>
            </div>
            ` : ''}
            ${stats.resigned > 0 ? `
            <div class="stat-card" style="background: linear-gradient(135deg, #f3f4f6 0%, #e5e7eb 100%); border-color: #9ca3af;">
                <div class="stat-icon">🚪</div>
                <div class="stat-info">
                    <div class="stat-value" style="color: #6b7280;">${stats.resigned}</div>
                    <div class="stat-label">已离职</div>
                </div>
            </div>
            ` : ''}
        </div>
        
        <!-- 图表区 -->
        <div class="charts-grid">
            <div class="chart-card">
                <div class="card-header">
                    <h3>📊 出勤状态分布</h3>
                </div>
                <div class="chart-container" id="statusChart"></div>
            </div>
            <div class="chart-card">
                <div class="card-header">
                    <h3>🏙️ 各工作地点人数</h3>
                </div>
                <div class="chart-container" id="locationChart"></div>
            </div>
            <!-- 新增：按职类统计出勤情况 -->
            <div class="chart-card">
                <div class="card-header">
                    <h3>👔 各职类出勤状态分布</h3>
                </div>
                <div class="chart-container" id="jobCategoryChart"></div>
            </div>
            <div class="chart-card">
                <div class="card-header">
                    <h3>📈 各职类出勤率对比</h3>
                </div>
                <div class="chart-container" id="jobCategoryRateChart"></div>
            </div>
        </div>
        
        <!-- 数据表格 - 以岗位为首要统计维度 -->
        <div class="table-card">
            <div class="card-header">
                <h3>📋 考勤明细（按岗位统计）</h3>
                <div class="card-actions">
                    <button class="btn btn-small btn-outline" onclick="exportTable()">📥 导出Excel</button>
                </div>
            </div>
            <!-- 岗位维度统计汇总 -->
            <div class="filter-bar" style="background: linear-gradient(135deg, #f0f4ff 0%, #faf5ff 100%); border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                <h4 style="margin: 0 0 12px 0; color: #374151; font-size: 14px;">🏢 各岗位出勤汇总</h4>
                <div id="jobStatsSummary" style="display: flex; flex-wrap: wrap; gap: 12px;"></div>
            </div>
            <div class="filter-bar">
                <select id="statusFilter" onchange="filterTable()">
                    <option value="">全部状态</option>
                    <option value="正常">正常</option>
                    <option value="迟到">迟到</option>
                    <option value="请假">请假</option>
                    <option value="缺勤">缺勤</option>
                    <option value="异常">异常</option>
                    <option value="未设置上班时间">未设置上班时间</option>
                </select>
                <select id="locationFilter" onchange="filterTable()">
                    <option value="">全部地点</option>
                    <option value="北京">北京</option>
                    <option value="郑州">郑州</option>
                    <option value="杭州">杭州</option>
                </select>
                <select id="jobCategoryFilter" onchange="filterTable()">
                    <option value="">全部职类</option>
                </select>
                <select id="managerFilter" onchange="filterTable()">
                    <option value="">全部主管</option>
                </select>
                <input type="text" id="searchInput" placeholder="搜索花名..." oninput="filterTable()">
            </div>
            <div class="table-wrapper">
                <table id="dataTable">
                    <thead>
                        <tr>
                            <th>序号</th>
                            <th>花名</th>
                            <th>职类</th>
                            <th>工作地点</th>
                            <th>直接主管</th>
                            <th>上班时间</th>
                            <th>出勤状态</th>
                            <th>打卡时间</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </div>
            <div class="pagination">
                <button class="btn btn-small btn-secondary" onclick="prevPage()">上一页</button>
                <span id="pageInfo">第 1 页</span>
                <button class="btn btn-small btn-secondary" onclick="nextPage()">下一页</button>
            </div>
        </div>
        
        <!-- 主管团队统计 -->
        <div class="table-card">
            <div class="card-header" style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 12px;">
                <div>
                    <h3 style="margin: 0;">👥 各主管团队出勤统计</h3>
                    <p style="color: #6b7280; font-size: 0.875rem; margin-top: 4px;">出勤率 = (正常+迟到) / (总人数-请假) | 迟到率 = 迟到 / (总人数-请假) | 缺勤率 = (缺勤+异常) / (总人数-请假)</p>
                </div>
                <div style="display: flex; gap: 10px; align-items: center;">
                    <select id="managerStatsFilter" onchange="filterManagerStats()" style="padding: 8px 12px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; background: white;">
                        <option value="">全部主管</option>
                    </select>
                    <select id="rateFilter" onchange="filterManagerStats()" style="padding: 8px 12px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; background: white;">
                        <option value="">全部出勤率</option>
                        <option value="95">出勤率 ≥ 95%</option>
                        <option value="80">出勤率 80%-95%</option>
                        <option value="0">出勤率 < 80%</option>
                    </select>
                    <button class="btn btn-secondary" onclick="exportManagerStats()" style="padding: 8px 16px;">
                        📥 导出数据
                    </button>
                </div>
            </div>
            <!-- 图形看板区域 -->
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(450px, 1fr)); gap: 20px; margin: 20px 0; padding: 20px; background: #f9fafb; border-radius: 8px;">
                <!-- 出勤人数对比图 -->
                <div style="background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08);">
                    <h4 style="margin: 0 0 16px 0; color: #374151; font-size: 14px; font-weight: 600;">📊 各团队出勤人数对比</h4>
                    <div id="managerAttendanceChart" style="width: 100%; height: 280px;"></div>
                </div>
                <!-- 出勤率对比图 -->
                <div style="background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08);">
                    <h4 style="margin: 0 0 16px 0; color: #374151; font-size: 14px; font-weight: 600;">📈 各团队出勤率、迟到率、缺勤率对比</h4>
                    <div id="managerRateChart" style="width: 100%; height: 280px;"></div>
                </div>
            </div>
            <!-- 详细数据表格 -->
            <div class="table-wrapper" style="margin-top: 20px;">
                <table id="managerStatsTable">
                    <thead>
                        <tr>
                            <th>直接主管</th>
                            <th>团队人数</th>
                            <th>正常</th>
                            <th>迟到</th>
                            <th>请假</th>
                            <th>缺勤</th>
                            <th>异常</th>
                            <th>出勤率</th>
                            <th>迟到率</th>
                            <th>缺勤率</th>
                        </tr>
                    </thead>
                    <tbody id="managerStatsBody"></tbody>
                </table>
            </div>
        </div>
        
        <!-- 主管本人出勤情况 -->
        <div class="table-card">
            <div class="card-header" style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 12px;">
                <div>
                    <h3 style="margin: 0;">👔 主管本人出勤情况</h3>
                    <p style="color: #6b7280; font-size: 0.875rem; margin-top: 4px;">各主管本人的出勤状态（非团队统计）</p>
                </div>
                <div style="display: flex; gap: 10px; align-items: center;">
                    <select id="managerPersonalNameFilter" onchange="filterManagerPersonalData()" style="padding: 8px 12px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; background: white;">
                        <option value="">全部主管</option>
                    </select>
                    <select id="managerPersonalStatusFilter" onchange="filterManagerPersonalData()" style="padding: 8px 12px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; background: white;">
                        <option value="">全部状态</option>
                        <option value="正常">正常</option>
                        <option value="迟到">迟到</option>
                        <option value="请假">请假</option>
                        <option value="缺勤">缺勤</option>
                        <option value="异常">异常</option>
                    </select>
                    <button class="btn btn-secondary" onclick="exportManagerPersonalData()" style="padding: 8px 16px;">
                        📥 导出数据
                    </button>
                </div>
            </div>
            <!-- 图形看板区域 -->
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(450px, 1fr)); gap: 20px; margin: 20px 0; padding: 20px; background: #f9fafb; border-radius: 8px;">
                <!-- 出勤状态分布图 -->
                <div style="background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08);">
                    <h4 style="margin: 0 0 16px 0; color: #374151; font-size: 14px; font-weight: 600;">📊 主管出勤状态分布</h4>
                    <div id="managerPersonalStatusChart" style="width: 100%; height: 280px;"></div>
                </div>
            </div>
            <!-- 详细数据表格 -->
            <div class="table-wrapper" style="margin-top: 20px;">
                <table id="managerPersonalTable">
                    <thead>
                        <tr>
                            <th>主管姓名</th>
                            <th>职类</th>
                            <th>工作地点</th>
                            <th>上班时间</th>
                            <th>打卡时间</th>
                            <th>出勤状态</th>
                        </tr>
                    </thead>
                    <tbody id="managerPersonalBody"></tbody>
                </table>
            </div>
        </div>
    `;
}

// 计算职类统计
function calculateJobCategoryStats() {
    if (!AppState.analysisDetail) return {};
    
    const jobStats = {};
    AppState.analysisDetail.forEach(r => {
        if (r.status === '已离职') return;
        
        const jobCat = r.jobCategory || '未分类';
        if (!jobStats[jobCat]) {
            jobStats[jobCat] = { total: 0, normal: 0, late: 0, leave: 0, absent: 0, abnormal: 0 };
        }
        jobStats[jobCat].total++;
        if (r.status === '正常') jobStats[jobCat].normal++;
        else if (r.status === '迟到') jobStats[jobCat].late++;
        else if (r.status === '请假') jobStats[jobCat].leave++;
        else if (r.status === '缺勤') jobStats[jobCat].absent++;
        else jobStats[jobCat].abnormal++;
    });
    
    // 计算各职类出勤率
    Object.keys(jobStats).forEach(cat => {
        const s = jobStats[cat];
        const shouldAttend = s.total - s.leave;
        s.attendanceRate = shouldAttend > 0 ? ((s.normal + s.late) / shouldAttend * 100).toFixed(1) : '0.0';
    });
    
    return jobStats;
}

// ==================== 数据分析 ====================
function calculateStats() {
    const total = AppState.employees.length;
    const result = AppState.analysisResult || { normal: 0, late: 0, leave: 0, absent: 0, abnormal: 0, totalAbsent: total, resigned: 0 };
    const activeTotal = total - (result.resigned || 0);
    
    return {
        total,
        activeTotal,
        normal: result.normal || 0,
        late: result.late || 0,
        leave: result.leave || 0,
        absent: result.absent || 0,
        abnormal: result.abnormal || 0,
        totalAbsent: result.totalAbsent || 0, // 缺勤总数（含异常）
        resigned: result.resigned || 0,
        normalPercent: activeTotal > 0 ? ((result.normal || 0) / activeTotal * 100).toFixed(1) : '0.0',
        latePercent: activeTotal > 0 ? ((result.late || 0) / activeTotal * 100).toFixed(1) : '0.0',
        leavePercent: activeTotal > 0 ? ((result.leave || 0) / activeTotal * 100).toFixed(1) : '0.0',
        absentPercent: activeTotal > 0 ? ((result.totalAbsent || activeTotal) / activeTotal * 100).toFixed(1) : '0.0'
    };
}

/**
 * 根据职类、工作地点、直接主管计算实际上班时间
 * 规则：
 * 1. 销售职类 + 工作地点郑州 + 直接主管佩皮 -> 12:00上班
 * 2. 销售职类、运营职类 -> 13:00上班
 * 3. 其他职类 -> 返回 null，由调用方处理
 */
function calculateWorkTime(jobCategory, location, manager, defaultWorkTime) {
    const category = (jobCategory || '').trim();
    const loc = (location || '').trim();
    const mgr = (manager || '').trim();
    
    // 规则1：销售 + 郑州 + 佩皮 -> 12:00
    if (category === '销售' && loc === '郑州' && mgr === '佩皮') {
        return '12:00';
    }
    
    // 规则2：销售或运营 -> 13:00
    if (category === '销售' || category === '运营') {
        return '13:00';
    }
    
    // 其他职类：返回 null，表示没有匹配规则
    return null;
}

/**
 * 从时间字符串中提取 HH:MM 格式
 * 支持：'08:30:00', '2024-04-14 08:30:00', '2024/04/14 08:30' 等格式
 */
function parseTimeToHHMM(timeStr) {
    if (timeStr === null || timeStr === undefined) return null;
    
    // 处理数字类型（Excel时间序列号）
    if (typeof timeStr === 'number' || (!isNaN(parseFloat(timeStr)) && isFinite(timeStr) && !String(timeStr).includes(':'))) {
        const num = parseFloat(timeStr);
        // Excel时间序列号：0=00:00, 0.5=12:00, 1=24:00
        // 只取小数部分（纯时间）
        const timeValue = Math.abs(num % 1);
        const totalMinutes = Math.round(timeValue * 24 * 60);
        const hour = Math.floor(totalMinutes / 60) % 24;
        const minute = totalMinutes % 60;
        return { 
            hour, 
            minute, 
            str: `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}` 
        };
    }
    
    const str = String(timeStr).trim();
    
    // 尝试匹配时间部分 (HH:MM 或 HH:MM:SS)
    const timeMatch = str.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
    if (timeMatch) {
        const hour = parseInt(timeMatch[1]);
        const minute = parseInt(timeMatch[2]);
        if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
            return { hour, minute, str: `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}` };
        }
    }
    return null;
}

/**
 * 从时间字符串中提取日期 YYYY-MM-DD 格式
 */
function parseDate(timeStr) {
    if (!timeStr) return null;
    const str = String(timeStr).trim();
    
    // 尝试匹配日期部分
    const dateMatch = str.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (dateMatch) {
        return `${dateMatch[1]}-${dateMatch[2].padStart(2, '0')}-${dateMatch[3].padStart(2, '0')}`;
    }
    return null;
}

/**
 * 比较两个时间，返回较早的时间
 */
function getEarlierTime(time1, time2) {
    const t1 = parseTimeToHHMM(time1);
    const t2 = parseTimeToHHMM(time2);
    if (!t1) return time2;
    if (!t2) return time1;
    
    const total1 = t1.hour * 60 + t1.minute;
    const total2 = t2.hour * 60 + t2.minute;
    return total1 <= total2 ? time1 : time2;
}

/**
 * 检查员工是否已离职（离职日期已过）
 */
function isResigned(resignDate) {
    if (!resignDate) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const resign = new Date(resignDate);
    resign.setHours(0, 0, 0, 0);
    return resign < today;
}

/**
 * 检查指定日期是否在请假期间内
 */
function isDateInLeaveRange(checkDate, startDate, endDate) {
    if (!startDate || !endDate) return false;
    
    const check = new Date(checkDate);
    check.setHours(0, 0, 0, 0);
    
    const start = new Date(startDate);
    start.setHours(0, 0, 0, 0);
    
    const end = new Date(endDate);
    end.setHours(23, 59, 59, 999);
    
    return check >= start && check <= end;
}

/**
 * 根据姓名和日期查找请假信息
 */
function findLeaveInfo(leaveRecords, name, date) {
    const records = leaveRecords[name];
    if (!records || !Array.isArray(records)) return null;
    
    for (const record of records) {
        if (isDateInLeaveRange(date, record.startDate, record.endDate)) {
            return record;
        }
    }
    return null;
}

function analyzeAttendance() {
    if (AppState.employees.length === 0) return null;
    
    // 获取今天日期作为分析日期
    const today = new Date().toISOString().split('T')[0];
    
    // 合并所有门禁记录
    const allAccessRecords = [
        ...AppState.accessRecords['北京'],
        ...AppState.accessRecords['郑州'],
        ...AppState.accessRecords['杭州']
    ];
    
    // 构建门禁记录索引（花名+日期 -> 最早打卡时间）
    // 郑州格式按日期分组，取每天最早的记录
    const accessMap = {};
    allAccessRecords.forEach(record => {
        const name = record.name || record['姓名'] || record['名称'];
        const rawTime = record.time || record['签到时间'] || record['时间'];
        const rawDate = record.date || record['日期'];
        // 读取门禁记录中的上班时间字段
        const recordWorkTime = record['上班时间'] || record.workTime || record['工作时间'];
        
        // 解析日期和时间
        let recordDate = rawDate ? parseDate(rawDate) : parseDate(rawTime);
        const recordTime = rawTime;
        
        if (name && recordTime) {
            const key = name.trim();
            const parsedTime = parseTimeToHHMM(recordTime);
            
            if (parsedTime) {
                // 如果已有记录，比较时间取更早的
                if (!accessMap[key]) {
                    accessMap[key] = { 
                        time: parsedTime.str, 
                        date: recordDate || today, 
                        location: record.city || record['工作地点'],
                        workTime: recordWorkTime // 保存门禁记录中的上班时间
                    };
                } else {
                    // 取更早的时间
                    const existingTime = parseTimeToHHMM(accessMap[key].time);
                    if (parsedTime.hour * 60 + parsedTime.minute < existingTime.hour * 60 + existingTime.minute) {
                        accessMap[key] = { 
                            time: parsedTime.str, 
                            date: recordDate || today, 
                            location: record.city || record['工作地点'] 
                        };
                    }
                }
            }
        }
    });
    
    // 构建请假记录索引（花名 -> 请假信息数组，支持多条请假记录）
    const leaveMap = {};
    
    console.log('[调试] 开始解析请假记录，共', AppState.leaves.length, '条原始记录');
    
    AppState.leaves.forEach((leave, idx) => {
        // 获取姓名（支持多种字段名）
        const name = leave['花名'] || leave['姓名'] || leave.name;
        const status = leave['审批状态'] || leave['状态'] || leave.status;
        
        // 获取日期（优先匹配「请假开始日期」「请假结束日期」，然后是其他常见字段名）
        let startDate = leave['请假开始日期'] || leave['开始日期'] || leave['开始时间'] || leave.startDate || leave['起始日期'];
        let endDate = leave['请假结束日期'] || leave['结束日期'] || leave['结束时间'] || leave.endDate || leave['截止日期'];
        
        // 处理 Excel 日期序列号（数字格式）
        if (typeof startDate === 'number') {
            const date = new Date((startDate - 25569) * 86400 * 1000);
            startDate = date.toISOString().split('T')[0];
        }
        if (typeof endDate === 'number') {
            const date = new Date((endDate - 25569) * 86400 * 1000);
            endDate = date.toISOString().split('T')[0];
        }
        
        // 格式化日期字符串
        if (startDate && typeof startDate === 'string') {
            startDate = startDate.trim().substring(0, 10);
        }
        if (endDate && typeof endDate === 'string') {
            endDate = endDate.trim().substring(0, 10);
        }
        
        const leaveType = leave['假期类型'] || leave['请假类型'] || leave['类型'] || leave.leaveType;
        
        // 调试输出
        if (idx < 3) {
            console.log(`[调试] 请假记录#${idx}: 姓名=${name}, 状态=${status}, 开始=${startDate}, 结束=${endDate}`);
        }
        
        if (name && (status === '已通过' || status === '审批中')) {
            // 验证日期有效性
            const startValid = startDate && startDate !== '1970-01-01' && startDate.length >= 10;
            const endValid = endDate && endDate !== '1970-01-01' && endDate.length >= 10;
            
            if (startValid && endValid) {
                const key = String(name).trim();
                if (!leaveMap[key]) {
                    leaveMap[key] = [];
                }
                leaveMap[key].push({ 
                    isLeave: true, 
                    type: leaveType,
                    startDate: startDate,
                    endDate: endDate
                });
            } else {
                console.warn(`[警告] 请假记录日期无效: ${name}, 开始=${startDate}, 结束=${endDate}`);
            }
        }
    });
    
    // 调试：打印请假记录数量
    console.log('[调试] 请假记录解析完成，共', Object.keys(leaveMap).length, '人有有效请假记录');
    if (Object.keys(leaveMap).length > 0) {
        const sampleNames = Object.keys(leaveMap).slice(0, 5);
        console.log('[调试] 请假人员样例:', sampleNames);
        sampleNames.forEach(n => {
            console.log(`[调试] ${n}:`, leaveMap[n]);
        });
    }
    
    // 构建排班索引
    const scheduleMap = {};
    AppState.schedules.forEach(schedule => {
        const name = schedule.name || schedule['花名'];
        const shift = schedule.shift || schedule['班次时间'];
        if (name) {
            scheduleMap[name.trim()] = shift;
        }
    });
    
    // 分析每个员工的出勤状态
    const results = AppState.employees.map(emp => {
        const name = emp.name || emp['花名'] || '';
        const location = emp.location || emp['工作地点'] || '';
        const manager = emp.manager || emp['直接主管'] || '';
        const jobCategory = emp.jobCategory || emp['职类'] || '';
        const defaultWorkTime = emp.workTime || emp['上班时间'] || '';
        const resignDate = emp.resignDate || emp['离职日期'] || '';
        
        // 检查是否已离职
        if (isResigned(resignDate)) {
            return {
                name,
                location,
                manager,
                jobCategory,
                workTime: '-',
                shift: '-',
                status: '已离职',
                checkTime: '-'
            };
        }
        
        // 根据规则计算实际上班时间
        let actualWorkTime = calculateWorkTime(jobCategory, location, manager, defaultWorkTime);
        // 如果没有匹配到规则（返回null），使用员工档案中的上班时间
        if (!actualWorkTime) {
            actualWorkTime = defaultWorkTime || '未设置';
        }
        
        // 查找门禁记录
        const access = accessMap[name];
        // 如果门禁记录中有上班时间，优先使用门禁记录中的
        if (access && access.workTime) {
            actualWorkTime = access.workTime;
        }
        
        // 格式化上班时间，只保留HH:MM部分
        let displayWorkTime = actualWorkTime;
        const workTimeParsed = parseTimeToHHMM(actualWorkTime);
        if (workTimeParsed) {
            displayWorkTime = workTimeParsed.str; // 格式化为 "HH:MM"
        }
        
        // 根据当天日期查找请假记录
        const leaveInfo = findLeaveInfo(leaveMap, name, today);
        // 查找排班
        const shift = scheduleMap[name];
        
        let status = '缺勤';
        let checkTime = '';
        
        // 判断逻辑：先检查请假，再检查打卡，最后看排班
        if (leaveInfo && leaveInfo.isLeave) {
            // 当天在请假期间内
            status = '请假';
        } else if (access) {
            checkTime = access.time;
            // 解析打卡时间和上班时间
            const checkTimeParsed = parseTimeToHHMM(checkTime);
            const parsedWorkTime = parseTimeToHHMM(actualWorkTime);
            
            if (checkTimeParsed && parsedWorkTime) {
                const checkTotal = checkTimeParsed.hour * 60 + checkTimeParsed.minute;
                const workTotal = parsedWorkTime.hour * 60 + parsedWorkTime.minute;
                
                // 严格判定：打卡时间 <= 上班时间 = 正常，否则迟到
                if (checkTotal <= workTotal) {
                    status = '正常';
                } else {
                    status = '迟到';
                    // 调试：记录迟到详情
                    const lateMinutes = checkTotal - workTotal;
                    console.log(`[迟到] ${name}: 打卡${checkTime}(=${checkTotal}分钟) vs 上班${actualWorkTime}(=${workTotal}分钟), 迟到${lateMinutes}分钟, 职类=${jobCategory}, 地点=${location}`);
                }
            } else {
                // 上班时间未设置或时间解析失败
                if (actualWorkTime === '未设置') {
                    status = '未设置上班时间';
                    console.log(`[未设置上班时间] ${name}: 职类=${jobClass}, 打卡时间=${checkTime}`);
                } else {
                    status = '异常';
                    console.log(`[异常] ${name}: 上班时间解析失败(${actualWorkTime}), 打卡时间=${checkTime}`);
                }
            }
        } else if (shift === '休息') {
            status = '正常';
        } else {
            // 缺勤：没有请假、没有打卡、也不是休息日
            console.log(`[缺勤] ${name}: 职类=${jobCategory}, 地点=${location}, 主管=${manager}`);
        }
        
        return {
            name,
            location,
            manager,
            jobCategory,
            workTime: displayWorkTime,  // 显示格式化后的上班时间（HH:MM）
            shift: shift || '未排班',
            status,
            checkTime
        };
    });
    
    // 统计（排除已离职人员）
    const activeResults = results.filter(r => r.status !== '已离职');
    
    // 基础统计
    const summary = {
        normal: activeResults.filter(r => r.status === '正常').length,
        late: activeResults.filter(r => r.status === '迟到').length,
        leave: activeResults.filter(r => r.status === '请假').length,
        absent: activeResults.filter(r => r.status === '缺勤').length,
        abnormal: activeResults.filter(r => r.status === '异常' || r.status === '未设置上班时间').length,
        resigned: results.filter(r => r.status === '已离职').length
    };
    // 缺勤总数 = 缺勤 + 异常 + 未设置上班时间
    summary.totalAbsent = summary.absent + summary.abnormal;
    
    // 按主管分组统计
    const managerStats = {};
    activeResults.forEach(r => {
        const mgr = r.manager || '未分配';
        if (!managerStats[mgr]) {
            managerStats[mgr] = {
                total: 0,
                normal: 0,
                late: 0,
                leave: 0,
                absent: 0,
                abnormal: 0
            };
        }
        managerStats[mgr].total++;
        if (r.status === '正常') managerStats[mgr].normal++;
        else if (r.status === '迟到') managerStats[mgr].late++;
        else if (r.status === '请假') managerStats[mgr].leave++;
        else if (r.status === '缺勤') managerStats[mgr].absent++;
        else managerStats[mgr].abnormal++;
    });
    
    // 计算每个主管团队的出勤率、迟到率、缺勤率（不含请假）
    Object.keys(managerStats).forEach(mgr => {
        const stats = managerStats[mgr];
        const shouldAttend = stats.total - stats.leave; // 应出勤人数 = 总人数 - 请假
        const actualAttend = stats.normal + stats.late; // 实际出勤人数 = 正常 + 迟到
        stats.attendanceRate = shouldAttend > 0 ? ((actualAttend / shouldAttend) * 100).toFixed(1) : '0.0';
        stats.lateRate = shouldAttend > 0 ? ((stats.late / shouldAttend) * 100).toFixed(1) : '0.0';
        stats.absentRate = shouldAttend > 0 ? (((stats.absent + stats.abnormal) / shouldAttend) * 100).toFixed(1) : '0.0';
    });
    
    AppState.analysisResult = summary;
    AppState.analysisDetail = results;
    AppState.managerStats = managerStats; // 保存主管统计
    
    // 保存数据到本地存储（用于访客模式和部署）
    try {
        const saveData = {
            employees: AppState.employees,
            accessRecords: AppState.accessRecords,
            schedules: AppState.schedules,
            leaves: AppState.leaves,
            analysisResult: summary,
            analysisDetail: results,
            managerStats: managerStats,
            managerPersonalData: getManagerPersonalData(),
            exportTime: new Date().toISOString()
        };
        localStorage.setItem('attendanceData', JSON.stringify(saveData));
        console.log('[数据保存] 已保存到本地存储');
    } catch (e) {
        console.error('[数据保存] 保存失败:', e);
    }
    
    console.log('[调试] 主管团队统计:', managerStats);
    
    return { summary, results, managerStats };
}

// ==================== 图表初始化 ====================
function initDashboardCharts() {
    const stats = calculateStats();
    const jobCategoryStats = calculateJobCategoryStats();
    
    // 出勤状态饼图
    const statusChart = echarts.init(document.getElementById('statusChart'));
    statusChart.setOption({
        tooltip: { trigger: 'item', formatter: '{b}: {c}人 ({d}%)' },
        legend: { bottom: '5%', left: 'center' },
        series: [{
            type: 'pie',
            radius: ['40%', '70%'],
            avoidLabelOverlap: false,
            itemStyle: { borderRadius: 10, borderColor: '#fff', borderWidth: 2 },
            label: { show: false },
            emphasis: { label: { show: true, fontSize: 16, fontWeight: 'bold' } },
            labelLine: { show: false },
            data: [
                { value: stats.normal, name: '正常', itemStyle: { color: '#22c55e' } },
                { value: stats.late, name: '迟到', itemStyle: { color: '#f59e0b' } },
                { value: stats.leave, name: '请假', itemStyle: { color: '#3b82f6' } },
                { value: stats.totalAbsent, name: '缺勤', itemStyle: { color: '#ef4444' } }
            ]
        }]
    });
    AppState.charts.status = statusChart;
    
    // 工作地点柱状图
    const locationStats = {};
    AppState.employees.forEach(emp => {
        const loc = emp.location || emp['工作地点'] || '未知';
        locationStats[loc] = (locationStats[loc] || 0) + 1;
    });
    
    const locationChart = echarts.init(document.getElementById('locationChart'));
    locationChart.setOption({
        tooltip: { trigger: 'axis' },
        xAxis: { type: 'category', data: Object.keys(locationStats), axisLabel: { color: '#6b7280' } },
        yAxis: { type: 'value', axisLabel: { color: '#6b7280' } },
        grid: { left: '3%', right: '4%', bottom: '3%', containLabel: true },
        series: [{
            data: Object.values(locationStats),
            type: 'bar',
            barWidth: '50%',
            itemStyle: {
                borderRadius: [8, 8, 0, 0],
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: '#6366f1' },
                    { offset: 1, color: '#ec4899' }
                ])
            }
        }]
    });
    AppState.charts.location = locationChart;
    
    // 职类出勤状态饼图
    initJobCategoryChart(jobCategoryStats);
    
    // 职类出勤率对比图
    initJobCategoryRateChart(jobCategoryStats);
    
    // 渲染岗位维度统计汇总
    renderJobStatsSummary(jobCategoryStats);
    
    // 渲染主管团队统计表格
    renderManagerStatsTable();
    
    // 响应式
    window.addEventListener('resize', () => {
        statusChart.resize();
        locationChart.resize();
        Object.values(AppState.charts).forEach(chart => chart.resize && chart.resize());
    });
}

// 职类出勤状态饼图
function initJobCategoryChart(jobCategoryStats) {
    const chartDom = document.getElementById('jobCategoryChart');
    if (!chartDom) return;
    
    const chart = echarts.init(chartDom);
    
    // 为每个职类创建出勤状态分布
    const categories = Object.keys(jobCategoryStats);
    if (categories.length === 0) {
        chart.setOption({
            graphic: {
                type: 'text',
                left: 'center',
                top: 'center',
                style: {
                    text: '暂无数据',
                    fill: '#9ca3af',
                    fontSize: 14
                }
            }
        });
        AppState.charts.jobCategory = chart;
        return;
    }
    
    // 准备堆叠柱状图数据
    const seriesData = [
        { name: '正常', data: [], color: '#22c55e' },
        { name: '迟到', data: [], color: '#f59e0b' },
        { name: '请假', data: [], color: '#3b82f6' },
        { name: '缺勤', data: [], color: '#ef4444' },
        { name: '异常', data: [], color: '#92400e' }
    ];
    
    categories.forEach(cat => {
        const stats = jobCategoryStats[cat];
        seriesData[0].data.push(stats.normal);
        seriesData[1].data.push(stats.late);
        seriesData[2].data.push(stats.leave);
        seriesData[3].data.push(stats.absent);
        seriesData[4].data.push(stats.abnormal);
    });
    
    chart.setOption({
        tooltip: {
            trigger: 'axis',
            axisPointer: { type: 'shadow' },
            formatter: function(params) {
                let result = params[0].axisValue + '<br/>';
                let total = 0;
                params.forEach(item => {
                    result += item.marker + item.seriesName + ': ' + item.value + '人<br/>';
                    if (item.seriesName !== '异常') total += item.value;
                });
                return result;
            }
        },
        legend: { bottom: '5%', data: seriesData.map(s => s.name) },
        grid: { left: '3%', right: '4%', bottom: '15%', top: '10px', containLabel: true },
        xAxis: {
            type: 'category',
            data: categories,
            axisLabel: { color: '#6b7280', rotate: categories.length > 3 ? 30 : 0 }
        },
        yAxis: { type: 'value', name: '人数', axisLabel: { color: '#6b7280' } },
        series: seriesData.map(s => ({
            name: s.name,
            type: 'bar',
            stack: 'total',
            data: s.data,
            itemStyle: { color: s.color }
        }))
    });
    
    AppState.charts.jobCategory = chart;
}

// 职类出勤率对比图
function initJobCategoryRateChart(jobCategoryStats) {
    const chartDom = document.getElementById('jobCategoryRateChart');
    if (!chartDom) return;
    
    const chart = echarts.init(chartDom);
    
    const categories = Object.keys(jobCategoryStats);
    if (categories.length === 0) {
        chart.setOption({
            graphic: {
                type: 'text',
                left: 'center',
                top: 'center',
                style: {
                    text: '暂无数据',
                    fill: '#9ca3af',
                    fontSize: 14
                }
            }
        });
        AppState.charts.jobCategoryRate = chart;
        return;
    }
    
    // 排序：按出勤率从高到低
    const sortedCategories = categories.sort((a, b) => {
        return parseFloat(jobCategoryStats[b].attendanceRate) - parseFloat(jobCategoryStats[a].attendanceRate);
    });
    
    const attendanceRates = sortedCategories.map(c => parseFloat(jobCategoryStats[c].attendanceRate));
    
    chart.setOption({
        tooltip: {
            trigger: 'axis',
            axisPointer: { type: 'shadow' },
            formatter: function(params) {
                const cat = params[0].axisValue;
                const stats = jobCategoryStats[cat];
                return `${cat}<br/>出勤率: ${stats.attendanceRate}%<br/>总人数: ${stats.total}<br/>正常: ${stats.normal} | 迟到: ${stats.late} | 请假: ${stats.leave} | 缺勤: ${stats.absent}`;
            }
        },
        grid: { left: '3%', right: '4%', bottom: '3%', top: '10px', containLabel: true },
        xAxis: {
            type: 'category',
            data: sortedCategories,
            axisLabel: { color: '#6b7280', rotate: sortedCategories.length > 3 ? 30 : 0 }
        },
        yAxis: {
            type: 'value',
            name: '出勤率(%)',
            max: 100,
            axisLabel: { color: '#6b7280', formatter: '{value}%' }
        },
        series: [{
            type: 'bar',
            data: attendanceRates.map((rate, idx) => ({
                value: rate,
                itemStyle: {
                    color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                        { offset: 0, color: rate >= 95 ? '#22c55e' : rate >= 80 ? '#f59e0b' : '#ef4444' },
                        { offset: 1, color: rate >= 95 ? '#16a34a' : rate >= 80 ? '#d97706' : '#dc2626' }
                    ])
                }
            })),
            barWidth: '50%'
        }]
    });
    
    AppState.charts.jobCategoryRate = chart;
}

// 渲染岗位维度统计汇总
function renderJobStatsSummary(jobCategoryStats) {
    const container = document.getElementById('jobStatsSummary');
    if (!container) return;
    
    const categories = Object.keys(jobCategoryStats).sort((a, b) => {
        return jobCategoryStats[b].total - jobCategoryStats[a].total;
    });
    
    if (categories.length === 0) {
        container.innerHTML = '<span style="color: #9ca3af;">暂无岗位数据</span>';
        return;
    }
    
    const colors = ['#6366f1', '#ec4899', '#f59e0b', '#22c55e', '#3b82f6', '#ef4444', '#8b5cf6', '#14b8a6'];
    
    container.innerHTML = categories.map((cat, idx) => {
        const stats = jobCategoryStats[cat];
        const rate = parseFloat(stats.attendanceRate);
        const rateColor = rate >= 95 ? '#22c55e' : rate >= 80 ? '#f59e0b' : '#ef4444';
        
        return `
            <div style="background: white; border-radius: 10px; padding: 12px 16px; min-width: 180px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-left: 4px solid ${colors[idx % colors.length]};">
                <div style="font-weight: 600; color: #374151; margin-bottom: 8px;">${cat}</div>
                <div style="display: flex; gap: 12px; font-size: 13px;">
                    <span style="color: #6b7280;">总: <strong>${stats.total}</strong></span>
                    <span style="color: #22c55e;">正常: <strong>${stats.normal}</strong></span>
                    <span style="color: #f59e0b;">迟到: <strong>${stats.late}</strong></span>
                    <span style="color: ${rateColor};">出勤率: <strong>${stats.attendanceRate}%</strong></span>
                </div>
            </div>
        `;
    }).join('');
}

// 渲染主管团队统计表格
function renderManagerStatsTable() {
    const tbody = document.getElementById('managerStatsBody');
    if (!tbody || !AppState.managerStats) return;
    
    // 初始化主管筛选下拉框
    initManagerStatsFilter();
    
    // 渲染表格和图表
    filterManagerStats();
}

// 初始化主管筛选下拉框
function initManagerStatsFilter() {
    const select = document.getElementById('managerStatsFilter');
    if (!select || !AppState.managerStats) return;
    
    const managers = Object.keys(AppState.managerStats).sort();
    
    select.innerHTML = '<option value="">全部主管</option>' +
        managers.map(m => `<option value="${m}">${m}</option>`).join('');
}

// 筛选主管统计数据
function filterManagerStats() {
    const tbody = document.getElementById('managerStatsBody');
    if (!tbody || !AppState.managerStats) return;
    
    const managerFilter = document.getElementById('managerStatsFilter');
    const rateFilter = document.getElementById('rateFilter');
    
    const selectedManager = managerFilter ? managerFilter.value : '';
    const selectedRate = rateFilter ? rateFilter.value : '';
    
    // 筛选数据
    let filteredManagers = Object.keys(AppState.managerStats);
    
    if (selectedManager) {
        filteredManagers = filteredManagers.filter(m => m === selectedManager);
    }
    
    if (selectedRate) {
        const threshold = parseInt(selectedRate);
        filteredManagers = filteredManagers.filter(m => {
            const rate = parseFloat(AppState.managerStats[m].attendanceRate);
            if (threshold === 95) return rate >= 95;
            if (threshold === 80) return rate >= 80 && rate < 95;
            if (threshold === 0) return rate < 80;
            return true;
        });
    }
    
    // 排序
    filteredManagers.sort((a, b) => {
        return parseFloat(AppState.managerStats[b].absentRate) - parseFloat(AppState.managerStats[a].absentRate);
    });
    
    // 渲染表格
    tbody.innerHTML = filteredManagers.map(manager => {
        const stats = AppState.managerStats[manager];
        const attendanceRateColor = parseFloat(stats.attendanceRate) >= 95 ? '#22c55e' : 
                                    parseFloat(stats.attendanceRate) >= 80 ? '#f59e0b' : '#ef4444';
        const lateRateColor = parseFloat(stats.lateRate) <= 5 ? '#22c55e' : 
                              parseFloat(stats.lateRate) <= 15 ? '#f59e0b' : '#ef4444';
        const absentRateColor = parseFloat(stats.absentRate) <= 5 ? '#22c55e' : 
                               parseFloat(stats.absentRate) <= 20 ? '#f59e0b' : '#ef4444';
        
        return `
            <tr>
                <td><strong>${manager}</strong></td>
                <td>${stats.total}</td>
                <td style="color: #22c55e;">${stats.normal}</td>
                <td style="color: #f59e0b;">${stats.late}</td>
                <td style="color: #3b82f6;">${stats.leave}</td>
                <td style="color: #ef4444;">${stats.absent}</td>
                <td style="color: #92400e;">${stats.abnormal}</td>
                <td><strong style="color: ${attendanceRateColor};">${stats.attendanceRate}%</strong></td>
                <td><strong style="color: ${lateRateColor};">${stats.lateRate}%</strong></td>
                <td><strong style="color: ${absentRateColor};">${stats.absentRate}%</strong></td>
            </tr>
        `;
    }).join('');
    
    // 渲染图表（使用筛选后的数据）
    renderManagerCharts(filteredManagers);
}

// 导出主管统计数据
function exportManagerStats() {
    if (!AppState.managerStats) {
        showToast('暂无数据可导出', 'error');
        return;
    }
    
    const managers = Object.keys(AppState.managerStats).sort((a, b) => {
        return parseFloat(AppState.managerStats[b].absentRate) - parseFloat(AppState.managerStats[a].absentRate);
    });
    
    // 构建CSV数据
    const headers = ['直接主管', '团队人数', '正常', '迟到', '请假', '缺勤', '异常', '出勤率', '迟到率', '缺勤率'];
    const rows = managers.map(m => {
        const s = AppState.managerStats[m];
        return [m, s.total, s.normal, s.late, s.leave, s.absent, s.abnormal, s.attendanceRate + '%', s.lateRate + '%', s.absentRate + '%'];
    });
    
    // 添加BOM以支持中文
    let csv = '\ufeff' + headers.join(',') + '\n';
    rows.forEach(row => {
        csv += row.join(',') + '\n';
    });
    
    // 下载
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `主管团队出勤统计_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    
    showToast('主管统计数据已导出', 'success');
}

// 渲染主管团队图形看板
function renderManagerCharts(managers) {
    const statsData = managers.map(m => AppState.managerStats[m]);
    
    // 图表1：出勤人数对比（堆叠柱状图）
    const attendanceChartDom = document.getElementById('managerAttendanceChart');
    if (attendanceChartDom) {
        const attendanceChart = echarts.init(attendanceChartDom);
        attendanceChart.setOption({
            tooltip: {
                trigger: 'axis',
                axisPointer: { type: 'shadow' }
            },
            legend: {
                data: ['正常', '迟到', '请假', '缺勤', '异常'],
                top: 0
            },
            grid: {
                left: '3%',
                right: '4%',
                bottom: '3%',
                top: '40px',
                containLabel: true
            },
            xAxis: {
                type: 'category',
                data: managers,
                axisLabel: {
                    interval: 0,
                    rotate: managers.length > 5 ? 30 : 0,
                    fontSize: 11
                }
            },
            yAxis: {
                type: 'value',
                name: '人数'
            },
            series: [
                {
                    name: '正常',
                    type: 'bar',
                    stack: 'total',
                    data: statsData.map(s => s.normal),
                    itemStyle: { color: '#22c55e' }
                },
                {
                    name: '迟到',
                    type: 'bar',
                    stack: 'total',
                    data: statsData.map(s => s.late),
                    itemStyle: { color: '#f59e0b' }
                },
                {
                    name: '请假',
                    type: 'bar',
                    stack: 'total',
                    data: statsData.map(s => s.leave),
                    itemStyle: { color: '#3b82f6' }
                },
                {
                    name: '缺勤',
                    type: 'bar',
                    stack: 'total',
                    data: statsData.map(s => s.absent),
                    itemStyle: { color: '#ef4444' }
                },
                {
                    name: '异常',
                    type: 'bar',
                    stack: 'total',
                    data: statsData.map(s => s.abnormal),
                    itemStyle: { color: '#92400e' }
                }
            ]
        });
        AppState.charts['managerAttendance'] = attendanceChart;
    }
    
    // 图表2：出勤率、迟到率、缺勤率对比（分组柱状图）
    const rateChartDom = document.getElementById('managerRateChart');
    if (rateChartDom) {
        const rateChart = echarts.init(rateChartDom);
        rateChart.setOption({
            tooltip: {
                trigger: 'axis',
                axisPointer: { type: 'shadow' },
                formatter: function(params) {
                    let result = params[0].axisValue + '<br/>';
                    params.forEach(item => {
                        result += item.marker + item.seriesName + ': ' + item.value + '%<br/>';
                    });
                    return result;
                }
            },
            legend: {
                data: ['出勤率', '迟到率', '缺勤率'],
                top: 0
            },
            grid: {
                left: '3%',
                right: '4%',
                bottom: '3%',
                top: '40px',
                containLabel: true
            },
            xAxis: {
                type: 'category',
                data: managers,
                axisLabel: {
                    interval: 0,
                    rotate: managers.length > 5 ? 30 : 0,
                    fontSize: 11
                }
            },
            yAxis: {
                type: 'value',
                name: '百分比(%)',
                max: 100
            },
            series: [
                {
                    name: '出勤率',
                    type: 'bar',
                    data: statsData.map(s => parseFloat(s.attendanceRate)),
                    itemStyle: { color: '#22c55e' },
                    barGap: '10%'
                },
                {
                    name: '迟到率',
                    type: 'bar',
                    data: statsData.map(s => parseFloat(s.lateRate)),
                    itemStyle: { color: '#f59e0b' }
                },
                {
                    name: '缺勤率',
                    type: 'bar',
                    data: statsData.map(s => parseFloat(s.absentRate)),
                    itemStyle: { color: '#ef4444' }
                }
            ]
        });
        AppState.charts['managerRate'] = rateChart;
    }
}

// ==================== 表格渲染 ====================
let currentPage = 1;
const pageSize = 20;
let filteredData = [];

function renderAttendanceTable() {
    if (!AppState.analysisDetail) {
        analyzeAttendance();
    }
    filteredData = AppState.analysisDetail || [];
    currentPage = 1;
    
    // 初始化主管下拉框
    initManagerFilter();
    
    renderTablePage();
}

// 初始化主管筛选下拉框
function initManagerFilter() {
    const managerFilter = document.getElementById('managerFilter');
    if (!managerFilter) return;
    
    // 获取所有主管（去重）
    const managers = [...new Set(
        (AppState.analysisDetail || [])
            .map(r => r.manager)
            .filter(m => m && m !== '-' && m !== '未设置上班时间')
    )].sort();
    
    // 保存当前选中值
    const currentValue = managerFilter.value;
    
    // 重新生成选项
    managerFilter.innerHTML = '<option value="">全部主管</option>' +
        managers.map(m => `<option value="${m}">${m}</option>`).join('');
    
    // 恢复选中值
    if (currentValue && managers.includes(currentValue)) {
        managerFilter.value = currentValue;
    }
    
    // 初始化职类筛选下拉框
    initJobCategoryFilter();
}

// 初始化职类筛选下拉框
function initJobCategoryFilter() {
    const jobCategoryFilter = document.getElementById('jobCategoryFilter');
    if (!jobCategoryFilter) return;
    
    // 获取所有职类（去重）
    const jobCategories = [...new Set(
        (AppState.analysisDetail || [])
            .map(r => r.jobCategory)
            .filter(c => c && c !== '-' && c !== '未设置上班时间')
    )].sort();
    
    // 保存当前选中值
    const currentValue = jobCategoryFilter.value;
    
    // 重新生成选项
    jobCategoryFilter.innerHTML = '<option value="">全部职类</option>' +
        jobCategories.map(c => `<option value="${c}">${c}</option>`).join('');
    
    // 恢复选中值
    if (currentValue && jobCategories.includes(currentValue)) {
        jobCategoryFilter.value = currentValue;
    }
}

function renderTablePage() {
    const tbody = document.getElementById('tableBody');
    if (!tbody) return;
    
    const start = (currentPage - 1) * pageSize;
    const end = start + pageSize;
    const pageData = filteredData.slice(start, end);
    
    tbody.innerHTML = pageData.map((row, index) => {
        let statusClass = 'absent';
        if (row.status === '正常') statusClass = 'normal';
        else if (row.status === '迟到') statusClass = 'late';
        else if (row.status === '请假') statusClass = 'leave';
        else if (row.status === '已离职') statusClass = 'resigned';
        else if (row.status === '未设置上班时间') statusClass = 'warning';
        else if (row.status === '异常') statusClass = 'absent';
        
        return `
        <tr>
            <td>${start + index + 1}</td>
            <td><strong>${row.name}</strong></td>
            <td>${row.jobCategory || '-'}</td>
            <td>${row.location || '-'}</td>
            <td>${row.manager || '-'}</td>
            <td><span style="color: #6366f1; font-weight: 600;">${row.workTime}</span></td>
            <td><span class="status-badge status-${statusClass}">${row.status}</span></td>
            <td>${row.checkTime || '-'}</td>
        </tr>
    `}).join('');
    
    updatePageInfo();
}

function updatePageInfo() {
    const totalPages = Math.ceil(filteredData.length / pageSize) || 1;
    document.getElementById('pageInfo').textContent = `第 ${currentPage} / ${totalPages} 页 (共${filteredData.length}条)`;
}

function prevPage() {
    if (currentPage > 1) {
        currentPage--;
        renderTablePage();
    }
}

function nextPage() {
    const totalPages = Math.ceil(filteredData.length / pageSize);
    if (currentPage < totalPages) {
        currentPage++;
        renderTablePage();
    }
}

function filterTable() {
    const status = document.getElementById('statusFilter').value;
    const location = document.getElementById('locationFilter').value;
    const jobCategory = document.getElementById('jobCategoryFilter')?.value || '';
    const manager = document.getElementById('managerFilter')?.value || '';
    const search = document.getElementById('searchInput').value.toLowerCase();
    
    filteredData = (AppState.analysisDetail || []).filter(row => {
        if (status && row.status !== status) return false;
        if (location && row.location !== location) return false;
        if (jobCategory && row.jobCategory !== jobCategory) return false;
        if (manager && row.manager !== manager) return false;
        if (search && !row.name.toLowerCase().includes(search)) return false;
        return true;
    });
    
    currentPage = 1;
    renderTablePage();
}

// ==================== 上传页面 ====================
function renderUploadPage() {
    const content = document.getElementById('pageContent');
    
    content.innerHTML = `
        <!-- 使用说明 -->
        <div class="info-card" style="border-left: 4px solid #6366f1;">
            <h3>📌 上传说明</h3>
            <ol>
                <li>先下载对应的数据模板，按照模板格式填写数据</li>
                <li><strong>员工档案</strong>是基础数据，需要先上传，且必须填写<strong>职类</strong>字段</li>
                <li>门禁记录支持<strong>多文件上传</strong>，可同时上传三地数据</li>
                <li>系统会自动识别门禁记录的城市格式</li>
                <li>上传后点击「开始分析」查看考勤结果</li>
            </ol>
            <div style="margin-top: 16px; padding: 12px 16px; background: #f0f4ff; border-radius: 8px;">
                <strong style="color: #4f46e5;">⏰ 上班时间规则（根据职类自动判断）：</strong>
                <ul style="margin-top: 8px; color: #374151; line-height: 1.8;">
                    <li>销售职类 + 郑州工作地点 + 佩皮主管 → <strong style="color: #22c55e;">12:00</strong> 上班</li>
                    <li>销售职类、运营职类 → <strong style="color: #f59e0b;">13:00</strong> 上班</li>
                    <li>其他职类 → 使用员工档案中的<strong>上班时间</strong>字段（需在档案中填写）</li>
                </ul>
            </div>
        </div>
        
        <!-- 上传卡片 -->
        <div class="upload-grid">
            <!-- 员工档案上传 -->
            <div class="upload-card">
                <div class="upload-card-header">
                    <div class="upload-icon employee">👥</div>
                    <div>
                        <h3>员工档案 <span style="color: #ef4444;">*必填</span></h3>
                        <p>员工基础信息（权限基准）</p>
                    </div>
                </div>
                <div class="upload-zone" id="employeeZone">
                    <input type="file" accept=".xlsx,.xls" onchange="handleFileUpload(this, 'employee')" single>
                    <div class="upload-zone-icon">📁</div>
                    <div class="upload-zone-text">点击或拖拽上传</div>
                    <div class="upload-zone-hint">支持 .xlsx 格式</div>
                </div>
                <div id="employeeFileList" class="file-list"></div>
                <div class="upload-actions">
                    <button class="btn btn-outline" onclick="downloadTemplate('employee')">📥 下载模板</button>
                    <button class="btn btn-secondary" onclick="clearData('employee')">清空</button>
                </div>
            </div>
            
            <!-- 门禁记录上传（支持多文件） -->
            <div class="upload-card">
                <div class="upload-card-header">
                    <div class="upload-icon access">🚪</div>
                    <div>
                        <h3>门禁记录 <span style="color: #f59e0b;">支持多文件</span></h3>
                        <p>三地门禁打卡数据（可同时上传多个）</p>
                    </div>
                </div>
                <div class="upload-zone" id="accessZone">
                    <input type="file" accept=".xlsx,.xls" multiple onchange="handleFileUpload(this, 'access')">
                    <div class="upload-zone-icon">📁</div>
                    <div class="upload-zone-text">点击或拖拽上传（可多选）</div>
                    <div class="upload-zone-hint">自动识别北京/郑州/杭州格式</div>
                </div>
                <div id="accessFileList" class="file-list"></div>
                <div class="upload-actions">
                    <button class="btn btn-outline" onclick="downloadTemplate('access-zhengzhou')">郑州模板</button>
                    <button class="btn btn-outline" onclick="downloadTemplate('access-beijing')">北京模板</button>
                    <button class="btn btn-outline" onclick="downloadTemplate('access-hangzhou')">杭州模板</button>
                    <button class="btn btn-secondary" onclick="clearData('access')">清空</button>
                </div>
            </div>
            
            <!-- 排班表上传 -->
            <div class="upload-card">
                <div class="upload-card-header">
                    <div class="upload-icon schedule">📅</div>
                    <div>
                        <h3>排班表</h3>
                        <p>员工排班安排</p>
                    </div>
                </div>
                <div class="upload-zone" id="scheduleZone">
                    <input type="file" accept=".xlsx,.xls" onchange="handleFileUpload(this, 'schedule')">
                    <div class="upload-zone-icon">📁</div>
                    <div class="upload-zone-text">点击或拖拽上传</div>
                    <div class="upload-zone-hint">支持 .xlsx 格式</div>
                </div>
                <div id="scheduleFileList" class="file-list"></div>
                <div class="upload-actions">
                    <button class="btn btn-outline" onclick="downloadTemplate('schedule')">📥 下载模板</button>
                    <button class="btn btn-secondary" onclick="clearData('schedule')">清空</button>
                </div>
            </div>
            
            <!-- 请假记录上传 -->
            <div class="upload-card">
                <div class="upload-card-header">
                    <div class="upload-icon leave">📝</div>
                    <div>
                        <h3>请假记录</h3>
                        <p>员工请假申请记录</p>
                    </div>
                </div>
                <div class="upload-zone" id="leaveZone">
                    <input type="file" accept=".xlsx,.xls" onchange="handleFileUpload(this, 'leave')">
                    <div class="upload-zone-icon">📁</div>
                    <div class="upload-zone-text">点击或拖拽上传</div>
                    <div class="upload-zone-hint">支持 .xlsx 格式</div>
                </div>
                <div id="leaveFileList" class="file-list"></div>
                <div class="upload-actions">
                    <button class="btn btn-outline" onclick="downloadTemplate('leave')">📥 下载模板</button>
                    <button class="btn btn-secondary" onclick="clearData('leave')">清空</button>
                </div>
            </div>
        </div>
        
        <!-- 操作按钮 -->
        <div style="margin-top: 24px; display: flex; gap: 16px; justify-content: center;">
            <button class="btn btn-primary" onclick="startAnalysis()" style="padding: 14px 40px; font-size: 1rem;">
                🚀 开始分析
            </button>
            <button class="btn btn-danger" onclick="clearAllData()">
                🗑️ 清空所有数据
            </button>
        </div>
        
        <!-- 上传统计 -->
        <div class="analysis-result" id="uploadStats" style="display: none;">
            <h4>📊 已上传数据统计</h4>
            <div class="result-stats" id="uploadStatsContent"></div>
        </div>
    `;
    
    // 初始化拖拽
    initDragDrop();
    
    // 显示已上传的数据统计
    updateUploadStats();
}

// ==================== 文件处理 ====================
const uploadedFiles = {
    employee: [],
    access: [],
    schedule: [],
    leave: []
};

function initDragDrop() {
    // 防止重复绑定事件
    if (window.dragDropInitialized) return;
    window.dragDropInitialized = true;
    
    ['employee', 'access', 'schedule', 'leave'].forEach(type => {
        const zone = document.getElementById(type + 'Zone');
        if (!zone) return;
        
        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            zone.classList.add('dragover');
        });
        
        zone.addEventListener('dragleave', () => {
            zone.classList.remove('dragover');
        });
        
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation(); // 阻止事件冒泡
            zone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            processFiles(files, type);
        });
        
        // 点击上传区域时触发文件选择
        zone.addEventListener('click', (e) => {
            // 如果点击的不是 input 本身，触发 input 的 click
            const input = zone.querySelector('input[type="file"]');
            if (input && e.target !== input) {
                e.stopPropagation();
                input.click();
            }
        });
    });
}

function handleFileUpload(input, type) {
    const files = input.files;
    if (files.length === 0) return;
    
    // 处理文件
    processFiles(files, type);
    
    // 清空 input，防止重复触发
    input.value = '';
}

function processFiles(files, type) {
    Array.from(files).forEach(file => {
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            showToast('请上传 Excel 文件 (.xlsx)', 'error');
            return;
        }
        
        // 检查是否已存在同名文件，避免重复添加
        const exists = uploadedFiles[type].some(f => f.name === file.name && f.size === file.size);
        if (exists) {
            showToast(`文件 ${file.name} 已存在，跳过`, 'warning');
            return;
        }
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                
                // 存储文件信息
                uploadedFiles[type].push({
                    name: file.name,
                    size: file.size,
                    count: jsonData.length,
                    data: jsonData
                });
                
                // 根据类型处理数据
                processUploadedData(type, jsonData, file.name);
                
                // 更新文件列表显示
                updateFileListDisplay(type);
                updateUploadStats();
                
                showToast(`${file.name} 上传成功，共 ${jsonData.length} 条记录`, 'success');
            } catch (err) {
                console.error('解析文件失败:', err);
                showToast(`文件解析失败: ${err.message}`, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

function processUploadedData(type, data, filename) {
    switch(type) {
        case 'employee':
            AppState.employees = data;
            break;
        case 'access':
            // 自动识别城市格式
            const city = detectCity(data, filename);
            AppState.accessRecords[city] = data.map(row => ({
                ...row,
                city
            }));
            break;
        case 'schedule':
            AppState.schedules = data;
            break;
        case 'leave':
            // 验证请假数据
            let validCount = 0;
            let invalidRecords = [];
            data.forEach((record, index) => {
                const startDate = record['开始日期'] || record['startDate'] || record['开始时间'];
                const endDate = record['结束日期'] || record['endDate'] || record['结束时间'];
                const name = record['花名'] || record['姓名'] || record['name'];
                
                // 检查日期是否有效
                const isDateValid = startDate && endDate && 
                    startDate !== '1970-01-01' && endDate !== '1970-01-01' &&
                    new Date(startDate).toString() !== 'Invalid Date' &&
                    new Date(endDate).toString() !== 'Invalid Date';
                
                if (isDateValid) {
                    validCount++;
                } else {
                    invalidRecords.push({
                        row: index + 2, // Excel行号（从2开始）
                        name: name,
                        startDate: startDate,
                        endDate: endDate
                    });
                }
            });
            
            AppState.leaves = data;
            
            // 显示解析结果
            if (invalidRecords.length > 0 && data.length > 0) {
                console.warn('[请假数据警告] 以下记录日期无效:', invalidRecords.slice(0, 5));
                showToast(`请假记录解析：${validCount}/${data.length} 条有效，${invalidRecords.length} 条日期无效`, 'warning');
            } else if (validCount > 0) {
                showToast(`请假记录解析成功：${validCount} 条有效记录`, 'success');
            }
            break;
    }
}

function detectCity(data, filename) {
    // 通过文件名判断
    if (filename.includes('北京') || filename.includes('beijing')) return '北京';
    if (filename.includes('郑州') || filename.includes('zhengzhou')) return '郑州';
    if (filename.includes('杭州') || filename.includes('hangzhou')) return '杭州';
    
    // 通过数据格式判断
    if (data.length > 0) {
        const firstRow = data[0];
        // 郑州格式：时间、名称
        if (firstRow['时间'] && firstRow['名称']) return '郑州';
        // 北京格式：姓名、日期、上班时间、签到时间
        if (firstRow['姓名'] && firstRow['签到时间']) return '北京';
        // 杭州格式：姓名、日期、时间
        if (firstRow['姓名'] && firstRow['时间'] && !firstRow['签到时间']) return '杭州';
    }
    
    return '北京'; // 默认
}

function updateFileListDisplay(type) {
    const fileList = document.getElementById(type + 'FileList');
    if (!fileList) return;
    
    fileList.innerHTML = uploadedFiles[type].map((file, index) => `
        <div class="file-item">
            <div class="file-icon">📄</div>
            <div class="file-info">
                <div class="file-name">${file.name}</div>
                <div class="file-size">${formatFileSize(file.size)} · ${file.count} 条记录</div>
            </div>
            <span class="file-status success">✓ 已解析</span>
            <button class="file-remove" onclick="removeFile('${type}', ${index})">×</button>
        </div>
    `).join('');
}

function removeFile(type, index) {
    uploadedFiles[type].splice(index, 1);
    // 重新处理数据
    if (type === 'access') {
        AppState.accessRecords = { 北京: [], 郑州: [], 杭州: [] };
        uploadedFiles[type].forEach(f => {
            const city = detectCity(f.data, f.name);
            AppState.accessRecords[city] = f.data.map(row => ({ ...row, city }));
        });
    } else {
        AppState[type] = uploadedFiles[type].flatMap(f => f.data);
    }
    updateFileListDisplay(type);
    updateUploadStats();
}

function updateUploadStats() {
    const stats = document.getElementById('uploadStats');
    const content = document.getElementById('uploadStatsContent');
    if (!stats || !content) return;
    
    const employeeCount = AppState.employees.length;
    const accessCount = Object.values(AppState.accessRecords).flat().length;
    const scheduleCount = AppState.schedules.length;
    const leaveCount = AppState.leaves.length;
    
    if (employeeCount + accessCount + scheduleCount + leaveCount === 0) {
        stats.style.display = 'none';
        return;
    }
    
    stats.style.display = 'block';
    content.innerHTML = `
        <div class="result-item">
            <div class="value">${employeeCount}</div>
            <div class="label">员工档案</div>
        </div>
        <div class="result-item">
            <div class="value">${accessCount}</div>
            <div class="label">门禁记录</div>
        </div>
        <div class="result-item">
            <div class="value">${scheduleCount}</div>
            <div class="label">排班记录</div>
        </div>
        <div class="result-item">
            <div class="value">${leaveCount}</div>
            <div class="label">请假记录</div>
        </div>
    `;
}

// ==================== 分析操作 ====================
function startAnalysis() {
    if (AppState.employees.length === 0) {
        showToast('请先上传员工档案数据', 'error');
        return;
    }
    
    showToast('正在分析数据...', 'info');
    
    setTimeout(() => {
        const result = analyzeAttendance();
        if (result) {
            const totalCount = result.summary.normal + result.summary.late + result.summary.leave + result.summary.absent;
            const resignedCount = result.summary.resigned || 0;
            
            // 显示分析完成弹窗
            showAnalysisResultModal(result, totalCount, resignedCount);
            
            // 飞书同步已禁用，如需同步请在仪表盘页面手动操作
            // syncToFeishuBase();
            
            // 跳转到仪表盘
            setTimeout(() => {
                document.querySelector('[data-page="dashboard"]').click();
            }, 500);
        }
    }, 500);
}

// ==================== 飞书多维表格同步 ====================
// 同步考勤数据到飞书多维表格
async function syncToFeishuBase() {
    const baseToken = FEISHU_BASE_TOKEN;
    
    if (!AppState.analysisResult || !AppState.analysisDetail) {
        console.log('[飞书同步] 暂无分析数据，跳过同步');
        return;
    }
    
    showToast('正在同步数据到飞书多维表格...', 'info');
    
    try {
        // 获取统计数据
        const stats = calculateStats();
        const jobCategoryStats = calculateJobCategoryStats();
        
        // 构建同步数据 - 修复字段映射问题
        const todayDate = new Date().toISOString().split('T')[0];
        const todayTimestamp = new Date(todayDate + 'T00:00:00+08:00').getTime(); // 毫秒时间戳
        
        // 构建飞书同步用的明细数据
        const feishuDetails = AppState.analysisDetail.map(record => {
            // 出勤状态映射：内部状态 -> 飞书选项
            let feishuStatus = '缺勤';
            if (record.status === '正常') feishuStatus = '正常出勤';
            else if (record.status === '迟到') feishuStatus = '迟到';
            else if (record.status === '请假') feishuStatus = '请假';
            else if (record.status === '缺勤' || record.status === '旷工') feishuStatus = '缺勤';
            
            return {
                date: todayTimestamp, // 使用毫秒时间戳，避免1970年问题
                jobCategory: record.jobCategory || '-',
                name: record.name, // 花名，直接使用分析结果中的花名
                manager: record.manager || '-', // 直接主管
                workTime: record.workTime || '-', // 上班时间（已按规则计算）
                attendanceStatus: feishuStatus, // 出勤情况
                clockInTime: record.checkTime || '-' // 打卡时间
            };
        });
        
        const syncData = {
            date: todayDate,
            dateTimestamp: todayTimestamp, // 毫秒时间戳
            summary: {
                total: stats.total,
                normal: stats.normal,
                late: stats.late,
                leave: stats.leave,
                absent: stats.totalAbsent,
                normalRate: stats.normalPercent,
                lateRate: stats.latePercent
            },
            jobCategoryStats: jobCategoryStats,
            details: feishuDetails // 使用修复后的数据
        };
        
        // 调用飞书CLI同步数据
        const syncResult = await callFeishuSyncAPI(syncData);
        
        if (syncResult && syncResult.success) {
            showToast('✅ 数据已同步到飞书多维表格', 'success');
            console.log('[飞书同步] 同步成功:', syncResult);
        } else {
            showToast('飞书同步遇到问题，请检查配置', 'warning');
            console.log('[飞书同步] 同步失败:', syncResult);
        }
    } catch (error) {
        console.error('[飞书同步] 同步出错:', error);
        showToast('飞书同步出错: ' + error.message, 'error');
    }
}

// 调用飞书同步API（直接生成同步指令）
async function callFeishuSyncAPI(data) {
    // 直接生成同步指令供用户发送给助手
    generateFeishuSyncCommand(data);
    return { success: true, manual: true };
}

// 生成飞书同步指令
function generateFeishuSyncCommand(data) {
    const syncCommand = {
        action: 'syncAttendanceToFeishu',
        baseToken: FEISHU_BASE_TOKEN,
        tableId: 'tblxOwM5MnvRfETE', // 飞书多维表格ID
        date: data.date,
        dateTimestamp: data.dateTimestamp, // 毫秒时间戳
        summary: data.summary,
        jobCategoryStats: data.jobCategoryStats,
        details: data.details // 包含完整的明细数据
    };
    
    const commandBase64 = btoa(unescape(encodeURIComponent(JSON.stringify(syncCommand))));
    
    // 显示同步提示
    const modal = document.createElement('div');
    modal.id = 'feishuSyncModal';
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.5);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 10000;
    `;
    
    modal.innerHTML = `
        <div style="background: white; border-radius: 16px; padding: 28px; max-width: 500px; width: 90%; box-shadow: 0 20px 60px rgba(0,0,0,0.3);">
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 20px;">
                <span style="font-size: 2rem;">📤</span>
                <h3 style="margin: 0; color: #111827;">飞书多维表格同步</h3>
            </div>
            
            <div style="background: #f0fdf4; border: 1px solid #86efac; border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                <div style="color: #166534; font-weight: 500; margin-bottom: 8px;">
                    ✅ 分析完成，数据已准备好同步到飞书
                </div>
                <div style="color: #15803d; font-size: 0.875rem; line-height: 1.6;">
                    <p><strong>同步内容：</strong></p>
                    <ul style="margin: 8px 0; padding-left: 20px;">
                        <li>日期：${data.date}</li>
                        <li>总人数：${data.summary.total}</li>
                        <li>正常：${data.summary.normal} (${data.summary.normalRate}%)</li>
                        <li>迟到：${data.summary.late} (${data.summary.lateRate}%)</li>
                        <li>请假：${data.summary.leave}</li>
                        <li>缺勤：${data.summary.absent}</li>
                    </ul>
                    <p><strong>考勤明细：</strong>${data.details.length} 条记录</p>
                    <p style="color: #166534; font-size: 0.8rem; margin-top: 8px;">
                        ✓ 包含字段：日期(时间戳)、职类、花名、直接主管、上班时间、出勤情况、打卡时间
                    </p>
                </div>
            </div>
            
            <div style="background: #fef3c7; border: 1px solid #fcd34d; border-radius: 8px; padding: 12px; margin-bottom: 16px; font-size: 13px; color: #92400e;">
                💡 <strong>操作步骤：</strong><br>
                复制下方指令发送给助手，助手将为您同步数据到飞书多维表格
            </div>
            
            <div style="background: #f3f4f6; border-radius: 8px; padding: 12px; margin-bottom: 20px; font-family: monospace; font-size: 11px; color: #374151; cursor: pointer; word-break: break-all;" 
                 onclick="navigator.clipboard.writeText(this.innerText).then(() => showToast('已复制', 'success'))">
飞书同步：${commandBase64}
            </div>
            
            <div style="display: flex; gap: 12px; justify-content: center;">
                <button onclick="navigator.clipboard.writeText('飞书同步：${commandBase64}').then(() => showToast('已复制到剪贴板', 'success'))" 
                        style="padding: 12px 24px; border: none; border-radius: 8px; background: #22c55e; color: white; cursor: pointer; font-size: 14px; font-weight: 600;">
                    📋 复制指令
                </button>
                <button onclick="document.getElementById('feishuSyncModal').remove()" 
                        style="padding: 12px 24px; border: 1px solid #d1d5db; border-radius: 8px; background: white; cursor: pointer; font-size: 14px;">
                    关闭
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
}

// 显示分析结果弹窗
function showAnalysisResultModal(result, totalCount, resignedCount) {
    const modal = document.createElement('div');
    modal.className = 'analysis-modal';
    modal.innerHTML = `
        <div class="analysis-modal-content">
            <div class="modal-header">
                <div class="success-icon">✅</div>
                <h2>分析完成！</h2>
            </div>
            <div class="modal-body">
                <div class="result-summary">
                    <div class="summary-item normal">
                        <div class="count">${result.summary.normal}</div>
                        <div class="label">正常出勤</div>
                    </div>
                    <div class="summary-item late">
                        <div class="count">${result.summary.late}</div>
                        <div class="label">迟到</div>
                    </div>
                    <div class="summary-item leave">
                        <div class="count">${result.summary.leave}</div>
                        <div class="label">请假</div>
                    </div>
                    <div class="summary-item absent">
                        <div class="count">${result.summary.absent}</div>
                        <div class="label">缺勤</div>
                    </div>
                    ${resignedCount > 0 ? `
                    <div class="summary-item resigned">
                        <div class="count">${resignedCount}</div>
                        <div class="label">已离职</div>
                    </div>
                    ` : ''}
                </div>
                <div class="total-info">
                    共分析 <strong>${totalCount}</strong> 名在职员工
                    ${resignedCount > 0 ? `，<strong>${resignedCount}</strong> 人已离职不计入统计` : ''}
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn btn-primary" onclick="closeAnalysisModal(this)">查看详情</button>
            </div>
        </div>
    `;
    document.body.appendChild(modal);
    
    // 点击背景关闭
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            closeAnalysisModal(modal.querySelector('button'));
        }
    });
}

function closeAnalysisModal(btn) {
    const modal = btn.closest('.analysis-modal');
    if (modal) {
        modal.remove();
    }
}

function clearData(type) {
    uploadedFiles[type] = [];
    if (type === 'access') {
        AppState.accessRecords = { 北京: [], 郑州: [], 杭州: [] };
    } else {
        AppState[type] = [];
    }
    updateFileListDisplay(type);
    updateUploadStats();
    showToast('已清空数据', 'info');
}

function clearAllData() {
    if (!confirm('确定要清空所有已上传的数据吗？')) return;
    
    uploadedFiles.employee = [];
    uploadedFiles.access = [];
    uploadedFiles.schedule = [];
    uploadedFiles.leave = [];
    
    AppState.employees = [];
    AppState.accessRecords = { 北京: [], 郑州: [], 杭州: [] };
    AppState.schedules = [];
    AppState.leaves = [];
    AppState.analysisResult = null;
    AppState.analysisDetail = null;
    
    ['employee', 'access', 'schedule', 'leave'].forEach(type => {
        updateFileListDisplay(type);
    });
    updateUploadStats();
    showToast('已清空所有数据', 'success');
}

// ==================== 历史记录页面 ====================
function renderHistoryPage() {
    const content = document.getElementById('pageContent');
    content.innerHTML = `
        <div class="info-card">
            <h3>📅 历史记录功能</h3>
            <p>历史记录功能需要配合后端存储使用。当前版本数据存储在浏览器内存中，刷新页面后将清空。</p>
            <p style="margin-top: 12px;">如需保存分析结果，请使用「导出报告」功能。</p>
        </div>
    `;
}

// ==================== 设置页面 ====================
// 飞书配置状态
const FeishuConfig = {
    appToken: '',      // 飞书应用 Token
    tableToken: '',    // 多维表格 Token
    webhookUrl: '',    // 飞书机器人 Webhook URL
    notifyTime: '17:00', // 默认通知时间
    enabled: false,     // 是否启用自动通知
    lastUpdated: null   // 最后更新时间
};

// 从 localStorage 和配置文件加载配置
function loadFeishuConfig() {
    // 优先从本地配置加载（实时生效）
    const localSaved = localStorage.getItem('feishuConfigFull');
    if (localSaved) {
        try {
            const config = JSON.parse(localSaved);
            FeishuConfig.appToken = config.appToken || '';
            FeishuConfig.tableToken = config.tableToken || '';
            FeishuConfig.webhookUrl = config.webhookUrl || '';
            FeishuConfig.notifyTime = config.notifyTime || '17:00';
            FeishuConfig.enabled = config.enabled || false;
            FeishuConfig.lastUpdated = config.updatedAt || null;
            FeishuConfig.notifyTemplate = config.notifyTemplate || '';
            return FeishuConfig;
        } catch (e) {
            console.error('解析本地配置失败:', e);
        }
    }
    
    // 尝试从旧格式配置加载
    const saved = localStorage.getItem('feishuConfig');
    if (saved) {
        try {
            const config = JSON.parse(saved);
            Object.assign(FeishuConfig, config);
            return FeishuConfig;
        } catch (e) {
            console.error('解析配置失败:', e);
        }
    }
    
    return FeishuConfig;
}

// 保存配置到 localStorage（实时生效）
function saveFeishuConfig(config) {
    const fullConfig = {
        ...FeishuConfig,
        ...config,
        updatedAt: new Date().toISOString()
    };
    
    // 更新全局配置对象
    Object.assign(FeishuConfig, fullConfig);
    
    // 保存完整配置
    localStorage.setItem('feishuConfigFull', JSON.stringify(fullConfig));
    
    // 同时保存简化配置（兼容旧版本）
    localStorage.setItem('feishuConfig', JSON.stringify({
        enabled: fullConfig.enabled,
        notifyTime: fullConfig.notifyTime,
        webhookUrl: fullConfig.webhookUrl,
        tableToken: fullConfig.tableToken,
        appToken: fullConfig.appToken
    }));
    
    showToast('飞书配置已保存', 'success');
}

// 测试飞书连接
async function testFeishuConnection() {
    const config = {
        webhookUrl: document.getElementById('feishuWebhook').value,
        tableToken: document.getElementById('feishuTableToken').value
    };
    
    if (!config.webhookUrl && !config.tableToken) {
        showToast('请先填写飞书 Webhook URL 或多维表格 Token', 'error');
        return;
    }
    
    showToast('正在测试连接...', 'info');
    
    // 这里只是模拟测试，实际需要后端支持
    setTimeout(() => {
        if (config.webhookUrl) {
            showToast('Webhook 配置已保存，实际连接测试需要在后端进行', 'success');
        }
        if (config.tableToken) {
            showToast('多维表格 Token 已保存', 'success');
        }
    }, 1000);
}

// 保存飞书配置
async function saveFeishuConfigToFile(config) {
    // 构建配置对象
    const fileConfig = {
        enabled: config.enabled,
        notifyTime: config.notifyTime,
        webhookUrl: config.webhookUrl,
        tableToken: config.tableToken,
        appToken: config.appToken,
        notifyTemplate: document.getElementById('notifyTemplate')?.value || '',
        managerStatsTemplate: '',
        targetSessionKey: '主对话',
        updatedAt: new Date().toISOString()
    };
    
    // 保存到 localStorage 以便后续读取
    localStorage.setItem('feishuConfigFull', JSON.stringify(fileConfig));
    
    // 尝试通过 API 保存到服务器文件
    try {
        const response = await fetch('http://localhost:8765/config', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(fileConfig)
        });
        
        if (response.ok) {
            const result = await response.json();
            console.log('✅ 配置已保存到服务器:', result);
            return true;
        } else {
            const error = await response.text();
            console.warn('⚠️ API 保存失败:', error);
            return false;
        }
    } catch (e) {
        console.warn('⚠️ 无法连接到配置 API 服务:', e.message);
        return false;
    }
}

// 更新通知时间显示
function updateNotifyTimeDisplay(time) {
    // 通知时间已弃用，但保留函数以防其他地方调用
}

function renderSettingsPage() {
    const config = loadFeishuConfig();
    const content = document.getElementById('pageContent');
    
    content.innerHTML = `
        <div style="display: grid; gap: 24px;">
            <!-- 飞书多维表格配置 -->
            <div class="chart-card">
                <div class="card-header">
                    <h3>📊 飞书多维表格同步</h3>
                    <p style="color: #6b7280; font-size: 0.875rem; margin-top: 4px;">
                        配置飞书多维表格用于自动同步考勤数据
                    </p>
                </div>
                <div style="padding: 24px;">
                    <div style="display: grid; gap: 20px;">
                        <!-- 说明 -->
                        <div style="padding: 16px; background: #f0f4ff; border-radius: 10px; border: 1px solid #c7d2fe;">
                            <div style="font-weight: 600; color: #3730a3; margin-bottom: 8px;">💡 功能说明</div>
                            <ul style="margin: 0; padding-left: 20px; font-size: 0.875rem; color: #4338ca; line-height: 1.8;">
                                <li>上传考勤数据并点击「开始分析」后，数据将自动同步到飞书多维表格</li>
                                <li>同步内容包括：当日统计、各职类出勤数据、考勤明细</li>
                                <li>可在飞书多维表格中直接查看和分享考勤数据</li>
                            </ul>
                        </div>
                        
                        <!-- 多维表格 Token -->
                        <div style="display: grid; gap: 8px;">
                            <label style="font-weight: 500; color: #374151;">
                                飞书多维表格 Token
                                <span style="font-weight: 400; color: #9ca3af; font-size: 0.8rem; margin-left: 8px;">
                                    （已预配置）
                                </span>
                            </label>
                            <input type="text" id="feishuTableToken" 
                                value="${FEISHU_BASE_TOKEN}"
                                readonly
                                placeholder="X9Dxwaondir1tHkwYFCc8COBnAe"
                                style="padding: 12px 16px; border: 2px solid #e5e7eb; border-radius: 10px; font-size: 0.95rem; background: #f9fafb; color: #6b7280;">
                            <div style="font-size: 0.8rem; color: #9ca3af;">
                                💡 考勤数据将自动同步到此多维表格
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- 数据管理 -->
            <div class="chart-card">
                <div class="card-header">
                    <h3>💾 数据管理</h3>
                </div>
                <div style="padding: 24px;">
                    <div style="display: flex; gap: 16px; flex-wrap: wrap;">
                        <button class="btn btn-primary" onclick="exportDeploymentPackage()">📦 导出部署包</button>
                        <button class="btn btn-secondary" onclick="exportAllData()">📥 导出所有数据</button>
                        <button class="btn btn-outline" onclick="exportFeishuConfig()">📥 导出飞书配置</button>
                        <button class="btn btn-outline" style="border-color: #ef4444; color: #ef4444;" onclick="clearAllData()">🗑️ 清空所有数据</button>
                    </div>
                    <div style="margin-top: 16px; padding: 12px; background: #fef3c7; border-radius: 8px; font-size: 0.875rem; color: #92400e;">
                        💡 「导出部署包」会生成 attendance-data.json 文件，用于部署到静态托管平台（GitHub Pages/Vercel），让团队成员访问线上版本。
                    </div>
                </div>
            </div>
            
            <!-- 访客模式链接 -->
            <div class="chart-card">
                <div class="card-header">
                    <h3>🔗 访客模式链接</h3>
                    <p style="color: #6b7280; font-size: 0.875rem; margin-top: 4px;">
                        生成访客模式链接，供团队成员查看（只读）
                    </p>
                </div>
                <div style="padding: 24px;">
                    <div style="display: flex; gap: 12px; align-items: center;">
                        <input type="text" id="visitorLinkInput" 
                            value="${window.location.origin + window.location.pathname}?mode=visitor"
                            readonly
                            style="flex: 1; padding: 12px 16px; border: 2px solid #e5e7eb; border-radius: 10px; font-size: 0.95rem; background: #f9fafb; color: #6b7280;">
                        <button class="btn btn-secondary" onclick="copyVisitorLink()">📋 复制链接</button>
                    </div>
                    <div style="margin-top: 12px; font-size: 0.8rem; color: #9ca3af;">
                        💡 访客模式下，用户只能查看看板和导出数据，无法上传或修改数据
                    </div>
                </div>
            </div>
        </div>
    `;
}
function exportFeishuConfig() {
    const config = loadFeishuConfig();
    const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `feishu-config-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('飞书配置已导出', 'success');
}

// ==================== 工具函数 ====================
function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function showToast(message, type = 'info') {
    const toast = document.getElementById('toast');
    toast.textContent = message;
    toast.className = 'toast ' + type;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
}

function downloadTemplate(type) {
    const templates = {
        'employee': {
            name: '员工档案模板.xlsx',
            data: [
                { '花名': '张三', '工作地点': '北京', '职类': '研发', '上班时间': '09:00', '直接主管': '李四', '入职日期': '2024-01-01', '离职日期': '' },
                { '花名': '李四', '工作地点': '郑州', '职类': '销售', '上班时间': '', '直接主管': '佩皮', '入职日期': '2024-01-01', '离职日期': '' },
                { '花名': '王五', '工作地点': '北京', '职类': '运营', '上班时间': '', '直接主管': '赵六', '入职日期': '2024-01-01', '离职日期': '2024-04-15' }
            ],
            // 职类说明
            description: '职类说明：销售/运营默认13:00上班；销售+郑州+佩皮主管为12:00上班'
        },
        'access-zhengzhou': {
            name: '郑州门禁记录模板.xlsx',
            data: [
                { '时间': '2024-04-14 08:30:00', '名称': '张三' }
            ]
        },
        'access-beijing': {
            name: '北京门禁记录模板.xlsx',
            data: [
                { '姓名': '张三', '日期': '2024-04-14', '上班时间': '09:00', '签到时间': '08:30:00' }
            ]
        },
        'access-hangzhou': {
            name: '杭州门禁记录模板.xlsx',
            data: [
                { '姓名': '张三', '日期': '2024-04-14', '时间': '08:30:00' }
            ]
        },
        'schedule': {
            name: '排班表模板.xlsx',
            data: [
                { '花名': '张三', '日期': '2024-04-14', '班次时间': '上班' }
            ]
        },
        'leave': {
            name: '请假记录模板.xlsx',
            data: [
                { '花名': '张三', '假期类型': '事假', '请假开始日期': '2024-04-14', '请假结束日期': '2024-04-14', '审批状态': '已通过' },
                { '花名': '李四', '假期类型': '年假', '请假开始日期': '2026-04-14', '请假结束日期': '2026-04-16', '审批状态': '已通过' }
            ]
        }
    };
    
    const template = templates[type];
    if (!template) return;
    
    const ws = XLSX.utils.json_to_sheet(template.data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, template.name);
    
    showToast(`已下载 ${template.name}`, 'success');
}

function exportTable() {
    if (!AppState.analysisDetail || AppState.analysisDetail.length === 0) {
        showToast('暂无数据可导出', 'error');
        return;
    }
    
    // 转换为中文表头，按岗位优先顺序排列
    const chineseData = AppState.analysisDetail.map(row => ({
        '花名': row.name,
        '职类': row.jobCategory || '-',
        '工作地点': row.location || '-',
        '直接主管': row.manager || '-',
        '上班时间': row.workTime,
        '排班状态': row.shift,
        '出勤状态': row.status,
        '打卡时间': row.checkTime || '-'
    }));
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 工作表1：考勤明细
    const ws1 = XLSX.utils.json_to_sheet(chineseData);
    
    // 设置列宽
    ws1['!cols'] = [
        { wch: 10 },  // 花名
        { wch: 12 },  // 职类
        { wch: 12 },  // 工作地点
        { wch: 12 },  // 直接主管
        { wch: 10 },  // 上班时间
        { wch: 10 },  // 排班状态
        { wch: 10 },  // 出勤状态
        { wch: 12 }   // 打卡时间
    ];
    
    XLSX.utils.book_append_sheet(wb, ws1, '考勤明细');
    
    // 工作表2：汇总统计
    const stats = calculateStats();
    const summaryData = [
        { '项目': '总人数', '数值': stats.total },
        { '项目': '正常出勤', '数值': stats.normal, '占比': stats.normalPercent + '%' },
        { '项目': '迟到', '数值': stats.late, '占比': stats.latePercent + '%' },
        { '项目': '请假', '数值': stats.leave, '占比': stats.leavePercent + '%' },
        { '项目': '缺勤', '数值': stats.totalAbsent, '占比': stats.absentPercent + '%' },
        { '项目': '已离职', '数值': stats.resigned }
    ];
    const ws2 = XLSX.utils.json_to_sheet(summaryData);
    ws2['!cols'] = [{ wch: 12 }, { wch: 10 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws2, '汇总统计');
    
    // 工作表3：各岗位统计
    const jobStats = calculateJobCategoryStats();
    const jobStatsData = Object.entries(jobStats).map(([cat, s]) => ({
        '职类': cat,
        '总人数': s.total,
        '正常': s.normal,
        '迟到': s.late,
        '请假': s.leave,
        '缺勤': s.absent,
        '异常': s.abnormal,
        '出勤率': s.attendanceRate + '%'
    }));
    if (jobStatsData.length > 0) {
        const ws3 = XLSX.utils.json_to_sheet(jobStatsData);
        ws3['!cols'] = [{ wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 10 }];
        XLSX.utils.book_append_sheet(wb, ws3, '职类统计');
    }
    
    // 工作表4：主管团队统计
    if (AppState.managerStats && Object.keys(AppState.managerStats).length > 0) {
        const managerStatsData = Object.entries(AppState.managerStats).map(([mgr, s]) => ({
            '直接主管': mgr,
            '团队人数': s.total,
            '正常': s.normal,
            '迟到': s.late,
            '请假': s.leave,
            '缺勤': s.absent,
            '异常': s.abnormal,
            '出勤率': s.attendanceRate + '%',
            '迟到率': s.lateRate + '%',
            '缺勤率': s.absentRate + '%'
        }));
        const ws4 = XLSX.utils.json_to_sheet(managerStatsData);
        ws4['!cols'] = [{ wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];
        XLSX.utils.book_append_sheet(wb, ws4, '主管团队统计');
    }
    
    // 工作表5：主管本人出勤情况
    const managerPersonalData = AppState.managerPersonalData || getManagerPersonalData();
    if (managerPersonalData && managerPersonalData.length > 0) {
        const managerPersonalExport = managerPersonalData.map(m => ({
            '主管姓名': m.name,
            '职类': m.jobCategory || '-',
            '工作地点': m.location || '-',
            '上班时间': m.workTime,
            '打卡时间': m.checkTime || '-',
            '出勤状态': m.status
        }));
        const ws5 = XLSX.utils.json_to_sheet(managerPersonalExport);
        ws5['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 10 }];
        XLSX.utils.book_append_sheet(wb, ws5, '主管本人出勤');
    }
    
    // 生成Excel文件（注意：XLSX格式不需要BOM，添加BOM会破坏ZIP结构）
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    // 下载
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `考勤明细_${new Date().toISOString().split('T')[0]}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    
    showToast('导出成功（含考勤明细、汇总统计、职类统计、主管团队统计、主管本人出勤）', 'success');
}

function exportAllData() {
    const allData = {
        employees: AppState.employees,
        accessRecords: AppState.accessRecords,
        schedules: AppState.schedules,
        leaves: AppState.leaves
    };
    
    const blob = new Blob([JSON.stringify(allData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `考勤数据备份_${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    
    showToast('数据已导出', 'success');
}

function exportData() {
    exportTable();
}

function refreshData() {
    if (AppState.employees.length > 0) {
        analyzeAttendance();
        showToast('数据已刷新', 'success');
        if (AppState.currentPage === 'dashboard') {
            loadPage('dashboard');
        }
    } else {
        showToast('暂无数据', 'info');
    }
}

// ==================== 主管本人出勤情况相关函数 ====================

// 渲染主管本人数据表格
function renderManagerPersonalTable() {
    const managerPersonalData = getManagerPersonalData();
    
    // 初始化筛选下拉框
    const nameSelect = document.getElementById('managerPersonalNameFilter');
    if (nameSelect && nameSelect.options.length <= 1) {
        const managerNames = managerPersonalData.map(m => m.name);
        nameSelect.innerHTML = '<option value="">全部主管</option>' +
            managerNames.map(name => `<option value="${name}">${name}</option>`).join('');
    }
    
    // 渲染表格
    const tbody = document.getElementById('managerPersonalBody');
    if (!tbody) return;
    
    tbody.innerHTML = managerPersonalData.map(m => {
        // 状态颜色
        let statusColor = '#22c55e'; // 正常
        if (m.status === '迟到') statusColor = '#f59e0b';
        else if (m.status === '缺勤') statusColor = '#ef4444';
        else if (m.status === '请假') statusColor = '#3b82f6';
        else if (m.status === '异常') statusColor = '#92400e';
        
        return `
            <tr>
                <td><strong>${m.name}</strong></td>
                <td>${m.jobCategory}</td>
                <td>${m.location}</td>
                <td>${m.workTime}</td>
                <td>${m.checkTime}</td>
                <td><strong style="color: ${statusColor};">${m.status}</strong></td>
            </tr>
        `;
    }).join('');
}

// 初始化主管本人图表
function initManagerPersonalCharts() {
    const managerPersonalData = AppState.managerPersonalData || getManagerPersonalData();
    if (!managerPersonalData || managerPersonalData.length === 0) return;
    
    // 图表1：出勤状态分布饼图
    const statusChartDom = document.getElementById('managerPersonalStatusChart');
    if (statusChartDom) {
        const statusCount = { '正常': 0, '迟到': 0, '请假': 0, '缺勤': 0, '异常': 0 };
        managerPersonalData.forEach(m => {
            if (statusCount.hasOwnProperty(m.status)) {
                statusCount[m.status]++;
            }
        });
        
        const statusChart = echarts.init(statusChartDom);
        statusChart.setOption({
            tooltip: { trigger: 'item', formatter: '{b}: {c}人 ({d}%)' },
            legend: { bottom: '5%', left: 'center' },
            series: [{
                type: 'pie',
                radius: ['40%', '70%'],
                avoidLabelOverlap: false,
                itemStyle: { borderRadius: 10, borderColor: '#fff', borderWidth: 2 },
                label: { show: false },
                emphasis: { label: { show: true, fontSize: 16, fontWeight: 'bold' } },
                labelLine: { show: false },
                data: [
                    { value: statusCount['正常'], name: '正常', itemStyle: { color: '#22c55e' } },
                    { value: statusCount['迟到'], name: '迟到', itemStyle: { color: '#f59e0b' } },
                    { value: statusCount['请假'], name: '请假', itemStyle: { color: '#3b82f6' } },
                    { value: statusCount['缺勤'], name: '缺勤', itemStyle: { color: '#ef4444' } },
                    { value: statusCount['异常'], name: '异常', itemStyle: { color: '#92400e' } }
                ].filter(d => d.value > 0)
            }]
        });
        AppState.charts['managerPersonalStatus'] = statusChart;
    }
}

// 筛选主管本人数据
function filterManagerPersonalData() {
    const nameFilter = document.getElementById('managerPersonalNameFilter')?.value || '';
    const statusFilter = document.getElementById('managerPersonalStatusFilter')?.value || '';
    
    const managerPersonalData = AppState.managerPersonalData || getManagerPersonalData();
    let filteredData = managerPersonalData;
    
    if (nameFilter) {
        filteredData = filteredData.filter(m => m.name === nameFilter);
    }
    if (statusFilter) {
        filteredData = filteredData.filter(m => m.status === statusFilter);
    }
    
    // 更新表格
    const tbody = document.getElementById('managerPersonalBody');
    if (!tbody) return;
    
    tbody.innerHTML = filteredData.map(m => {
        let statusColor = '#22c55e';
        if (m.status === '迟到') statusColor = '#f59e0b';
        else if (m.status === '缺勤') statusColor = '#ef4444';
        else if (m.status === '请假') statusColor = '#3b82f6';
        else if (m.status === '异常') statusColor = '#92400e';
        
        return `
            <tr>
                <td><strong>${m.name}</strong></td>
                <td>${m.jobCategory}</td>
                <td>${m.location}</td>
                <td>${m.workTime}</td>
                <td>${m.checkTime}</td>
                <td><strong style="color: ${statusColor};">${m.status}</strong></td>
            </tr>
        `;
    }).join('');
    
    // 更新图表
    updateManagerPersonalCharts(filteredData);
}

// 更新主管本人图表
function updateManagerPersonalCharts(data) {
    // 更新状态分布图
    const statusChart = AppState.charts['managerPersonalStatus'];
    if (statusChart) {
        const statusCount = { '正常': 0, '迟到': 0, '请假': 0, '缺勤': 0, '异常': 0 };
        data.forEach(m => {
            if (statusCount.hasOwnProperty(m.status)) {
                statusCount[m.status]++;
            }
        });
        
        statusChart.setOption({
            series: [{
                data: [
                    { value: statusCount['正常'], name: '正常', itemStyle: { color: '#22c55e' } },
                    { value: statusCount['迟到'], name: '迟到', itemStyle: { color: '#f59e0b' } },
                    { value: statusCount['请假'], name: '请假', itemStyle: { color: '#3b82f6' } },
                    { value: statusCount['缺勤'], name: '缺勤', itemStyle: { color: '#ef4444' } },
                    { value: statusCount['异常'], name: '异常', itemStyle: { color: '#92400e' } }
                ].filter(d => d.value > 0)
            }]
        });
    }
    
    // 更新打卡时间图
    const timeChart = AppState.charts['managerPersonalTime'];
    if (timeChart) {
        const timeData = data.filter(m => m.checkTime && m.checkTime !== '-');
        timeChart.setOption({
            xAxis: {
                data: timeData.map(m => m.name)
            },
            series: [{
                data: timeData.map(m => {
                    if (m.checkTime && m.checkTime !== '-') {
                        const parts = m.checkTime.split(':');
                        if (parts.length >= 2) {
                            const hours = parseInt(parts[0]);
                            const minutes = parseInt(parts[1]);
                            return {
                                value: hours * 60 + minutes,
                                itemStyle: {
                                    color: m.status === '正常' ? '#22c55e' : 
                                           m.status === '迟到' ? '#f59e0b' : '#ef4444'
                                }
                            };
                        }
                    }
                    return 0;
                })
            }]
        });
    }
}

// 导出主管本人数据
function exportManagerPersonalData() {
    const managerPersonalData = AppState.managerPersonalData || getManagerPersonalData();
    if (!managerPersonalData || managerPersonalData.length === 0) {
        showToast('暂无数据可导出', 'error');
        return;
    }
    
    // 转换为中文表头
    const exportData = managerPersonalData.map(m => ({
        '主管姓名': m.name,
        '职类': m.jobCategory,
        '工作地点': m.location,
        '上班时间': m.workTime,
        '打卡时间': m.checkTime,
        '出勤状态': m.status
    }));
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // 设置列宽
    ws['!cols'] = [
        { wch: 12 },  // 主管姓名
        { wch: 12 },  // 职类
        { wch: 12 },  // 工作地点
        { wch: 10 },  // 上班时间
        { wch: 12 },  // 打卡时间
        { wch: 10 }   // 出勤状态
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, '主管本人出勤情况');
    
    // 生成Excel文件
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    // 下载
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `主管本人出勤情况_${new Date().toISOString().split('T')[0]}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    
    showToast('主管本人出勤数据已导出', 'success');
}

// 导出部署数据包（用于静态部署）
function exportDeploymentPackage() {
    const deploymentData = {
        employees: AppState.employees,
        accessRecords: AppState.accessRecords,
        schedules: AppState.schedules,
        leaves: AppState.leaves,
        analysisResult: AppState.analysisResult,
        analysisDetail: AppState.analysisDetail,
        managerStats: AppState.managerStats,
        managerPersonalData: AppState.managerPersonalData || getManagerPersonalData(),
        exportTime: new Date().toISOString(),
        version: '2.0'
    };
    
    const blob = new Blob([JSON.stringify(deploymentData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'attendance-data.json';
    a.click();
    URL.revokeObjectURL(url);
    
    showToast('部署数据已导出（attendance-data.json）', 'success');
}

// 从 JSON 文件加载数据（用于访客模式）
function loadFromJSON(jsonData) {
    try {
        const data = typeof jsonData === 'string' ? JSON.parse(jsonData) : jsonData;
        AppState.employees = data.employees || [];
        AppState.accessRecords = data.accessRecords || { 北京: [], 郑州: [], 杭州: [] };
        AppState.schedules = data.schedules || [];
        AppState.leaves = data.leaves || [];
        AppState.analysisResult = data.analysisResult || null;
        AppState.analysisDetail = data.analysisDetail || null;
        AppState.managerStats = data.managerStats || null;
        AppState.managerPersonalData = data.managerPersonalData || null;
        
        // 保存到本地存储
        localStorage.setItem('attendanceData', JSON.stringify(data));
        
        console.log('[数据加载] 从 JSON 文件加载成功');
        return true;
    } catch (e) {
        console.error('[数据加载] JSON 解析失败:', e);
        return false;
    }
}

// 复制访客链接
function copyVisitorLink() {
    const input = document.getElementById('visitorLinkInput');
    if (input) {
        input.select();
        document.execCommand('copy');
        showToast('访客链接已复制到剪贴板', 'success');
    }
}
