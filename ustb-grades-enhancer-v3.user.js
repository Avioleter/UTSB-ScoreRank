// ==UserScript==
// @name         USTB教务系统 - 成绩增强
// @namespace    https://github.com/ustb-grade-enhancer
// @version      3.0.0
// @description  右下角面板展示排名+平时/考试成绩，支持导出Excel，计算加权平均分
// @author       Avioleter
// @match        https://byyt.ustb.edu.cn/cjgl/*
// @grant        GM_download
// @require      https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js
// @run-at       document-start
// @license      MIT
// ==/UserScript==

(function () {
    'use strict';

    // ======== 配置 ========
    const CFG = {
        debug: true,
        detailConcurrency: 4,
    };

    // ======== 状态 ========
    let courses      = [];
    let fetching     = false;
    let detailLoaded = false;

    const log  = (...a) => CFG.debug && console.log('[USTB]', ...a);
    const warn = (...a) => console.warn('[USTB]', ...a);

    // ======== API 拦截 ========
    function hookXHR() {
        const _open = XMLHttpRequest.prototype.open;
        const _send = XMLHttpRequest.prototype.send;
        XMLHttpRequest.prototype.open = function (m, u) { this.__u = u; return _open.apply(this, arguments); };
        XMLHttpRequest.prototype.send = function (body) {
            const x = this;
            x.addEventListener('load', () => {
                if (!x.__u || x.status !== 200) return;
                try {
                    const d = JSON.parse(x.responseText);
                    if (d?.code === 200 && d.content?.list) gotCourses(d.content.list, d.content.total);
                } catch (e) {}
            });
            return _send.apply(this, arguments);
        };
    }

    function hookFetch() {
        const _fetch = window.fetch;
        window.fetch = function (input, init) {
            return _fetch.apply(this, arguments).then(r => {
                const url = typeof input === 'string' ? input : (input?.url || '');
                r.clone().text().then(t => {
                    try {
                        const d = JSON.parse(t);
                        if (d?.code === 200 && d.content?.list) gotCourses(d.content.list, d.content.total);
                    } catch (e) {}
                }).catch(() => {});
                return r;
            });
        };
    }

    function gotCourses(list, total) {
        log('捕获', list.length, '条, 共', total, '条');
        const ids = new Set(courses.map(c => c.id));
        let added = 0;
        for (const c of list) {
            if (!ids.has(c.id)) { courses.push(c); ids.add(c.id); added++; }
        }
        if (added) log('新增', added, '条, 总计', courses.length, '条');

        updatePanel();

        // 首次获取后自动拉取分项成绩
        if (!detailLoaded && courses.length > 0) {
            detailLoaded = true;
            setTimeout(fetchAllDetails, 600);
        }
    }

    // ======== 分项成绩获取 ========
    async function fetchAllDetails(silent) {
        if (fetching) return;
        fetching = true;
        if (!silent) log('🔍 获取分项成绩...');
        updatePanel();

        let cursor = 0;
        const total = courses.length;
        let ok = 0;
        const workers = Array.from(
            { length: Math.min(CFG.detailConcurrency, total) },
            async () => {
                while (cursor < total) {
                    const idx = cursor++;
                    const course = courses[idx];
                    try {
                        const detail = await fetchDetail(course);
                        if (detail.length > 0) {
                            course.__detail = detail;
                            course.__pscj  = extractScore(detail, '平时');
                            course.__kscj  = extractScore(detail, '考试', '期末');
                            course.__sycj  = extractScore(detail, '实验');
                            ok++;
                        } else {
                            course.__detail = [];
                        }
                    } catch (e) {
                        course.__detail = null;
                    }
                    updatePanel();
                }
            }
        );
        await Promise.all(workers);

        fetching = false;
        const msg = `✅ ${ok}/${total} 门`;
        if (!silent) log('分项成绩:', msg);
        updatePanel(msg);
    }

    async function fetchDetail(course) {
        if (!course.rwid || !course.id) return [];
        const resp = await fetch('/cjgl/grcjcx/seeFx', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
            body: new URLSearchParams({ rwid: course.rwid, cjid: course.id }).toString(),
        });
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        const rows = await resp.json();
        if (!Array.isArray(rows)) return [];
        return rows.filter(row => !course.glcjid || row.GLCJID === course.glcjid);
    }

    function extractScore(detail, ...names) {
        if (!Array.isArray(detail)) return null;
        for (const part of detail) {
            const label = (part.FXMC || part.FXDM || '').toLowerCase();
            for (const name of names) {
                if (label.includes(name.toLowerCase())) {
                    const s = parseFloat(part.DF);
                    if (!isNaN(s)) return s;
                }
            }
        }
        return null;
    }

    // ======== 加权平均分 ========
    /**
     * 判断课程是否计入加权平均分
     * 计入：必修课、实验课、平台课（学分 > 0）
     * 不计入：素质拓展、专业拓展、0学分课程等
     */
    function isWeightedCourse(course) {
        const xf = parseFloat(course.xf) || 0;
        if (xf <= 0) return false; // 0 学分不计入

        // 取课程性质和类别（不同教务版本字段不同，两者都检查）
        const typeStr = [course.kcxz, course.kclb, course.kcxzmc, course.kclbmc]
            .filter(Boolean)
            .join(' ');

        // 明确不计入的类型
        if (/素质拓展|专业拓展|公共选修|通识选修|全校任选|校际选修|创新拓展|个性拓展/i.test(typeStr)) {
            return false;
        }

        // 明确计入的类型：必修、实验、平台
        if (/必修|实验|平台|专业核心|学科基础|集中实践|实践环节|毕业设计|毕业论文|实习/i.test(typeStr)) {
            return true;
        }

        // 未识别的类型——字段可能为空或取法不同，保守计入（避免漏算必修课）
        return true;
    }

    function calcAvg() {
        let sw = 0, sc = 0;
        let included = 0, excluded = 0;
        for (const c of courses) {
            const xf = parseFloat(c.xf) || 0;
            const cj = parseFloat(c.zzcj) || parseFloat(c.zpcj) || parseFloat(c.xscj) || 0;
            if (xf <= 0 || cj <= 0) continue; // 无学分或无成绩直接跳过

            if (isWeightedCourse(c)) {
                sw += cj * xf;
                sc += xf;
                included++;
            } else {
                excluded++;
            }
        }

        // 首次计算时输出统计信息
        if (included + excluded > 0 && !calcAvg._logged) {
            calcAvg._logged = true;
            log(`📊 加权计算: ${included} 门计入, ${excluded} 门排除（素质拓展/专业拓展等不计入）`);
        }

        return sc > 0 ? (sw / sc).toFixed(2) : '--';
    }
    calcAvg._logged = false;

    // ======== 排名百分位 ========
    function rankInfo(course) {
        const r = parseInt(course.pm, 10);
        const t = parseInt(course.zrs, 10);
        if (!r || !t || t <= 0) return null;
        return { rank: r, total: t, pct: (r / t * 100).toFixed(1) };
    }

    // ======== UI 构建 ========
    function buildUI() {
        document.querySelectorAll('#ustb-root').forEach(el => el.remove());

        const root = document.createElement('div');
        root.id = 'ustb-root';
        root.innerHTML = `
        <style>
            #ustb-root, #ustb-root * { box-sizing: border-box; }
            #ustb-root { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", sans-serif; }
            #ustb-float-btn {
                position: fixed; bottom: 30px; right: 30px; z-index: 2147483646;
                width: 46px; height: 46px; border-radius: 50%; border: 0; cursor: pointer;
                background: #1890ff; color: #fff; font-size: 20px;
                box-shadow: 0 4px 14px rgba(24,144,255,0.35);
                display: flex; align-items: center; justify-content: center;
                transition: transform 0.15s, box-shadow 0.15s;
            }
            #ustb-float-btn:hover { transform: scale(1.08); box-shadow: 0 6px 20px rgba(24,144,255,0.45); }
            #ustb-overlay {
                display: none; position: fixed; inset: 0; z-index: 2147483647;
                background: rgba(15,23,42,0.45); align-items: center; justify-content: center;
            }
            #ustb-overlay.open { display: flex; }
            #ustb-panel {
                width: min(1100px, 95vw); max-height: 90vh; background: #fff;
                border-radius: 12px; box-shadow: 0 20px 60px rgba(0,0,0,0.25);
                display: flex; flex-direction: column; overflow: hidden;
            }
            #ustb-panel-header {
                display: flex; align-items: center; justify-content: space-between;
                padding: 16px 20px; border-bottom: 1px solid #eee;
                flex-shrink: 0;
            }
            #ustb-panel-header h2 { margin: 0; font-size: 18px; color: #1a1a2e; }
            #ustb-panel-header .ustb-actions { display: flex; gap: 8px; }
            #ustb-panel .ustb-btn {
                border: 1px solid #d9d9d9; border-radius: 6px; padding: 6px 14px;
                background: #fff; cursor: pointer; font-size: 13px; color: #555;
                transition: all 0.15s;
            }
            #ustb-panel .ustb-btn:hover { border-color: #1890ff; color: #1890ff; }
            #ustb-panel .ustb-btn.primary { background: #1890ff; color: #fff; border-color: #1890ff; }
            #ustb-panel .ustb-btn.primary:hover { background: #40a9ff; }
            #ustb-summary {
                display: flex; gap: 16px; padding: 14px 20px; flex-shrink: 0;
                flex-wrap: wrap;
            }
            #ustb-summary .ustb-stat {
                background: #f7f9fc; border-radius: 8px; padding: 10px 16px;
                min-width: 100px;
            }
            #ustb-summary .ustb-stat .label { font-size: 11px; color: #999; }
            #ustb-summary .ustb-stat .value { font-size: 20px; font-weight: 700; color: #1a1a2e; margin-top: 2px; }
            #ustb-panel-body {
                overflow: auto; flex: 1; padding: 0 8px 8px;
            }
            #ustb-table {
                width: 100%; border-collapse: collapse; font-size: 13px;
                min-width: 800px;
            }
            #ustb-table th, #ustb-table td {
                border: 1px solid #eee; padding: 9px 10px; text-align: left;
                vertical-align: middle; line-height: 1.5;
            }
            #ustb-table thead th {
                position: sticky; top: 0; z-index: 2;
                background: #fafafa; font-weight: 600; color: #555;
            }
            #ustb-table thead th:first-child { border-radius: 6px 0 0 0; }
            #ustb-table thead th:last-child  { border-radius: 0 6px 0 0; }
            #ustb-table .ustb-rank-badge {
                display: inline-block; padding: 2px 8px; border-radius: 4px;
                font-weight: 600; font-size: 12px; white-space: nowrap;
            }
            #ustb-table .ustb-rank-gold   { background: #fff7e6; color: #ad6800; }
            #ustb-table .ustb-rank-blue   { background: #e6f7ff; color: #1890ff; }
            #ustb-table .ustb-part-tag {
                display: inline-block; margin: 1px 3px; padding: 2px 7px;
                border-radius: 4px; background: #f0f5ff; color: #2b6cb0;
                font-size: 12px; white-space: nowrap;
            }
            #ustb-table .ustb-empty { color: #ccc; }
            #ustb-status-line {
                padding: 8px 20px; font-size: 12px; color: #999;
                border-top: 1px solid #f0f0f0; flex-shrink: 0;
            }
            @media (max-width: 700px) {
                #ustb-panel { width: 100vw; max-height: 100vh; border-radius: 0; }
                #ustb-panel-header { padding: 12px 14px; }
                #ustb-panel-body { padding: 0 4px 4px; }
                #ustb-summary { padding: 10px 14px; gap: 8px; }
            }
        </style>
        <button id="ustb-float-btn" title="成绩增强">📊</button>
        <div id="ustb-overlay" role="dialog" aria-modal="true">
            <div id="ustb-panel">
                <div id="ustb-panel-header">
                    <h2>📊 成绩增强</h2>
                    <div class="ustb-actions">
                        <button class="ustb-btn" id="ustb-btn-refresh">🔄 刷新</button>
                        <button class="ustb-btn primary" id="ustb-btn-detail">🔍 获取平时/考试成绩</button>
                        <button class="ustb-btn" id="ustb-btn-export">📥 导出 Excel</button>
                        <button class="ustb-btn" id="ustb-btn-close">✕ 关闭</button>
                    </div>
                </div>
                <div id="ustb-summary"></div>
                <div id="ustb-panel-body">
                    <table id="ustb-table">
                        <thead>
                            <tr>
                                <th style="width:100px">学期</th>
                                <th style="width:200px">课程名称</th>
                                <th style="width:58px">学分</th>
                                <th style="width:58px">总评</th>
                                <th style="width:120px">排名</th>
                                <th>分项成绩</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>
                </div>
                <div id="ustb-status-line"></div>
            </div>
        </div>`;
        document.documentElement.appendChild(root);

        // 事件绑定
        const overlay   = document.getElementById('ustb-overlay');
        const floatBtn  = document.getElementById('ustb-float-btn');

        floatBtn.addEventListener('click', () => { overlay.classList.add('open'); updatePanel(); });
        document.getElementById('ustb-btn-close').addEventListener('click', () => overlay.classList.remove('open'));
        overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.classList.remove('open'); });
        document.addEventListener('keydown', (e) => { if (e.key === 'Escape') overlay.classList.remove('open'); });

        document.getElementById('ustb-btn-refresh').addEventListener('click', () => {
            courses = []; detailLoaded = false; calcAvg._logged = false;
            updatePanel('已重置，切换页面或筛选条件重新加载数据');
            log('已刷新');
        });

        document.getElementById('ustb-btn-detail').addEventListener('click', () => {
            if (!fetching && courses.length > 0) fetchAllDetails();
        });

        document.getElementById('ustb-btn-export').addEventListener('click', exportExcel);
    }

    // ======== 面板更新 ========
    function updatePanel(statusMsg) {
        const tbody = document.querySelector('#ustb-table tbody');
        const summary = document.getElementById('ustb-summary');
        const statusLine = document.getElementById('ustb-status-line');
        if (!tbody || !summary) return; // UI 还没初始化

        // 底栏状态
        if (statusMsg) {
            statusLine.textContent = statusMsg;
        } else if (fetching) {
            statusLine.textContent = '⏳ 正在获取分项成绩...';
        } else if (courses.some(c => c.__pscj !== undefined)) {
            const n = courses.filter(c => c.__pscj !== null && c.__pscj !== undefined).length;
            statusLine.textContent = `✅ 已获取 ${n}/${courses.length} 门课程的分项成绩`;
        } else if (detailLoaded && !fetching) {
            statusLine.textContent = '💡 点击"获取平时/考试成绩"查询分项明细';
        } else {
            statusLine.textContent = courses.length > 0
                ? `已加载 ${courses.length} 门课程（分项成绩待获取）`
                : '等待数据加载...';
        }

        // 统计卡片
        const avg = calcAvg();
        const ranked = courses.filter(c => parseInt(c.pm) > 0).length;
        const hasDetail = courses.filter(c => c.__pscj !== null && c.__pscj !== undefined || c.__kscj !== null && c.__kscj !== undefined).length;
        summary.innerHTML = `
            <div class="ustb-stat"><div class="label">课程总数</div><div class="value">${courses.length}</div></div>
            <div class="ustb-stat"><div class="label">加权平均分</div><div class="value" style="color:#1890ff">${avg}</div></div>
            <div class="ustb-stat"><div class="label">有排名课程</div><div class="value">${ranked}</div></div>
            <div class="ustb-stat"><div class="label">分项成绩</div><div class="value" style="color:${hasDetail ? '#52c41a' : '#faad14'}">${hasDetail}/${courses.length}</div></div>
        `;

        // 课程表格
        if (courses.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:40px;color:#999;">暂无数据，请在成绩页面加载课程后再打开面板</td></tr>';
            return;
        }

        const rows = courses.map(c => {
            const ri = rankInfo(c);
            let rankHTML;
            if (ri) {
                const cls = ri.rank <= 3 ? 'ustb-rank-gold' : 'ustb-rank-blue';
                rankHTML = `<span class="ustb-rank-badge ${cls}">第 ${ri.rank}/${ri.total}<br>前 ${ri.pct}%</span>`;
            } else {
                rankHTML = '<span class="ustb-empty">—</span>';
            }

            // 分项成绩
            let partsHTML = '';
            const parts = [];
            if (c.__pscj !== null && c.__pscj !== undefined) parts.push(`<span class="ustb-part-tag">平时 <b>${c.__pscj}</b></span>`);
            if (c.__kscj !== null && c.__kscj !== undefined) parts.push(`<span class="ustb-part-tag">考试 <b>${c.__kscj}</b></span>`);
            if (c.__sycj !== null && c.__sycj !== undefined) parts.push(`<span class="ustb-part-tag">实验 <b>${c.__sycj}</b></span>`);

            if (parts.length > 0) {
                partsHTML = parts.join('');
            } else if (Array.isArray(c.__detail) && c.__detail.length > 0) {
                // 有原始数据但没匹配到平时/考试/实验，展示全部
                partsHTML = c.__detail.map(p => {
                    const name = p.FXMC || p.FXDM || '?';
                    const s = parseFloat(p.DF);
                    return !isNaN(s) ? `<span class="ustb-part-tag">${name} <b>${s}</b></span>` : '';
                }).join('');
            } else if (c.__detail === null) {
                partsHTML = '<span class="ustb-empty" style="color:#ff7875">获取失败</span>';
            } else if (c.__detail && c.__detail.length === 0) {
                partsHTML = '<span class="ustb-empty">无分项</span>';
            } else {
                partsHTML = '<span class="ustb-empty">待获取</span>';
            }

            const score = c.zzcj || c.zpcj || c.xscj || '—';

            return `<tr>
                <td>${esc(c.xnxqmc)}</td>
                <td><b>${esc(c.kcmc)}</b><br><span style="font-size:11px;color:#999">${esc(c.kcdm)}</span></td>
                <td>${esc(c.xf)}</td>
                <td><b>${esc(score)}</b></td>
                <td>${rankHTML}</td>
                <td>${partsHTML}</td>
            </tr>`;
        }).join('');

        tbody.innerHTML = rows;
    }

    function esc(v) {
        const s = String(v ?? '');
        return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    // ======== Excel 导出 ========
    function exportExcel() {
        if (courses.length === 0) { alert('暂无数据'); return; }
        if (typeof XLSX === 'undefined') { alert('Excel 库加载中，请稍后'); return; }

        const data = courses.map(c => ({
            '学年学期': c.xnxqmc || '',
            '课程代码': c.kcdm || '',
            '课程名称': c.kcmc || '',
            '课程性质': c.kcxz || '',
            '课程类别': c.kclb || '',
            '学分':     c.xf || '',
            '平时分':   (c.__pscj != null) ? c.__pscj : '',
            '考试分':   (c.__kscj != null) ? c.__kscj : '',
            '实验分':   (c.__sycj != null) ? c.__sycj : '',
            '总评成绩': c.zzcj || c.zpcj || c.xscj || '',
            '排名':     c.pm || '',
            '总人数':   c.zrs || '',
            '开课学院': c.yxmc || '',
            '考核方式': c.khfs || '',
            '补考重修': c.bkcx || '',
        }));
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '成绩');
        const d = new Date();
        XLSX.writeFile(wb, `成绩_${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}.xlsx`);
        log('Excel 已导出');
    }

    // ======== 调试 ========
    window.__USTB__ = {
        courses,
        dump:    () => console.table(courses.map(c => ({
            kcmc: c.kcmc, pm: c.pm, zrs: c.zrs, zzcj: c.zzcj,
            平时分: c.__pscj, 考试分: c.__kscj, 实验分: c.__sycj,
        }))),
        raw:     (i = 0) => { if (courses[i]) { log(courses[i].kcmc); console.table(courses[i].__detail); } },
        detail:  () => fetchAllDetails(),
    };

    // ======== 启动 ========
    hookXHR();
    hookFetch();

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    function init() {
        log('🚀 v3.0.0 — 面板式成绩增强');
        buildUI();
    }
})();
