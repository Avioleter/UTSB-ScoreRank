// ==UserScript==
// @name         USTB教务系统 - 功能拓展
// @namespace    https://github.com/ustb-grade-enhancer
// @version      2.4.0
// @description  成绩备注列自动显示排名，支持导出 Excel，显示加权平均分
// @author       USTB Student
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
        badge: { bg: '#1890ff', fg: '#fff', radius: '4px', size: '12px', pad: '2px 8px' },
        top3:  { bg: '#faad14', fg: '#fff' },
    };

    // ======== 状态 ========
    let courses    = [];       // 课程数据
    let remarkIdx  = -1;       // 备注列索引缓存
    let injTimer   = null;     // 注入防抖
    let injecting  = false;    // MutationObserver 防回环

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
                if (x.__u && x.status === 200) {
                    try { const d = JSON.parse(x.responseText); if (d?.code === 200 && d.content?.list) gotData(d); } catch (e) {}
                }
            });
            return _send.apply(this, arguments);
        };
    }

    function hookFetch() {
        const _fetch = window.fetch;
        window.fetch = function (input, init) {
            return _fetch.apply(this, arguments).then(r => {
                r.clone().text().then(t => {
                    try { const d = JSON.parse(t); if (d?.code === 200 && d.content?.list) gotData(d); } catch (e) {}
                }).catch(() => {});
                return r;
            });
        };
    }

    function gotData(data) {
        const list = data.content.list;
        log('捕获', list.length, '条, 共', data.content.total, '条');

        const ids = new Set(courses.map(c => c.id));
        let added = 0;
        for (const c of list) {
            if (!ids.has(c.id)) { courses.push(c); ids.add(c.id); added++; }
        }
        if (added) log('新增', added, '条, 总计', courses.length, '条');

        clearTimeout(injTimer);
        injTimer = setTimeout(inject, 300);
        updateAvg();
    }

    // ======== DOM 注入 ========
    function findTable() {
        const t = document.querySelector('tbody.ivu-table-tbody');
        return t && t.querySelectorAll('tr.ivu-table-row').length > 0 ? t : null;
    }

    function findRemarkIdx() {
        if (remarkIdx >= 0) return remarkIdx;
        const ths = document.querySelectorAll('thead th, .ivu-table-header th');
        let idx = 0;
        for (const th of ths) {
            if ((th.textContent || '').includes('备注')) { remarkIdx = idx; return idx; }
            idx++;
        }
        remarkIdx = idx > 0 ? idx - 1 : 0;
        return remarkIdx;
    }

    function badgeHTML(pm, zrs) {
        if (pm === null || pm === undefined || pm === '') return null;
        const n = parseInt(pm, 10);
        const bg = (!isNaN(n) && n <= 3) ? CFG.top3.bg : CFG.badge.bg;
        const fg = (!isNaN(n) && n <= 3) ? CFG.top3.fg : CFG.badge.fg;
        return `<span class="ustb-badge" style="
            display:inline-block;background:${bg};color:${fg};
            border-radius:${CFG.badge.radius};font-size:${CFG.badge.size};
            padding:${CFG.badge.pad};font-weight:600;white-space:nowrap;
        ">排名 ${pm}/${zrs}</span>`;
    }

    function inject() {
        const tbody = findTable();
        if (!tbody) return;
        injecting = true;

        // 清除旧徽章
        tbody.querySelectorAll('.ustb-badge').forEach(el => el.remove());

        const ri = findRemarkIdx();
        if (ri < 0) { injecting = false; return; }

        const pool = [...courses];
        let cnt = 0;

        for (const row of tbody.querySelectorAll('tr.ivu-table-row')) {
            const cells = row.querySelectorAll('td');
            if (cells.length === 0 || ri >= cells.length) continue;
            const rowText = (row.textContent || '').trim();

            // 最长匹配
            let best = null, bestLen = 0;
            for (const c of pool) {
                const name = (c.kcmc || '').trim();
                if (name && rowText.includes(name) && name.length > bestLen) {
                    best = c; bestLen = name.length;
                }
            }
            if (!best) continue;
            pool.splice(pool.indexOf(best), 1);

            const b = badgeHTML(best.pm, best.zrs);
            if (!b) continue;

            cells[ri].insertAdjacentHTML('beforeend', b);
            cnt++;
        }

        if (cnt) log('注入:', cnt, '个徽章');
        injecting = false;
    }

    function retryInject(max, ms) {
        let n = 0;
        function f() {
            if (findTable() && courses.length > 0) inject();
            else if (n < max) { n++; setTimeout(f, ms); }
        }
        f();
    }

    // ======== 面板 ========
    function createPanel() {
        // 清理旧版本残留
        document.querySelectorAll('#ustb-btn, #ustb-panel').forEach(el => el.remove());

        const html = `
        <div id="ustb-btn" style="position:fixed;bottom:30px;right:30px;z-index:99999;
            width:44px;height:44px;border-radius:50%;background:#1890ff;color:#fff;
            cursor:pointer;display:flex;align-items:center;justify-content:center;
            font-size:18px;box-shadow:0 4px 12px rgba(0,0,0,0.3);user-select:none;"
            title="成绩增强">📊</div>
        <div id="ustb-panel" style="position:fixed;bottom:84px;right:30px;z-index:99998;
            display:none;background:#fff;border-radius:8px;padding:16px;width:200px;
            box-shadow:0 8px 24px rgba(0,0,0,0.2);
            font-family:system-ui,sans-serif;font-size:13px;">
            <div style="font-weight:600;margin-bottom:8px;color:#333;">📊 成绩增强</div>
            <div style="background:#f5f5f5;border-radius:6px;padding:8px 10px;margin-bottom:10px;">
                <span style="color:#999;font-size:11px;">加权平均分</span>
                <span id="ustb-avg" style="float:right;font-size:18px;font-weight:700;color:#1890ff;">--</span>
            </div>
            <button id="ustb-refresh" class="ustb-pbtn" style="display:block;width:100%;margin-bottom:6px;
                padding:7px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;
                cursor:pointer;font-size:13px;">🔄 刷新排名</button>
            <button id="ustb-export" class="ustb-pbtn" style="display:block;width:100%;
                padding:7px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;
                cursor:pointer;font-size:13px;">📥 导出 Excel</button>
        </div>`;
        document.body.insertAdjacentHTML('beforeend', html);

        const btn   = document.getElementById('ustb-btn');
        const panel = document.getElementById('ustb-panel');
        let visible = false;

        btn.addEventListener('click', () => {
            visible = !visible;
            panel.style.display = visible ? 'block' : 'none';
            if (visible) updateAvg();
        });

        document.getElementById('ustb-refresh').addEventListener('click', () => {
            courses = []; remarkIdx = -1;
            clearTimeout(injTimer);
            injTimer = setTimeout(inject, 500);
            updateAvg();
            log('已刷新');
        });

        document.getElementById('ustb-export').addEventListener('click', exportExcel);

        // hover
        panel.querySelectorAll('.ustb-pbtn').forEach(b => {
            b.addEventListener('mouseenter', () => { b.style.borderColor = '#1890ff'; b.style.color = '#1890ff'; });
            b.addEventListener('mouseleave', () => { b.style.borderColor = '#d9d9d9'; b.style.color = '#333'; });
        });
    }

    // ======== 加权平均分 ========
    function calcAvg() {
        let sw = 0, sc = 0;
        for (const c of courses) {
            const xf = parseFloat(c.xf) || 0;
            const cj = parseFloat(c.zzcj) || parseFloat(c.zpcj) || parseFloat(c.xscj) || 0;
            if (xf > 0 && cj > 0) { sw += cj * xf; sc += xf; }
        }
        return sc > 0 ? (sw / sc).toFixed(2) : '--';
    }

    function updateAvg() {
        const el = document.getElementById('ustb-avg');
        if (el) el.textContent = calcAvg();
    }

    // ======== Excel 导出 ========
    function exportExcel() {
        if (courses.length === 0) { alert('暂无数据'); return; }
        if (typeof XLSX === 'undefined') { alert('Excel 库加载中，请稍后'); return; }

        const data = courses.map(c => ({
            '学年学期': c.xnxqmc || '', '课程代码': c.kcdm || '', '课程名称': c.kcmc || '',
            '课程性质': c.kcxz || '', '课程类别': c.kclb || '', '学分': c.xf || '',
            '成绩': c.zzcj || c.zpcj || c.xscj || '', '排名': c.pm || '', '总人数': c.zrs || '',
            '开课学院': c.yxmc || '', '考核方式': c.khfs || '', '补考重修': c.bkcx || '',
        }));
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '成绩');
        const d = new Date();
        XLSX.writeFile(wb, `成绩_${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}.xlsx`);
        log('Excel 已导出');
    }

    // ======== 调试 ========
    window.__USTB__ = { courses, inject, exportExcel, calcAvg, dump: () => console.table(courses) };

    // ======== 启动 ========
    hookXHR();
    hookFetch();

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    function init() {
        log('🚀 v2.4.0 — 排名+导出+加权平均分');
        createPanel();

        setTimeout(() => retryInject(20, 500), 800);

        new MutationObserver(() => {
            if (injecting) return;
            clearTimeout(injTimer);
            injTimer = setTimeout(inject, 400);
        }).observe(document.body, { childList: true, subtree: true });
    }
})();
