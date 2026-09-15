/* CPK-PCBA - 前端交互逻辑 */

// 拼接应用根路径 (含 /cpk 前缀) 与接口地址
function api(path) {
    const base = (window.BASE_URL || "").replace(/\/+$/, "");
    return base + (path.startsWith("/") ? path : "/" + path);
}

document.addEventListener("DOMContentLoaded", () => {
    // ===== 元素引用 =====
    const uploadForm = document.getElementById("uploadForm");
    const fileInput = document.getElementById("fileInput");
    const productSelect = document.getElementById("productSelect");
    const uploadBtn = document.getElementById("uploadBtn");
    const statusAlert = document.getElementById("statusAlert");

    const overviewSection = document.getElementById("overviewSection");
    const statCards = document.getElementById("statCards");
    const fileBadge = document.getElementById("fileBadge");

    const sidebarFilter = document.getElementById("sidebarFilter");
    const expandAllBtn = document.getElementById("expandAllBtn");
    const selectAllBtn = document.getElementById("selectAllBtn");
    const unselectAllBtn = document.getElementById("unselectAllBtn");

    const chartsSection = document.getElementById("chartsSection");
    const chartsContainer = document.getElementById("chartsContainer");

    const gridBtn = document.getElementById("gridBtn");
    const listBtn = document.getElementById("listBtn");

    // 图表放大 Modal
    const chartZoomModal = document.getElementById("chartZoomModal");
    const zoomModalTitle = document.getElementById("zoomModalTitle");
    const zoomChartBody = document.getElementById("zoomChartBody");
    const zoomDownloadBtn = document.getElementById("zoomDownloadBtn");
    let zoomChartInstance = null;  // 放大弹窗中的 ECharts 实例

    // 如果是统计页面，跳过首页特定逻辑
    if (!uploadForm && !chartsContainer) {
        return;
    }

    // 导航栏"数据统计"按钮仅在解析出数据后启用
    function enableStatsNav() {
        const btn = document.getElementById("navStatsBtn");
        if (btn) {
            btn.classList.remove("disabled");
            btn.removeAttribute("aria-disabled");
            btn.title = "查看数据统计";
        }
    }

    // 当前数据存储
    let chartInstances = [];
    let wrapperElements = [];   // 按 chartIdx 对应 DOM wrapper (.chart-wrapper) + echartsInstance
    let currentCharts = [];     // 当前全部图表数据 (与 chartsContainer 顺序一致)
    let currentTestItems = [];
    // 可见状态:
    //   visibleChartIdx: Set<number>    哪些图表显示 (一级)
    //   visibleSeries[chartIdx]: Set<seriesName>   每个图表内哪些 series 显示 (二级)
    let visibleChartIdx = new Set();
    let visibleSeries = {};
    // 滚动联动: 当前视口内的 chartIdx
    let scrollIntersectionObserver = null;
    let scrollActiveChartIdx = -1;

    // ====================================================================
    //  0. 从 session 恢复状态 (从统计页返回时自动加载)
    // ====================================================================
    const restoreState = window.__RESTORE_STATE__;
    if (restoreState && restoreState.has_data) {
        fetch(api("/api/session-charts"))
            .then(r => r.json())
            .then(data => {
                if (!data.has_data || !data.charts || data.charts.length === 0) return;

                fileBadge.textContent = data.filename || "";
                fileBadge.classList.remove("d-none");
                // 恢复产品类别下拉
                const opt = Array.from(productSelect.options).find(o => o.text === data.product_category);
                if (opt) productSelect.value = opt.value;

                currentCharts = data.charts;
                currentTestItems = [];
                currentCharts.forEach(c => {
                    if (c.series) currentTestItems.push(...c.series.map(s => s.name));
                    else if (c.item) currentTestItems.push(c.item);
                });

                visibleChartIdx = new Set(currentCharts.map((_, i) => i));
                visibleSeries = {};
                currentCharts.forEach((c, i) => {
                    const names = c.series && c.series.length
                        ? c.series.map(s => s.name)
                        : [c.item];
                    visibleSeries[i] = new Set(names);
                });

                renderSidebarTree(currentCharts);
                renderCharts(currentCharts);
                overviewSection.classList.remove("d-none");
                enableStatsNav();
            })
            .catch(() => { /* 静默失败 */ });
    }

    // ====================================================================
    //  1. 上传并解析 Excel
    // ====================================================================
    uploadForm.addEventListener("submit", async (e) => {
        e.preventDefault();

        if (!fileInput.files.length) {
            showAlert("请先选择 Excel 文件", "warning");
            return;
        }

        const formData = new FormData();
        formData.append("file", fileInput.files[0]);
        formData.append("product_category", productSelect.value);

        uploadBtn.classList.add("btn-loading");
        uploadBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> 解析中...';

        try {
            const resp = await fetch(api("/api/upload"), { method: "POST", body: formData });
            const data = await resp.json();

            if (!resp.ok) {
                showAlert(data.error || "上传失败", "danger");
                return;
            }

            showAlert(
                `解析成功: ${data.filename}  产品 ${data.product_count || 0} / 项目 ${data.analysis_item_count || 0} / 图表 ${data.chart_count || 0}`,
                "success"
            );
            currentTestItems = data.test_items || [];
            currentCharts = data.charts || [];

            // 默认全部显示
            visibleChartIdx = new Set(currentCharts.map((_, i) => i));
            visibleSeries = {};
            currentCharts.forEach((c, i) => {
                const names = c.series && c.series.length
                    ? c.series.map(s => s.name)
                    : [c.item];
                visibleSeries[i] = new Set(names);
            });

            renderSidebarTree(currentCharts);
            renderCharts(currentCharts);
            renderOverview(data);
            enableStatsNav();
        } catch (err) {
            showAlert("请求失败: " + err.message, "danger");
        } finally {
            uploadBtn.classList.remove("btn-loading");
            uploadBtn.innerHTML = '<i class="bi bi-upload"></i> 上传并分析';
        }
    });

    // ====================================================================
    //  2. 左侧筛选树 (一级=图表名, 二级=图表内测试项目)
    // ====================================================================

    // 去掉项目名末尾的 "_数字及其后续" (例: "LED_颜色检测_34" -> "LED_颜色检测")
    // 规则: 仅当下划线后紧跟数字时才裁剪
    function stripItemSuffix(name) {
        if (name == null) return "";
        return String(name).replace(/_\d+.*$/, "").trim();
    }

    // 提取图表的主项目名: 若 series 多个, 取共同前缀 (去掉 _数字 后缀)
    // 例: ["LED_颜色检测_34", "LED_颜色检测_35"] -> "LED_颜色检测"
    function Analyzer_extractMainName(chart) {
        const series = chart.series && chart.series.length ? chart.series : [];
        if (series.length === 0) return stripItemSuffix(chart.item || "未知");
        if (series.length === 1) return stripItemSuffix(series[0].name);
        // 取所有 series.name 的共同字符串前缀, 然后去掉末尾的 _数字 分隔符
        const names = series.map(s => stripItemSuffix(s.name || ""));
        let prefix = names[0];
        for (let i = 1; i < names.length; i++) {
            let j = 0;
            while (j < prefix.length && j < names[i].length && prefix[j] === names[i][j]) j++;
            prefix = prefix.substring(0, j);
        }
        // 去掉末尾的 _数字 / _[ 等
        prefix = prefix.replace(/[_\[（(]+$/, "").trim();
        return prefix || names[0];
    }

    // 格式化 [下限,上限] 字符串
    function Analyzer_formatSpecBracket(sl, su) {
        const hasL = sl != null && sl !== "";
        const hasU = su != null && su !== "";
        if (!hasL && !hasU) return "";
        const lv = hasL ? sl : "-";
        const uv = hasU ? su : "-";
        return `[${lv}, ${uv}]`;
    }

    // 格式化数值: 保留原始精度, 不自动缩减位数
    function Analyzer_formatValue(v) {
        if (v == null) return "-";
        if (typeof v === "number") {
            // 保持数字的完整精度, 不缩减尾零
            return String(v);
        }
        return String(v);
    }

    function renderSidebarTree(charts) {
        if (!charts.length) {
            sidebarFilter.innerHTML = '<div class="text-muted small p-3">暂无项目…</div>';
            return;
        }
        sidebarFilter.innerHTML = "";

        const ul = document.createElement("ul");
        ul.className = "list-group list-group-flush rounded-0";

        charts.forEach((chart, cIdx) => {
            // 一级: 图表名 (按规则: 主项目名(项数)/[下限,上限] 换行显示)
            const li = document.createElement("li");
            li.className = "list-group-item p-0 border-0";

            // 解析 chart.item 生成 主项目名(项数) + [下限,上限] 两行
            // chart.item 格式: "下限 ~ 上限 (N项)" / "≤ 上限 (N项)" / "≥ 下限 (N项)"
            // 主项目名: 从 series 中提取共同前缀 (去掉 _数字 后缀)
            const seriesCount = (chart.series && chart.series.length) || 1;
            const mainName = Analyzer_extractMainName(chart);
            const specBracket = Analyzer_formatSpecBracket(
                chart.spec_lower_str != null ? chart.spec_lower_str : Analyzer_formatValue(chart.spec_lower),
                chart.spec_upper_str != null ? chart.spec_upper_str : Analyzer_formatValue(chart.spec_upper)
            );

            // 一级行
            const header = document.createElement("label");
            header.className = "d-flex align-items-start gap-2 py-2 px-3 w-100 user-select-none sidebar-chart-header";
            header.setAttribute("data-chart-idx", cIdx);
            header.style.cursor = "pointer";
            const cbChart = document.createElement("input");
            cbChart.type = "checkbox";
            cbChart.className = "form-check-input m-0 mt-1";
            cbChart.checked = visibleChartIdx.has(cIdx);
            cbChart.addEventListener("change", () => toggleChartVisible(cIdx, cbChart.checked));
            // 序号徽章 (需求1: 筛选菜单加序号, 与图表序号一致)
            const idxBadge = document.createElement("span");
            idxBadge.className = "badge bg-secondary rounded-pill sidebar-idx-badge";
            idxBadge.textContent = String(cIdx + 1);
            idxBadge.style.minWidth = "1.6rem";
            idxBadge.style.textAlign = "center";
            // title 容器: 单行布局 (项目名 + 规格区间 在同一行)
            const titleWrap = document.createElement("div");
            titleWrap.className = "flex-grow-1 min-w-0 d-flex align-items-center gap-2 flex-wrap";
            const titleMain = document.createElement("span");
            titleMain.className = "fw-semibold small text-truncate";
            const titleMainText = mainName + (seriesCount > 1 ? ` (${seriesCount}项)` : "");
            titleMain.title = titleMainText + (specBracket ? ` ${specBracket}` : "");
            titleMain.textContent = titleMainText;
            titleWrap.appendChild(titleMain);
            if (specBracket) {
                const titleSpec = document.createElement("span");
                titleSpec.className = "text-muted small";
                titleSpec.style.fontSize = "0.72rem";
                titleSpec.style.whiteSpace = "nowrap";
                titleSpec.textContent = specBracket;
                titleWrap.appendChild(titleSpec);
            }
            // 展开/折叠图标 (默认折叠: 朝右)
            const carets = document.createElement("i");
            carets.className = "bi bi-chevron-down text-muted toggle-caret mt-1";
            carets.style.transition = "transform .15s";
            carets.style.transform = "rotate(-90deg)";

            header.appendChild(cbChart);
            header.appendChild(idxBadge);
            header.appendChild(titleWrap);
            header.appendChild(carets);

            // 二级: 子项目列表 (默认折叠) - 子项目完整显示, 不裁剪后缀
            const seriesList = chart.series && chart.series.length
                ? chart.series
                : [{ name: chart.item, item_note: (chart.meta && chart.meta.item_note) || "" }];
            const childUl = document.createElement("ul");
            childUl.className = "list-group list-group-flush ps-4 pe-2 pb-2 collapse";

            seriesList.forEach(s => {
                const sname = s.name;
                const sNote = s.item_note || "";
                const cli = document.createElement("li");
                cli.className = "list-group-item p-0 border-0 sidebar-sub-item";
                const clbl = document.createElement("label");
                clbl.className = "d-flex align-items-center gap-2 py-1 w-100 user-select-none sidebar-sub-item-label";
                clbl.style.cursor = "pointer";
                const cb = document.createElement("input");
                cb.type = "checkbox";
                cb.className = "form-check-input m-0";
                cb.checked = visibleSeries[cIdx] && visibleSeries[cIdx].has(sname);
                cb.addEventListener("change", () => toggleSeriesVisible(cIdx, sname, cb.checked));
                const sn = document.createElement("span");
                sn.className = "flex-grow-1 small text-truncate";
                sn.style.maxWidth = "210px";
                // 子项目完整显示, 不裁剪后缀
                sn.title = sname + (sNote ? ` (${sNote})` : "");
                sn.textContent = sname;
                // 项目备注 (非空时显示为浅色小字)
                if (sNote) {
                    const noteSpan = document.createElement("span");
                    noteSpan.className = "text-muted";
                    noteSpan.style.fontSize = "0.7rem";
                    noteSpan.style.whiteSpace = "nowrap";
                    noteSpan.textContent = ` · ${sNote}`;
                    sn.appendChild(noteSpan);
                }
                // 跳转定位图标
                const go = document.createElement("i");
                go.className = "bi bi-geo-alt text-muted";
                go.title = "定位到图表";
                go.style.cursor = "pointer";
                go.addEventListener("click", (ev) => {
                    ev.preventDefault();
                    ev.stopPropagation();
                    scrollToChart(cIdx);
                });
                clbl.appendChild(cb);
                clbl.appendChild(sn);
                clbl.appendChild(go);
                cli.appendChild(clbl);
                childUl.appendChild(cli);
            });

            // 一级点击非 checkbox/caret 区域 = 折叠/展开; 点击 checkbox 仅切换显示
            carets.addEventListener("click", (ev) => {
                ev.preventDefault();
                ev.stopPropagation();
                toggleCollapse(childUl, carets);
            });
            titleWrap.addEventListener("click", (ev) => {
                if (ev.target === cbChart) return;
                ev.preventDefault();
                ev.stopPropagation();
                toggleCollapse(childUl, carets);
            });

            li.appendChild(header);
            li.appendChild(childUl);
            ul.appendChild(li);
        });

        sidebarFilter.appendChild(ul);

        // 展开/折叠全部 (默认折叠: allExpanded=false, 第一次点击=展开全部)
        let allExpanded = false;
        expandAllBtn.onclick = () => {
            allExpanded = !allExpanded;
            sidebarFilter.querySelectorAll(".collapse").forEach(el => {
                if (allExpanded) el.classList.add("show");
                else el.classList.remove("show");
            });
            sidebarFilter.querySelectorAll(".toggle-caret").forEach(i => {
                i.style.transform = allExpanded ? "rotate(0deg)" : "rotate(-90deg)";
            });
        };

        // 需求4: 全选 / 全取消 (同时勾选一级图表 + 二级 series)
        selectAllBtn.onclick = () => {
            currentCharts.forEach((chart, cIdx) => {
                visibleChartIdx.add(cIdx);
                const snames = chart.series && chart.series.length
                    ? chart.series.map(s => s.name)
                    : [chart.item];
                if (!visibleSeries[cIdx]) visibleSeries[cIdx] = new Set();
                snames.forEach(n => visibleSeries[cIdx].add(n));
            });
            // 同步 UI checkbox
            sidebarFilter.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = true; });
            applyVisibility();
        };
        unselectAllBtn.onclick = () => {
            currentCharts.forEach((chart, cIdx) => {
                visibleChartIdx.delete(cIdx);
                const snames = chart.series && chart.series.length
                    ? chart.series.map(s => s.name)
                    : [chart.item];
                if (visibleSeries[cIdx]) snames.forEach(n => visibleSeries[cIdx].delete(n));
            });
            sidebarFilter.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = false; });
            applyVisibility();
        };
    }

    function toggleCollapse(el, caretIcon) {
        const show = el.classList.toggle("show");
        if (caretIcon) caretIcon.style.transform = show ? "rotate(0deg)" : "rotate(-90deg)";
    }

    function toggleChartVisible(cIdx, checked) {
        if (checked) visibleChartIdx.add(cIdx);
        else visibleChartIdx.delete(cIdx);
        applyVisibility();
    }

    function toggleSeriesVisible(cIdx, sname, checked) {
        if (!visibleSeries[cIdx]) visibleSeries[cIdx] = new Set();
        if (checked) visibleSeries[cIdx].add(sname);
        else visibleSeries[cIdx].delete(sname);
        // 若该图表无任何 series 勾选, 则也隐藏该图表
        const anyChild = visibleSeries[cIdx] && visibleSeries[cIdx].size > 0;
        if (anyChild) visibleChartIdx.add(cIdx);
        else visibleChartIdx.delete(cIdx);
        applyVisibility();
    }

    function applyVisibility() {
        // 1. wrapper 显示/隐藏
        wrapperElements.forEach((w, i) => {
            if (!w) return;
            if (visibleChartIdx.has(i)) {
                w.classList.remove("d-none");
            } else {
                w.classList.add("d-none");
            }
        });
        // 2. echarts 内显示/隐藏 series (按 legend selected)
        currentCharts.forEach((c, i) => {
            const inst = chartInstances[i];
            if (!inst) return;
            const names = c.series && c.series.length
                ? c.series.map(s => s.name)
                : [c.item];
            const selected = {};
            names.forEach(n => { selected[n] = !!(visibleSeries[i] && visibleSeries[i].has(n)); });
            inst.dispatchAction({ type: "legendUnSelect" });
            inst.dispatchAction({ type: "legendSelect", selected });
            // 仅当图表本身配置了 legend 时才更新 selected, 避免给无图例的图表凭空创建图例
            const instOpt = inst.getOption();
            if (instOpt.legend && instOpt.legend.length && instOpt.legend[0].show !== false) {
                inst.setOption({ legend: { selected: selected } });
            }
        });
    }

    // ====================================================================
    //  滚动联动: 页面滚动时侧边栏筛选同步高亮 + 自动滑到可视区
    // ====================================================================
    function setupScrollSync() {
        // 清理旧的 Observer
        if (scrollIntersectionObserver) {
            scrollIntersectionObserver.disconnect();
            scrollIntersectionObserver = null;
        }
        scrollActiveChartIdx = -1;

        // 清除所有旧的高亮样式
        document.querySelectorAll(".sidebar-chart-active").forEach(el => el.classList.remove("sidebar-chart-active"));
        document.querySelectorAll(".sidebar-chart-header-active").forEach(el => el.classList.remove("sidebar-chart-header-active"));

        if (!wrapperElements.length) return;

        // 用 IntersectionObserver 观察每个 chart-wrapper, 找出当前视口顶部的图表
        const io = new IntersectionObserver((entries) => {
            // 计算进入视口且最靠近视口顶部的图表
            const viewportH = window.innerHeight || document.documentElement.clientHeight;
            let bestIdx = -1;
            let bestDist = Infinity;
            wrapperElements.forEach((w, i) => {
                if (!w || w.classList.contains("d-none")) return;
                const r = w.getBoundingClientRect();
                // 图表中间点位置与视口中间的距离
                const mid = r.top + r.height / 2;
                const viewMid = viewportH / 2;
                const d = Math.abs(mid - viewMid);
                // 只考虑有部分进入视口的
                if (r.bottom >= 0 && r.top <= viewportH) {
                    if (d < bestDist) {
                        bestDist = d;
                        bestIdx = i;
                    }
                }
            });

            if (bestIdx >= 0 && bestIdx !== scrollActiveChartIdx) {
                scrollActiveChartIdx = bestIdx;
                highlightSidebarItem(bestIdx);
                scrollSidebarToItem(bestIdx);
            }
        }, {
            // 以视口为基准, 把观察阈值设置得低一些, 保证进入视口就响应
            root: null,
            threshold: [0, 0.05, 0.1, 0.25, 0.5, 0.75, 1.0],
        });

        wrapperElements.forEach(w => { if (w) io.observe(w); });
        scrollIntersectionObserver = io;
    }

    // 按 data-chart-idx 查找侧边栏一级 header (避免 querySelector 属性选择器在某些环境下的兼容问题)
    function findSidebarHeader(cIdx) {
        var need = String(cIdx);
        var headers = sidebarFilter.querySelectorAll(".sidebar-chart-header");
        for (var i = 0; i < headers.length; i++) {
            if (headers[i].getAttribute("data-chart-idx") === need) return headers[i];
        }
        return null;
    }

    // 高亮侧边栏指定 chartIdx 的一级项 (蓝色背景或边框)
    function highlightSidebarItem(cIdx) {
        var target = findSidebarHeader(cIdx);
        var actives = sidebarFilter.querySelectorAll(".sidebar-chart-header-active");
        for (var i = 0; i < actives.length; i++) {
            if (actives[i] !== target) actives[i].classList.remove("sidebar-chart-header-active");
        }
        if (target) target.classList.add("sidebar-chart-header-active");
    }

    // 将侧边栏指定 chartIdx 项滚动到可视区 (不超出侧边栏容器范围)
    var sidebarScrollRaf = 0;
    function scrollSidebarToItem(cIdx) {
        var target = findSidebarHeader(cIdx);
        if (!target) return;
        // 用 requestAnimationFrame 节流, 避免连续滚动抖动
        if (sidebarScrollRaf) cancelAnimationFrame(sidebarScrollRaf);
        sidebarScrollRaf = requestAnimationFrame(function () {
            var parentRect = sidebarFilter.getBoundingClientRect();
            var targetRect = target.getBoundingClientRect();
            var outside = (targetRect.top < parentRect.top + 4) || (targetRect.bottom > parentRect.bottom - 4);
            if (outside) {
                // 使用 getBoundingClientRect 差值 + 当前 scrollTop 得到目标位置 (offsetTop 会因多层相对定位父元素失真)
                var relativeTop = targetRect.top - parentRect.top + sidebarFilter.scrollTop;
                var wantTop = relativeTop - sidebarFilter.clientHeight / 2 + targetRect.height / 2;
                sidebarFilter.scrollTo({ top: Math.max(0, wantTop), behavior: "smooth" });
            }
        });
    }

    function scrollToChart(cIdx) {
        const w = wrapperElements[cIdx];
        if (!w) return;
        // 先保证它显示
        if (!visibleChartIdx.has(cIdx)) {
            visibleChartIdx.add(cIdx);
            applyVisibility();
        }
        w.scrollIntoView({ behavior: "smooth", block: "start" });
        w.animate(
            [{ boxShadow: "0 0 0 3px #0d6efd88" }, { boxShadow: "none" }],
            { duration: 1500, easing: "ease-out" }
        );
    }

    // ====================================================================
    //  3. 渲染 ECharts 曲线图 (功能4)
    // ====================================================================
    function renderCharts(charts) {
        chartInstances.forEach(c => c && c.dispose());
        chartInstances = [];
        wrapperElements = [];
        chartsContainer.innerHTML = "";

        if (!charts.length) {
            chartsContainer.innerHTML = '<div class="empty-state"><i class="bi bi-inbox"></i><p>暂无图表数据</p></div>';
            chartsSection.classList.remove("d-none");
            return;
        }

        chartsSection.classList.remove("d-none");

        charts.forEach(chart => {
            const meta = chart.meta || {};
            const s = chart.stats || {};
            const noteText = meta.item_note ? `<span class="text-muted small">${meta.item_note}</span>` : "";
            const paramText = meta.param_variable ? `<span class="badge bg-light text-dark border">${meta.param_variable}</span>` : "";
            let passBadge = "";
            if (s.pass_rate != null) {
                const cls = s.pass_rate >= 95 ? "bg-success" : (s.pass_rate >= 80 ? "bg-warning" : "bg-danger");
                passBadge = `<span class="badge ${cls}">合格率 ${s.pass_rate}%</span>`;
            }
            const specText = (chart.spec_upper != null || chart.spec_lower != null)
                ? `<span class="badge bg-light text-dark border">规格 ${chart.spec_lower_str != null ? chart.spec_lower_str : (chart.spec_lower ?? "-")} ~ ${chart.spec_upper_str != null ? chart.spec_upper_str : (chart.spec_upper ?? "-")}</span>`
                : "";

            const wrapper = document.createElement("div");
            wrapper.className = "chart-wrapper mb-3";
            // 需求1: 图表标题显示筛选栏一致的一级公共"项目名称", 不要把 series 全部列出
            // 同时加序号 (与筛选菜单一致, 1-based)
            const chartIdx1 = chartsContainer.children.length + 1;
            const chartTitleText = Analyzer_extractMainName(chart);
            wrapper.innerHTML = `
                <div class="chart-box">
                    <div class="chart-header flex-wrap">
                        <div>
                            <span class="badge bg-secondary rounded-pill chart-idx-badge me-2">${chartIdx1}</span>
                            <strong><i class="bi bi-graph-up"></i> ${chartTitleText}</strong>
                            ${noteText} ${paramText} ${specText}
                        </div>
                        <div class="d-flex align-items-center gap-1">
                            ${passBadge}
                            <span class="badge bg-secondary ms-1">${chart.values ? chart.values.length : 0} 点</span>
                            <button class="btn btn-sm btn-outline-secondary ms-1 zoom-btn" title="放大查看">
                                <i class="bi bi-zoom-in"></i>
                            </button>
                        </div>
                    </div>
                    <div class="chart-body"></div>
                </div>
            `;
            chartsContainer.appendChild(wrapper);
            wrapperElements.push(wrapper);

            const chartDom = wrapper.querySelector(".chart-body");
            const instance = echarts.init(chartDom);
            instance.setOption(buildOption(chart, chartDom.clientWidth));
            chartInstances.push(instance);

            // 放大按钮: 点击后在 modal 中渲染同一图表的大尺寸版本
            const chartIdx = chartsContainer.children.length - 1;
            const zoomBtn = wrapper.querySelector(".zoom-btn");
            if (zoomBtn) {
                zoomBtn.addEventListener("click", () => openZoomModal(chartIdx));
            }
        });

        applyVisibility();
        setupScrollSync();  // 初始化滚动联动 (页面滚动时侧边栏筛选同步)
    }

    // ====================================================================
    //  图表放大 Modal
    // ====================================================================
    let _currentZoomChartIdx = -1;
    let _zoomResizeHandler = null;

    function openZoomModal(chartIdx) {
        const chart = currentCharts[chartIdx];
        if (!chart) return;

        // 保存当前图表索引 (事件回调中使用)
        _currentZoomChartIdx = chartIdx;

        // 标题: #序号 主项目名 + 规格区间
        const mainName = Analyzer_extractMainName(chart);
        const seriesCount = (chart.series && chart.series.length) || 1;
        const specBracket = Analyzer_formatSpecBracket(
            chart.spec_lower_str != null ? chart.spec_lower_str : Analyzer_formatValue(chart.spec_lower),
            chart.spec_upper_str != null ? chart.spec_upper_str : Analyzer_formatValue(chart.spec_upper)
        );
        const titleParts = [`#${chartIdx + 1}`, mainName];
        if (seriesCount > 1) titleParts.push(`(${seriesCount}项)`);
        if (specBracket) titleParts.push(specBracket);
        zoomModalTitle.textContent = titleParts.join(" ");

        // 销毁旧实例, 防止残留尺寸污染
        if (zoomChartInstance) {
            zoomChartInstance.dispose();
            zoomChartInstance = null;
        }
        // 清空 DOM 容器内容 (残留 canvas 会干扰尺寸计算)
        zoomChartBody.innerHTML = "";

        // 显示 Modal
        const modal = bootstrap.Modal.getOrCreateInstance(chartZoomModal);
        modal.show();

        // 下载按钮绑定当前图表
        zoomDownloadBtn.onclick = () => {
            if (zoomChartInstance) {
                const url = zoomChartInstance.getDataURL({
                    type: "png", pixelRatio: 2, backgroundColor: "#fff"
                });
                const a = document.createElement("a");
                a.href = url;
                a.download = `${mainName || "chart"}.png`;
                a.click();
            }
        };

        // 仅首次绑定 shown/hidden 与 resize 事件 (按经验: 避免多次 setTimeout 赌时机)
        if (!chartZoomModal._zoomBound) {
            chartZoomModal._zoomBound = true;

            // shown: Modal 动画结束, 容器真实尺寸已知 → 再 init ECharts
            chartZoomModal.addEventListener("shown.bs.modal", () => {
                renderZoomChart();
            });

            // hidden: 销毁实例 + 移除 resize 监听
            chartZoomModal.addEventListener("hidden.bs.modal", () => {
                if (zoomChartInstance) {
                    zoomChartInstance.dispose();
                    zoomChartInstance = null;
                }
                zoomChartBody.innerHTML = "";
                if (_zoomResizeHandler) {
                    window.removeEventListener("resize", _zoomResizeHandler);
                    _zoomResizeHandler = null;
                }
            });
        } else {
            // 已绑定事件, 在 shown 事件中自动 renderZoomChart (通过 _currentZoomChartIdx 获取当前图表)
        }
    }

    // 在 Modal shown 之后调用 (容器有真实尺寸)
    function renderZoomChart() {
        const chart = currentCharts[_currentZoomChartIdx];
        if (!chart) return;

        // 先清旧实例
        if (zoomChartInstance) {
            zoomChartInstance.dispose();
            zoomChartInstance = null;
        }

        // 取真实容器尺寸 (debug 用)
        const bodyW = zoomChartBody.clientWidth;
        const bodyH = zoomChartBody.clientHeight;

        // init
        zoomChartInstance = echarts.init(zoomChartBody);

        // 构造 option
        const opt = buildOption(chart, bodyW);
        // 底部 dataZoom slider + inside 滚轮缩放
        opt.dataZoom = [
            { type: "inside", start: 0, end: 100 },
            { type: "slider", start: 0, end: 100, height: 24, bottom: 10 },
        ];
        // 全屏版: grid 给底部 slider 留足够空间; 左右留边距防止标签被裁
        opt.grid = {
            ...(opt.grid || {}),
            left: 70,
            right: 50,
            top: 60,
            bottom: 80,
            containLabel: false,
        };
        // 全屏版: x/y axis name 使用稍大字号
        if (opt.xAxis) {
            opt.xAxis = Array.isArray(opt.xAxis) ? opt.xAxis.map(fullscreenAxisStyle) : fullscreenAxisStyle(opt.xAxis);
        }
        if (opt.yAxis) {
            opt.yAxis = Array.isArray(opt.yAxis) ? opt.yAxis.map(fullscreenAxisStyle) : fullscreenAxisStyle(opt.yAxis);
        }

        zoomChartInstance.setOption(opt, true);  // true = 不合并, 完全覆盖
        // init 后立即 resize (确保按真实容器尺寸重排)
        zoomChartInstance.resize();

        // 监听 window resize → 同步放大图表尺寸
        if (_zoomResizeHandler) window.removeEventListener("resize", _zoomResizeHandler);
        _zoomResizeHandler = () => {
            if (zoomChartInstance) zoomChartInstance.resize();
        };
        window.addEventListener("resize", _zoomResizeHandler);
    }

    // 全屏版坐标轴样式: 字号稍大, 增加间距
    function fullscreenAxisStyle(axis) {
        if (!axis) return axis;
        return {
            ...axis,
            nameTextStyle: { ...(axis.nameTextStyle || {}), fontSize: Math.max(14, (axis.nameTextStyle && axis.nameTextStyle.fontSize) || 13) },
            axisLabel: { ...(axis.axisLabel || {}), fontSize: 13 },
        };
    }

    // 计算友好的 Y 轴边界, 确保规格上下限恰好对齐刻度位置
    function computeNiceAxisBounds(yMin, yMax, specLower, specUpper) {
        if (yMin == null || yMax == null || yMax <= yMin) {
            return { min: yMin, max: yMax, interval: null };
        }

        // 计算浮点值的小数位数
        function decimals(v) {
            if (v == null || isNaN(v)) return 0;
            const s = String(v);
            const dot = s.indexOf('.');
            return dot >= 0 ? s.length - dot - 1 : 0;
        }

        // 四舍五入到指定小数位
        function roundTo(v, d) {
            const p = Math.pow(10, d);
            return Math.round(v * p) / p;
        }

        // 计算所有涉及值的最大精度
        const allVals = [yMin, yMax, specLower, specUpper].filter(v => v != null && !isNaN(v));
        // 确保 maxDec 至少为 3, 避免小数值被圆整为 0
        const maxDec = Math.max(3, allVals.length ? Math.max(...allVals.map(decimals)) : 3);

        // 将边界值圆整到最大精度
        let niceMin = roundTo(yMin, maxDec);
        let niceMax = roundTo(yMax, maxDec);

        // 如果有规格上下限, 确保它们恰好落在刻度位置
        if (specLower != null && specUpper != null && specUpper > specLower) {
            // 关键修复: 先圆整 specLower/specUpper 消除浮点误差
            // 例: 1.515 - 1.485 = 0.030000000000000042 (IEEE 754 误差)
            //     roundTo 后 -> 0.03, 后续 interval 计算才正确
            const sLower = roundTo(specLower, maxDec);
            const sUpper = roundTo(specUpper, maxDec);
            const specRange = roundTo(sUpper - sLower, maxDec);
            // 选择 3-6 个刻度间隔, 让规格值恰好对齐
            const candidates = [3, 4, 5, 6];
            let bestInterval = null;
            for (const n of candidates) {
                const interval = roundTo(specRange / n, maxDec);
                // 关键验证: interval × n 必须等于 specRange (圆整后)
                // 否则规格上下限无法恰好对齐刻度
                const check = roundTo(interval * n, maxDec);
                if (Math.abs(check - specRange) > 1e-12) continue;
                const id = decimals(interval);
                if (id <= maxDec + 2) {
                    bestInterval = interval;
                    break;
                }
            }
            // 如果所有候选都无法整除, 使用 specRange 本身作为 interval (仅2个刻度: 上下限)
            if (bestInterval == null) {
                bestInterval = specRange;
            }

            // 调整 niceMin 使 specLower 恰好落在刻度
            // specLower = niceMin + k * bestInterval => niceMin = specLower - k * bestInterval
            // 选择 k 使 niceMin 尽量接近原始 yMin
            const kMin = Math.round((sLower - niceMin) / bestInterval);
            niceMin = roundTo(sLower - kMin * bestInterval, maxDec);

            // 调整 niceMax 使 specUpper 也恰好落在刻度
            const kMax = Math.round((niceMax - sUpper) / bestInterval);
            niceMax = roundTo(sUpper + kMax * bestInterval, maxDec);

            // 确保包含所有数据
            if (niceMin > roundTo(yMin, maxDec)) {
                const k2 = Math.ceil((niceMin - roundTo(yMin, maxDec)) / bestInterval);
                niceMin = roundTo(niceMin - k2 * bestInterval, maxDec);
            }
            if (niceMax < roundTo(yMax, maxDec)) {
                const k2 = Math.ceil((roundTo(yMax, maxDec) - niceMax) / bestInterval);
                niceMax = roundTo(niceMax + k2 * bestInterval, maxDec);
            }

            // 关键修复: 向上下各扩展一个 interval, 使规格线不紧贴 grid 边缘
            // 否则 markLine 恰好在 yMin/yMax 位置会被 grid 边界裁剪, 视觉上偏移
            niceMin = roundTo(niceMin - bestInterval, maxDec);
            niceMax = roundTo(niceMax + bestInterval, maxDec);

            return { min: niceMin, max: niceMax, interval: bestInterval };
        }

        // 仅有单边规格或无规格: 使用简单的友好间隔
        const range = niceMax - niceMin;
        const rawInterval = range / 5;
        const magnitude = Math.pow(10, Math.floor(Math.log10(rawInterval)));
        const residual = rawInterval / magnitude;
        let niceInterval;
        if (residual <= 1.5) niceInterval = magnitude;
        else if (residual <= 3) niceInterval = 2 * magnitude;
        else if (residual <= 7) niceInterval = 5 * magnitude;
        else niceInterval = 10 * magnitude;

        // 调整边界为间隔的整数倍
        niceMin = roundTo(Math.floor(niceMin / niceInterval) * niceInterval, maxDec);
        niceMax = roundTo(Math.ceil(niceMax / niceInterval) * niceInterval, maxDec);

        // 关键修复: 单边规格时也确保 markLine 不紧贴 grid 边缘
        // 如果 specUpper 恰好等于 niceMax, 向上扩展一个 interval
        if (specUpper != null && specLower == null) {
            if (Math.abs(niceMax - specUpper) < 1e-12) {
                niceMax = roundTo(niceMax + niceInterval, maxDec);
            }
        }
        // 如果 specLower 恰好等于 niceMin, 向下扩展一个 interval
        if (specLower != null && specUpper == null) {
            if (Math.abs(niceMin - specLower) < 1e-12) {
                niceMin = roundTo(niceMin - niceInterval, maxDec);
            }
        }

        return { min: niceMin, max: niceMax, interval: roundTo(niceInterval, maxDec) };
    }

    // 基于字符串精度圆整规格值, 避免 IEEE 754 浮点误差
    // 例: roundSpecValue(1.5149999999999998, "1.515") -> 1.515
    function roundSpecValue(v, str) {
        if (v == null) return null;
        if (isNaN(v)) return null;
        if (str != null) {
            const dot = str.indexOf('.');
            const decimals = dot >= 0 ? str.length - dot - 1 : 0;
            const p = Math.pow(10, decimals);
            return Math.round(v * p) / p;
        }
        return v;
    }

    function buildOption(chart, containerWidth) {
        // 构建参考线:
        //  - 规格上限/下限: 绿色实线 (只画 1 次, 放到第一个 series 上)
        //  - 均值: 每个 series 一条, 粉红色虚线, 标签含系列名前缀
        const SPEC_COLOR = "#198754";
        const PALETTE = [
            "#5470c6", "#91cc75", "#fac858", "#ee6666",
            "#73c0de", "#3ba272", "#fc8452", "#9a60b4", "#ea7ccc"
        ];

        const specMarkLine = [];
        const suStr = chart.spec_upper_str != null ? chart.spec_upper_str : Analyzer_formatValue(chart.spec_upper);
        const slStr = chart.spec_lower_str != null ? chart.spec_lower_str : Analyzer_formatValue(chart.spec_lower);
        const suVal = roundSpecValue(chart.spec_upper, chart.spec_upper_str);
        const slVal = roundSpecValue(chart.spec_lower, chart.spec_lower_str);
        // 获取 X 轴数据长度, 用于两点 coord 格式 markLine
        const xDataLen = Math.max(1, (chart.x_axis || []).length - 1);
        // 规格上下限是否重合 (如 1/1), 重合时只画一条线并合并标签, 避免重叠
        const specSame = (suVal != null && slVal != null && Math.abs(suVal - slVal) < 1e-9);
        if (specSame) {
            specMarkLine.push([
                {coord: [0, suVal], name: "规格", lineStyle: { color: SPEC_COLOR, type: "solid", width: 2 },
                 label: { formatter: "规格 " + suStr, position: "insideStartTop", color: SPEC_COLOR, fontWeight: "bold", distance: 8 }},
                {coord: [xDataLen, suVal]}
            ]);
        } else {
            if (suVal != null) {
                specMarkLine.push([
                    {coord: [0, suVal], name: "规格上限", lineStyle: { color: SPEC_COLOR, type: "solid", width: 2 },
                     label: { formatter: "上限 " + suStr, position: "insideStartTop", color: SPEC_COLOR, fontWeight: "bold", distance: 6 }},
                    {coord: [xDataLen, suVal]}
                ]);
            }
            if (slVal != null) {
                specMarkLine.push([
                    {coord: [0, slVal], name: "规格下限", lineStyle: { color: SPEC_COLOR, type: "solid", width: 2 },
                     label: { formatter: "下限 " + slStr, position: "insideEndBottom", color: SPEC_COLOR, fontWeight: "bold", distance: 6 }},
                    {coord: [xDataLen, slVal]}
                ]);
            }
        }

        const meta = chart.meta || {};
        const yName = meta.param_variable || "测试值";

        const seriesList = chart.series && chart.series.length
            ? chart.series
            : [{ name: chart.item, values: chart.values, stats: chart.stats }];

        const allValues = seriesList.flatMap(s => s.values || []);
        // 过滤 null/NaN 后计算最值, 避免 Math.min/max 遇到 NaN 返回 NaN
        const numericValues = allValues.filter(v => v != null && !isNaN(v));
        const dataMin = numericValues.length ? Math.min(...numericValues) : null;
        const dataMax = numericValues.length ? Math.max(...numericValues) : null;
        const lowerCandidates = [dataMin, slVal].filter(v => v != null && !isNaN(v));
        const upperCandidates = [dataMax, suVal].filter(v => v != null && !isNaN(v));
        const yMin = lowerCandidates.length ? Math.min(...lowerCandidates) : undefined;
        const yMax = upperCandidates.length ? Math.max(...upperCandidates) : undefined;

        // 检测数字量/枚举图表: 规格与数据均为整数, 且取值范围小 (如 0/1, 1/1 等)
        function isInteger(v) {
            return Math.abs(v - Math.round(v)) < 1e-9;
        }
        const allInts = numericValues.length > 0 && numericValues.every(isInteger);
        const specInts = (suVal == null || isInteger(suVal)) && (slVal == null || isInteger(slVal));
        const valRange = (dataMax != null && dataMin != null) ? (dataMax - dataMin) : 0;
        const isDigital = allInts && specInts && valRange <= 10;

        // Y 轴边界与间隔
        let yMinAdj, yMaxAdj, yInterval, yMinFn, yMaxFn;
        if (isDigital) {
            // 数字量: 使用整数边界 (ECharts 对非整数 min/max 会重算刻度间隔), 上下各留 1 个整数余量
            const intMin = Math.min(...[dataMin, slVal].filter(v => v != null));
            const intMax = Math.max(...[dataMax, suVal].filter(v => v != null));
            yMinAdj = Math.max(0, Math.floor(intMin) - 1);
            yMaxAdj = Math.ceil(intMax) + 1;
            yInterval = 1;
            // 使用函数形式确保 ECharts 不自动覆盖 (所有数据相同时 ECharts 可能忽略静态 min/max)
            yMinFn = function () { return yMinAdj; };
            yMaxFn = function () { return yMaxAdj; };
        } else {
            // 使用友好边界算法, 确保规格上下限恰好对齐 Y 轴刻度
            const niceBounds = computeNiceAxisBounds(yMin, yMax, slVal, suVal);
            yMinAdj = niceBounds.min;
            yMaxAdj = niceBounds.max;
            yInterval = niceBounds.interval;
        }

        // 计算规格值的小数位数, 均值精度与规格保持一致
        const _suDec = suStr ? (suStr.split(".")[1] || "").length : 0;
        const _slDec = slStr ? (slStr.split(".")[1] || "").length : 0;
        const _specDec = Math.max(3, _suDec, _slDec);

        const seriesArr = seriesList.map((s, idx) => {
            const color = PALETTE[idx % PALETTE.length];
            const perSeriesMarkLine = [];
            if (idx === 0 && specMarkLine.length) {
                perSeriesMarkLine.push(...specMarkLine);
            }
            if (s.stats && s.stats.mean !== undefined) {
                const meanStr = s.stats.mean.toFixed(_specDec);
                const labelText = seriesList.length > 1
                    ? `${s.name} 均值 ${meanStr}`
                    : "均值 " + meanStr;
                // 判断均值是否与规格线重合/极近, 若是则将均值标签放到底部避免与规格标签重叠
                const nearUpper = suVal != null && Math.abs(s.stats.mean - suVal) < 1e-6;
                const nearLower = slVal != null && Math.abs(s.stats.mean - slVal) < 1e-6;
                const nearSpec = nearUpper || nearLower;
                let meanPos, meanOffset;
                if (isDigital) {
                    // 数字量图表: 均值标签放底部, 与顶部的规格标签分开
                    meanPos = "insideEndBottom";
                    meanOffset = [0, 0];
                } else if (nearSpec) {
                    meanPos = "insideEndBottom";
                    meanOffset = [0, 0];
                } else {
                    meanPos = (idx % 2 === 0) ? "insideEndTop" : "insideStartTop";
                    meanOffset = [0, 0];
                }
                perSeriesMarkLine.push([
                    {coord: [0, s.stats.mean], name: `${s.name}均值`, lineStyle: { color: "#e91e63", type: "dashed", width: 2 },
                     label: { formatter: labelText, position: meanPos, color: "#e91e63", fontWeight: "bold", distance: 8, offset: meanOffset }},
                    {coord: [xDataLen, s.stats.mean]}
                ]);
            }
            return {
                name: s.name,
                type: "line",
                data: s.values,
                // 数字量图表禁用平滑曲线, 避免显示不存在的中间过渡值
                smooth: !isDigital,
                symbol: "circle",
                symbolSize: isDigital ? 4 : 5,
                lineStyle: { width: 2, color: color },
                itemStyle: { color: color },
                areaStyle: (!isDigital && seriesList.length === 1) ? { opacity: 0.1, color: color } : undefined,
                markLine: perSeriesMarkLine.length ? { data: perSeriesMarkLine, symbol: "none" } : undefined,
            };
        });

        const seriesNames = seriesList.map(s => s.name);
        const selected = {};
        seriesNames.forEach(n => { selected[n] = true; });

        // 图例项目超过 3 行时, 隐藏多余的项目, 保留最多 3 行, 超标的优先显示
        let legendData = null;   // null = 不过滤(显示全部); [] = 全部隐藏; [...] = 仅显示这些
        let legendRowHeight = 0; // 图例单行高度 (px), 用于限制总高度不超过 3 行
        if (seriesList.length > 1 && containerWidth > 0) {
            // 估算每行可容纳的图例项数 (平均每项约 110px: 色块25 + 文字 + 间距, 名称较长取偏大值)
            const avgItemWidth = 110;
            const itemsPerRow = Math.max(1, Math.floor(containerWidth / avgItemWidth));
            const maxDisplay = itemsPerRow * 3;   // 最多保留 3 行
            // 单行图例高度 ≈ itemHeight(14) + itemGap(8) + 文字行高(16) ≈ 38px
            legendRowHeight = 38;
            if (seriesList.length > maxDisplay) {
                // 判断项目是否"超标": 合格率 < 100 或 (CPK 存在且 <= 1)
                const isBad = (s) => {
                    const pr = s.stats && s.stats.pass_rate;
                    const cpk = s.stats && s.stats.cpk;
                    const passOk = (pr === 100 || pr >= 100);
                    const cpkOk = (cpk == null || cpk > 1);
                    return !(passOk && cpkOk);
                };
                const badItems = seriesList.filter(isBad);
                const goodItems = seriesList.filter(s => !isBad(s));
                // 超标项目优先显示, 剩余位置用良好项目填充
                const displayList = [];
                for (const s of badItems) {
                    if (displayList.length >= maxDisplay) break;
                    displayList.push(s);
                }
                for (const s of goodItems) {
                    if (displayList.length >= maxDisplay) break;
                    displayList.push(s);
                }
                legendData = displayList.map(s => s.name);
            }
        }
        // 图例是否显示: 多系列且(不过滤 或 过滤后仍有项目)
        const showLegend = seriesList.length > 1 && !(legendData !== null && legendData.length === 0);

        const option = {
            tooltip: {
                trigger: "axis",
                formatter: function (params) {
                    if (!params || !params.length) return "";
                    const idx = params[0].dataIndex;
                    const snList = chart.sn_list || [];
                    const sn = (snList[idx] !== undefined && snList[idx] !== "")
                        ? snList[idx]
                        : (chart.x_axis[idx] || "");
                    let html = `SN: ${sn}<br/>`;
                    params.forEach(p => {
                        const val = (p.value !== null && p.value !== undefined) ? p.value : "-";
                        html += `${p.marker} ${p.seriesName}: ${val}<br/>`;
                    });
                    return html;
                },
                valueFormatter: function (v) {
                    return Analyzer_formatValue(v);
                },
            },
            grid: {
                left: 80, right: 80, bottom: 65,
                top: showLegend
                    ? (legendData ? legendRowHeight * 3 + 10 : 40)   // 过滤场景下图例固定3行, 预留空间
                    : 25,
            },
            xAxis: {
                type: "category",
                data: chart.x_axis,
                axisLabel: { rotate: 30, fontSize: 10 },
                name: "时间",
                nameLocation: "middle",
                nameGap: 35,
            },
            yAxis: {
                type: "value",
                name: yName,
                min: yMinFn || yMinAdj,
                max: yMaxFn || yMaxAdj,
                interval: yInterval != null ? yInterval : undefined,
                minInterval: isDigital ? 1 : undefined,
                axisLabel: {
                    formatter: function (v) {
                        // 数字量图表只显示整数刻度标签
                        if (isDigital && Math.abs(v - Math.round(v)) > 1e-9) return "";
                        return Analyzer_formatValue(v);
                    },
                },
            },
            dataZoom: [
                { type: "inside", start: 0, end: 100 },
                { type: "slider", start: 0, end: 100, height: 18, bottom: 8 },
            ],
            series: seriesArr,
            toolbox: {
                feature: {
                    saveAsImage: { title: "保存图片" },
                    dataZoom: { title: { zoom: "缩放", back: "还原" } },
                },
                right: 10,
            },
        };
        if (showLegend) {
            const legendCfg = { top: 0, selected: selected, itemGap: 8 };
            if (legendData) {
                legendCfg.data = legendData;
                // 过滤场景下固定图例高度为 3 行, 超出部分隐藏, 避免遮挡曲线
                legendCfg.height = legendRowHeight * 3;
            }
            option.legend = legendCfg;
        }
        return option;
    }

    // ====================================================================
    //  4. 数据概览统计卡片
    // ====================================================================
    function renderOverview(data) {
        overviewSection.classList.remove("d-none");
        fileBadge.textContent = `${data.filename} · ${data.product_category} · ${data.total_rows} 条`;

        const pc = data.product_count || 0;
        const aic = data.analysis_item_count || 0;
        const cc = data.chart_count || 0;
        const tr = data.total_rows || 0;

        const cards = [
            { label: "产品数量", value: pc, unit: "台", icon: "bi-box-seam", cls: "bg-primary" },
            { label: "分析项目数量", value: aic, unit: "项", icon: "bi-list-check", cls: "bg-info" },
            { label: "图表数量", value: cc, unit: "张", icon: "bi-graph-up", cls: "bg-success" },
            { label: "记录总数", value: tr, unit: "条", icon: "bi-table", cls: "bg-secondary" },
        ];

        const cardsHtml = cards.map(c => `
            <div class="col-sm-6 col-lg-3">
                <div class="d-flex align-items-center gap-3 p-3 bg-white border rounded shadow-sm h-100">
                    <div class="d-inline-flex align-items-center justify-content-center text-white rounded-circle ${c.cls}"
                         style="width:52px;height:52px;">
                        <i class="bi ${c.icon} fs-4"></i>
                    </div>
                    <div>
                        <div class="text-muted small">${c.label}</div>
                        <div class="fs-3 fw-bold lh-1 mt-1">${c.value}</div>
                        <div class="text-muted small">${c.unit}</div>
                    </div>
                </div>
            </div>
        `).join("");

        // 公共信息: Program Name / Model 提升到 navbar 和 card-header 显示, 其余留在公共信息卡片
        const ci = data.common_info || {};
        const ciKeys = Object.keys(ci);
        // 优先 Program Name (旧格式), 回退 Model (新格式)
        const PROGRAM_NAME_KEY = ci["Program Name"] ? "Program Name" : (ci["Model"] ? "Model" : "Program Name");
        const programNameValue = ci[PROGRAM_NAME_KEY] || "";

        // 1. 填充 navbar 中的 Program Name (使用 border 容器, 与公共信息卡片样式一致)
        const navProgramName = document.getElementById("navProgramName");
        if (navProgramName) {
            if (programNameValue) {
                navProgramName.innerHTML = `
                    <div class="border rounded p-2 h-100 bg-light">
                        <div class="text-muted" style="font-size:0.75rem;">${PROGRAM_NAME_KEY}</div>
                        <div class="fw-semibold text-truncate" title="${programNameValue}">${programNameValue}</div>
                    </div>
                `;
            } else {
                navProgramName.innerHTML = "";
            }
        }

        // 2. 填充 card-header 中的 Program Name (label + value)
        const headerProgramName = document.getElementById("headerProgramName");
        if (headerProgramName) {
            if (programNameValue) {
                headerProgramName.classList.remove("d-none");
                headerProgramName.innerHTML = `
                    <span class="text-muted" style="font-size:0.75rem;">${PROGRAM_NAME_KEY}</span>
                    <span class="fw-semibold text-truncate ms-1" title="${programNameValue}">${programNameValue}</span>
                `;
            } else {
                headerProgramName.classList.add("d-none");
                headerProgramName.innerHTML = "";
            }
        }

        // 3. 其余公共信息 (工单号/测试员/工装编号等) 仍在公共信息卡片
        const otherKeys = ciKeys.filter(k => k !== PROGRAM_NAME_KEY);
        let commonHtml = "";
        if (otherKeys.length) {
            commonHtml = `
                <div class="p-3 bg-white border rounded shadow-sm">
                    <div class="text-muted small mb-2"><i class="bi bi-info-circle"></i> 公共信息</div>
                    <div class="row g-2">
                        ${otherKeys.map(k => `
                            <div class="col-md-3 col-sm-6">
                                <div class="border rounded p-2 h-100 bg-light">
                                    <div class="text-muted" style="font-size:0.75rem;">${k}</div>
                                    <div class="fw-semibold text-truncate" title="${ci[k]}">${ci[k]}</div>
                                </div>
                            </div>
                        `).join("")}
                    </div>
                </div>
            `;
        }

        statCards.innerHTML = cardsHtml;
        const commonInfoSection = document.getElementById("commonInfoSection");
        if (commonInfoSection) {
            commonInfoSection.innerHTML = commonHtml;
        }
    }

    // ====================================================================
    //  5. 视图切换: 网格 / 列表
    // ====================================================================
    gridBtn.addEventListener("click", () => {
        chartsContainer.classList.add("view-grid");
        chartsContainer.classList.remove("view-list");
        gridBtn.classList.add("active");
        listBtn.classList.remove("active");
        resizeCharts();
    });

    listBtn.addEventListener("click", () => {
        chartsContainer.classList.add("view-list");
        chartsContainer.classList.remove("view-grid");
        listBtn.classList.add("active");
        gridBtn.classList.remove("active");
        resizeCharts();
    });

    function resizeCharts() {
        chartInstances.forEach(c => c && c.resize());
    }
    window.addEventListener("resize", resizeCharts);

    // ====================================================================
    //  6. 产品配置管理 (功能6)
    // ====================================================================
    const configModal = document.getElementById("configModal");
    const configTableBody = document.getElementById("configTableBody");
    const configForm = document.getElementById("configForm");
    const loadDefaultBtn = document.getElementById("loadDefaultBtn");

    configModal && configModal.addEventListener("show.bs.modal", loadConfigList);

    async function loadConfigList() {
        try {
            const resp = await fetch(api("/api/configs"));
            const data = await resp.json();
            const configs = data.configs || [];
            if (!configs.length) {
                configTableBody.innerHTML = '<tr><td colspan="3" class="text-center text-muted">暂无保存的配置</td></tr>';
                return;
            }
            configTableBody.innerHTML = configs.map(c => `
                <tr>
                    <td>${c.name}</td>
                    <td>${c.test_items.join(", ") || "-"}</td>
                    <td>
                        <button class="btn btn-sm btn-outline-primary me-1" onclick="loadConfigToForm('${c.name}')">
                            <i class="bi bi-pencil"></i>
                        </button>
                        <button class="btn btn-sm btn-outline-danger" onclick="deleteConfig('${c.name}')">
                            <i class="bi bi-trash"></i>
                        </button>
                    </td>
                </tr>
            `).join("");
        } catch (err) {
            configTableBody.innerHTML = '<tr><td colspan="3" class="text-danger">加载失败</td></tr>';
        }
    }

    window.loadConfigToForm = async function (name) {
        const resp = await fetch(api(`/api/config/${encodeURIComponent(name)}`));
        const cfg = await resp.json();
        document.getElementById("cfgName").value = cfg.category_name || "";
        document.getElementById("cfgSheet").value = cfg.excel_config?.sheet_name ?? 0;
        document.getElementById("cfgSkipTop").value = cfg.excel_config?.skip_top_rows ?? 0;
        document.getElementById("cfgHeaderRows").value = cfg.excel_config?.header_rows ?? 5;
        document.getElementById("cfgFixedHeaderRow").value = cfg.excel_config?.fixed_columns_header_row ?? -1;
        const fx = cfg.excel_config?.fixed_columns || {};
        document.getElementById("col_sequence").value = fx.sequence || "序列号";
        document.getElementById("col_product_model").value = fx.product_model || "产品型号";
        document.getElementById("col_work_order").value = fx.work_order || "";
        document.getElementById("col_tester").value = fx.tester || "";
        document.getElementById("col_fixture").value = fx.fixture || "";
        document.getElementById("col_result").value = fx.result || "";
        document.getElementById("col_time").value = fx.time || "";
        const mr = cfg.excel_config?.meta_row_mapping || {};
        document.getElementById("row_item_name").value = mr.item_name ?? 0;
        document.getElementById("row_item_note").value = mr.item_note ?? 1;
        document.getElementById("row_param_variable").value = mr.param_variable ?? 2;
        document.getElementById("row_spec_upper").value = mr.spec_upper ?? 3;
        document.getElementById("row_spec_lower").value = mr.spec_lower ?? 4;
        document.getElementById("cfgItems").value = (cfg.test_items || []).join(", ");
        document.getElementById("cfgSpecInfinity").value = (cfg.excel_config?.spec_infinity_values || []).join(", ");
        document.getElementById("cfgTitle").value = cfg.chart_config?.title || "测试分析曲线图";

        const tab = new bootstrap.Tab(document.querySelector('[data-bs-target="#tab-edit"]'));
        tab.show();
    };

    window.deleteConfig = async function (name) {
        if (!confirm(`确定删除配置 "${name}" 吗？`)) return;
        const resp = await fetch(api(`/api/config/${encodeURIComponent(name)}`), { method: "DELETE" });
        if (resp.ok) {
            showAlert("配置已删除", "success");
            loadConfigList();
            location.reload();
        }
    };

    loadDefaultBtn && loadDefaultBtn.addEventListener("click", () => {
        document.getElementById("cfgSheet").value = 0;
        document.getElementById("cfgSkipTop").value = 0;
        document.getElementById("cfgHeaderRows").value = 5;
        document.getElementById("cfgFixedHeaderRow").value = -1;
        document.getElementById("col_sequence").value = "序列号";
        document.getElementById("col_product_model").value = "产品型号";
        document.getElementById("col_work_order").value = "";
        document.getElementById("col_tester").value = "";
        document.getElementById("col_fixture").value = "";
        document.getElementById("col_result").value = "";
        document.getElementById("col_time").value = "";
        document.getElementById("row_item_name").value = 0;
        document.getElementById("row_item_note").value = 1;
        document.getElementById("row_param_variable").value = 2;
        document.getElementById("row_spec_upper").value = 3;
        document.getElementById("row_spec_lower").value = 4;
        document.getElementById("cfgSpecInfinity").value = "";
        document.getElementById("cfgTitle").value = "测试分析曲线图";
    });

    configForm && configForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        const name = document.getElementById("cfgName").value.trim();
        if (!name) return;
        const payload = {
            category_name: name,
            excel_config: {
                sheet_name: document.getElementById("cfgSheet").value,
                skip_top_rows: parseInt(document.getElementById("cfgSkipTop").value) || 0,
                header_rows: parseInt(document.getElementById("cfgHeaderRows").value) || 5,
                fixed_columns_header_row: parseInt(document.getElementById("cfgFixedHeaderRow").value),
                fixed_columns: {
                    sequence: document.getElementById("col_sequence").value || "序列号",
                    product_model: document.getElementById("col_product_model").value || "产品型号",
                    work_order: document.getElementById("col_work_order").value,
                    tester: document.getElementById("col_tester").value,
                    fixture: document.getElementById("col_fixture").value,
                    result: document.getElementById("col_result").value,
                    time: document.getElementById("col_time").value,
                },
                meta_row_mapping: {
                    item_name: parseInt(document.getElementById("row_item_name").value) || 0,
                    item_note: parseInt(document.getElementById("row_item_note").value) || 1,
                    param_variable: parseInt(document.getElementById("row_param_variable").value) || 2,
                    spec_upper: parseInt(document.getElementById("row_spec_upper").value) || 3,
                    spec_lower: parseInt(document.getElementById("row_spec_lower").value) || 4,
                },
                spec_infinity_values: document.getElementById("cfgSpecInfinity").value
                    .split(",").map(s => s.trim()).filter(Boolean)
                    .map(s => isNaN(Number(s)) ? s : Number(s)),
            },
            test_items: document.getElementById("cfgItems").value.split(",").map(s => s.trim()).filter(Boolean),
            chart_config: {
                title: document.getElementById("cfgTitle").value || "测试分析曲线图",
                y_axis_name: "测试值",
            },
        };
        const resp = await fetch(api("/api/config"), {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
        });
        if (resp.ok) {
            showAlert("配置已保存", "success");
            loadConfigList();
            setTimeout(() => location.reload(), 600);
        } else {
            const err = await resp.json().catch(() => ({}));
            showAlert("保存失败: " + (err.error || resp.statusText), "danger");
        }
    });

    // ====================================================================
    //  7. 工具: 状态提示
    // ====================================================================
    function showAlert(msg, type = "info") {
        statusAlert.className = `alert alert-${type} mb-3 fade show`;
        statusAlert.classList.remove("d-none");
        statusAlert.textContent = msg;
        window.scrollTo({ top: 0, behavior: "smooth" });
        setTimeout(() => {
            statusAlert.classList.add("d-none");
        }, 4500);
    }

    // 启动后默认选中网格视图
    chartsContainer.classList.add("view-grid");

    // ====================================================================
    //  8. 左侧筛选栏水平拉伸 (拖拽分隔条调整宽度)
    // ====================================================================
    const sidebarResizer = document.getElementById("sidebarResizer");
    const sidebarAside = document.querySelector(".sidebar-aside");
    if (sidebarResizer && sidebarAside) {
        let isDragging = false;
        let startX = 0;
        let startWidth = 0;

        sidebarResizer.addEventListener("mousedown", (e) => {
            isDragging = true;
            startX = e.clientX;
            startWidth = sidebarAside.offsetWidth;
            sidebarResizer.classList.add("dragging");
            document.body.style.cursor = "col-resize";
            document.body.style.userSelect = "none";
            e.preventDefault();
        });

        document.addEventListener("mousemove", (e) => {
            if (!isDragging) return;
            const delta = e.clientX - startX;
            let newWidth = startWidth + delta;
            const minWidth = 240;
            const maxWidth = Math.min(window.innerWidth - 400, 1000);
            if (newWidth < minWidth) newWidth = minWidth;
            if (newWidth > maxWidth) newWidth = maxWidth;
            sidebarAside.style.width = newWidth + "px";
        });

        document.addEventListener("mouseup", () => {
            if (!isDragging) return;
            isDragging = false;
            sidebarResizer.classList.remove("dragging");
            document.body.style.cursor = "";
            document.body.style.userSelect = "";
            // 拖拽结束后重绘所有图表 (尺寸可能变化)
            chartInstances.forEach(c => c && c.resize());
        });

        // 双击重置为默认宽度
        sidebarResizer.addEventListener("dblclick", () => {
            sidebarAside.style.width = "320px";
            chartInstances.forEach(c => c && c.resize());
        });
    }
});
