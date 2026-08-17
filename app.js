// State
let DASHBOARD_DATA = null;
let ACTIVE_WP = 'executive';
let ACTIVE_SUB = 'mdr';

// Chart instances
let progressChartInstance = null;
let delayChartInstance = null;
let wpProgressChartInstance = null;

const WP_NAMES = {
    executive: 'Executive Overview across All Work Packages',
    wp1_topside_excl: 'WP-1 Topside Excl. Structure (`00210 EDSR...xlsm`)',
    wp1_topside_structure: 'WP-1 Topside Structure (`Z1F Topsides...xlsx`)',
    wp1_jacket: 'WP-1 Jacket (`Z1F Jacket...xlsx`)',
    wp2_pipeline: 'WP-2 Pipeline (`Att.2. WP2...xlsx`)'
};

const WP_SHORT_NAMES = {
    executive: 'All Work Packages Combined',
    wp1_topside_excl: 'WP-1 Topside Excl. Structure',
    wp1_topside_structure: 'WP-1 Topside Structure',
    wp1_jacket: 'WP-1 Jacket',
    wp2_pipeline: 'WP-2 Pipeline'
};

document.addEventListener('DOMContentLoaded', () => {
    initClock();
    setupEventListeners();
    fetchDashboardData();
});

function initClock() {
    const badge = document.getElementById('clockBadge');
    const update = () => {
        const now = new Date();
        badge.textContent = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    };
    update();
    setInterval(update, 1000);
}

function setupEventListeners() {
    // WP Tab switching
    document.querySelectorAll('.wp-tab').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const wp = e.currentTarget.getAttribute('data-wp');
            if (wp === ACTIVE_WP) return;
            document.querySelectorAll('.wp-tab').forEach(b => b.classList.remove('active'));
            e.currentTarget.classList.add('active');
            ACTIVE_WP = wp;
            populateDisciplineFilter();
            renderDashboard();
        });
    });

    // Sub-tab switching
    document.querySelectorAll('.sub-tab').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const sub = e.currentTarget.getAttribute('data-sub');
            if (sub === ACTIVE_SUB) return;
            document.querySelectorAll('.sub-tab').forEach(b => b.classList.remove('active'));
            e.currentTarget.classList.add('active');
            ACTIVE_SUB = sub;
            renderTable();
        });
    });

    // Filters (Attach both input and change events for cross-browser dropdown support)
    ['searchInput', 'disciplineSelect', 'milestoneSelect', 'statusSelect'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', () => renderTable());
            el.addEventListener('change', () => renderTable());
        }
    });

    document.getElementById('resetFiltersBtn').addEventListener('click', () => {
        document.getElementById('searchInput').value = '';
        document.getElementById('disciplineSelect').value = 'ALL';
        document.getElementById('milestoneSelect').value = 'ALL';
        document.getElementById('statusSelect').value = 'ALL';
        renderTable();
    });

    // Modal triggers
    const modal = document.getElementById('uploadModal');
    document.getElementById('openUploadModalBtn').addEventListener('click', () => {
        modal.classList.add('active');
        updateModalFilePreviews();
    });

    ['closeUploadModalBtn', 'closeUploadModalBtnBottom'].forEach(id => {
        document.getElementById(id).addEventListener('click', () => {
            modal.classList.remove('active');
        });
    });

    // Drag and Drop
    const dropzone = document.getElementById('dropzone');
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
    });
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleMultiUpload(files);
        }
    });

    // File inputs
    document.getElementById('fileInput').addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleMultiUpload(e.target.files);
        }
    });

    document.querySelectorAll('.single-wp-upload').forEach(input => {
        input.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                const wpKey = e.target.getAttribute('data-wp');
                uploadFileToServer(e.target.files[0], wpKey);
            }
        });
    });
}

async function fetchDashboardData(isRefresh = false) {
    const statusText = document.getElementById('statusText');
    const statusPill = document.getElementById('statusPill');
    
    if (!isRefresh) statusText.textContent = 'Fetching cut-offs...';
    try {
        const response = await fetch('/api/data');
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        DASHBOARD_DATA = await response.json();
        
        statusText.textContent = `🟢 Connected | Today: ${DASHBOARD_DATA.today}`;
        document.getElementById('cutoffDateSpan').textContent = DASHBOARD_DATA.today;
        
        updateTabBadges();
        populateDisciplineFilter();
        renderDashboard();
        updateModalFilePreviews();
        
        if (isRefresh) {
            showToast('✅ Dashboard updated with latest weekly cut-off files!', 'success');
        }
    } catch (err) {
        console.error('Error fetching dashboard data:', err);
        statusText.textContent = '🔴 Server Error / Offline';
        showToast('Error connecting to local server: ' + err.message, 'error');
    }
}

function updateTabBadges() {
    if (!DASHBOARD_DATA) return;
    ['executive', 'wp1_topside_excl', 'wp1_topside_structure', 'wp1_jacket', 'wp2_pipeline'].forEach(key => {
        const badge = document.getElementById(`badge-${key}`);
        if (badge && DASHBOARD_DATA[key]) {
            const kpi = DASHBOARD_DATA[key].kpi;
            badge.textContent = `${kpi.delayed_count} Delayed`;
        }
    });
}

function updateModalFilePreviews() {
    if (!DASHBOARD_DATA) return;
    ['wp1_topside_excl', 'wp1_topside_structure', 'wp1_jacket', 'wp2_pipeline'].forEach(key => {
        const fileElem = document.getElementById(`file-${key}`);
        const statusElem = document.getElementById(`status-${key}`);
        if (fileElem && DASHBOARD_DATA[key]) {
            const fname = DASHBOARD_DATA[key].filename || 'No file mapped';
            fileElem.textContent = fname;
            statusElem.textContent = fname === 'Not Found' ? 'Missing' : 'Active';
            statusElem.style.background = fname === 'Not Found' ? 'var(--accent-rose-bg)' : 'var(--accent-emerald-bg)';
            statusElem.style.color = fname === 'Not Found' ? 'var(--accent-rose)' : 'var(--accent-emerald)';
        }
    });
}

function populateDisciplineFilter() {
    const select = document.getElementById('disciplineSelect');
    const currentVal = select.value;
    select.innerHTML = '<option value="ALL">All Disciplines</option>';
    
    if (!DASHBOARD_DATA || !DASHBOARD_DATA[ACTIVE_WP]) return;
    const discMap = DASHBOARD_DATA[ACTIVE_WP].discipline_summary || {};
    const disciplines = Object.keys(discMap).sort();
    
    disciplines.forEach(disc => {
        const opt = document.createElement('option');
        opt.value = disc;
        opt.textContent = `${disc} (${discMap[disc].total} docs)`;
        select.appendChild(opt);
    });
    
    if (disciplines.includes(currentVal)) {
        select.value = currentVal;
    }
}

function renderDashboard() {
    if (!DASHBOARD_DATA || !DASHBOARD_DATA[ACTIVE_WP]) return;
    
    const currentSection = DASHBOARD_DATA[ACTIVE_WP];
    const kpi = currentSection.kpi;
    
    // Banner update
    document.getElementById('sectionTitle').textContent = WP_NAMES[ACTIVE_WP];
    if (ACTIVE_WP === 'executive') {
        const fnames = currentSection.filenames || {};
        document.getElementById('activeFilename').textContent = Object.values(fnames).join(' | ');
    } else {
        document.getElementById('activeFilename').textContent = currentSection.filename || 'Unknown file';
    }
    
    // KPI Cards Dual Metrics (Gates vs MDR Docs Basis)
    document.getElementById('kpiTotalDocs').textContent = kpi.total_docs.toLocaleString();
    if (document.getElementById('kpiTotalGates')) {
        document.getElementById('kpiTotalGates').textContent = (kpi.total_docs * 3).toLocaleString();
    }
    if (document.getElementById('kpiTotalOnTime')) {
        document.getElementById('kpiTotalOnTimeGates').textContent = (kpi.on_time_gates_count || 0).toLocaleString();
        document.getElementById('kpiTotalOnTime').textContent = (kpi.on_time_count || 0).toLocaleString();
    }
    document.getElementById('kpiTotalDelayed').textContent = (kpi.delayed_count || 0).toLocaleString();
    if (document.getElementById('kpiTotalDelayedDocs')) {
        document.getElementById('kpiTotalDelayedDocs').textContent = (kpi.delayed_docs_count || 0).toLocaleString();
    }
    document.getElementById('kpiType1').textContent = (kpi.delayed_type1_count || 0).toLocaleString();
    if (document.getElementById('kpiType1Docs')) {
        document.getElementById('kpiType1Docs').textContent = (kpi.delayed_type1_docs_count || 0).toLocaleString();
    }
    document.getElementById('kpiType2').textContent = (kpi.delayed_type2_count || 0).toLocaleString();
    if (document.getElementById('kpiType2Docs')) {
        document.getElementById('kpiType2Docs').textContent = (kpi.delayed_type2_docs_count || 0).toLocaleString();
    }
    document.getElementById('kpiLookahead').textContent = (kpi.lookahead_count || 0).toLocaleString();
    if (document.getElementById('kpiLookaheadDocs')) {
        document.getElementById('kpiLookaheadDocs').textContent = (kpi.lookahead_docs_count || 0).toLocaleString();
    }
    document.getElementById('kpiSlippingFooter').textContent = `${kpi.lookahead_slipping_count} due gates slipping from Plan`;
    
    // Sub-tab counts (Unique MDR Documents Basis)
    document.getElementById('subCountMdr').textContent = kpi.total_docs.toLocaleString();
    if (document.getElementById('subCountDelay3_1') && document.getElementById('subCountDelay3_2')) {
        document.getElementById('subCountDelay3_1').textContent = `3.1: ${(kpi.delayed_type1_docs_count || 0).toLocaleString()}`;
        document.getElementById('subCountDelay3_2').textContent = `3.2: ${(kpi.delayed_type2_docs_count || 0).toLocaleString()}`;
    } else if (document.getElementById('subCountDelay')) {
        document.getElementById('subCountDelay').textContent = (kpi.delayed_docs_count || 0).toLocaleString();
    }
    document.getElementById('subCountLookahead').textContent = (kpi.lookahead_docs_count || 0).toLocaleString();
    
    // Status Boxes: PENDING FINAL APPROVAL and COMPLETE (Between Row 1 & Row 2)
    const pendingTotal = kpi.pending_approval_count !== undefined ? kpi.pending_approval_count : 0;
    const pendingEng = (kpi.pending_breakdown && kpi.pending_breakdown.ENG !== undefined) ? kpi.pending_breakdown.ENG : 0;
    const pendingTbe = (kpi.pending_breakdown && kpi.pending_breakdown.TBE !== undefined) ? kpi.pending_breakdown.TBE : 0;
    const pendingMr = (kpi.pending_breakdown && kpi.pending_breakdown.MR !== undefined) ? kpi.pending_breakdown.MR : 0;
    
    if (document.getElementById('boxPendingApprovalTotal')) {
        document.getElementById('boxPendingApprovalTotal').textContent = `${pendingTotal.toLocaleString()} Docs`;
        document.getElementById('boxPendingEng').textContent = pendingEng.toLocaleString();
        document.getElementById('boxPendingTbe').textContent = pendingTbe.toLocaleString();
        document.getElementById('boxPendingMr').textContent = pendingMr.toLocaleString();
    }
    
    const compTotal = kpi.complete_count !== undefined ? kpi.complete_count : 0;
    const compPct = kpi.complete_pct !== undefined ? kpi.complete_pct : (kpi.total_docs > 0 ? (compTotal / kpi.total_docs * 100).toFixed(1) : 0);
    const compEng = (kpi.complete_breakdown && kpi.complete_breakdown.ENG !== undefined) ? kpi.complete_breakdown.ENG : 0;
    const compTbe = (kpi.complete_breakdown && kpi.complete_breakdown.TBE !== undefined) ? kpi.complete_breakdown.TBE : 0;
    const compMr = (kpi.complete_breakdown && kpi.complete_breakdown.MR !== undefined) ? kpi.complete_breakdown.MR : 0;
    
    const compEngPct = (kpi.complete_pct_breakdown && kpi.complete_pct_breakdown.ENG !== undefined) ? kpi.complete_pct_breakdown.ENG : 0;
    const compTbePct = (kpi.complete_pct_breakdown && kpi.complete_pct_breakdown.TBE !== undefined) ? kpi.complete_pct_breakdown.TBE : 0;
    const compMrPct = (kpi.complete_pct_breakdown && kpi.complete_pct_breakdown.MR !== undefined) ? kpi.complete_pct_breakdown.MR : 0;
    
    if (document.getElementById('boxCompleteTotal')) {
        document.getElementById('boxCompleteTotal').textContent = `${compTotal.toLocaleString()} Docs (${compPct}%)`;
        document.getElementById('boxCompleteEng').textContent = `${compEng.toLocaleString()} (${compEngPct}%)`;
        document.getElementById('boxCompleteTbe').textContent = `${compTbe.toLocaleString()} (${compTbePct}%)`;
        document.getElementById('boxCompleteMr').textContent = `${compMr.toLocaleString()} (${compMrPct}%)`;
    }
    
    // Overall Progress & SPI updates
    const planPct = kpi.plan_progress_pct || 0.0;
    const actPct = kpi.actual_progress_pct || 0.0;
    const forePct = kpi.forecast_progress_pct || 0.0;
    const varPct = kpi.variance_pct || (actPct - planPct);
    const spiVal = kpi.spi || (planPct > 0 ? actPct / planPct : 1.00);
    
    if (document.getElementById('kpiPlanProgress')) {
        document.getElementById('kpiPlanProgress').textContent = `${planPct.toFixed(2)}%`;
        document.getElementById('kpiActualProgress').textContent = `${actPct.toFixed(2)}%`;
        document.getElementById('kpiForecastProgress').textContent = `${forePct.toFixed(2)}%`;
        
        const varElem = document.getElementById('kpiProgressVariance');
        varElem.textContent = `${varPct >= 0 ? '+' : ''}${varPct.toFixed(2)}% variance`;
        if (varPct < 0) {
            varElem.classList.add('negative');
        } else {
            varElem.classList.remove('negative');
        }
        
        document.getElementById('kpiSpi').textContent = spiVal.toFixed(2);
        const spiBadge = document.getElementById('kpiSpiBadge');
        spiBadge.classList.remove('ahead', 'behind');
        if (spiVal >= 1.01) {
            spiBadge.textContent = '🚀 Ahead of Schedule';
            spiBadge.classList.add('ahead');
        } else if (spiVal >= 0.98) {
            spiBadge.textContent = '⚡ On Schedule';
            spiBadge.classList.add('ahead');
        } else {
            spiBadge.textContent = '⚠️ Behind Schedule';
            spiBadge.classList.add('behind');
        }
    }
    
    // Executive First Tab Breakdown Table & Chart visibility
    const execContainer = document.getElementById('executiveBreakdownContainer');
    if (execContainer) {
        if (ACTIVE_WP === 'executive') {
            execContainer.style.display = 'grid';
            renderExecutiveBreakdown();
        } else {
            execContainer.style.display = 'none';
        }
    }
    
    renderCharts();
    renderTable();
}

function renderExecutiveBreakdown() {
    if (!DASHBOARD_DATA || !DASHBOARD_DATA.executive) return;
    const wpSummary = DASHBOARD_DATA.executive.wp_summary || [];
    const tbody = document.getElementById('wpSummaryTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';
    
    // Render the 4 work package rows
    wpSummary.forEach(item => {
        const kpi = item.kpi || {};
        const tr = document.createElement('tr');
        const planP = (kpi.plan_progress_pct || 0.0).toFixed(2);
        const foreP = (kpi.forecast_progress_pct || 0.0).toFixed(2);
        const actP = (kpi.actual_progress_pct || 0.0).toFixed(2);
        const varP = (kpi.variance_pct || 0.0).toFixed(2);
        const spi = (kpi.spi || 1.00).toFixed(2);
        
        let statusBadge = '<span class="spi-badge ahead">On Time</span>';
        if (kpi.spi >= 1.01) statusBadge = '<span class="spi-badge ahead">Ahead</span>';
        else if (kpi.spi < 0.98) statusBadge = '<span class="spi-badge behind">Behind</span>';
        
        tr.innerHTML = `
            <td style="font-weight: 600; color: #38bdf8;">${item.wp_name}</td>
            <td>${(kpi.total_docs || 0).toLocaleString()}</td>
            <td>${planP}%</td>
            <td>${foreP}%</td>
            <td style="font-weight: 700; color: #10b981;">${actP}%</td>
            <td style="color: ${varP >= 0 ? '#10b981' : '#f43f5e'}; font-weight: 600;">${varP >= 0 ? '+' : ''}${varP}%</td>
            <td style="font-weight: 700;">${spi}</td>
            <td>${statusBadge}</td>
        `;
        tr.style.cursor = 'pointer';
        tr.addEventListener('click', () => {
            const tabBtn = document.querySelector(`.wp-tab[data-wp="${item.wp_key}"]`);
            if (tabBtn) tabBtn.click();
        });
        tbody.appendChild(tr);
    });
    
    // Add Total Executive Row
    const execKpi = DASHBOARD_DATA.executive.kpi || {};
    const trTotal = document.createElement('tr');
    trTotal.className = 'wp-row-total';
    const totalPlan = (execKpi.plan_progress_pct || 0.0).toFixed(2);
    const totalFore = (execKpi.forecast_progress_pct || 0.0).toFixed(2);
    const totalAct = (execKpi.actual_progress_pct || 0.0).toFixed(2);
    const totalVar = (execKpi.variance_pct || 0.0).toFixed(2);
    const totalSpi = (execKpi.spi || 1.00).toFixed(2);
    
    let totalBadge = '<span class="spi-badge ahead">On Time</span>';
    if (execKpi.spi >= 1.01) totalBadge = '<span class="spi-badge ahead">Ahead</span>';
    else if (execKpi.spi < 0.98) totalBadge = '<span class="spi-badge behind">Behind</span>';
    
    trTotal.innerHTML = `
        <td style="color: #ffffff;">⭐ TOTAL PROJECT EXECUTIVE SUMMARY</td>
        <td>${(execKpi.total_docs || 0).toLocaleString()}</td>
        <td>${totalPlan}%</td>
        <td>${totalFore}%</td>
        <td style="color: #10b981;">${totalAct}%</td>
        <td style="color: ${totalVar >= 0 ? '#10b981' : '#f43f5e'};">${totalVar >= 0 ? '+' : ''}${totalVar}%</td>
        <td>${totalSpi}</td>
        <td>${totalBadge}</td>
    `;
    tbody.appendChild(trTotal);
    
    renderExecutiveWpChart(wpSummary, execKpi);
}

function renderExecutiveWpChart(wpSummary, execKpi) {
    const canvas = document.getElementById('wpProgressChart');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (wpProgressChartInstance) wpProgressChartInstance.destroy();
    
    const labels = wpSummary.map(item => item.wp_name.replace('WP-1 ', '').replace('WP-2 ', ''));
    labels.push('Total Project');
    
    const planData = wpSummary.map(item => item.kpi.plan_progress_pct || 0.0);
    planData.push(execKpi.plan_progress_pct || 0.0);
    
    const foreData = wpSummary.map(item => item.kpi.forecast_progress_pct || 0.0);
    foreData.push(execKpi.forecast_progress_pct || 0.0);
    
    const actData = wpSummary.map(item => item.kpi.actual_progress_pct || 0.0);
    actData.push(execKpi.actual_progress_pct || 0.0);
    
    wpProgressChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Baseline Plan %',
                    data: planData,
                    backgroundColor: 'rgba(56, 189, 248, 0.75)',
                    borderColor: '#38bdf8',
                    borderWidth: 1,
                    borderRadius: 4
                },
                {
                    label: 'Forecast Expected %',
                    data: foreData,
                    backgroundColor: 'rgba(168, 85, 247, 0.75)',
                    borderColor: '#a855f7',
                    borderWidth: 1,
                    borderRadius: 4
                },
                {
                    label: 'Achieved Actual %',
                    data: actData,
                    backgroundColor: 'rgba(16, 185, 129, 0.9)',
                    borderColor: '#10b981',
                    borderWidth: 1,
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                    labels: { color: '#e2e8f0', font: { family: 'Outfit', weight: 600 } }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `${context.dataset.label}: ${context.raw.toFixed(2)}%`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: { color: '#94a3b8', font: { family: 'Outfit', weight: 600 } },
                    grid: { color: 'rgba(51, 65, 85, 0.3)' }
                },
                y: {
                    beginAtZero: true,
                    max: 100,
                    ticks: {
                        color: '#94a3b8',
                        callback: function(val) { return val + '%'; }
                    },
                    grid: { color: 'rgba(51, 65, 85, 0.3)' }
                }
            }
        }
    });
}

function renderCharts() {
    if (!DASHBOARD_DATA || !DASHBOARD_DATA[ACTIVE_WP]) return;
    const discMap = DASHBOARD_DATA[ACTIVE_WP].discipline_summary || {};
    const disciplines = Object.keys(discMap).sort();
    
    // Prepare Stacked Bar Chart data
    const completeCounts = disciplines.map(d => discMap[d].complete || 0);
    const pendingApprovalCounts = disciplines.map(d => discMap[d].pending_approval || 0);
    const ifaCounts = disciplines.map(d => discMap[d].ifa_sub || 0);
    const ifrCounts = disciplines.map(d => discMap[d].ifr_sub || 0);
    const notSubCounts = disciplines.map(d => discMap[d].not_submitted || 0);
    
    // 1. Progress Chart
    const ctxProg = document.getElementById('progressChart').getContext('2d');
    if (progressChartInstance) progressChartInstance.destroy();
    progressChartInstance = new Chart(ctxProg, {
        type: 'bar',
        data: {
            labels: disciplines.map(d => d.length > 18 ? d.substring(0, 18) + '...' : d),
            datasets: [
                { label: 'COMPLETE', data: completeCounts, backgroundColor: '#10b981', stack: 'Stack 0' },
                { label: 'PENDING FINAL APPROVAL', data: pendingApprovalCounts, backgroundColor: '#8b5cf6', stack: 'Stack 0' },
                { label: 'IFA Submitted', data: ifaCounts, backgroundColor: '#0284c7', stack: 'Stack 0' },
                { label: 'IFR Submitted', data: ifrCounts, backgroundColor: '#06b6d4', stack: 'Stack 0' },
                { label: 'Not Yet Submitted', data: notSubCounts, backgroundColor: '#64748b', stack: 'Stack 0' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: '#f8fafc', font: { family: 'Outfit', size: 12 } } },
                tooltip: { backgroundColor: '#0f172a', titleColor: '#06b6d4', bodyColor: '#f8fafc', borderColor: '#334155', borderWidth: 1 }
            },
            scales: {
                x: { stacked: true, grid: { color: 'rgba(51, 65, 85, 0.3)' }, ticks: { color: '#94a3b8' } },
                y: { stacked: true, grid: { color: 'rgba(51, 65, 85, 0.3)' }, ticks: { color: '#94a3b8' } }
            }
        }
    });

    // 2. Delay & Lookahead Chart
    const ctxDelay = document.getElementById('delayChart').getContext('2d');
    if (delayChartInstance) delayChartInstance.destroy();
    delayChartInstance = new Chart(ctxDelay, {
        type: 'bar',
        data: {
            labels: disciplines.map(d => d.length > 18 ? d.substring(0, 18) + '...' : d),
            datasets: [
                { label: 'Type 3.1 (Late)', data: disciplines.map(d => discMap[d].type3_1), backgroundColor: '#b45309' },
                { label: 'Type 3.2 (Overdue)', data: disciplines.map(d => discMap[d].type3_2), backgroundColor: '#ef4444' },
                { label: '14-Day Lookahead Risk', data: disciplines.map(d => discMap[d].lookahead || 0), backgroundColor: '#eab308' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: '#f8fafc', font: { family: 'Outfit', size: 12 } } },
                tooltip: { backgroundColor: '#0f172a', titleColor: '#ef4444', bodyColor: '#f8fafc', borderColor: '#334155', borderWidth: 1 }
            },
            scales: {
                x: { grid: { color: 'rgba(51, 65, 85, 0.3)' }, ticks: { color: '#94a3b8' } },
                y: { grid: { color: 'rgba(51, 65, 85, 0.3)' }, ticks: { color: '#94a3b8' } }
            }
        }
    });
}

function renderTable() {
    if (!DASHBOARD_DATA || !DASHBOARD_DATA[ACTIVE_WP]) return;
    const currentSection = DASHBOARD_DATA[ACTIVE_WP];
    const wrapper = document.getElementById('dataTableWrapper');
    
    // Filter parameters
    const searchVal = document.getElementById('searchInput').value.toLowerCase().trim();
    const discVal = document.getElementById('disciplineSelect').value;
    const milestoneVal = document.getElementById('milestoneSelect').value;
    const statusVal = document.getElementById('statusSelect').value;
    
    if (ACTIVE_SUB === 'mdr') {
        renderMdrTable(currentSection.docs, wrapper, searchVal, discVal, statusVal, currentSection.delayed_list, milestoneVal, currentSection.lookahead_list);
    } else if (ACTIVE_SUB === 'delay') {
        renderDelayTable(currentSection.delayed_list, wrapper, searchVal, discVal, milestoneVal, statusVal);
    } else if (ACTIVE_SUB === 'lookahead') {
        renderLookaheadTable(currentSection.lookahead_list, wrapper, searchVal, discVal, milestoneVal, statusVal);
    }
}

function renderMdrTable(docs, wrapper, searchVal, discVal, statusVal, delayedList = [], milestoneVal = 'ALL', lookaheadList = []) {
    const matchingDelays = milestoneVal === 'ALL' ? delayedList : delayedList.filter(x => x.milestone.includes(milestoneVal));
    const delayedDocNos = new Set(matchingDelays.map(x => x.doc_no));
    const type3_1DocNos = new Set(matchingDelays.filter(x => x.delay_type_code === '3.1').map(x => x.doc_no));
    const type3_2DocNos = new Set(matchingDelays.filter(x => x.delay_type_code === '3.2').map(x => x.doc_no));
    const critDocNos = new Set(matchingDelays.filter(x => x.urgency === 'critical').map(x => x.doc_no));
    const lookaheadDocNos = new Set((milestoneVal === 'ALL' ? lookaheadList : lookaheadList.filter(x => x.milestone.includes(milestoneVal))).map(x => x.doc_no));

    let filtered = docs.filter(d => {
        if (discVal !== 'ALL' && d.discipline !== discVal) return false;
        if (searchVal && !(`${d.doc_no} ${d.title} ${d.discipline} ${d.status}`.toLowerCase().includes(searchVal))) return false;
        if (milestoneVal !== 'ALL') {
            const cleanMs = milestoneVal.toLowerCase();
            const hasGateEntry = (
                (d.status && d.status.toLowerCase().includes(cleanMs)) ||
                delayedList.some(x => x.doc_no === d.doc_no && x.milestone.toLowerCase().includes(cleanMs)) ||
                lookaheadList.some(x => x.doc_no === d.doc_no && x.milestone.toLowerCase().includes(cleanMs)) ||
                (cleanMs.includes('ifr') && d.ifr_plan && d.ifr_plan !== '-') ||
                (cleanMs.includes('ifa') && d.ifa_plan && d.ifa_plan !== '-') ||
                (cleanMs.includes('afc sub') && d.afc_sub_plan && d.afc_sub_plan !== '-') ||
                (cleanMs.includes('afc app') && d.afc_app_plan && d.afc_app_plan !== '-') ||
                (cleanMs === 'afc' && d.afc_plan && d.afc_plan !== '-')
            );
            if (!hasGateEntry) return false;
        }
        if (statusVal === 'COMPLETE' && d.status !== 'COMPLETE') return false;
        if (statusVal === 'PENDING_APPROVAL' && d.status !== 'PENDING FINAL APPROVAL') return false;
        if (statusVal === 'IFA_SUB' && d.status !== 'IFA Submitted') return false;
        if (statusVal === 'IFR_SUB' && d.status !== 'IFR Submitted') return false;
        if (statusVal === 'NOT_SUB' && d.status !== 'Not Yet Submitted') return false;
        if (statusVal === 'ON_TIME' && delayedDocNos.has(d.doc_no)) return false;
        if (statusVal === 'DELAYED' && !delayedDocNos.has(d.doc_no)) return false;
        if (statusVal === 'TYPE3_1' && !type3_1DocNos.has(d.doc_no)) return false;
        if (statusVal === 'TYPE3_2' && !type3_2DocNos.has(d.doc_no)) return false;
        if (statusVal === 'CRITICAL' && !critDocNos.has(d.doc_no)) return false;
        if (statusVal === 'LOOKAHEAD' && !lookaheadDocNos.has(d.doc_no)) return false;
        return true;
    });

    if (filtered.length === 0) {
        wrapper.innerHTML = '<div class="loading-spinner">No deliverables matching active filters found.</div>';
        return;
    }

    let html = `
        <table>
            <thead>
                <tr>
                    <th>Doc Number</th>
                    <th>Deliverable Title</th>
                    <th>Discipline</th>
                    <th>Current Status</th>
                    <th>IFR Plan</th>
                    <th>IFR Forecast</th>
                    <th>IFR Actual</th>
                    <th>IFA Plan</th>
                    <th>IFA Forecast</th>
                    <th>IFA Actual</th>
                    <th>AFC Sub (10%)</th>
                    <th>AFC Sub Act</th>
                    <th>AFC App (20%)</th>
                    <th>AFC App Act (AP)</th>
                </tr>
            </thead>
            <tbody>
    `;

    // Limit render to first 250 items for ultra smooth UI performance if too many
    const displayDocs = filtered.slice(0, 250);
    displayDocs.forEach(d => {
        let badgeClass = 'badge-not-sub';
        let statusIcon = '';
        if (d.status === 'COMPLETE') {
            badgeClass = 'badge-complete';
            statusIcon = '✅ ';
        } else if (d.status === 'PENDING FINAL APPROVAL') {
            badgeClass = 'badge-pending-approval';
            statusIcon = '⏳ ';
        } else if (d.status === 'IFA Submitted') {
            badgeClass = 'badge-ifa-sub';
            statusIcon = '🔵 ';
        } else if (d.status === 'IFR Submitted') {
            badgeClass = 'badge-ifr-sub';
            statusIcon = '🔷 ';
        }
        
        const afcSubDate = d.afc_sub_date || d.afc_submit_date || '-';
        const afcSubPlan = d.afc_sub_plan || d.afc_plan || '-';
        const afcAppDate = d.afc_app_date || (d.status === 'COMPLETE' ? d.afc_submit_date : '-');
        const afcAppPlan = d.afc_app_plan || d.afc_plan || '-';
        
        html += `
            <tr>
                <td class="doc-no-cell">${d.doc_no}</td>
                <td class="doc-title-cell" title="${d.title}">${d.title}</td>
                <td><span class="chart-tag">${d.discipline}</span></td>
                <td><span class="badge-status ${badgeClass}">${statusIcon}${d.status}</span></td>
                <td>${d.ifr_plan || '-'}</td>
                <td>${d.ifr_forecast || '-'}</td>
                <td style="color: ${d.ifr_submit_date !== '-' && d.ifr_submit_date ? 'var(--accent-emerald)' : ''}">${d.ifr_submit_date || '-'}</td>
                <td>${d.ifa_plan || '-'}</td>
                <td>${d.ifa_forecast || '-'}</td>
                <td style="color: ${d.ifa_submit_date !== '-' && d.ifa_submit_date ? 'var(--accent-emerald)' : ''}">${d.ifa_submit_date || '-'}</td>
                <td>${afcSubPlan}</td>
                <td style="color: ${afcSubDate !== '-' && afcSubDate ? 'var(--accent-emerald)' : ''}">${afcSubDate}</td>
                <td>${afcAppPlan}</td>
                <td style="color: ${afcAppDate !== '-' && afcAppDate ? 'var(--accent-emerald)' : ''}; font-weight: ${afcAppDate !== '-' ? '700' : 'normal'}">${afcAppDate}</td>
            </tr>
        `;
    });

    html += `</tbody></table>`;
    if (filtered.length > 250) {
        html += `<div style="padding: 12px; text-align: center; color: var(--text-secondary); font-size: 13px;">Showing first 250 rows of ${filtered.length}. Use Search or Filter to refine.</div>`;
    }
    wrapper.innerHTML = html;
}

function renderDelayTable(delayedList, wrapper, searchVal, discVal, milestoneVal, statusVal) {
    if (statusVal === 'ON_TIME') {
        wrapper.innerHTML = '<div class="loading-spinner" style="color: var(--accent-emerald); font-weight: 600;">⚡ You selected the "On-Time / On-Schedule" filter. No delayed breaches exist for on-time items! Please switch to the "Master Document Register (MDR)" tab above to inspect your On-Time deliverables.</div>';
        return;
    }

    let filtered = delayedList.filter(d => {
        if (discVal !== 'ALL' && d.discipline !== discVal) return false;
        if (milestoneVal !== 'ALL' && !d.milestone.includes(milestoneVal)) return false;
        if (searchVal && !(`${d.doc_no} ${d.title} ${d.discipline}`.toLowerCase().includes(searchVal))) return false;
        if (statusVal === 'TYPE3_1' && d.delay_type_code !== '3.1') return false;
        if (statusVal === 'TYPE3_2' && d.delay_type_code !== '3.2') return false;
        if (statusVal === 'CRITICAL' && d.urgency !== 'critical') return false;
        return true;
    });

    if (filtered.length === 0) {
        wrapper.innerHTML = '<div class="loading-spinner" style="color: var(--accent-emerald);">✅ Great job! No overdue or delayed review gates match current criteria.</div>';
        return;
    }

    let html = `
        <table>
            <thead>
                <tr>
                    <th>Doc Number</th>
                    <th>Deliverable Title</th>
                    <th>Discipline</th>
                    <th>Gate Milestone</th>
                    <th>Breach Classification</th>
                    <th>Days Delayed</th>
                    <th>Urgent Severity</th>
                    <th>Baseline Plan</th>
                    <th>Baseline Forecast</th>
                    <th>Actual Submission</th>
                </tr>
            </thead>
            <tbody>
    `;

    const displayList = filtered.slice(0, 250);
    displayList.forEach(d => {
        const typeBadge = d.delay_type_code === '3.1' ? 'badge-type3_1' : 'badge-type3_2';
        const urgClass = `urgency-${d.urgency}`;
        const cleanMilestone = d.milestone.replace(' Gate', '');
        
        html += `
            <tr>
                <td class="doc-no-cell">${d.doc_no}</td>
                <td class="doc-title-cell" title="${d.title}">${d.title}</td>
                <td><span class="chart-tag">${d.discipline}</span></td>
                <td><strong>${cleanMilestone}</strong></td>
                <td><span class="badge-status ${typeBadge}">${d.delay_type}</span></td>
                <td style="font-weight: 800; font-size: 15px; color: var(--accent-rose);">${d.delay_days} Days</td>
                <td><span class="badge-urgency ${urgClass}">${d.urgency.toUpperCase()}</span></td>
                <td>${d.plan_date}</td>
                <td>${d.forecast_date}</td>
                <td style="font-weight: 600;">${d.submit_date}</td>
            </tr>
        `;
    });

    html += `</tbody></table>`;
    if (filtered.length > 250) {
        html += `<div style="padding: 12px; text-align: center; color: var(--text-secondary); font-size: 13px;">Showing top 250 delayed gates of ${filtered.length}.</div>`;
    }
    wrapper.innerHTML = html;
}

function renderLookaheadTable(lookaheadList, wrapper, searchVal, discVal, milestoneVal, statusVal) {
    if (statusVal === 'ON_TIME' || statusVal === 'DELAYED' || statusVal === 'TYPE3_1' || statusVal === 'TYPE3_2' || statusVal === 'CRITICAL') {
        wrapper.innerHTML = `<div class="loading-spinner" style="color: var(--accent-yellow); font-weight: 600;">⚡ You are currently on the 14-Day Lookahead tab, but selected a delay filter ("${statusVal}"). Please switch to the "Overdue & Delayed Table" tab above to view those delay items, or reset Breach Status to "All Statuses" / "14-Day Lookahead Warnings".</div>`;
        return;
    }

    let filtered = lookaheadList.filter(d => {
        if (discVal !== 'ALL' && d.discipline !== discVal) return false;
        if (milestoneVal !== 'ALL' && !d.milestone.includes(milestoneVal)) return false;
        if (searchVal && !(`${d.doc_no} ${d.title} ${d.discipline}`.toLowerCase().includes(searchVal))) return false;
        return true;
    });

    if (filtered.length === 0) {
        wrapper.innerHTML = '<div class="loading-spinner">No deliverables due within the upcoming 14 days lookahead window.</div>';
        return;
    }

    let html = `
        <table>
            <thead>
                <tr>
                    <th>Doc Number</th>
                    <th>Deliverable Title</th>
                    <th>Discipline</th>
                    <th>Upcoming Gate</th>
                    <th>Days Remaining</th>
                    <th>Slippage Warning</th>
                    <th>Baseline Plan</th>
                    <th>Forecast Due Date</th>
                </tr>
            </thead>
            <tbody>
    `;

    filtered.forEach(d => {
        const slipBadge = d.slipping 
            ? '<span class="badge-urgency urgency-critical">⚠️ SLIPPING FROM PLAN</span>' 
            : '<span class="badge-status badge-submitted">⚡ ON SCHEDULE</span>';
        const remColor = d.days_remaining <= 7 ? 'var(--accent-rose)' : 'var(--accent-amber)';
        
        html += `
            <tr>
                <td class="doc-no-cell">${d.doc_no}</td>
                <td class="doc-title-cell" title="${d.title}">${d.title}</td>
                <td><span class="chart-tag">${d.discipline}</span></td>
                <td><strong>${d.milestone}</strong></td>
                <td style="font-weight: 800; font-size: 15px; color: ${remColor};">${d.days_remaining} Days left</td>
                <td>${slipBadge}</td>
                <td>${d.plan_date}</td>
                <td style="font-weight: 700; color: var(--accent-amber);">${d.forecast_date}</td>
            </tr>
        `;
    });

    html += `</tbody></table>`;
    wrapper.innerHTML = html;
}

async function handleMultiUpload(files) {
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const fname = file.name.toLowerCase();
        let targetWp = null;
        if (fname.includes('00210') || fname.includes('excl')) targetWp = 'wp1_topside_excl';
        else if (fname.includes('topsides') || fname.includes('structure')) targetWp = 'wp1_topside_structure';
        else if (fname.includes('jacket')) targetWp = 'wp1_jacket';
        else if (fname.includes('wp2') || fname.includes('pipeline')) targetWp = 'wp2_pipeline';
        
        if (!targetWp) {
            // Default or ask user fallback assignment
            targetWp = 'wp1_topside_structure';
        }
        await uploadFileToServer(file, targetWp);
    }
}

async function uploadFileToServer(file, wpKey) {
    const progBox = document.getElementById('uploadProgressBox');
    const progFill = document.getElementById('progressBarFill');
    const progText = document.getElementById('progressText');
    const progPct = document.getElementById('progressPercent');
    
    progBox.style.display = 'block';
    progText.textContent = `Uploading ${file.name} for ${WP_SHORT_NAMES[wpKey] || wpKey}...`;
    progFill.style.width = '30%';
    progPct.textContent = '30%';

    const formData = new FormData();
    formData.append('file', file);
    formData.append('wp_key', wpKey);

    try {
        progFill.style.width = '70%';
        progPct.textContent = '70%';
        
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const errJson = await response.json().catch(() => ({}));
            throw new Error(errJson.error || `Upload failed (HTTP ${response.status})`);
        }
        
        progFill.style.width = '100%';
        progPct.textContent = '100%';
        
        setTimeout(() => { progBox.style.display = 'none'; }, 1500);
        await fetchDashboardData(true);
    } catch (err) {
        console.error('Upload Error:', err);
        progText.textContent = `Upload failed: ${err.message}`;
        progFill.style.background = 'var(--accent-rose)';
        showToast('Upload error: ' + err.message, 'error');
    }
}

function showToast(msg, type = 'success') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
        <span style="font-size: 18px;">${type === 'success' ? '✅' : '🚨'}</span>
        <span>${msg}</span>
    `;
    container.appendChild(toast);
    setTimeout(() => {
        toast.remove();
    }, 4500);
}
