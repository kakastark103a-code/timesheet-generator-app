document.addEventListener('DOMContentLoaded', () => {
    // State management
    const state = {
        selectedMonth: getInitialMonth(),
        domains: [],
        selectedDomainKeys: [],
        publicHolidays: [],
        autoSGHolidays: true,
        currentMemberDomain: 'cbg',
        membersByDomain: {}
    };

    // Theme & Language Controls
    const btnToggleTheme = document.getElementById('btnToggleTheme');
    const themeIcon = document.getElementById('themeIcon');
    const themeLabel = document.getElementById('themeLabel');
    
    const btnToggleLang = document.getElementById('btnToggleLang');
    const langFlag = document.getElementById('langFlag');
    const langLabel = document.getElementById('langLabel');

    let currentTheme = localStorage.getItem('app_theme') || 'dark';
    let currentLang = localStorage.getItem('app_lang') || 'vi';

    const i18nData = {
        vi: {
            headerSubtitle: "Khởi tạo & Review file Excel Timesheet tự động theo từng Month & Domain (Thêm/Bớt Member & Ghi chú điều chỉnh tùy chọn)",
            tabGenerator: "✨ Tạo Timesheet",
            tabIndividual: "👤 Timesheet Cá Nhân",
            tabReview: "🔍 Review & Audit",
            
            // Step 1
            step1Title: "1. Chọn Tháng & Domain Timesheet",
            lblMonth: "Tháng & Năm Mục Tiêu",
            btnPrevMonth: "← Tháng trước",
            btnCurrentMonth: "Tháng này",
            btnNextMonth: "Tháng sau →",
            lblSelectDomains: "Chọn Các Domain Cần Tạo:",
            btnSelectAllDomains: "Chọn tất cả",
            btnUploadTemplate: "📤 Upload Template Domain (.xlsx)",
            uploadHintText: "💡 Upload file Excel Template (.xlsx) của Domain bạn cần tạo timesheet để hệ thống tự động nhận diện Domain & Nhân viên.",
            
            // Step 2
            step2Title: "2. Danh Sách Nhân Viên & Ngày Lễ Singapore",
            btnAddMember: "+ Thêm Nhân Viên",
            lblManageDomain: "Quản lý nhân viên thuộc Domain:",
            thName: "Họ tên",
            thTeam: "Team",
            thLead: "Lead",
            thRole: "Vị trí",
            thLeave: "Phép năm",
            thBal: "Dư phép",
            thActions: "Thao tác",
            lblAutoSG: "🇸🇬 Tự động nạp Ngày lễ Singapore",
            lblAutoSGDesc: "Nhận diện chuẩn ngày lễ Singapore & tự dời bù nếu rơi vào Chủ Nhật",
            btnAddHoliday: "+ Thêm ngày lễ",
            
            // Step 3
            step3Title: "3. Ghi Chú & Điều Chỉnh Tùy Chọn (Overrides)",
            btnLoadNotes: "Load Mẫu Comment Ghi Chú",
            notesHelpText: "Dán bảng ghi chú / comment điều chỉnh số dư phép hoặc đổi loại công việc của ngày cụ thể:",
            notesHint: "💡 Hỗ trợ tự động nhân đôi dòng cho <code>Weekday support</code> (1 dòng Project Task 8h + 1 dòng Weekday support 4h chữ đỏ), <code>Weekend support</code>, <code>PH Support</code> và cập nhật dư phép <code>Leave Balance upto</code>.",
            lblStatsTitle: "Thống Kê Tháng & Xem Trước Ngày",
            statWorkingDays: "Ngày làm việc (Mon-Fri)",
            statWorkingDaysSub: "Tự động chọn cho từng Resource",
            statWeekendDays: "Ngày nghỉ cuối tuần",
            statWeekendDaysSub: "Đã tự động loại bỏ (Sat/Sun)",
            statHolidays: "Ngày lễ (Public Holiday)",
            statHolidaysSub: "Số ngày công 0 giờ",
            statTotalDays: "Tổng số ngày trong tháng",
            statTotalDaysSub: "Tất cả các ngày calendar",
            btnGenerate: "Tạo & Tải Xuống File Excel Timesheet",
            
            // Tab 3 Individual
            indivCardTitle: "1. Chọn Nhân Viên & Tháng Mục Tiêu",
            lblSelectMember: "Chọn Nhân Viên Cần Tạo Timesheet:",
            indivNotesTitle: "2. Ghi Chú Hoạt Động / OT / Phép Cho Nhân Viên",
            indivNotesHelp: "Nhập các hoạt động OT, Weekday support, Weekend support hoặc Ngày nghỉ phép riêng của nhân sự này:",
            indivPreviewTitle: "Lịch Hoạt Động Tháng Của",
            btnGenIndiv: "🚀 Tạo & Tải File Timesheet Cho Nhân Viên Này (.xlsx)",
            
            // Tab 2 Review
            reviewUploadTitle: "Review Nhiều File Timesheet Đã Điền (.xlsx)",
            btnReviewUpload: "📂 Chọn Một Hoặc Nhiều File Timesheet To Review (.xlsx)",
            reviewUploadHint: "Kéo thả hoặc giữ Ctrl / Cmd để chọn đồng thời nhiều file Excel Timesheet (VD: Timesheet các domain CBG, DXP, Digital...).",
            kpiFiles: "TỔNG SỐ TỆP REVIEWED",
            kpiFilesSub: "File Excel đã được quét",
            kpiHeadcount: "TỔNG NHÂN SỰ ĐƯỢC RÀ SOÁT",
            kpiHeadcountSub: "Số lượng Headcount",
            kpiErrors: "SỐ LỖI & VI PHẠM PHÁT HIỆN",
            kpiErrorsSub: "Worktype sai, âm phép, giờ nghỉ...",
            kpiCompliance: "TỶ LỆ TUÂN THỦ CLIENT RULE",
            kpiComplianceSub: "Độ chuẩn xác dữ liệu",
            sectionDetectedErrors: "⚠️ Danh Sách Lỗi Violations Phát Hiện & Đề Xuất Auto-Fix",
            btnSelectAllFixes: "Tích Tất Cả Fixes",
            sectionExcelViewer: "Trình Xem Dữ Liệu Trang Tính (Excel Viewer)",
            sheetTimesheet: "Sheet Timesheet",
            sheetBalance: "Sheet Balance",
            sheetSummary: "Sheet Summary",
            fixBarTitle: "⚡ Phê Duyệt & Tự Động Khắc Phục Lỗi",
            fixBarSub: "Áp dụng các Fix Option đã tích chọn trực tiếp vào file Excel và xuất bản sạch chuẩn 100% Client Rule.",
            btnApplyFixes: "🚀 Phê Duyệt Fix & Tải File Sạch (.xlsx / .zip)",
            
            // Modal & General
            modalTitleAdd: "Thêm Nhân Viên Mới",
            modalTitleEdit: "Chỉnh Sửa Thông Tin Nhân Viên",
            btnCancel: "Hủy Bỏ",
            btnSave: "Lưu Thay Đổi"
        },
        en: {
            headerSubtitle: "Automated Excel Timesheet Generation & Audit Review across Domains (Member management & Custom overrides)",
            tabGenerator: "✨ Generator",
            tabIndividual: "👤 Member Generator",
            tabReview: "🔍 Review & Audit",
            
            // Step 1
            step1Title: "1. Target Month & Domain Selection",
            lblMonth: "Target Month & Year",
            btnPrevMonth: "← Prev Month",
            btnCurrentMonth: "Current Month",
            btnNextMonth: "Next Month →",
            lblSelectDomains: "Select Domains To Generate:",
            btnSelectAllDomains: "Select All",
            btnUploadTemplate: "📤 Upload Domain Template (.xlsx)",
            uploadHintText: "💡 Upload your Domain Excel Template file (.xlsx) to automatically parse Domain & Member list.",
            
            // Step 2
            step2Title: "2. Members List & Singapore Public Holidays",
            btnAddMember: "+ Add Member",
            lblManageDomain: "Manage Members for Domain:",
            thName: "Name",
            thTeam: "Team",
            thLead: "Lead",
            thRole: "Role",
            thLeave: "Total Leave",
            thBal: "Leave Bal.",
            thActions: "Actions",
            lblAutoSG: "🇸🇬 Auto-load Singapore Public Holidays",
            lblAutoSGDesc: "Detect Singapore Public Holidays & handle Sunday substitute holidays",
            btnAddHoliday: "+ Add Holiday",
            
            // Step 3
            step3Title: "3. Notes & Custom Overrides",
            btnLoadNotes: "Load Sample Notes",
            notesHelpText: "Paste notes / comment table for OT activities, weekend support, or leave balance updates:",
            notesHint: "💡 Automatically duplicates row for <code>Weekday support</code> (1 Project Task 8h row + 1 Weekday support 4h red row), <code>Weekend support</code>, <code>PH Support</code> and updates <code>Leave Balance upto</code>.",
            lblStatsTitle: "Month Statistics & Calendar Preview",
            statWorkingDays: "Working Days (Mon-Fri)",
            statWorkingDaysSub: "Auto-computed per Resource",
            statWeekendDays: "Weekend Days",
            statWeekendDaysSub: "Excluded automatically (Sat/Sun)",
            statHolidays: "Public Holidays",
            statHolidaysSub: "0 hours actual time",
            statTotalDays: "Total Calendar Days",
            statTotalDaysSub: "Days in selected month",
            btnGenerate: "Generate & Download Excel Timesheets",
            
            // Tab 3 Individual
            indivCardTitle: "1. Target Member & Month Selection",
            lblSelectMember: "Select Member To Generate:",
            indivNotesTitle: "2. Member Activities / OT / Leave Notes",
            indivNotesHelp: "Enter OT activities, weekend support, or individual leave for this member:",
            indivPreviewTitle: "Monthly Activity Schedule for",
            btnGenIndiv: "🚀 Generate & Download Member Timesheet (.xlsx)",
            
            // Tab 2 Review
            reviewUploadTitle: "Upload Members' Timesheet Files for AI Review & Audit",
            btnReviewUpload: "📂 Select One or Multiple Timesheet Files To Review (.xlsx)",
            reviewUploadHint: "Drag and drop or hold Ctrl / Cmd to select multiple Excel Timesheet files simultaneously (Batch .xlsx upload).",
            kpiFiles: "TOTAL REVIEWED FILES",
            kpiFilesSub: "Audited Excel Files",
            kpiHeadcount: "TOTAL HEADCOUNT AUDITED",
            kpiHeadcountSub: "Total Member Count",
            kpiErrors: "VIOLATIONS & ERRORS FOUND",
            kpiErrorsSub: "Invalid worktypes, negative leaves...",
            kpiCompliance: "CLIENT RULE COMPLIANCE",
            kpiComplianceSub: "Data Accuracy Rate",
            sectionDetectedErrors: "⚠️ Detected Violations & Proposed Auto-Fixes",
            btnSelectAllFixes: "Select All Fixes",
            sectionExcelViewer: "Live Excel Sheet Inspector",
            sheetTimesheet: "Sheet Timesheet",
            sheetBalance: "Sheet Balance",
            sheetSummary: "Sheet Summary",
            fixBarTitle: "⚡ Approve & Auto-Fix Timesheet Files",
            fixBarSub: "Apply selected fix options directly into Excel files and download clean Client-ready files.",
            btnApplyFixes: "🚀 Approve Fixes & Download Clean Files (.xlsx / .zip)",
            
            // Modal & General
            modalTitleAdd: "Add New Member",
            modalTitleEdit: "Edit Member Information",
            btnCancel: "Cancel",
            btnSave: "Save Changes"
        }
    };

    function applyTheme(theme) {
        currentTheme = theme;
        localStorage.setItem('app_theme', theme);
        if (theme === 'light') {
            document.body.classList.add('light-theme');
            if (themeIcon) themeIcon.textContent = '☀️';
            if (themeLabel) themeLabel.textContent = 'Light';
        } else {
            document.body.classList.remove('light-theme');
            if (themeIcon) themeIcon.textContent = '🌙';
            if (themeLabel) themeLabel.textContent = 'Dark';
        }
    }

    const monthNames = {
        vi: ['Tháng 1 (Jan)', 'Tháng 2 (Feb)', 'Tháng 3 (Mar)', 'Tháng 4 (Apr)', 'Tháng 5 (May)', 'Tháng 6 (Jun)', 'Tháng 7 (Jul)', 'Tháng 8 (Aug)', 'Tháng 9 (Sep)', 'Tháng 10 (Oct)', 'Tháng 11 (Nov)', 'Tháng 12 (Dec)'],
        en: ['Month 1 (Jan)', 'Month 2 (Feb)', 'Month 3 (Mar)', 'Month 4 (Apr)', 'Month 5 (May)', 'Month 6 (Jun)', 'Month 7 (Jul)', 'Month 8 (Aug)', 'Month 9 (Sep)', 'Month 10 (Oct)', 'Month 11 (Nov)', 'Month 12 (Dec)']
    };

    function updateMonthDropdownOptions(lang) {
        const selectMonthNum = document.getElementById('selectMonthNum');
        const indivSelectMonthNum = document.getElementById('indivSelectMonthNum');
        const names = monthNames[lang] || monthNames.vi;
        
        [selectMonthNum, indivSelectMonthNum].forEach(sel => {
            if (!sel) return;
            const currentVal = sel.value;
            sel.innerHTML = names.map((n, idx) => {
                const val = String(idx + 1).padStart(2, '0');
                return `<option value="${val}">${n}</option>`;
            }).join('');
            sel.value = currentVal;
        });
    }

    function applyLang(lang) {
        currentLang = lang;
        localStorage.setItem('app_lang', lang);
        if (lang === 'en') {
            if (langFlag) langFlag.textContent = '🇬🇧';
            if (langLabel) langLabel.textContent = 'EN';
        } else {
            if (langFlag) langFlag.textContent = '🇻🇳';
            if (langLabel) langLabel.textContent = 'VI';
        }

        const dict = i18nData[lang];
        if (!dict) return;

        document.querySelectorAll('[data-i18n]').forEach(el => {
            const key = el.getAttribute('data-i18n');
            if (dict[key]) {
                if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
                    el.placeholder = dict[key];
                } else {
                    el.innerHTML = dict[key];
                }
            }
        });

        const btnGen = document.getElementById('btnGenerate');
        if (btnGen && dict.btnGenerate) {
            btnGen.innerHTML = `
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="7 10 12 15 17 10"></polyline>
                    <line x1="12" y1="15" x2="12" y2="3"></line>
                </svg>
                ${dict.btnGenerate}
            `;
        }

        updateMonthDropdownOptions(lang);

        if (state.membersByDomain[state.currentMemberDomain]) {
            renderMembersTable(state.membersByDomain[state.currentMemberDomain]);
        }

        if (isIndivInitialized) {
            populateIndivDomains();
        }

        const aiMsgContent = document.querySelector('#aiChatMessages .ai-msg.bot .ai-msg-content');
        if (aiMsgContent) {
            if (lang === 'en') {
                aiMsgContent.innerHTML = `
                    Hello! I am your <strong>AI Support Assistant</strong>. I can help you:<br>
                    • 🔍 <strong>Audit Timesheets</strong>: Auto-detect negative leave balances & invalid OT.<br>
                    • 📝 <strong>Natural Language Fixes</strong>: Enter requests like <em>"Fix leave for Ha to 12"</em>, <em>"Add 4h OT for Mai on Jul 1"</em>.<br>
                    • 💡 <strong>Automatic updates</strong> directly into your system!
                `;
            } else {
                aiMsgContent.innerHTML = `
                    Xin chào! Tôi là <strong>AI Support Assistant</strong>. Tôi có thể giúp bạn:<br>
                    • 🔍 <strong>Rà soát lỗi Timesheet</strong>: Tự động tìm số dư phép âm, sai loại OT.<br>
                    • 📝 <strong>Điều chỉnh tự động</strong>: Nhập lệnh như <em>"Sửa phép cho Hà về 12"</em>, <em>"Thêm OT 4h cho Mai ngày 1/7"</em>.<br>
                    • 💡 <strong>Tự động áp dụng ghi chú</strong> vào hệ thống!
                `;
            }
        }

        const aiInput = document.getElementById('aiChatInput');
        if (aiInput) {
            aiInput.placeholder = lang === 'en' ? "Type request or paste comment notes..." : "Nhập yêu cầu hoặc dán ghi chú...";
        }
    }

    if (btnToggleTheme) {
        btnToggleTheme.addEventListener('click', () => {
            applyTheme(currentTheme === 'dark' ? 'light' : 'dark');
        });
    }

    if (btnToggleLang) {
        btnToggleLang.addEventListener('click', () => {
            applyLang(currentLang === 'vi' ? 'en' : 'vi');
        });
    }

    // DOM Elements
    const monthSelect = document.getElementById('monthSelect');
    const btnPrevMonth = document.getElementById('btnPrevMonth');
    const btnCurrentMonth = document.getElementById('btnCurrentMonth');
    const btnNextMonth = document.getElementById('btnNextMonth');
    
    const domainList = document.getElementById('domainList');
    const btnSelectAllDomains = document.getElementById('btnSelectAllDomains');
    const btnTriggerUpload = document.getElementById('btnTriggerUpload');
    const templateFileInput = document.getElementById('templateFileInput');
    const uploadStatusText = document.getElementById('uploadStatusText');
    
    const memberDomainSelect = document.getElementById('memberDomainSelect');
    const membersTableBody = document.getElementById('membersTableBody');
    const btnAddNewMember = document.getElementById('btnAddNewMember');
    
    const memberModal = document.getElementById('memberModal');
    const modalTitle = document.getElementById('modalTitle');
    const btnModalClose = document.getElementById('btnModalClose');
    const btnModalCancel = document.getElementById('btnModalCancel');
    const btnModalSave = document.getElementById('btnModalSave');
    const editMemberIndex = document.getElementById('editMemberIndex');
    
    const inputMemberName = document.getElementById('inputMemberName');
    const inputFsoftAccount = document.getElementById('inputFsoftAccount');
    const inputMemberLead = document.getElementById('inputMemberLead');
    const inputMemberTeam = document.getElementById('inputMemberTeam');
    const inputMemberLocation = document.getElementById('inputMemberLocation');
    const inputTotalLeave = document.getElementById('inputTotalLeave');
    const inputLeaveBalanceUpto = document.getElementById('inputLeaveBalanceUpto');
    
    const commentNotesText = document.getElementById('commentNotesText');
    const btnLoadExampleNotes = document.getElementById('btnLoadExampleNotes');
    
    const chkAutoSGHolidays = document.getElementById('chkAutoSGHolidays');
    const phDatePicker = document.getElementById('phDatePicker');
    const btnAddPH = document.getElementById('btnAddPH');
    const phTagsContainer = document.getElementById('phTagsContainer');
    
    const btnGenerate = document.getElementById('btnGenerate');
    const loadingOverlay = document.getElementById('loadingOverlay');
    const loadingText = document.getElementById('loadingText');
    
    const lblMonthTitle = document.getElementById('lblMonthTitle');
    const statWorkingDays = document.getElementById('statWorkingDays');
    const statWeekendDays = document.getElementById('statWeekendDays');
    const statHolidays = document.getElementById('statHolidays');
    const statTotalDays = document.getElementById('statTotalDays');
    const daysGrid = document.getElementById('daysGrid');

    function getInitialMonth() {
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        return `${year}-${month}`;
    }

    const selectMonthNum = document.getElementById('selectMonthNum');
    const selectYearNum = document.getElementById('selectYearNum');

    function syncTab1DropdownsFromState() {
        if (!state.selectedMonth) return;
        const parts = state.selectedMonth.split('-');
        if (selectYearNum) selectYearNum.value = parts[0] || '2026';
        if (selectMonthNum) selectMonthNum.value = parts[1] || '08';
    }

    function syncStateFromTab1Dropdowns() {
        if (!selectMonthNum || !selectYearNum) return;
        state.selectedMonth = `${selectYearNum.value}-${selectMonthNum.value}`;
        if (monthSelect) monthSelect.value = state.selectedMonth;
        updatePreview();
    }

    function init() {
        monthSelect.value = state.selectedMonth;
        syncTab1DropdownsFromState();
        loadDomains();
        updatePreview();
        setupEventListeners();
        applyTheme(currentTheme);
        applyLang(currentLang);
    }

    function setupEventListeners() {
        if (selectMonthNum) selectMonthNum.addEventListener('change', syncStateFromTab1Dropdowns);
        if (selectYearNum) selectYearNum.addEventListener('change', syncStateFromTab1Dropdowns);

        if (indivSelectMonthNum) indivSelectMonthNum.addEventListener('change', syncTab3MonthValue);
        if (indivSelectYearNum) indivSelectYearNum.addEventListener('change', syncTab3MonthValue);

        monthSelect.addEventListener('change', (e) => {
            if (e.target.value) {
                state.selectedMonth = e.target.value;
                syncTab1DropdownsFromState();
                updatePreview();
            }
        });

        btnPrevMonth.addEventListener('click', () => changeMonth(-1));
        btnNextMonth.addEventListener('click', () => changeMonth(1));
        btnCurrentMonth.addEventListener('click', () => {
            state.selectedMonth = getInitialMonth();
            monthSelect.value = state.selectedMonth;
            syncTab1DropdownsFromState();
            updatePreview();
        });

        chkAutoSGHolidays.addEventListener('change', (e) => {
            state.autoSGHolidays = e.target.checked;
            updatePreview();
        });

        btnSelectAllDomains.addEventListener('click', toggleSelectAllDomains);

        btnTriggerUpload.addEventListener('click', () => templateFileInput.click());
        templateFileInput.addEventListener('change', handleTemplateUpload);

        memberDomainSelect.addEventListener('change', (e) => {
            state.currentMemberDomain = e.target.value;
            loadMembersForDomain(state.currentMemberDomain);
        });

        btnAddNewMember.addEventListener('click', openAddMemberModal);
        btnModalClose.addEventListener('click', closeMemberModal);
        btnModalCancel.addEventListener('click', closeMemberModal);
        btnModalSave.addEventListener('click', saveMemberModal);

        btnLoadExampleNotes.addEventListener('click', loadExampleNotes);

        btnAddPH.addEventListener('click', addPublicHoliday);
        phDatePicker.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') addPublicHoliday();
        });

        btnGenerate.addEventListener('click', generateTimesheets);
    }

    function changeMonth(delta) {
        const [y, m] = state.selectedMonth.split('-').map(Number);
        const date = new Date(y, m - 1 + delta, 1);
        const newY = date.getFullYear();
        const newM = String(date.getMonth() + 1).padStart(2, '0');
        state.selectedMonth = `${newY}-${newM}`;
        monthSelect.value = state.selectedMonth;
        syncTab1DropdownsFromState();
        updatePreview();
    }

    async function loadDomains() {
        try {
            const res = await fetch('/api/domains');
            const data = await res.json();
            state.domains = data.domains || [];
            
            const availableKeys = state.domains.filter(d => d.available).map(d => d.key);
            if (!state.selectedDomainKeys.length) {
                state.selectedDomainKeys = [...availableKeys];
            } else {
                state.selectedDomainKeys = state.selectedDomainKeys.filter(k => availableKeys.includes(k));
                availableKeys.forEach(k => {
                    if (!state.selectedDomainKeys.includes(k)) state.selectedDomainKeys.push(k);
                });
            }
                
            renderDomainList();
            renderMemberDomainDropdown();
            
            if (availableKeys.length > 0) {
                state.currentMemberDomain = availableKeys[0];
                memberDomainSelect.value = state.currentMemberDomain;
                loadMembersForDomain(state.currentMemberDomain);
            }
        } catch (err) {
            console.error('Failed to load domains:', err);
            domainList.innerHTML = `<div class="error-msg">Không thể tải danh sách Domain</div>`;
        }
    }

    function renderDomainList() {
        if (!state.domains.length) {
            domainList.innerHTML = `<div class="empty-msg">Không có domain nào</div>`;
            return;
        }

        domainList.innerHTML = state.domains.map(domain => {
            const isChecked = state.selectedDomainKeys.includes(domain.key);
            const disabledAttr = domain.available ? '' : 'disabled';
            const badge = domain.is_custom ? ' <span class="ph-tag" style="font-size: 0.7rem; padding: 0.1rem 0.4rem;">Custom</span>' : '';
            return `
                <label class="domain-item ${isChecked ? 'active' : ''}">
                    <input type="checkbox" value="${domain.key}" ${isChecked ? 'checked' : ''} ${disabledAttr}>
                    <div class="domain-info">
                        <div class="domain-name">${domain.name}${badge}</div>
                        <div class="domain-file">${domain.template}</div>
                    </div>
                </label>
            `;
        }).join('');

        domainList.querySelectorAll('input[type="checkbox"]').forEach(chk => {
            chk.addEventListener('change', (e) => {
                const key = e.target.value;
                if (e.target.checked) {
                    if (!state.selectedDomainKeys.includes(key)) state.selectedDomainKeys.push(key);
                } else {
                    state.selectedDomainKeys = state.selectedDomainKeys.filter(k => k !== key);
                }
                btnSelectAllDomains.textContent = state.selectedDomainKeys.length === state.domains.length ? 'Bỏ chọn tất cả' : 'Chọn tất cả';
            });
        });
    }

    function renderMemberDomainDropdown() {
        memberDomainSelect.innerHTML = state.domains.map(d => `
            <option value="${d.key}">${d.name}</option>
        `).join('');
    }

    function toggleSelectAllDomains() {
        const availableKeys = state.domains.filter(d => d.available).map(d => d.key);
        if (state.selectedDomainKeys.length === availableKeys.length) {
            state.selectedDomainKeys = [];
        } else {
            state.selectedDomainKeys = [...availableKeys];
        }
        renderDomainList();
        btnSelectAllDomains.textContent = state.selectedDomainKeys.length === availableKeys.length ? 'Bỏ chọn tất cả' : 'Chọn tất cả';
    }

    // Member Management
    async function loadMembersForDomain(domainKey) {
        if (state.membersByDomain[domainKey]) {
            renderMembersTable(state.membersByDomain[domainKey]);
            return;
        }

        try {
            const res = await fetch(`/api/members?domain=${domainKey}`);
            const data = await res.json();
            const members = data.members || [];
            state.membersByDomain[domainKey] = members;
            renderMembersTable(members);
        } catch (err) {
            console.error('Failed to load members:', err);
            membersTableBody.innerHTML = `<tr><td colspan="7">Lỗi tải danh sách nhân viên</td></tr>`;
        }
    }

    function renderMembersTable(members) {
        const isEn = currentLang === 'en';
        if (!members || !members.length) {
            const emptyMsg = isEn ? "No members found in this domain." : "Chưa có nhân viên nào trong domain này";
            membersTableBody.innerHTML = `<tr><td colspan="7" style="text-align:center; color: var(--text-dim);">${emptyMsg}</td></tr>`;
            return;
        }

        const titleEdit = isEn ? "Edit" : "Sửa";
        const titleDelete = isEn ? "Delete" : "Xóa";

        membersTableBody.innerHTML = members.map((m, idx) => `
            <tr>
                <td><strong>${m.name}</strong> ${m.fsoft_account ? `<span style="color:var(--text-dim)">(${m.fsoft_account})</span>` : ''}</td>
                <td>${m.team || '--'}</td>
                <td><strong style="color:#a5b4fc">${m.lead || '--'}</strong></td>
                <td><span class="ph-tag" style="font-size:0.75rem; padding:0.1rem 0.5rem;">${m.location || 'Offshore'}</span></td>
                <td>${m.total_leave || 14}</td>
                <td><strong style="color:var(--accent)">${m.leave_balance_upto || 10}</strong></td>
                <td>
                    <div class="member-actions">
                        <button type="button" class="action-icon-btn" onclick="window.editMember(${idx})" title="${titleEdit}">✏️</button>
                        <button type="button" class="action-icon-btn" onclick="window.deleteMember(${idx})" title="${titleDelete}">🗑️</button>
                    </div>
                </td>
            </tr>
        `).join('');
    }

    window.editMember = function(idx) {
        const members = state.membersByDomain[state.currentMemberDomain] || [];
        const m = members[idx];
        if (!m) return;

        editMemberIndex.value = idx;
        modalTitle.textContent = currentLang === 'en' ? 'Edit Member Information' : 'Chỉnh Sửa Thông Tin Nhân Viên';
        inputMemberName.value = m.name || '';
        inputFsoftAccount.value = m.fsoft_account || '';
        inputMemberLead.value = m.lead || '';
        inputMemberTeam.value = m.team || '';
        inputMemberLocation.value = m.location || 'Offshore';
        inputTotalLeave.value = m.total_leave || 14;
        inputLeaveBalanceUpto.value = m.leave_balance_upto || 10;

        memberModal.classList.add('active');
    };

    window.deleteMember = function(idx) {
        const members = state.membersByDomain[state.currentMemberDomain] || [];
        if (!members[idx]) return;
        const memberName = members[idx].name;
        const isEn = currentLang === 'en';
        const title = isEn ? '⚠️ Delete Member Confirmation' : '⚠️ Xác Nhận Xóa Nhân Viên';
        const msg = isEn ? `Are you sure you want to remove member "${memberName}" from the list?` : `Bạn có chắc chắn muốn xóa nhân viên "${memberName}" khỏi danh sách?`;
        const okText = isEn ? 'Delete Member' : 'Xác Nhận Xóa';
        const cancelText = isEn ? 'Cancel' : 'Hủy Bỏ';

        showConfirmDialog(
            title,
            msg,
            () => {
                members.splice(idx, 1);
                renderMembersTable(members);
                syncMembersToBackend(state.currentMemberDomain, members);
                showNotification(isEn ? `Deleted member ${memberName}` : `Đã xóa nhân viên ${memberName}`, 'warning');
            },
            okText,
            cancelText
        );
    };

    function openAddMemberModal() {
        editMemberIndex.value = -1;
        modalTitle.textContent = currentLang === 'en' ? 'Add New Member' : 'Thêm Nhân Viên Mới';
        inputMemberName.value = '';
        inputFsoftAccount.value = '';
        inputMemberLead.value = 'Xinjian';
        inputMemberTeam.value = 'EAI';
        inputMemberLocation.value = 'Offshore';
        inputTotalLeave.value = 14;
        inputLeaveBalanceUpto.value = 10;
        memberModal.classList.add('active');
    }

    function closeMemberModal() {
        memberModal.classList.remove('active');
    }

    function saveMemberModal() {
        const name = inputMemberName.value.trim();
        if (!name) {
            showNotification('Vui lòng nhập Họ tên nhân viên', 'warning');
            return;
        }

        const idx = parseInt(editMemberIndex.value);
        const members = state.membersByDomain[state.currentMemberDomain] || [];

        const memberObj = {
            name: name,
            fsoft_account: inputFsoftAccount.value.trim(),
            team: inputMemberTeam.value.trim(),
            lead: inputMemberLead.value.trim(),
            location: inputMemberLocation.value,
            vendor: 'FPT',
            total_leave: parseFloat(inputTotalLeave.value) || 14,
            leave_balance_upto: parseFloat(inputLeaveBalanceUpto.value) || 10
        };

        if (idx >= 0 && idx < members.length) {
            members[idx] = memberObj;
        } else {
            members.push(memberObj);
        }

        state.membersByDomain[state.currentMemberDomain] = members;
        renderMembersTable(members);
        syncMembersToBackend(state.currentMemberDomain, members);
        closeMemberModal();
    }

    async function syncMembersToBackend(domainKey, members) {
        try {
            await fetch('/api/members', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ domain: domainKey, members: members })
            });
        } catch (err) {
            console.error('Failed to sync members:', err);
        }
    }

    function loadExampleNotes() {
        commentNotesText.value = `Member name\tFsoft Account\tProject name\tNote
Dao Manh Ha\tHaDM\tCBG CRM_OM_ AD\t1-Jul OT 4h để support activities
Duong Xuan Tung\tTungDX9\tCBG CRM_OM_ AD\t10-Jul OT 8h Weekend support
Duong Thi Tuyet Mai\tMaiDTT23\tDigital Domain\t7-May có 2 line Project Task => Cần update Work Item Type 1 line 7-May về Weekday support (4h)
Do Phu Tung\tTungDP2\tFlutter\tCần update Leave Balance upto Apr 26 về 12`;
    }

    // Handle Upload Template File
    async function handleTemplateUpload(e) {
        const files = Array.from(e.target.files || []);
        if (files.length === 0) return;

        const invalidFiles = files.filter(f => !f.name.endsWith('.xlsx'));
        if (invalidFiles.length > 0) {
            showNotification('Vui lòng chỉ chọn các tệp định dạng Excel (.xlsx).', 'warning');
            return;
        }

        const formData = new FormData();
        files.forEach(file => {
            formData.append('files', file);
        });
        uploadStatusText.textContent = `⏳ Đang upload ${files.length} file template...`;

        try {
            const res = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });

            const contentType = res.headers.get('content-type') || '';
            if (!contentType.includes('application/json')) {
                throw new Error(`Server lỗi (HTTP ${res.status}). Vui lòng kiểm tra dung lượng file.`);
            }

            const data = await res.json();
            if (!res.ok) throw new Error(data.error || 'Upload thất bại');

            const count = data.filenames ? data.filenames.length : 1;
            uploadStatusText.textContent = `✅ Đã upload thành công ${count} template domain!`;
            setTimeout(() => { uploadStatusText.textContent = ''; }, 4000);

            await loadDomains();
        } catch (err) {
            console.error('Upload error:', err);
            uploadStatusText.textContent = `❌ Lỗi: ${err.message}`;
        } finally {
            templateFileInput.value = '';
        }
    }


    // Public Holidays Handling
    function addPublicHoliday() {
        const dateVal = phDatePicker.value;
        if (!dateVal) return;

        if (!state.publicHolidays.includes(dateVal)) {
            state.publicHolidays.push(dateVal);
            state.publicHolidays.sort();
            updatePreview();
        }
        phDatePicker.value = '';
    }

    function removePublicHoliday(dateVal) {
        state.publicHolidays = state.publicHolidays.filter(d => d !== dateVal);
        updatePreview();
    }

    function renderPHTags(daysData) {
        const holidayDays = daysData ? daysData.filter(d => d.is_holiday) : [];
        if (!holidayDays.length) {
            phTagsContainer.innerHTML = `<span class="empty-ph-notice">Không có ngày lễ nào trong tháng này.</span>`;
            return;
        }

        phTagsContainer.innerHTML = holidayDays.map(d => {
            const isCustom = state.publicHolidays.includes(d.date);
            const tagClass = isCustom ? 'ph-tag' : 'ph-tag sg-holiday';
            const icon = isCustom ? '🎉' : '🇸🇬';
            const label = d.holiday_name ? `${icon} ${d.date} (${d.holiday_name})` : `${icon} ${d.date}`;
            
            return `
                <span class="${tagClass}">
                    ${label}
                    ${isCustom ? `<span class="ph-remove" data-date="${d.date}">×</span>` : ''}
                </span>
            `;
        }).join('');

        phTagsContainer.querySelectorAll('.ph-remove').forEach(btn => {
            btn.addEventListener('click', (e) => {
                removePublicHoliday(e.target.dataset.date);
            });
        });
    }

    // Fetch month breakdown preview
    async function updatePreview() {
        try {
            const res = await fetch('/api/preview', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    month: state.selectedMonth,
                    public_holidays: state.publicHolidays,
                    auto_sg_holidays: state.autoSGHolidays
                })
            });

            if (!res.ok) return;

            const data = await res.json();
            lblMonthTitle.textContent = data.month_name;
            statWorkingDays.textContent = data.working_days_count;
            statWeekendDays.textContent = data.weekend_days_count;
            statHolidays.textContent = data.holidays_count;
            statTotalDays.textContent = data.total_days;

            renderPHTags(data.days);
            renderDaysGrid(data.days);
        } catch (err) {
            console.error('Failed to update preview:', err);
        }
    }

    function renderDaysGrid(days) {
        if (!days || !days.length) return;

        daysGrid.innerHTML = days.map(d => {
            let statusText = 'Làm việc (8h)';
            let cellClass = 'working';

            if (d.is_weekend) {
                statusText = 'Cuối tuần (Bỏ qua)';
                cellClass = 'weekend';
            } else if (d.is_holiday) {
                statusText = d.holiday_name ? `🇸🇬 ${d.holiday_name} (0h)` : 'Public Holiday (0h)';
                cellClass = 'holiday';
            }

            return `
                <div class="day-cell ${cellClass}">
                    <div class="day-date">${d.day_num} ${d.weekday_name}</div>
                    <div class="day-status">${statusText}</div>
                </div>
            `;
        }).join('');
    }

    // Generate & Download Excel Timesheets
    async function generateTimesheets() {
        if (!state.selectedDomainKeys.length) {
            alert('Vui lòng chọn ít nhất 1 Domain để tạo timesheet.');
            return;
        }

        showLoading('Đang xử lý và khởi tạo file Excel Timesheet...');

        try {
            const response = await fetch('/api/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    month: state.selectedMonth,
                    domains: state.selectedDomainKeys,
                    public_holidays: state.publicHolidays,
                    auto_sg_holidays: state.autoSGHolidays,
                    comment_notes: commentNotesText.value,
                    custom_members: state.membersByDomain
                })
            });

            if (!response.ok) {
                const errData = await response.json().catch(() => ({}));
                throw new Error(errData.error || 'Có lỗi xảy ra khi tạo tệp Timesheet');
            }

            const contentDisposition = response.headers.get('Content-Disposition');
            let filename = `FPT_Timesheets_${state.selectedMonth}.xlsx`;
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename="?([^"]+)"?/);
                if (filenameMatch && filenameMatch[1]) {
                    filename = filenameMatch[1];
                }
            }

            const blob = await response.blob();
            const downloadUrl = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(downloadUrl);

        } catch (err) {
            console.error('Generation error:', err);
            alert(`Lỗi: ${err.message}`);
        } finally {
            hideLoading();
        }
    }

    // AI Chat & Audit Handling
    const btnRunAudit = document.getElementById('btnRunAudit');
    const auditAnomaliesContainer = document.getElementById('auditAnomaliesContainer');
    const reviewTablesSection = document.getElementById('reviewTablesSection');
    const reviewBalanceTableBody = document.getElementById('reviewBalanceTableBody');
    
    const btnToggleAIChat = document.getElementById('btnToggleAIChat');
    const btnCloseAIChat = document.getElementById('btnCloseAIChat');
    const aiChatPanel = document.getElementById('aiChatPanel');
    const aiChatMessages = document.getElementById('aiChatMessages');
    const aiChatInput = document.getElementById('aiChatInput');
    const btnSendAIChat = document.getElementById('btnSendAIChat');

    if (btnRunAudit) btnRunAudit.addEventListener('click', runTimesheetAudit);
    if (btnToggleAIChat) btnToggleAIChat.addEventListener('click', () => aiChatPanel.classList.toggle('active'));
    if (btnCloseAIChat) btnCloseAIChat.addEventListener('click', () => aiChatPanel.classList.remove('active'));
    if (btnSendAIChat) btnSendAIChat.addEventListener('click', sendAIChatMsg);
    if (aiChatInput) {
        aiChatInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') sendAIChatMsg();
        });
    }

    async function runTimesheetAudit() {
        const domainKey = state.selectedDomainKeys[0] || 'cbg';
        auditAnomaliesContainer.innerHTML = `<div class="empty-ph-notice">⏳ AI đang phân tích và rà soát dữ liệu Timesheet...</div>`;

        try {
            const res = await fetch('/api/review', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    month: state.selectedMonth,
                    domain: domainKey,
                    comment_notes: commentNotesText.value,
                    public_holidays: state.publicHolidays,
                    auto_sg_holidays: state.autoSGHolidays,
                    custom_members: state.membersByDomain
                })
            });

            if (!res.ok) throw new Error('Không thể rà soát dữ liệu');
            const data = await res.json();

            renderAuditAnomalies(data.anomalies || []);
            renderReviewBalanceTable(data.balance || []);
            reviewTablesSection.style.display = 'block';
        } catch (err) {
            auditAnomaliesContainer.innerHTML = `<div class="anomaly-item danger">Lỗi rà soát: ${err.message}</div>`;
        }
    }

    function renderAuditAnomalies(anomalies) {
        if (!anomalies.length) {
            auditAnomaliesContainer.innerHTML = `<div class="anomaly-item success">✅ Dữ liệu Timesheet hợp lệ 100%!</div>`;
            return;
        }

        auditAnomaliesContainer.innerHTML = anomalies.map(a => `
            <div class="anomaly-item ${a.type}">
                <div><strong>[${a.category}]</strong> ${a.message}</div>
                ${a.action_hint ? `<div class="anomaly-hint">💡 Gợi ý: ${a.action_hint}</div>` : ''}
            </div>
        `).join('');
    }

    function renderReviewBalanceTable(balanceList) {
        if (!balanceList.length) {
            reviewBalanceTableBody.innerHTML = `<tr><td colspan="5">Không có dữ liệu số dư phép</td></tr>`;
            return;
        }

        reviewBalanceTableBody.innerHTML = balanceList.map(b => {
            const isNegative = typeof b.balance_in_month === 'number' && b.balance_in_month < 0;
            const badge = isNegative ? `<span class="ph-tag" style="background:rgba(239,68,68,0.2); color:#fca5a5;">⚠️ Bị Âm Phép</span>` : `<span class="ph-tag sg-holiday">Hợp Lệ</span>`;
            return `
                <tr>
                    <td><strong>${b.name}</strong></td>
                    <td>${b.total_leave || 14}</td>
                    <td>${b.balance_upto || 10}</td>
                    <td><strong style="color:${isNegative ? 'var(--danger)' : 'var(--secondary)'}">${b.balance_in_month}</strong></td>
                    <td>${badge}</td>
                </tr>
            `;
        }).join('');
    }

    async function sendAIChatMsg() {
        const text = aiChatInput.value.trim();
        if (!text) return;

        appendChatMessage('user', text);
        aiChatInput.value = '';

        const domainKey = state.selectedDomainKeys[0] || 'cbg';

        try {
            const res = await fetch('/api/ai-chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    prompt: text,
                    comment_notes: commentNotesText.value,
                    domain: domainKey
                })
            });

            if (!res.ok) throw new Error('AI Server không phản hồi');
            const data = await res.json();

            appendChatMessage('bot', data.reply);

            if (data.updated_notes && data.updated_notes !== commentNotesText.value) {
                commentNotesText.value = data.updated_notes;
                runTimesheetAudit();
            }
        } catch (err) {
            appendChatMessage('bot', `⚠️ Lỗi kết nối AI Assistant: ${err.message}`);
        }
    }

    function appendChatMessage(sender, msgHtml) {
        const msgDiv = document.createElement('div');
        msgDiv.className = `ai-msg ${sender}`;
        msgDiv.innerHTML = `<div class="ai-msg-content">${msgHtml.replace(/\n/g, '<br>')}</div>`;
        aiChatMessages.appendChild(msgDiv);
        aiChatMessages.scrollTop = aiChatMessages.scrollHeight;
    }

    window.sendQuickAIChat = function(promptText) {
        if (!aiChatPanel.classList.contains('active')) {
            aiChatPanel.classList.add('active');
        }
        aiChatInput.value = promptText;
        sendAIChatMsg();
    };

    // Main Tab Switching (3 Tabs)
    const tabBtnGenerator = document.getElementById('tabBtnGenerator');
    const tabBtnIndividual = document.getElementById('tabBtnIndividual');
    const tabBtnReview = document.getElementById('tabBtnReview');

    const tabGeneratorView = document.getElementById('tabGeneratorView');
    const tabIndividualView = document.getElementById('tabIndividualView');
    const tabReviewView = document.getElementById('tabReviewView');

    if (tabBtnGenerator) tabBtnGenerator.addEventListener('click', () => switchMainTab('generator'));
    if (tabBtnIndividual) tabBtnIndividual.addEventListener('click', () => switchMainTab('individual'));
    if (tabBtnReview) tabBtnReview.addEventListener('click', () => switchMainTab('review'));

    function switchMainTab(tabName) {
        [tabBtnGenerator, tabBtnIndividual, tabBtnReview].forEach(b => b && b.classList.remove('active'));
        [tabGeneratorView, tabIndividualView, tabReviewView].forEach(v => v && (v.style.display = 'none'));

        if (tabName === 'generator') {
            if (tabBtnGenerator) tabBtnGenerator.classList.add('active');
            if (tabGeneratorView) tabGeneratorView.style.display = 'block';
        } else if (tabName === 'individual') {
            if (tabBtnIndividual) tabBtnIndividual.classList.add('active');
            if (tabIndividualView) tabIndividualView.style.display = 'block';
            initIndividualTab();
        } else {
            if (tabBtnReview) tabBtnReview.classList.add('active');
            if (tabReviewView) tabReviewView.style.display = 'block';
        }
    }

    // Custom Toast Notification System (replaces native alert())
    function showNotification(msg, type = 'info', duration = 4000) {
        const toastContainer = document.getElementById('toastContainer');
        if (!toastContainer) {
            console.log(`[Toast ${type}]: ${msg}`);
            return;
        }

        const iconMap = {
            success: '✅',
            warning: '⚠️',
            danger: '❌',
            info: '💡'
        };

        const toast = document.createElement('div');
        toast.className = `custom-toast ${type}`;
        toast.innerHTML = `
            <div class="toast-content">
                <span>${iconMap[type] || '💡'}</span>
                <span>${msg}</span>
            </div>
            <button class="toast-close-btn">&times;</button>
        `;

        toast.querySelector('.toast-close-btn').addEventListener('click', () => {
            toast.style.animation = 'toastFadeOut 0.25s forwards';
            setTimeout(() => toast.remove(), 250);
        });

        toastContainer.appendChild(toast);

        setTimeout(() => {
            if (toast.parentNode) {
                toast.style.animation = 'toastFadeOut 0.25s forwards';
                setTimeout(() => toast.remove(), 250);
            }
        }, duration);
    }
    window.showNotification = showNotification;

    // Custom Glassmorphic Confirm Dialog System (replaces native confirm())
    function showConfirmDialog(title, message, onConfirm, okText = "Xác Nhận", cancelText = "Hủy Bỏ") {
        const confirmModal = document.getElementById('confirmModal');
        const confirmModalTitle = document.getElementById('confirmModalTitle');
        const confirmModalMessage = document.getElementById('confirmModalMessage');
        const btnConfirmOk = document.getElementById('btnConfirmOk');
        const btnConfirmCancel = document.getElementById('btnConfirmCancel');
        const btnConfirmClose = document.getElementById('btnConfirmClose');

        if (!confirmModal) {
            if (window.confirm(message)) onConfirm();
            return;
        }

        confirmModalTitle.textContent = title;
        confirmModalMessage.textContent = message;
        btnConfirmOk.textContent = okText;
        btnConfirmCancel.textContent = cancelText;

        confirmModal.classList.add('active');

        function cleanup() {
            confirmModal.classList.remove('active');
            btnConfirmOk.removeEventListener('click', handleOk);
            btnConfirmCancel.removeEventListener('click', handleCancel);
            btnConfirmClose.removeEventListener('click', handleCancel);
        }

        function handleOk() {
            cleanup();
            if (onConfirm) onConfirm();
        }

        function handleCancel() {
            cleanup();
        }

        btnConfirmOk.addEventListener('click', handleOk);
        btnConfirmCancel.addEventListener('click', handleCancel);
        btnConfirmClose.addEventListener('click', handleCancel);
    }
    window.showConfirmDialog = showConfirmDialog;

    // Tab 3: Individual Member Timesheet Generator Logic
    const indivSelectMonthNum = document.getElementById('indivSelectMonthNum');
    const indivSelectYearNum = document.getElementById('indivSelectYearNum');
    const indivMonthSelect = document.getElementById('indivMonthSelect');
    
    const indivDomainSelect = document.getElementById('indivDomainSelect');
    const indivMemberSelect = document.getElementById('indivMemberSelect');
    
    const customMemberInputGroup = document.getElementById('customMemberInputGroup');
    const inputCustomMemberName = document.getElementById('inputCustomMemberName');
    const inputCustomMemberAccount = document.getElementById('inputCustomMemberAccount');
    const inputCustomMemberTeam = document.getElementById('inputCustomMemberTeam');
    const inputCustomMemberLeave = document.getElementById('inputCustomMemberLeave');
    
    const badgeMemberName = document.getElementById('badgeMemberName');
    const badgeMemberAccount = document.getElementById('badgeMemberAccount');
    const badgeMemberTeam = document.getElementById('badgeMemberTeam');
    const badgeMemberRole = document.getElementById('badgeMemberRole');
    const badgeMemberTotalLeave = document.getElementById('badgeMemberTotalLeave');
    const badgeMemberLeaveBal = document.getElementById('badgeMemberLeaveBal');
    const lblIndivMemberName = document.getElementById('lblIndivMemberName');

    const indivCommentNotes = document.getElementById('indivCommentNotes');
    const indivDaysGrid = document.getElementById('indivDaysGrid');
    const btnGenerateIndividual = document.getElementById('btnGenerateIndividual');

    let indivMembersList = [];
    let currentIndivMember = null;
    let isIndivInitialized = false;

    function initIndividualTab() {
        if (indivSelectMonthNum && indivSelectYearNum) {
            const parts = (state.selectedMonth || '2026-08').split('-');
            indivSelectYearNum.value = parts[0] || '2026';
            indivSelectMonthNum.value = parts[1] || '08';
        }
        syncTab3MonthValue();

        if (!isIndivInitialized) {
            populateIndivDomains();
            isIndivInitialized = true;
        } else {
            updateIndivCalendarPreview();
        }
    }

    function syncTab3MonthValue() {
        if (!indivSelectMonthNum || !indivSelectYearNum) return;
        const m = indivSelectMonthNum.value;
        const y = indivSelectYearNum.value;
        const monthVal = `${y}-${m}`;
        if (indivMonthSelect) indivMonthSelect.value = monthVal;
        updateIndivCalendarPreview();
    }

    function populateIndivDomains() {
        if (!state.domains || !state.domains.length) return;
        const isEn = currentLang === 'en';
        const allDomOpt = isEn ? `-- All / Blank Domains --` : `-- Tất cả / Để trống Domains (All Domains) --`;
        const customDomOpt = isEn ? `➕ Enter New Custom Domain...` : `➕ Nhập tên Domain mới (Custom Domain)...`;

        let html = `<option value="">${allDomOpt}</option>`;
        html += `<option value="__custom_domain__">${customDomOpt}</option>`;
        html += state.domains.map(d => `<option value="${d.key}">${d.name}</option>`).join('');
        indivDomainSelect.innerHTML = html;
        onIndivDomainChange(indivDomainSelect.value);
    }

    function onIndivDomainChange(domainKey) {
        if (domainKey === '__custom_domain__') {
            if (customDomainInputGroup) customDomainInputGroup.style.display = 'block';
            loadIndivMembersForDomain('');
        } else {
            if (customDomainInputGroup) customDomainInputGroup.style.display = 'none';
            loadIndivMembersForDomain(domainKey);
        }
    }

    async function loadIndivMembersForDomain(domainKey) {
        try {
            indivMembersList = [];

            if (!domainKey || domainKey === '' || domainKey === '__custom_domain__') {
                const fetchPromises = state.domains.map(async d => {
                    const res = await fetch(`/api/members?domain=${d.key}`);
                    const data = await res.json();
                    return (data.members || []).map(m => ({ ...m, domainKey: d.key, domainName: d.name }));
                });
                const results = await Promise.all(fetchPromises);
                results.forEach(mList => indivMembersList.push(...mList));
            } else {
                const res = await fetch(`/api/members?domain=${domainKey}`);
                const data = await res.json();
                const dInfo = state.domains.find(d => d.key === domainKey);
                const dName = dInfo ? dInfo.name : domainKey;
                indivMembersList = (data.members || []).map(m => ({ ...m, domainKey, domainName: dName }));
            }

            const isEn = currentLang === 'en';
            const allMemberOpt = isEn ? `-- All / Blank (Template Sample) --` : `-- Tất cả / Để trống (Template mẫu) --`;
            const customMemberOpt = isEn ? `➕ Enter New Custom Member Name...` : `➕ Nhập tên nhân sự mới (Custom Member Name)...`;

            let html = `<option value="">${allMemberOpt}</option>`;
            html += `<option value="__custom__">${customMemberOpt}</option>`;
            
            if (indivMembersList.length > 0) {
                html += indivMembersList.map((m, idx) => 
                    `<option value="${idx}">${m.name} (${m.account || 'N/A'}) - [${m.domainName || m.team || 'Offshore'}]</option>`
                ).join('');
            }
            
            indivMemberSelect.innerHTML = html;
            selectIndivMember('');
        } catch (err) {
            console.error('Failed to load members for individual tab:', err);
        }
    }

    function selectIndivMember(idxVal) {
        if (idxVal === '__custom__') {
            if (customMemberInputGroup) customMemberInputGroup.style.display = 'block';
            updateIndivMemberFromCustomInputs();
            return;
        } else {
            if (customMemberInputGroup) customMemberInputGroup.style.display = 'none';
        }

        if (idxVal === '' || idxVal === null || idxVal === undefined) {
            currentIndivMember = null;
            updateIndivMemberBadge(null);
            updateIndivCalendarPreview();
            return;
        }

        const idx = parseInt(idxVal);
        if (indivMembersList[idx]) {
            currentIndivMember = indivMembersList[idx];
            updateIndivMemberBadge(currentIndivMember);
            updateIndivCalendarPreview();
        }
    }

    function updateIndivMemberFromCustomInputs() {
        const cName = inputCustomMemberName ? inputCustomMemberName.value.trim() : '';
        const cAcc = inputCustomMemberAccount ? inputCustomMemberAccount.value.trim() : '';
        const cTeam = inputCustomMemberTeam ? inputCustomMemberTeam.value.trim() : 'Offshore';
        const cLeave = inputCustomMemberLeave ? parseFloat(inputCustomMemberLeave.value) || 10 : 10;

        currentIndivMember = {
            name: cName || 'Custom Member',
            account: cAcc || 'CustomAcc',
            team: cTeam,
            lead: '',
            location: 'Offshore',
            total_leave: 14,
            balance_upto: cLeave,
            domainKey: indivDomainSelect.value || 'cbg'
        };

        updateIndivMemberBadge(currentIndivMember);
    }

    if (inputCustomMemberName) inputCustomMemberName.addEventListener('input', updateIndivMemberFromCustomInputs);
    if (inputCustomMemberAccount) inputCustomMemberAccount.addEventListener('input', updateIndivMemberFromCustomInputs);
    if (inputCustomMemberTeam) inputCustomMemberTeam.addEventListener('input', updateIndivMemberFromCustomInputs);
    if (inputCustomMemberLeave) inputCustomMemberLeave.addEventListener('input', updateIndivMemberFromCustomInputs);

    function updateIndivMemberBadge(m) {
        if (!m) {
            badgeMemberName.textContent = 'Template Mẫu (Blank Template)';
            badgeMemberAccount.textContent = 'Sample';
            badgeMemberTeam.textContent = 'Offshore';
            badgeMemberRole.textContent = 'Resource';
            badgeMemberTotalLeave.textContent = '14';
            badgeMemberLeaveBal.textContent = '10';
            lblIndivMemberName.textContent = 'Template Mẫu';
            return;
        }

        badgeMemberName.textContent = m.name;
        badgeMemberAccount.textContent = m.account || '--';
        badgeMemberTeam.textContent = `${m.team || 'Offshore'} (${m.domainName || ''})`;
        badgeMemberRole.textContent = m.lead ? `Lead: ${m.lead}` : 'Member';
        badgeMemberTotalLeave.textContent = m.total_leave || 14;
        badgeMemberLeaveBal.textContent = m.balance_upto || 10;
        lblIndivMemberName.textContent = m.name;
    }

    async function updateIndivCalendarPreview() {
        const monthVal = indivMonthSelect.value || state.selectedMonth;
        try {
            const res = await fetch('/api/preview', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    month: monthVal,
                    auto_sg_holidays: true,
                    public_holidays: state.publicHolidays
                })
            });
            const data = await res.json();
            if (!res.ok) return;

            renderIndivCalendar(data.days);
        } catch (e) {
            console.error('Failed to update indiv preview:', e);
        }
    }

    function renderIndivCalendar(days) {
        if (!days || !days.length || !indivDaysGrid) return;
        indivDaysGrid.innerHTML = days.map(d => {
            let statusText = 'Project Task (8h)';
            let cellClass = 'working';

            if (d.is_weekend) {
                statusText = 'Cuối tuần (Bỏ qua)';
                cellClass = 'weekend';
            } else if (d.is_holiday) {
                statusText = d.holiday_name ? `🇸🇬 ${d.holiday_name} (0h)` : 'Public Holiday (0h)';
                cellClass = 'holiday';
            }

            return `
                <div class="day-cell ${cellClass}">
                    <div class="day-date">${d.day_num} ${d.weekday_name}</div>
                    <div class="day-status">${statusText}</div>
                </div>
            `;
        }).join('');
    }

    if (indivDomainSelect) {
        indivDomainSelect.addEventListener('change', (e) => onIndivDomainChange(e.target.value));
    }
    if (indivMemberSelect) {
        indivMemberSelect.addEventListener('change', (e) => selectIndivMember(e.target.value));
    }
    if (indivMonthSelect) {
        indivMonthSelect.addEventListener('change', () => updateIndivCalendarPreview());
    }

    if (btnGenerateIndividual) {
        btnGenerateIndividual.addEventListener('click', generateIndividualTimesheet);
    }

    async function generateIndividualTimesheet() {
        const monthVal = indivMonthSelect.value || state.selectedMonth;
        let domainVal = indivDomainSelect.value || 'cbg';
        if (domainVal === '__custom_domain__') {
            const customDomText = inputCustomDomainName ? inputCustomDomainName.value.trim() : '';
            domainVal = customDomText || 'Custom_Domain';
        }

        let memberToGen = currentIndivMember;
        if (indivMemberSelect.value === '__custom__') {
            updateIndivMemberFromCustomInputs();
            memberToGen = currentIndivMember;
        }

        const genName = memberToGen ? memberToGen.name : `Template mẫu (${domainVal.toUpperCase()})`;
        showLoading(`Đang khởi tạo Timesheet cá nhân cho ${genName}...`);

        try {
            const response = await fetch('/api/generate-individual', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    month: monthVal,
                    domain: domainVal,
                    member: memberToGen,
                    comment_notes: indivCommentNotes.value,
                    auto_sg_holidays: true,
                    public_holidays: state.publicHolidays
                })
            });

            if (!response.ok) {
                const errData = await response.json().catch(() => ({}));
                throw new Error(errData.error || 'Lỗi khi tạo Timesheet cá nhân');
            }

            const contentDisposition = response.headers.get('Content-Disposition');
            let filename = `FPT_Timesheet_${memberToGen ? (memberToGen.account || 'Member') : 'Sample'}.xlsx`;
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename="?([^"]+)"?/);
                if (filenameMatch && filenameMatch[1]) {
                    filename = filenameMatch[1];
                }
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);

            showNotification(`✅ Đã tạo & tải xuống file ${filename} thành công!`, 'success');
        } catch (err) {
            showNotification(`❌ Lỗi: ${err.message}`, 'danger');
        } finally {
            hideLoading();
        }
    }

    // Tab 2: Multi-File Upload & Interactive Review Workspace
    const btnTriggerReviewUpload = document.getElementById('btnTriggerReviewUpload');
    const reviewUploadFileInput = document.getElementById('reviewUploadFileInput');
    const reviewUploadStatusText = document.getElementById('reviewUploadStatusText');
    
    const reviewMetricsBar = document.getElementById('reviewMetricsBar');
    const kpiTotalFiles = document.getElementById('kpiTotalFiles');
    const kpiTotalMembers = document.getElementById('kpiTotalMembers');
    const kpiTotalErrors = document.getElementById('kpiTotalErrors');
    const kpiComplianceRate = document.getElementById('kpiComplianceRate');

    const fileTabsNavigation = document.getElementById('fileTabsNavigation');
    const activeFileWorkspace = document.getElementById('activeFileWorkspace');
    const fileDiffsListContainer = document.getElementById('fileDiffsListContainer');
    const btnToggleSelectAllFixes = document.getElementById('btnToggleSelectAllFixes');

    const liveExcelSheetThead = document.getElementById('liveExcelSheetThead');
    const liveExcelSheetTbody = document.getElementById('liveExcelSheetTbody');
    const sheetTabBtns = document.querySelectorAll('.sheet-tab-btn');

    const autoFixActionBar = document.getElementById('autoFixActionBar');
    const btnApplyApprovedFixes = document.getElementById('btnApplyApprovedFixes');

    let currentReviewReports = [];
    let activeFileIdx = 0;
    let activeSheetName = 'Timesheet';
    let allFixesSelected = true;

    if (btnTriggerReviewUpload) btnTriggerReviewUpload.addEventListener('click', () => reviewUploadFileInput.click());
    if (reviewUploadFileInput) reviewUploadFileInput.addEventListener('change', handleMultiReviewFileUpload);
    if (btnApplyApprovedFixes) btnApplyApprovedFixes.addEventListener('click', handleApplyApprovedFixes);
    if (btnToggleSelectAllFixes) btnToggleSelectAllFixes.addEventListener('click', toggleSelectAllFixes);

    sheetTabBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            sheetTabBtns.forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            activeSheetName = e.target.getAttribute('data-sheet');
            renderLiveSheetTable();
        });
    });

    async function handleMultiReviewFileUpload(e) {
        const files = Array.from(e.target.files || []);
        if (files.length === 0) return;

        reviewUploadStatusText.textContent = `⏳ AI đang chuẩn bị quét và phân tích ${files.length} tệp Timesheet...`;

        const allReports = [];
        let errorCount = 0;
        let lastErrorMsg = '';

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            reviewUploadStatusText.textContent = `⏳ AI đang quét và phân tích tệp (${i + 1}/${files.length}): ${file.name}...`;

            const formData = new FormData();
            formData.append('files', file);

            try {
                const res = await fetch('/api/review-multi-upload', {
                    method: 'POST',
                    body: formData
                });

                const contentType = res.headers.get('content-type') || '';
                if (!contentType.includes('application/json')) {
                    if (res.status === 413) {
                        throw new Error(`Tệp ${file.name} quá lớn (vượt quá 4.5MB giới hạn Vercel).`);
                    }
                    throw new Error(`Server báo lỗi (HTTP ${res.status}) khi rà soát ${file.name}.`);
                }

                const data = await res.json();
                if (!res.ok) {
                    throw new Error(data.error || `Review tệp ${file.name} thất bại.`);
                }

                if (data.reports && data.reports.length > 0) {
                    allReports.push(...data.reports);
                }
            } catch (err) {
                console.error(`Error reviewing file ${file.name}:`, err);
                errorCount++;
                lastErrorMsg = err.message;
            }
        }

        if (allReports.length === 0) {
            reviewUploadStatusText.textContent = `❌ Lỗi rà soát: ${lastErrorMsg || 'Không thể đọc được dữ liệu các tệp đã chọn.'}`;
            return;
        }

        currentReviewReports = allReports;
        activeFileIdx = 0;

        if (errorCount > 0) {
            reviewUploadStatusText.textContent = `⚠️ Đã rà soát ${allReports.length}/${files.length} tệp (Lỗi ${errorCount} tệp: ${lastErrorMsg})`;
        } else {
            reviewUploadStatusText.textContent = `✅ Đã hoàn tất rà soát ${allReports.length} tệp Timesheet!`;
        }

        renderExecutiveKPIMetrics();
        renderFileTabsNavigation();
        renderActiveFileWorkspace();
    }


    function renderExecutiveKPIMetrics() {
        let totalFiles = currentReviewReports.length;
        let totalMembers = 0;
        let totalErrors = 0;

        currentReviewReports.forEach(rep => {
            const sumLen = (rep.summary || []).length;
            totalMembers += sumLen || 1;

            const errs = (rep.anomalies || []).filter(a => a.type === 'danger' || a.type === 'warning');
            totalErrors += errs.length;
        });

        kpiTotalFiles.textContent = totalFiles;
        kpiTotalMembers.textContent = totalMembers;
        kpiTotalErrors.textContent = totalErrors;

        const compliance = totalMembers > 0 ? Math.max(0, Math.round(((totalMembers - totalErrors) / totalMembers) * 100)) : 100;
        kpiComplianceRate.textContent = `${compliance}%`;

        reviewMetricsBar.style.display = 'grid';
    }

    function renderFileTabsNavigation() {
        if (!currentReviewReports.length) {
            fileTabsNavigation.style.display = 'none';
            return;
        }

        fileTabsNavigation.innerHTML = currentReviewReports.map((rep, idx) => {
            const errs = (rep.anomalies || []).filter(a => a.type === 'danger' || a.type === 'warning').length;
            const badgeClass = errs > 0 ? 'bg-danger' : 'sg-holiday';
            const badgeText = errs > 0 ? `⚠️ ${errs} Lỗi` : `✅ Clean`;
            const isActive = idx === activeFileIdx ? 'active' : '';

            return `
                <button type="button" class="tab-btn file-tab-btn ${isActive}" onclick="switchActiveReviewFile(${idx})">
                    📄 ${rep.filename}
                    <span class="ph-tag ${badgeClass}" style="font-size:0.75rem; margin-left:0.3rem;">${badgeText}</span>
                </button>
            `;
        }).join('');

        fileTabsNavigation.style.display = 'flex';
    }

    window.switchActiveReviewFile = function(idx) {
        activeFileIdx = idx;
        renderFileTabsNavigation();
        renderActiveFileWorkspace();
    };

    function renderActiveFileWorkspace() {
        const report = currentReviewReports[activeFileIdx];
        if (!report) {
            activeFileWorkspace.style.display = 'none';
            autoFixActionBar.style.display = 'none';
            return;
        }

        activeFileWorkspace.style.display = 'block';
        renderFileDiffsList(report);
        renderLiveSheetTable();

        const hasAnyFixable = currentReviewReports.some(r => (r.anomalies || []).some(a => a.fix_action));
        autoFixActionBar.style.display = hasAnyFixable ? 'block' : 'none';
    }

    function renderFileDiffsList(report) {
        const anomalies = report.anomalies || [];
        const fileToken = report.file_token;

        if (!anomalies.length) {
            fileDiffsListContainer.innerHTML = `<div class="anomaly-item success">✅ Tệp này hợp lệ 100%! Không có lỗi nào.</div>`;
            return;
        }

        fileDiffsListContainer.innerHTML = anomalies.map(a => {
            const isFixable = a.fix_action !== null && a.fix_action !== undefined;

            let customControlHtml = '';
            if (isFixable) {
                const actType = a.fix_action.type;
                if (['replace_invalid_worktype', 'fix_weekend_worktype'].includes(actType)) {
                    customControlHtml = `
                        <div style="margin-top:0.5rem; display:flex; align-items:center; gap:0.5rem; background:rgba(0,0,0,0.25); padding:0.4rem 0.6rem; border-radius:var(--radius-sm); width:fit-content;">
                            <small style="color:var(--text-muted); font-weight:600;">✏️ Sửa thành loại:</small>
                            <select class="form-control custom-fix-input" data-fix-id="${a.id}" style="padding:0.2rem 0.5rem; font-size:0.8rem; height:auto; width:auto; display:inline-block; border-color:rgba(16,206,103,0.4);">
                                <option value="Project Task" ${a.proposed_value === 'Project Task' ? 'selected' : ''}>Project Task</option>
                                <option value="Weekend support" ${a.proposed_value === 'Weekend support' ? 'selected' : ''}>Weekend support</option>
                                <option value="Weekday support">Weekday support</option>
                                <option value="PH Support">PH Support</option>
                                <option value="Leave">Leave</option>
                                <option value="Public Holiday">Public Holiday</option>
                            </select>
                        </div>
                    `;
                } else if (['set_ph_hours_zero', 'set_leave_hours_zero', 'set_project_task_hours_8', 'set_ot_hours_default'].includes(actType)) {
                    const proposedHrs = parseFloat(a.proposed_value) || 0;
                    customControlHtml = `
                        <div style="margin-top:0.5rem; display:flex; align-items:center; gap:0.5rem; background:rgba(0,0,0,0.25); padding:0.4rem 0.6rem; border-radius:var(--radius-sm); width:fit-content;">
                            <small style="color:var(--text-muted); font-weight:600;">✏️ Sửa thành số giờ:</small>
                            <select class="form-control custom-fix-input" data-fix-id="${a.id}" style="padding:0.2rem 0.5rem; font-size:0.8rem; height:auto; width:auto; display:inline-block; border-color:rgba(16,206,103,0.4);">
                                <option value="0" ${proposedHrs === 0 ? 'selected' : ''}>0h (Chuẩn PH/Leave)</option>
                                <option value="8" ${proposedHrs === 8 ? 'selected' : ''}>8h (Chuẩn Project Task)</option>
                                <option value="4" ${proposedHrs === 4 ? 'selected' : ''}>4h (Nửa ngày / OT)</option>
                                <option value="2" ${proposedHrs === 2 ? 'selected' : ''}>2h (OT)</option>
                                <option value="6" ${proposedHrs === 6 ? 'selected' : ''}>6h (OT)</option>
                            </select>
                        </div>
                    `;
                } else if (actType === 'adjust_leave_balance') {
                    const proposedUpto = parseFloat(a.proposed_value) || 10;
                    customControlHtml = `
                        <div style="margin-top:0.5rem; display:flex; align-items:center; gap:0.5rem; background:rgba(0,0,0,0.25); padding:0.4rem 0.6rem; border-radius:var(--radius-sm); width:fit-content;">
                            <small style="color:var(--text-muted); font-weight:600;">✏️ Sửa Balance Upto thành:</small>
                            <input type="number" step="0.5" class="form-control custom-fix-input" data-fix-id="${a.id}" value="${proposedUpto}" style="padding:0.2rem 0.5rem; font-size:0.8rem; height:auto; width:80px; display:inline-block; border-color:rgba(16,206,103,0.4);">
                        </div>
                    `;
                }
            }

            return `
                <div class="anomaly-item ${a.type}" style="margin-bottom: 0.9rem; background: rgba(15,23,42,0.65); border: 1px solid var(--card-border); border-radius: var(--radius-sm); padding: 0.9rem;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:0.5rem;">
                        <div style="flex:1;">
                            <div style="display:flex; align-items:center; gap:0.5rem; margin-bottom:0.3rem;">
                                <span class="ph-tag ${a.type === 'danger' ? '' : (a.type === 'warning' ? '' : 'sg-holiday')}" style="font-size:0.75rem; font-weight:700;">${a.category}</span>
                                <small style="color:var(--text-muted); font-weight:600;">Sheet: ${a.sheet} | Row ${a.row_index}</small>
                            </div>
                            <div style="font-size:0.88rem; margin-bottom:0.5rem;">${a.message}</div>
                            
                            ${a.current_value && a.proposed_value ? `
                                <div style="display:flex; align-items:center; gap:0.6rem; font-size:0.82rem; background:rgba(0,0,0,0.3); padding:0.4rem 0.75rem; border-radius:var(--radius-sm); width:fit-content;">
                                    <span style="color:var(--danger); text-decoration:line-through;">${a.current_value}</span>
                                    <span style="color:var(--text-muted);">➔</span>
                                    <span style="color:var(--primary); font-weight:700;">${a.proposed_value}</span>
                                </div>
                            ` : ''}

                            ${customControlHtml}
                        </div>

                        ${isFixable ? `
                            <label style="display:flex; align-items:center; gap:0.4rem; cursor:pointer; background:rgba(16,206,103,0.15); border:1px solid rgba(16,206,103,0.4); padding:0.45rem 0.8rem; border-radius:var(--radius-sm); white-space:nowrap; font-size:0.8rem; font-weight:700; color:#fff;">
                                <input type="checkbox" class="fix-checkbox" data-file-token="${fileToken}" data-fix-id="${a.id}" data-file-idx="${activeFileIdx}" ${allFixesSelected ? 'checked' : ''}>
                                <span>Approve Fix</span>
                            </label>
                        ` : ''}
                    </div>
                </div>
            `;
        }).join('');
    }

    function toggleSelectAllFixes() {
        allFixesSelected = !allFixesSelected;
        btnToggleSelectAllFixes.textContent = allFixesSelected ? 'Bỏ Tích Tất Cả' : 'Tích Tất Cả Fixes';
        document.querySelectorAll('.fix-checkbox').forEach(chk => chk.checked = allFixesSelected);
    }

    function renderLiveSheetTable() {
        const report = currentReviewReports[activeFileIdx];
        if (!report) return;

        if (activeSheetName === 'Timesheet') {
            const tsSample = report.timesheet_sample || [];
            liveExcelSheetThead.innerHTML = `
                <tr>
                    <th>Ngày</th>
                    <th>Nhân sự</th>
                    <th>Work Item Type</th>
                    <th>Task Description</th>
                    <th>Actual Time (Số giờ)</th>
                </tr>
            `;
            liveExcelSheetTbody.innerHTML = tsSample.map(t => {
                const wtype_raw = str(t.work_item_type).trim();
                const wtype_lower = wtype_raw.toLowerCase();
                const hours = typeof t.hours === 'number' ? t.hours : (parseFloat(t.hours) || 0);

                const ACCEPTED = ['project task', 'leave', 'public holiday', 'ph support', 'weekend support', 'weekday support'];

                // 1. Invalid Worktype
                const isInvalidWorktype = !ACCEPTED.includes(wtype_lower);

                // 2. Public Holiday Hours Error (Rule: PH Actual Time MUST = 0)
                const isPHErr = wtype_lower === 'public holiday' && hours !== 0;

                // 3. Leave Hours Error (Rule: Leave Actual Time MUST = 0)
                const isLeaveErr = wtype_lower === 'leave' && hours > 0;

                // 4. Project Task Hours Error (Rule: Project Task Actual Time MUST = 8 or 2/4/6, CANNOT = 0)
                const isTaskErr = wtype_lower === 'project task' && hours === 0;

                // 5. OT Hours Error (Rule: OT Actual Time CANNOT = 0)
                const isOTErr = ['ph support', 'weekend support', 'weekday support'].includes(wtype_lower) && hours === 0;

                // 6. Weekend Worktype Error
                let isWeekendErr = false;
                if (t.date) {
                    try {
                        const dtObj = new Date(t.date);
                        if ((dtObj.getDay() === 0 || dtObj.getDay() === 6) && !['weekend support', 'ph support'].includes(wtype_lower)) {
                            isWeekendErr = true;
                        }
                    } catch(e){}
                }

                const isErr = isInvalidWorktype || isPHErr || isLeaveErr || isTaskErr || isOTErr || isWeekendErr;

                let errReason = '';
                if (isPHErr) errReason = `⚠️ LỖI: Public Holiday số giờ phải = 0h (Hiện tại: ${hours}h)`;
                else if (isLeaveErr) errReason = `⚠️ LỖI: Dòng Leave số giờ phải = 0h (Hiện tại: ${hours}h)`;
                else if (isTaskErr) errReason = `⚠️ LỖI: Project Task số giờ phải = 8h (Hiện tại: 0h)`;
                else if (isOTErr) errReason = `⚠️ LỖI: Dòng OT số giờ phải khác 0h`;
                else if (isInvalidWorktype) errReason = `⚠️ LỖI: Worktype '${wtype_raw}' không hợp lệ`;
                else if (isWeekendErr) errReason = `⚠️ LỖI: Làm ngày cuối tuần cần đổi sang Weekend support`;

                return `
                    <tr style="${isErr ? 'background:rgba(239,68,68,0.22); border-left:4px solid var(--danger);' : ''}">
                        <td>${t.date}</td>
                        <td><strong>${t.name}</strong></td>
                        <td>
                            <span class="ph-tag ${isInvalidWorktype || isWeekendErr ? '' : (wtype_raw === 'Project Task' ? 'sg-holiday' : '')}" style="${isInvalidWorktype || isWeekendErr ? 'background:rgba(239,68,68,0.3); color:#fca5a5;' : ''}">
                                ${t.work_item_type}
                            </span>
                        </td>
                        <td>${t.task || '--'}</td>
                        <td>
                            <strong style="color:${isPHErr || isLeaveErr || isTaskErr || isOTErr ? 'var(--danger)' : 'var(--text-main)'}">
                                ${t.hours}h
                            </strong>
                            ${errReason ? `<div style="font-size:0.75rem; color:#fca5a5; font-weight:700; margin-top:0.25rem; font-family:sans-serif;">${errReason}</div>` : ''}
                        </td>
                    </tr>
                `;
            }).join('');
        } else if (activeSheetName === 'Balance') {
            const balList = report.balance || [];
            liveExcelSheetThead.innerHTML = `
                <tr>
                    <th>Họ và tên</th>
                    <th>Tổng phép năm</th>
                    <th>Dư phép tích lũy</th>
                    <th>Dư phép cuối tháng</th>
                    <th>Trạng thái</th>
                </tr>
            `;
            liveExcelSheetTbody.innerHTML = balList.map(b => {
                const isNeg = typeof b.balance_in_month === 'number' && b.balance_in_month < 0;
                return `
                    <tr style="${isNeg ? 'background:rgba(239,68,68,0.22); border-left:4px solid var(--danger);' : ''}">
                        <td><strong>${b.name}</strong></td>
                        <td>${b.total_leave || 14}</td>
                        <td>${b.balance_upto || 10}</td>
                        <td><strong style="color:${isNeg ? 'var(--danger)' : 'var(--primary)'}">${b.balance_in_month}</strong></td>
                        <td>${isNeg ? '<span class="ph-tag">⚠️ Âm Phép</span>' : '<span class="ph-tag sg-holiday">Hợp Lệ</span>'}</td>
                    </tr>
                `;
            }).join('');
        } else {
            const sumList = report.summary || [];
            liveExcelSheetThead.innerHTML = `
                <tr>
                    <th>STT</th>
                    <th>Họ và tên</th>
                    <th>Team</th>
                    <th>Vị trí</th>
                    <th>Số ngày công</th>
                    <th>Tổng OT</th>
                    <th>Số ngày nghỉ</th>
                </tr>
            `;
            liveExcelSheetTbody.innerHTML = sumList.map(s => `
                <tr>
                    <td>${s.no || ''}</td>
                    <td><strong>${s.name}</strong></td>
                    <td>${s.team || ''}</td>
                    <td>${s.location || ''}</td>
                    <td><strong>${s.working_days || 0}</strong></td>
                    <td>${s.total_ot || 0}</td>
                    <td>${s.leaves || 0}</td>
                </tr>
            `).join('');
        }
    }

    async function handleApplyApprovedFixes() {
        if (!currentReviewReports.length) return;

        const approvedFilesMap = {};

        document.querySelectorAll('.fix-checkbox:checked').forEach(chk => {
            const fileToken = chk.getAttribute('data-file-token');
            const fileIdx = parseInt(chk.getAttribute('data-file-idx'));
            const fixId = chk.getAttribute('data-fix-id');

            const report = currentReviewReports[fileIdx];
            if (!report) return;

            const anomaly = (report.anomalies || []).find(a => a.id === fixId);
            if (!anomaly || !anomaly.fix_action) return;

            // Check if user edited the proposed fix value!
            const customInput = document.querySelector(`.custom-fix-input[data-fix-id="${fixId}"]`);
            let customVal = null;
            if (customInput) {
                customVal = customInput.value;
            }

            const fixActionToApply = { ...anomaly.fix_action };
            if (customVal !== null) {
                fixActionToApply.custom_value = customVal;
            }

            if (!approvedFilesMap[fileToken]) {
                approvedFilesMap[fileToken] = {
                    file_token: fileToken,
                    filename: report.filename,
                    fix_actions: []
                };
            }
            approvedFilesMap[fileToken].fix_actions.push(anomaly.fix_action);
        });

        const approved_files = Object.values(approvedFilesMap);
        if (approved_files.length === 0) {
            showNotification('Vui lòng tích chọn ít nhất 1 gợi ý sửa đổi (Approve Fix) để áp dụng.', 'warning');
            return;
        }

        showLoading('⚡ Đang tự động sửa lỗi theo phê duyệt và khởi tạo file sạch...');

        try {
            const res = await fetch('/api/apply-fixes-download', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ approved_files })
            });

            if (!res.ok) {
                const contentType = res.headers.get('content-type') || '';
                let errMsg = 'Lỗi áp dụng fix';
                if (contentType.includes('application/json')) {
                    const errData = await res.json().catch(() => ({}));
                    errMsg = errData.error || errMsg;
                } else {
                    errMsg = `Server báo lỗi (HTTP ${res.status})`;
                }
                throw new Error(errMsg);
            }

            const blob = await res.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = approved_files.length === 1 ? `Fixed_${approved_files[0].filename}` : `Fixed_Timesheets_Batch.zip`;
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);

            reviewUploadStatusText.textContent = `✅ Đã sửa lỗi & tải xuống file sạch thành công!`;
        } catch (err) {
            showNotification(`Lỗi: ${err.message}`, 'danger');
        } finally {
            hideLoading();
        }
    }

    function str(val) {
        return val ? String(val) : '';
    }

    function showLoading(msg) {
        loadingText.textContent = msg || 'Đang khởi tạo các file Excel Timesheet...';
        loadingOverlay.classList.add('active');
    }

    function hideLoading() {
        loadingOverlay.classList.remove('active');
    }

    init();
});
