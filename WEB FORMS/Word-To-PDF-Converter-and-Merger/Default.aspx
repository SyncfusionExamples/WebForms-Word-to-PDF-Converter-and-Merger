<%@ Page Title="Word to PDF Converter and Merger" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Word_To_PDF_Converter_and_Merger._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <section class="doc-suite-container">

        <!-- Page Header -->
        <div class="doc-header-section">
            <h1 class="doc-header-title">Word to PDF Converter and Merger</h1>
            <p class="doc-header-subtitle">Convert, combine, and export documents seamlessly.</p>
        </div>

        <!-- Tab Navigation -->
        <div class="doc-nav-tabs" role="tablist">
            <button type="button" class="doc-tab-btn active" id="tabWordToPdf" role="tab" aria-selected="true">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                    <polyline points="14 2 14 8 20 8"></polyline>
                </svg>
                Word to PDF
            </button>
            <button type="button" class="doc-tab-btn" id="tabMergeToPdf" role="tab" aria-selected="false">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <rect x="3" y="3" width="7" height="7"></rect>
                    <rect x="14" y="3" width="7" height="7"></rect>
                    <rect x="3" y="14" width="7" height="7"></rect>
                    <rect x="14" y="14" width="7" height="7"></rect>
                </svg>
                Merge to PDF
            </button>
        </div>

        <!-- ============== WORD TO PDF PANEL ============== -->
        <asp:Panel ID="pnlWordToPdf" runat="server" CssClass="doc-card">

            <div id="wpDropzone" class="doc-dropzone" tabindex="0">
                <div class="dropzone-icon-wrap">
                    <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                        <polyline points="17 8 12 3 7 8"></polyline>
                        <line x1="12" y1="3" x2="12" y2="15"></line>
                    </svg>
                </div>
                <p class="dropzone-title">Drag and drop Word documents here</p>
                <p class="dropzone-sub">Supports <strong>.doc</strong>, <strong>.docx</strong>, <strong>.rtf</strong>, and <strong>.html</strong> files</p>
                <div class="dropzone-divider">OR</div>
                <button type="button" class="btn-primary-action" id="wpBrowseBtn">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
                    </svg>
                    Browse Files
                </button>
                <asp:FileUpload ID="fuWordFiles" runat="server" AllowMultiple="true"
                    CssClass="d-none" ClientIDMode="Static"
                    accept=".doc,.docx,.rtf,.htm,.html" />
            </div>

            <div id="wpAlertHost" runat="server"></div>

            <!-- Uploaded Files Section -->
            <div class="files-section" id="wpFilesSection" style="display: none;">
                <div class="text-center mb-4">
                    <asp:Button ID="btnConvertWordToPdf" runat="server"
                        CssClass="btn-primary-action btn-convert-lg"
                        Text="Convert to PDF"
                        OnClick="btnConvertWordToPdf_Click"
                        Enabled="false"
                        ClientIDMode="Static"
                        UseSubmitBehavior="true" />
                </div>
                <div class="files-section-header">
                    <h3 class="files-section-title">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
                        </svg>
                        Uploaded Files
                    </h3>
                </div>
                <div id="wpFilesList" class="files-list-container"></div>
            </div>

        </asp:Panel>

        <!-- ============== MERGE TO PDF PANEL ============== -->
        <asp:Panel ID="pnlMergeToPdf" runat="server" CssClass="doc-card" style="display: none;">

            <div id="mpDropzone" class="doc-dropzone" tabindex="0">
                <div class="dropzone-icon-wrap">
                    <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <rect x="3" y="3" width="7" height="7"></rect>
                        <rect x="14" y="3" width="7" height="7"></rect>
                        <rect x="3" y="14" width="7" height="7"></rect>
                        <rect x="14" y="14" width="7" height="7"></rect>
                    </svg>
                </div>
                <p class="dropzone-title">Drag and drop Word/PDF files here</p>
                <p class="dropzone-sub">Supports <strong>.doc</strong>, <strong>.docx</strong>, <strong>.rtf</strong>, <strong>.html</strong>, and <strong>.pdf</strong> files</p>
                <div class="dropzone-divider">OR</div>
                <button type="button" class="btn-primary-action" id="mpBrowseBtn">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
                    </svg>
                    Browse Files
                </button>
                <asp:FileUpload ID="fuMergeFiles" runat="server" AllowMultiple="true"
                    CssClass="d-none" ClientIDMode="Static"
                    accept=".doc,.docx,.rtf,.htm,.html,.pdf" />
            </div>

            <div id="mpAlertHost" runat="server"></div>

            <!-- Sortable Uploaded Files Section -->
            <div class="files-section" id="mpFilesSection" style="display: none;">
                <div class="text-center mb-4">
                    <asp:Button ID="btnMergeToPdf" runat="server"
                        CssClass="btn-primary-action btn-convert-lg"
                        Text="Merge to PDF"
                        OnClick="btnMergeToPdf_Click"
                        Enabled="false"
                        ClientIDMode="Static"
                        UseSubmitBehavior="true" />
                </div>
                <div class="files-section-header">
                    <h3 class="files-section-title">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <line x1="8" y1="6" x2="21" y2="6"></line>
                            <line x1="8" y1="12" x2="21" y2="12"></line>
                            <line x1="8" y1="18" x2="21" y2="18"></line>
                            <line x1="3" y1="6" x2="3.01" y2="6"></line>
                            <line x1="3" y1="12" x2="3.01" y2="12"></line>
                            <line x1="3" y1="18" x2="3.01" y2="18"></line>
                        </svg>
                        Uploaded Files (Sortable)
                    </h3>
                    <small class="text-muted">Drag to reorder or use arrows</small>
                </div>
                <div id="mpFilesList" class="files-list-container"></div>
                <div id="mpMergeHint" class="custom-alert custom-alert-warning mt-3" style="display: none;">
                    <div><strong>Add at least one more file to enable merging.</strong> Merging requires two or more documents.</div>
                </div>
            </div>

        </asp:Panel>

    </section>

    <!-- Loading Overlay -->
    <div id="loadingOverlay" class="loading-overlay">
        <div class="loading-card">
            <div class="loading-spinner"></div>
            <h4 class="loading-title" id="loadingTitle">Processing...</h4>
            <p class="loading-desc" id="loadingDesc">Please wait while your documents are being converted.</p>
        </div>
    </div>

    <!-- Hidden field used to convey the merge order to the server -->
    <asp:HiddenField ID="hfMergeOrder" runat="server" ClientIDMode="Static" />

    <script type="text/javascript">
        (function () {
            // =============== COMMON UTILITIES ===============
            function escapeHtml(str) {
                if (str == null) return '';
                return String(str).replace(/[&<>"']/g, function (m) {
                    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
                });
            }

            function formatBytes(bytes) {
                if (!bytes) return '0 B';
                var units = ['B', 'KB', 'MB', 'GB'];
                var i = 0;
                var val = bytes;
                while (val >= 1024 && i < units.length - 1) {
                    val = val / 1024;
                    i++;
                }
                return val.toFixed(val >= 10 ? 0 : 1) + ' ' + units[i];
            }

            function syncFileInput(inputEl, filesArray) {
                var dt;
                try {
                    dt = new DataTransfer();
                } catch (e) {
                    return;
                }
                filesArray.forEach(function (f) {
                    dt.items.add(f);
                });
                inputEl.files = dt.files;
            }

            function showAlert(host, message, type) {
                if (!host) return;
                var cssClass = 'custom-alert custom-alert-' + (type || 'info');
                host.innerHTML = '<div class="' + cssClass + '"><div><strong>' + escapeHtml(message) + '</strong></div>' +
                    '<button type="button" class="btn-close" onclick="this.parentNode.remove()"></button></div>';
            }

            function clearAlert(host) {
                if (host) host.innerHTML = '';
            }

            // =============== TAB SWITCHING ===============
            var tabWordToPdf = document.getElementById('tabWordToPdf');
            var tabMergeToPdf = document.getElementById('tabMergeToPdf');
            var pnlWordToPdf = document.getElementById('<%= pnlWordToPdf.ClientID %>');
            var pnlMergeToPdf = document.getElementById('<%= pnlMergeToPdf.ClientID %>');

            function activateTab(which) {
                if (which === 'word') {
                    tabWordToPdf.classList.add('active');
                    tabMergeToPdf.classList.remove('active');
                    pnlWordToPdf.style.display = '';
                    pnlMergeToPdf.style.display = 'none';
                    tabWordToPdf.setAttribute('aria-selected', 'true');
                    tabMergeToPdf.setAttribute('aria-selected', 'false');
                } else {
                    tabMergeToPdf.classList.add('active');
                    tabWordToPdf.classList.remove('active');
                    pnlWordToPdf.style.display = 'none';
                    pnlMergeToPdf.style.display = '';
                    tabWordToPdf.setAttribute('aria-selected', 'false');
                    tabMergeToPdf.setAttribute('aria-selected', 'true');
                }
            }

            tabWordToPdf.addEventListener('click', function () { activateTab('word'); });
            tabMergeToPdf.addEventListener('click', function () { activateTab('merge'); });

            // =============== WORD TO PDF LOGIC ===============
            var wpFiles = [];
            var wpDropzone = document.getElementById('wpDropzone');
            var wpBrowseBtn = document.getElementById('wpBrowseBtn');
            var wpInput = document.getElementById('<%= fuWordFiles.ClientID %>');
            var wpFilesList = document.getElementById('wpFilesList');
            var wpFilesSection = document.getElementById('wpFilesSection');
            var wpAlertHost = document.getElementById('<%= wpAlertHost.ClientID %>');
            var wpConvertBtn = document.getElementById('btnConvertWordToPdf');
            var wpValidExts = ['.doc', '.docx', '.rtf', '.htm', '.html'];

            function isValidWpExt(name) {
                name = name.toLowerCase();
                return wpValidExts.some(function (ext) { return name.endsWith(ext); });
            }

            function renderWpFiles() {
                wpFilesList.innerHTML = '';
                if (!wpFiles.length) {
                    wpFilesSection.style.display = 'none';
                    wpConvertBtn.disabled = true;
                    return;
                }
                wpFilesSection.style.display = '';
                wpConvertBtn.disabled = false;
                wpFiles.forEach(function (f, idx) {
                    var row = document.createElement('div');
                    row.className = 'file-row';
                    row.innerHTML =
                        '<div class="file-info-group">' +
                            '<span class="file-type-icon word">' +
                                '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
                                '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>' +
                                '<polyline points="14 2 14 8 20 8"></polyline></svg>' +
                            '</span>' +
                            '<div class="file-details">' +
                                '<span class="file-name">' + escapeHtml(f.name) + '</span>' +
                                '<span class="file-meta">' + formatBytes(f.size) + '</span>' +
                            '</div>' +
                        '</div>' +
                        '<div class="file-actions">' +
                            '<button type="button" class="btn-icon-action btn-icon-remove" data-idx="' + idx + '" title="Remove">' +
                                '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
                                '<polyline points="3 6 5 6 21 6"></polyline>' +
                                '<path d="M19 6l-2 14a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2L5 6"></path></svg>' +
                            '</button>' +
                        '</div>';
                    wpFilesList.appendChild(row);
                });

                Array.prototype.forEach.call(wpFilesList.querySelectorAll('.btn-icon-remove'), function (btn) {
                    btn.addEventListener('click', function () {
                        var i = parseInt(btn.getAttribute('data-idx'), 10);
                        wpFiles.splice(i, 1);
                        syncFileInput(wpInput, wpFiles);
                        clearAlert(wpAlertHost);
                        renderWpFiles();
                    });
                });
            }

            function handleWpSelected(fileList) {
                if (!fileList || !fileList.length) return;
                var invalid = [];
                for (var i = 0; i < fileList.length; i++) {
                    var f = fileList[i];
                    if (!isValidWpExt(f.name)) {
                        invalid.push(f.name);
                        continue;
                    }
                    var exists = wpFiles.some(function (x) { return x.name === f.name && x.size === f.size; });
                    if (exists) continue;
                    wpFiles.push(f);
                }
                syncFileInput(wpInput, wpFiles);
                clearAlert(wpAlertHost);
                if (invalid.length) {
                    showAlert(wpAlertHost, 'Unsupported file format: ' + invalid.join(', ') + '. Please upload only .doc, .docx, .rtf, or .html files.', 'danger');
                }
                renderWpFiles();
            }

            wpBrowseBtn.addEventListener('click', function (e) {
                e.stopPropagation();
                wpInput.click();
            });
            wpDropzone.addEventListener('click', function (e) {
                if (e.target === wpBrowseBtn || (wpBrowseBtn.contains && wpBrowseBtn.contains(e.target))) return;
                wpInput.click();
            });
            wpInput.addEventListener('change', function () { handleWpSelected(wpInput.files); });

            ['dragenter', 'dragover'].forEach(function (ev) {
                wpDropzone.addEventListener(ev, function (e) {
                    e.preventDefault(); e.stopPropagation();
                    wpDropzone.classList.add('drag-active');
                });
            });
            ['dragleave', 'drop'].forEach(function (ev) {
                wpDropzone.addEventListener(ev, function (e) {
                    e.preventDefault(); e.stopPropagation();
                    wpDropzone.classList.remove('drag-active');
                });
            });
            wpDropzone.addEventListener('drop', function (e) { handleWpSelected(e.dataTransfer.files); });

            if (wpConvertBtn) {
                wpConvertBtn.addEventListener('click', function () {
                    if (!wpFiles.length) return;
                    showLoading('Converting...', 'Please wait while your documents are being converted.');
                });
            }

            // =============== MERGE TO PDF LOGIC ===============
            var mpFiles = [];
            var mpDropzone = document.getElementById('mpDropzone');
            var mpBrowseBtn = document.getElementById('mpBrowseBtn');
            var mpInput = document.getElementById('<%= fuMergeFiles.ClientID %>');
            var mpFilesList = document.getElementById('mpFilesList');
            var mpFilesSection = document.getElementById('mpFilesSection');
            var mpAlertHost = document.getElementById('<%= mpAlertHost.ClientID %>');
            var mpMergeBtn = document.getElementById('btnMergeToPdf');
            var mpMergeHint = document.getElementById('mpMergeHint');
            var mpOrder = document.getElementById('hfMergeOrder');
            var mpValidExts = ['.doc', '.docx', '.rtf', '.htm', '.html', '.pdf'];

            function isValidMpExt(name) {
                name = name.toLowerCase();
                return mpValidExts.some(function (ext) { return name.endsWith(ext); });
            }

            function getMpFileType(name) {
                var n = name.toLowerCase();
                if (n.endsWith('.pdf')) return 'pdf';
                return 'word';
            }

            function updateMergeOrder() {
                if (!mpOrder) return;
                mpOrder.value = mpFiles.map(function (f) { return f.name; }).join('|');
            }

            function renderMpFiles() {
                mpFilesList.innerHTML = '';
                if (!mpFiles.length) {
                    mpFilesSection.style.display = 'none';
                    if (mpMergeHint) mpMergeHint.style.display = 'none';
                    mpMergeBtn.disabled = true;
                    updateMergeOrder();
                    return;
                }
                mpFilesSection.style.display = '';
                mpMergeBtn.disabled = mpFiles.length < 2;
                if (mpMergeHint) {
                    mpMergeHint.style.display = mpFiles.length === 1 ? '' : 'none';
                }
                updateMergeOrder();

                mpFiles.forEach(function (f, idx) {
                    var ftype = getMpFileType(f.name);
                    var row = document.createElement('div');
                    row.className = 'file-row';
                    row.setAttribute('draggable', 'true');
                    row.setAttribute('data-idx', idx);
                    var iconHtml = ftype === 'pdf'
                        ? '<span class="file-type-icon pdf"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg></span>'
                        : '<span class="file-type-icon word"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg></span>';

                    var upDisabled = idx === 0 ? 'disabled' : '';
                    var downDisabled = idx === mpFiles.length - 1 ? 'disabled' : '';

                    row.innerHTML =
                        '<div class="file-info-group">' +
                            '<span class="file-badge-index">' + (idx + 1) + '</span>' +
                            iconHtml +
                            '<div class="file-details">' +
                                '<span class="file-name">' + escapeHtml(f.name) + '</span>' +
                                '<span class="file-meta">' + formatBytes(f.size) + '</span>' +
                            '</div>' +
                        '</div>' +
                        '<div class="file-actions">' +
                            '<button type="button" class="btn-icon-action btn-move-up" data-idx="' + idx + '" ' + upDisabled + ' title="Move Up">' +
                                '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"></polyline></svg>' +
                            '</button>' +
                            '<button type="button" class="btn-icon-action btn-move-down" data-idx="' + idx + '" ' + downDisabled + ' title="Move Down">' +
                                '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>' +
                            '</button>' +
                            '<span class="drag-handle" title="Drag to reorder">' +
                                '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="6" r="1"></circle><circle cx="9" cy="12" r="1"></circle><circle cx="9" cy="18" r="1"></circle><circle cx="15" cy="6" r="1"></circle><circle cx="15" cy="12" r="1"></circle><circle cx="15" cy="18" r="1"></circle></svg>' +
                            '</span>' +
                            '<button type="button" class="btn-icon-action btn-icon-remove" data-idx="' + idx + '" title="Remove">' +
                                '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
                                '<polyline points="3 6 5 6 21 6"></polyline>' +
                                '<path d="M19 6l-2 14a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2L5 6"></path></svg>' +
                            '</button>' +
                        '</div>';
                    mpFilesList.appendChild(row);
                });

                Array.prototype.forEach.call(mpFilesList.querySelectorAll('.btn-move-up'), function (btn) {
                    btn.addEventListener('click', function () {
                        var i = parseInt(btn.getAttribute('data-idx'), 10);
                        if (i > 0) {
                            var tmp = mpFiles[i];
                            mpFiles[i] = mpFiles[i - 1];
                            mpFiles[i - 1] = tmp;
                            syncFileInput(mpInput, mpFiles);
                            renderMpFiles();
                        }
                    });
                });
                Array.prototype.forEach.call(mpFilesList.querySelectorAll('.btn-move-down'), function (btn) {
                    btn.addEventListener('click', function () {
                        var i = parseInt(btn.getAttribute('data-idx'), 10);
                        if (i < mpFiles.length - 1) {
                            var tmp = mpFiles[i];
                            mpFiles[i] = mpFiles[i + 1];
                            mpFiles[i + 1] = tmp;
                            syncFileInput(mpInput, mpFiles);
                            renderMpFiles();
                        }
                    });
                });
                Array.prototype.forEach.call(mpFilesList.querySelectorAll('.btn-icon-remove'), function (btn) {
                    btn.addEventListener('click', function () {
                        var i = parseInt(btn.getAttribute('data-idx'), 10);
                        mpFiles.splice(i, 1);
                        syncFileInput(mpInput, mpFiles);
                        clearAlert(mpAlertHost);
                        renderMpFiles();
                    });
                });

                // Drag-and-drop reordering
                var dragSrcIdx = null;
                Array.prototype.forEach.call(mpFilesList.querySelectorAll('.file-row'), function (row) {
                    row.addEventListener('dragstart', function (e) {
                        dragSrcIdx = parseInt(row.getAttribute('data-idx'), 10);
                        row.classList.add('dragging');
                        try { e.dataTransfer.setData('text/plain', String(dragSrcIdx)); } catch (_) { }
                    });
                    row.addEventListener('dragend', function () {
                        row.classList.remove('dragging');
                        dragSrcIdx = null;
                    });
                    row.addEventListener('dragover', function (e) {
                        e.preventDefault(); e.stopPropagation();
                    });
                    row.addEventListener('drop', function (e) {
                        e.preventDefault(); e.stopPropagation();
                        var targetIdx = parseInt(row.getAttribute('data-idx'), 10);
                        if (dragSrcIdx == null || isNaN(targetIdx) || dragSrcIdx === targetIdx) return;
                        var item = mpFiles.splice(dragSrcIdx, 1)[0];
                        mpFiles.splice(targetIdx, 0, item);
                        syncFileInput(mpInput, mpFiles);
                        renderMpFiles();
                    });
                });
            }

            function handleMpSelected(fileList) {
                if (!fileList || !fileList.length) return;
                var invalid = [];
                for (var i = 0; i < fileList.length; i++) {
                    var f = fileList[i];
                    if (!isValidMpExt(f.name)) {
                        invalid.push(f.name);
                        continue;
                    }
                    var exists = mpFiles.some(function (x) { return x.name === f.name && x.size === f.size; });
                    if (exists) continue;
                    mpFiles.push(f);
                }
                syncFileInput(mpInput, mpFiles);
                clearAlert(mpAlertHost);
                if (invalid.length) {
                    showAlert(mpAlertHost, 'Unsupported file format: ' + invalid.join(', ') + '. Please upload only .doc, .docx, .rtf, .html, or .pdf files.', 'danger');
                }
                renderMpFiles();
            }

            mpBrowseBtn.addEventListener('click', function (e) {
                e.stopPropagation();
                mpInput.click();
            });
            mpDropzone.addEventListener('click', function (e) {
                if (e.target === mpBrowseBtn || (mpBrowseBtn.contains && mpBrowseBtn.contains(e.target))) return;
                mpInput.click();
            });
            mpInput.addEventListener('change', function () { handleMpSelected(mpInput.files); });

            ['dragenter', 'dragover'].forEach(function (ev) {
                mpDropzone.addEventListener(ev, function (e) {
                    e.preventDefault(); e.stopPropagation();
                    mpDropzone.classList.add('drag-active');
                });
            });
            ['dragleave', 'drop'].forEach(function (ev) {
                mpDropzone.addEventListener(ev, function (e) {
                    e.preventDefault(); e.stopPropagation();
                    mpDropzone.classList.remove('drag-active');
                });
            });
            mpDropzone.addEventListener('drop', function (e) { handleMpSelected(e.dataTransfer.files); });

            if (mpMergeBtn) {
                mpMergeBtn.addEventListener('click', function () {
                    if (mpFiles.length < 2) {
                        showAlert(mpAlertHost, 'Please add at least two files to merge.', 'danger');
                        return;
                    }
                    showLoading('Merging...', 'Please wait while your documents are being merged.');
                });
            }

            // =============== LOADING OVERLAY ===============
            var overlay = document.getElementById('loadingOverlay');
            var overlayTitle = document.getElementById('loadingTitle');
            var overlayDesc = document.getElementById('loadingDesc');
            var pollInterval = null;

            function showLoading(title, desc) {
                if (overlayTitle) overlayTitle.textContent = title || 'Processing...';
                if (overlayDesc) overlayDesc.textContent = desc || 'Please wait while your documents are being processed.';
                overlay.classList.add('show');

                if (pollInterval) clearInterval(pollInterval);
                pollInterval = setInterval(function () {
                    tryHideOverlay();
                }, 500);
            }

            function tryHideOverlay() {
                if (document.cookie.indexOf('wp_download_done=') !== -1) {
                    overlay.classList.remove('show');
                    document.cookie = 'wp_download_done=; expires=Thu, 01 Jan 1970 00:00:00 GMT; path=/';
                    if (pollInterval) {
                        clearInterval(pollInterval);
                        pollInterval = null;
                    }
                }
            }
            window.addEventListener('pageshow', tryHideOverlay);
            window.addEventListener('focus', tryHideOverlay);

            // Initial render
            try { renderWpFiles(); } catch (e) { }
            try { renderMpFiles(); } catch (e) { }
        })();
    </script>

</asp:Content>
