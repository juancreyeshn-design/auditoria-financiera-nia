/*
 * Visor interactivo de cédulas del Expediente de Auditoría (87 cédulas).
 * Intercepta los enlaces hacia el repositorio GitHub
 * "juancreyeshn-design/expediente-auditoria-nia" y los abre en un
 * panel superpuesto dentro de la misma página del curso, en vez de
 * navegar a GitHub — así el usuario nunca pierde el módulo/tema en el
 * que estaba (scroll, progreso de autoevaluación, etc.).
 */
(function () {
  "use strict";

  var REPO_OWNER = "juancreyeshn-design";
  var REPO_NAME = "expediente-auditoria-nia";
  var REPO_ROOT_HTTPS = "https://github.com/" + REPO_OWNER + "/" + REPO_NAME;
  var RAW_BASE = "https://raw.githubusercontent.com/" + REPO_OWNER + "/" + REPO_NAME + "/main/";
  var API_TREE = "https://api.github.com/repos/" + REPO_OWNER + "/" + REPO_NAME + "/git/trees/main?recursive=1";
  var SHEETJS_SRC = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";

  var sheetJsPromise = null;
  var treePromise = null;
  var fileCache = Object.create(null);

  function loadSheetJs() {
    if (sheetJsPromise) return sheetJsPromise;
    sheetJsPromise = new Promise(function (resolve, reject) {
      if (window.XLSX) return resolve(window.XLSX);
      var s = document.createElement("script");
      s.src = SHEETJS_SRC;
      s.onload = function () { resolve(window.XLSX); };
      s.onerror = function () { reject(new Error("No se pudo cargar el lector de Excel (SheetJS).")); };
      document.head.appendChild(s);
    });
    return sheetJsPromise;
  }

  function injectStyles() {
    var css = ""
      + "#cv-overlay{position:fixed;inset:0;z-index:9999;display:none;background:rgba(11,29,44,.55);"
      + "align-items:center;justify-content:center;padding:16px;box-sizing:border-box;}"
      + "#cv-overlay.cv-open{display:flex;}"
      + "#cv-panel{background:var(--blanco,#fff);border-radius:var(--radio,10px);width:min(980px,100%);"
      + "max-height:88vh;display:flex;flex-direction:column;box-shadow:0 12px 40px rgba(11,29,44,.35);"
      + "font-family:\"Segoe UI\",Arial,Helvetica,sans-serif;overflow:hidden;}"
      + "#cv-header{background:linear-gradient(135deg,var(--azul-oscuro,#0b3d63),var(--azul,#124a76));"
      + "color:#fff;padding:14px 18px;display:flex;align-items:center;gap:12px;flex-shrink:0;}"
      + "#cv-header-titles{flex:1;min-width:0;}"
      + "#cv-title{font-weight:700;font-size:1rem;margin:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}"
      + "#cv-subtitle{font-size:.76rem;color:var(--dorado-claro,#f7e7c1);margin:2px 0 0;}"
      + "#cv-close{background:rgba(255,255,255,.15);border:none;color:#fff;width:32px;height:32px;border-radius:8px;"
      + "font-size:1.1rem;cursor:pointer;flex-shrink:0;line-height:1;}"
      + "#cv-close:hover{background:rgba(255,255,255,.3);}"
      + "#cv-nav{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;"
      + "background:var(--azul-claro,#e8f1f8);padding:8px 14px;flex-shrink:0;border-bottom:1px solid var(--borde,#dbe4ec);}"
      + "#cv-nav-index{background:var(--blanco,#fff);border:1px solid var(--borde,#dbe4ec);color:var(--azul-oscuro,#0b3d63);"
      + "padding:6px 12px;font-size:.8rem;font-weight:700;cursor:pointer;border-radius:8px;}"
      + "#cv-nav-index:hover{background:var(--azul-oscuro,#0b3d63);color:#fff;}"
      + "#cv-nav-pager{display:flex;align-items:center;gap:10px;}"
      + "#cv-nav-pager button{background:var(--blanco,#fff);border:1px solid var(--borde,#dbe4ec);"
      + "color:var(--azul-oscuro,#0b3d63);padding:6px 12px;font-size:.8rem;font-weight:600;cursor:pointer;border-radius:8px;}"
      + "#cv-nav-pager button:hover:not(:disabled){background:var(--azul-oscuro,#0b3d63);color:#fff;}"
      + "#cv-nav-pager button:disabled{opacity:.4;cursor:default;}"
      + "#cv-nav-position{font-size:.76rem;color:var(--texto-suave,#5a6b78);font-weight:600;white-space:nowrap;}"
      + "@media (max-width:600px){#cv-nav{justify-content:flex-start;}#cv-nav-pager{width:100%;justify-content:space-between;}}"
      + "#cv-tabs{display:flex;gap:4px;background:var(--fondo,#f4f7fa);padding:8px 12px 0;flex-shrink:0;"
      + "overflow-x:auto;border-bottom:1px solid var(--borde,#dbe4ec);}"
      + "#cv-tabs button{border:1px solid var(--borde,#dbe4ec);border-bottom:none;background:#fff;"
      + "color:var(--texto-suave,#5a6b78);padding:7px 14px;font-size:.82rem;font-weight:600;cursor:pointer;"
      + "border-radius:8px 8px 0 0;white-space:nowrap;}"
      + "#cv-tabs button.cv-active{background:var(--blanco,#fff);color:var(--azul-oscuro,#0b3d63);"
      + "border-color:var(--borde,#dbe4ec);position:relative;top:1px;}"
      + "#cv-body{flex:1;overflow:auto;padding:16px 18px;background:#fff;}"
      + "#cv-body table{border-collapse:collapse;font-size:.82rem;min-width:100%;}"
      + "#cv-body table td,#cv-body table th{border:1px solid var(--borde,#dbe4ec);padding:5px 9px;"
      + "vertical-align:top;white-space:pre-wrap;}"
      + "#cv-body table tr:first-child td{background:var(--azul-claro,#e8f1f8);font-weight:700;color:var(--azul-oscuro,#0b3d63);}"
      + "#cv-status{padding:40px 10px;text-align:center;color:var(--texto-suave,#5a6b78);font-size:.9rem;}"
      + "#cv-status.cv-error{color:#b02a2a;}"
      + "#cv-footer{border-top:1px solid var(--borde,#dbe4ec);padding:10px 18px;display:flex;gap:16px;"
      + "align-items:center;flex-wrap:wrap;flex-shrink:0;background:var(--fondo,#f4f7fa);}"
      + "#cv-footer a{font-size:.8rem;font-weight:700;color:var(--azul-oscuro,#0b3d63);text-decoration:none;}"
      + "#cv-footer a:hover{text-decoration:underline;}"
      + "#cv-footer span{font-size:.74rem;color:var(--texto-suave,#5a6b78);}"
      + ".cv-tree-folder{margin:0 0 14px;}"
      + ".cv-tree-folder h4{margin:0 0 6px;font-size:.85rem;color:var(--azul-oscuro,#0b3d63);"
      + "border-bottom:2px solid var(--dorado,#d4a017);padding-bottom:4px;}"
      + ".cv-tree-file{display:block;width:100%;text-align:left;background:none;border:none;padding:5px 4px;"
      + "font-size:.82rem;color:var(--texto,#1c2b36);cursor:pointer;border-radius:6px;}"
      + ".cv-tree-file:hover{background:var(--azul-claro,#e8f1f8);color:var(--azul-oscuro,#0b3d63);}"
      + "@media (max-width:600px){#cv-title{font-size:.9rem;}#cv-body{padding:12px;}}";
    var style = document.createElement("style");
    style.id = "cv-styles";
    style.textContent = css;
    document.head.appendChild(style);
  }

  function buildModal() {
    var overlay = document.createElement("div");
    overlay.id = "cv-overlay";
    overlay.innerHTML =
      '<div id="cv-panel" role="dialog" aria-modal="true" aria-labelledby="cv-title">' +
      '  <div id="cv-header">' +
      '    <div id="cv-header-titles">' +
      '      <p id="cv-title">Expediente de auditoría</p>' +
      '      <p id="cv-subtitle">expediente-auditoria-nia</p>' +
      "    </div>" +
      '    <button id="cv-close" title="Cerrar" aria-label="Cerrar">✕</button>' +
      "  </div>" +
      '  <div id="cv-nav">' +
      '    <button id="cv-nav-index" type="button">☰ Índice completo</button>' +
      '    <div id="cv-nav-pager">' +
      '      <button id="cv-nav-prev" type="button" disabled>← Anterior</button>' +
      '      <span id="cv-nav-position"></span>' +
      '      <button id="cv-nav-next" type="button" disabled>Siguiente →</button>' +
      "    </div>" +
      "  </div>" +
      '  <div id="cv-tabs"></div>' +
      '  <div id="cv-body"><div id="cv-status">Cargando…</div></div>' +
      '  <div id="cv-footer">' +
      "    <span>Contenido de solo lectura. El progreso de este módulo no se pierde al cerrar este panel.</span>" +
      "  </div>" +
      "</div>";
    document.body.appendChild(overlay);

    overlay.addEventListener("click", function (e) {
      if (e.target === overlay) closeModal();
    });
    document.getElementById("cv-close").addEventListener("click", closeModal);
    document.getElementById("cv-nav-index").addEventListener("click", function () {
      openFolderBrowser();
    });
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && overlay.classList.contains("cv-open")) closeModal();
    });
    return overlay;
  }

  var overlayEl, titleEl, subtitleEl, tabsEl, bodyEl;
  var navPagerEl, navPrevBtn, navNextBtn, navPositionEl;
  var scrollLockY = 0;
  var orderedFilesPromise = null;

  function ensureModal() {
    if (overlayEl) return;
    injectStyles();
    overlayEl = buildModal();
    titleEl = document.getElementById("cv-title");
    subtitleEl = document.getElementById("cv-subtitle");
    tabsEl = document.getElementById("cv-tabs");
    bodyEl = document.getElementById("cv-body");
    navPagerEl = document.getElementById("cv-nav-pager");
    navPrevBtn = document.getElementById("cv-nav-prev");
    navNextBtn = document.getElementById("cv-nav-next");
    navPositionEl = document.getElementById("cv-nav-position");
  }

  function getOrderedFiles() {
    if (orderedFilesPromise) return orderedFilesPromise;
    orderedFilesPromise = fetchTree().then(function (groups) {
      var folders = Object.keys(groups).sort();
      var files = [];
      folders.forEach(function (folder) {
        groups[folder].slice().sort().forEach(function (path) { files.push(path); });
      });
      return files;
    });
    return orderedFilesPromise;
  }

  function setPagerVisible(visible) {
    if (!navPagerEl) return;
    navPagerEl.style.display = visible ? "flex" : "none";
  }

  function updatePagerForFile(path) {
    setPagerVisible(true);
    navPositionEl.textContent = "";
    navPrevBtn.disabled = true;
    navNextBtn.disabled = true;
    navPrevBtn.onclick = null;
    navNextBtn.onclick = null;
    getOrderedFiles()
      .then(function (files) {
        var idx = files.indexOf(path);
        if (idx === -1) return;
        navPositionEl.textContent = "Cédula " + (idx + 1) + " de " + files.length;
        if (idx > 0) {
          navPrevBtn.disabled = false;
          navPrevBtn.onclick = function () { openFileViewer(files[idx - 1]); };
        }
        if (idx < files.length - 1) {
          navNextBtn.disabled = false;
          navNextBtn.onclick = function () { openFileViewer(files[idx + 1]); };
        }
      })
      .catch(function () { /* sin índice disponible: el pager queda deshabilitado */ });
  }

  function openModal() {
    ensureModal();
    scrollLockY = window.scrollY;
    overlayEl.classList.add("cv-open");
    document.body.style.position = "fixed";
    document.body.style.top = "-" + scrollLockY + "px";
    document.body.style.left = "0";
    document.body.style.right = "0";
  }

  function closeModal() {
    if (!overlayEl) return;
    overlayEl.classList.remove("cv-open");
    document.body.style.position = "";
    document.body.style.top = "";
    document.body.style.left = "";
    document.body.style.right = "";
    window.scrollTo(0, scrollLockY);
  }

  function setStatus(html, isError) {
    tabsEl.innerHTML = "";
    bodyEl.innerHTML = '<div id="cv-status"' + (isError ? ' class="cv-error"' : "") + ">" + html + "</div>";
  }

  function friendlyName(path) {
    var file = path.split("/").pop().replace(/\.xlsx$/i, "");
    return file.replace(/^([A-Z]{2}-[A-Z0-9]+)_/, "$1 · ").replace(/_/g, " ");
  }

  function fetchFileBytes(path) {
    if (fileCache[path]) return fileCache[path];
    var url = RAW_BASE + path.split("/").map(encodeURIComponent).join("/");
    fileCache[path] = fetch(url, { cache: "force-cache" }).then(function (r) {
      if (!r.ok) throw new Error("GitHub respondió " + r.status + " para este archivo.");
      return r.arrayBuffer();
    });
    return fileCache[path];
  }

  function renderWorkbook(workbook) {
    var sheetNames = workbook.SheetNames;
    tabsEl.innerHTML = "";
    sheetNames.forEach(function (name, i) {
      var btn = document.createElement("button");
      btn.textContent = name;
      if (i === 0) btn.className = "cv-active";
      btn.addEventListener("click", function () {
        Array.prototype.forEach.call(tabsEl.children, function (b) { b.className = ""; });
        btn.className = "cv-active";
        showSheet(workbook, name);
      });
      tabsEl.appendChild(btn);
    });
    showSheet(workbook, sheetNames[0]);
  }

  function showSheet(workbook, name) {
    var ws = workbook.Sheets[name];
    if (!ws) {
      bodyEl.innerHTML = '<div id="cv-status">Esta hoja no tiene contenido.</div>';
      return;
    }
    var html = window.XLSX.utils.sheet_to_html(ws, { id: "cv-sheet-table" });
    bodyEl.innerHTML = html;
  }

  function openFileViewer(path, displayTitle) {
    openModal();
    titleEl.textContent = displayTitle || friendlyName(path);
    subtitleEl.textContent = path;
    setStatus("Cargando cédula…");
    updatePagerForFile(path);

    Promise.all([loadSheetJs(), fetchFileBytes(path)])
      .then(function (results) {
        var XLSX = results[0];
        var buf = results[1];
        var workbook = XLSX.read(buf, { type: "array", cellStyles: true });
        renderWorkbook(workbook);
      })
      .catch(function (err) {
        setStatus(
          "No se pudo cargar esta cédula (" + (err && err.message ? err.message : "error de red") + "). " +
            "Cierre este panel e intente de nuevo.",
          true
        );
      });
  }

  function fetchTree() {
    if (treePromise) return treePromise;
    treePromise = fetch(API_TREE, { cache: "force-cache" }).then(function (r) {
      if (!r.ok) throw new Error("GitHub respondió " + r.status + " al listar el expediente.");
      return r.json();
    }).then(function (data) {
      var files = (data.tree || []).filter(function (n) {
        return n.type === "blob" && /\.xlsx$/i.test(n.path);
      }).map(function (n) { return n.path; });
      var groups = {};
      files.forEach(function (p) {
        var parts = p.split("/");
        var folder = parts.length > 1 ? parts[0] : "(raíz)";
        (groups[folder] = groups[folder] || []).push(p);
      });
      return groups;
    });
    return treePromise;
  }

  function openFolderBrowser() {
    openModal();
    titleEl.textContent = "Expediente de auditoría — 87 cédulas";
    subtitleEl.textContent = REPO_OWNER + "/" + REPO_NAME;
    setPagerVisible(false);
    setStatus("Cargando índice del expediente…");

    fetchTree()
      .then(function (groups) {
        tabsEl.innerHTML = "";
        var folders = Object.keys(groups).sort();
        var container = document.createElement("div");
        folders.forEach(function (folder) {
          var section = document.createElement("div");
          section.className = "cv-tree-folder";
          var h4 = document.createElement("h4");
          h4.textContent = folder + " (" + groups[folder].length + ")";
          section.appendChild(h4);
          groups[folder].forEach(function (path) {
            var b = document.createElement("button");
            b.type = "button";
            b.className = "cv-tree-file";
            b.textContent = friendlyName(path);
            b.addEventListener("click", function () { openFileViewer(path); });
            section.appendChild(b);
          });
          container.appendChild(section);
        });
        bodyEl.innerHTML = "";
        bodyEl.appendChild(container);
      })
      .catch(function (err) {
        setStatus(
          "No se pudo cargar el índice (" + (err && err.message ? err.message : "error de red") + "). " +
            "Cierre este panel e intente de nuevo.",
          true
        );
      });
  }

  function pathFromBlobHref(href) {
    var marker = "/blob/main/";
    var idx = href.indexOf(marker);
    if (idx === -1) return null;
    return decodeURIComponent(href.slice(idx + marker.length));
  }

  function isRepoRootHref(href) {
    return href === REPO_ROOT_HTTPS || href === REPO_ROOT_HTTPS + "/";
  }

  document.addEventListener("click", function (e) {
    var a = e.target.closest ? e.target.closest("a") : null;
    if (!a) return;
    var href = a.getAttribute("href") || "";
    if (href.indexOf(REPO_ROOT_HTTPS) !== 0) return;

    var path = pathFromBlobHref(href);
    if (path && /\.xlsx$/i.test(path)) {
      e.preventDefault();
      openFileViewer(path, a.textContent.trim());
      return;
    }
    if (isRepoRootHref(href)) {
      e.preventDefault();
      openFolderBrowser();
    }
  });
})();
