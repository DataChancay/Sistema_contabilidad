const viewLinks = Array.from(document.querySelectorAll("[data-view-link]"));
const views = Array.from(document.querySelectorAll("[data-view]"));
const exportForm = document.querySelector("#export-form");
const startInput = document.querySelector("#start-date");
const endInput = document.querySelector("#end-date");
const exportButton = document.querySelector("#export-button");
const exportButtonLabel = document.querySelector(".button-label");
const message = document.querySelector("#form-message");
const animationPanel = document.querySelector("#export-animation-panel");
const animationContainer = document.querySelector("#export-animation");
const animationText = document.querySelector("#export-animation-text");
const sourceInputs = Array.from(document.querySelectorAll("input[name='source']"));
const consolidadoForm = document.querySelector("#consolidado-form");
const consolidadoMessage = document.querySelector("#consolidado-message");
const consolidadoButton = document.querySelector("#consolidado-button");
const consolidadoButtonLabel = document.querySelector(".consolidado-button-label");
const consolidadoSummaryPanel = document.querySelector("#consolidado-summary");
const consolidadoDownloadButton = document.querySelector("#download-consolidado-button");
const consolidadoFileInputs = Array.from(document.querySelectorAll("#consolidado-form input[type='file']"));
const consolidadoTypeInputs = Array.from(document.querySelectorAll("input[name='tipo_consolidado']"));
const xafiroUploadCard = document.querySelector("#upload-xafiro");
const xafiroFileInput = document.querySelector("#file-xafiro");
const xafiroSummaryItem = document.querySelector("#summary-xafiro");
const consolidadoEyebrow = document.querySelector("#consolidado-eyebrow");
const consolidadoSubtitle = document.querySelector("#consolidado-subtitle");
const consolidadoProfileType = document.querySelector("#consolidado-profile-type");
const csrfToken = document.querySelector("meta[name='csrf-token']")?.content ?? "";
if (!exportForm ||
    !startInput ||
    !endInput ||
    !exportButton ||
    !exportButtonLabel ||
    !message ||
    !animationPanel ||
    !animationContainer ||
    !animationText ||
    sourceInputs.length === 0 ||
    !consolidadoForm ||
    !consolidadoMessage ||
    !consolidadoButton ||
    !consolidadoButtonLabel ||
    !consolidadoSummaryPanel ||
    !consolidadoDownloadButton ||
    consolidadoFileInputs.length === 0 ||
    consolidadoTypeInputs.length === 0 ||
    !xafiroUploadCard ||
    !xafiroFileInput ||
    !xafiroSummaryItem ||
    !consolidadoEyebrow ||
    !consolidadoSubtitle ||
    !consolidadoProfileType) {
    throw new Error("No se pudo inicializar la aplicaci?n.");
}
let currentAnimation = null;
let consolidadoBlob = null;
let consolidadoFilename = "consolidado_resort.xlsx";
const getConsolidadoType = () => {
    return consolidadoTypeInputs.find((input) => input.checked)?.value === "asociacion" ? "asociacion" : "resort";
};
const toLocalIsoDate = (value) => {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
};
const showMessage = (element, text, kind) => {
    element.textContent = text;
    element.className = `form-message ${kind}`;
    element.hidden = false;
};
const clearMessage = (element) => {
    element.hidden = true;
    element.textContent = "";
    element.className = "form-message";
};
const setActiveView = (viewName) => {
    views.forEach((view) => {
        const active = view.dataset.view === viewName;
        view.hidden = !active;
        view.classList.toggle("active", active);
    });
    viewLinks.forEach((link) => {
        const active = link.dataset.viewLink === viewName;
        link.classList.toggle("active", active);
        if (active)
            link.setAttribute("aria-current", "page");
        else
            link.removeAttribute("aria-current");
    });
};
const viewFromHash = () => (window.location.hash === "#consolidado" ? "consolidado" : "exportables");
viewLinks.forEach((link) => {
    link.addEventListener("click", () => {
        const target = link.dataset.viewLink === "consolidado" ? "consolidado" : "exportables";
        setActiveView(target);
    });
});
window.addEventListener("hashchange", () => setActiveView(viewFromHash()));
setActiveView(viewFromHash());
const today = new Date();
const thirtyDaysAgo = new Date(today);
thirtyDaysAgo.setDate(today.getDate() - 30);
endInput.value = toLocalIsoDate(today);
startInput.value = toLocalIsoDate(thirtyDaysAgo);
endInput.max = toLocalIsoDate(today);
const playExportAnimation = (state, text) => {
    const path = state === "loading" ? animationPanel.dataset.loadingAnimation : animationPanel.dataset.successAnimation;
    animationPanel.hidden = false;
    animationText.textContent = text;
    currentAnimation?.destroy();
    animationContainer.replaceChildren();
    if (!path || !window.lottie)
        return;
    currentAnimation = window.lottie.loadAnimation({
        container: animationContainer,
        renderer: "svg",
        loop: state === "loading",
        autoplay: true,
        path,
    });
};
const hideExportAnimation = () => {
    currentAnimation?.destroy();
    currentAnimation = null;
    animationContainer.replaceChildren();
    animationPanel.hidden = true;
};
const setExportLoading = (loading) => {
    exportButton.disabled = loading;
    exportButton.classList.toggle("loading", loading);
    exportButtonLabel.textContent = loading ? "Generando reportes..." : "Exportar reportes";
};
const getSelectedSources = () => {
    return sourceInputs
        .filter((input) => input.checked)
        .map((input) => {
        if (input.value === "neo")
            return "neo";
        if (input.value === "culqi")
            return "culqi";
        if (input.value === "mifact")
            return "mifact";
        return "xafiro";
    });
};
const getExportEndpoint = (source) => {
    if (source === "neo")
        return "/api/neo/export";
    if (source === "culqi")
        return "/api/culqi/export";
    if (source === "mifact")
        return "/api/mifact/export";
    return "/api/xafiro/export";
};
const getSourceLabel = (source) => {
    if (source === "neo")
        return "Neo";
    if (source === "culqi")
        return "Culqi";
    if (source === "mifact")
        return "Mifact";
    return "Xafiro";
};
const updateSelectedSource = () => {
    sourceInputs.forEach((input) => input.closest(".source-card")?.classList.toggle("selected", input.checked));
    clearMessage(message);
};
const getFilename = (response, fallback) => {
    const disposition = response.headers.get("Content-Disposition") ?? "";
    const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
    const basicMatch = disposition.match(/filename=["']?([^"';]+)["']?/i);
    if (utf8Match?.[1])
        return decodeURIComponent(utf8Match[1]);
    if (basicMatch?.[1])
        return basicMatch[1].trim();
    return fallback;
};
const downloadBlob = (blob, filename) => {
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(objectUrl);
};
const exportSource = async (source) => {
    const response = await fetch(getExportEndpoint(source), {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-CSRFToken": csrfToken },
        body: JSON.stringify({ fecha_inicio: startInput.value, fecha_fin: endInput.value }),
    });
    if (!response.ok) {
        const data = (await response.json().catch(() => ({})));
        throw new Error(`${getSourceLabel(source)}: ${data.error ?? "No se pudo generar el reporte."}`);
    }
    const blob = await response.blob();
    downloadBlob(blob, getFilename(response, `facturacion_${source}_${startInput.value}_${endInput.value}.xlsx`));
    return source;
};
exportForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearMessage(message);
    if (!startInput.value || !endInput.value) {
        showMessage(message, "Selecciona una fecha inicial y una fecha final.", "error");
        return;
    }
    if (startInput.value > endInput.value) {
        showMessage(message, "La fecha inicial no puede ser posterior a la fecha final.", "error");
        startInput.focus();
        return;
    }
    const selectedSources = getSelectedSources();
    if (selectedSources.length === 0) {
        showMessage(message, "Selecciona al menos un sistema para exportar.", "error");
        return;
    }
    setExportLoading(true);
    playExportAnimation("loading", selectedSources.length > 1 ? "Generando exportables seleccionados..." : "Generando exportable...");
    try {
        const results = await Promise.allSettled(selectedSources.map(exportSource));
        const exported = results.filter((result) => result.status === "fulfilled").map((result) => getSourceLabel(result.value));
        const failed = results.filter((result) => result.status === "rejected").map((result) => (result.reason instanceof Error ? result.reason.message : "Ocurri\u00f3 un error inesperado."));
        if (failed.length > 0) {
            hideExportAnimation();
            throw new Error(exported.length > 0 ? `Se descarg? ${exported.join(" y ")}. Fall?: ${failed.join(" ")}` : failed.join(" "));
        }
        playExportAnimation("success", `Exportables listos: ${exported.join(" y ")}.`);
        showMessage(message, `Reportes descargados correctamente: ${exported.join(" y ")}.`, "success");
    }
    catch (error) {
        showMessage(message, error instanceof Error ? error.message : "Ocurri\u00f3 un error inesperado.", "error");
    }
    finally {
        setExportLoading(false);
    }
});
const resetConsolidadoDownload = () => {
    consolidadoBlob = null;
    consolidadoFilename = `consolidado_${getConsolidadoType()}.xlsx`;
    consolidadoDownloadButton.disabled = true;
};
const updateConsolidadoType = () => {
    const isAssociation = getConsolidadoType() === "asociacion";
    xafiroUploadCard.hidden = isAssociation;
    xafiroFileInput.disabled = isAssociation;
    xafiroFileInput.required = !isAssociation;
    xafiroSummaryItem.hidden = isAssociation;
    consolidadoEyebrow.textContent = isAssociation ? "ASOCIACIÓN" : "RESORT";
    consolidadoProfileType.textContent = isAssociation ? "Asociación" : "Resort";
    consolidadoSubtitle.textContent = isAssociation
        ? "Cruza comprobantes de Mifact contra operaciones de CulqiLink y CulqiFull."
        : "Cruza comprobantes de Xafiro y Mifact contra operaciones de CulqiLink y CulqiFull.";
    clearMessage(consolidadoMessage);
    renderSummary(null);
    resetConsolidadoDownload();
};
const updateFileLabels = () => {
    consolidadoFileInputs.forEach((input) => {
        const label = document.querySelector(`[data-file-name='${input.name}']`);
        label.textContent = input.files?.[0]?.name ?? "Sin archivo cargado";
        input.closest(".upload-card")?.classList.toggle("selected", Boolean(input.files?.length));
    });
};
const setConsolidadoLoading = (loading) => {
    consolidadoButton.disabled = loading;
    consolidadoButton.classList.toggle("loading", loading);
    consolidadoButtonLabel.textContent = loading ? "Procesando..." : "Procesar consolidaci\u00f3n";
};
const parseSummary = (response) => {
    const raw = response.headers.get("X-Consolidado-Summary");
    if (!raw)
        return null;
    try {
        return JSON.parse(decodeURIComponent(raw));
    }
    catch (_error) {
        return null;
    }
};
const renderSummary = (summary) => {
    if (!summary) {
        consolidadoSummaryPanel.hidden = true;
        return;
    }
    Object.entries(summary).forEach(([key, value]) => {
        const target = consolidadoSummaryPanel.querySelector(`[data-summary='${key}']`);
        if (target)
            target.textContent = String(value);
    });
    consolidadoSummaryPanel.hidden = false;
};
const validateConsolidadoFiles = () => {
    const requiredInputs = consolidadoFileInputs.filter((input) => !input.disabled);
    const missing = requiredInputs.filter((input) => !input.files || input.files.length === 0);
    if (missing.length > 0) {
        const messageText = getConsolidadoType() === "asociacion"
            ? "Carga los tres archivos: Mifact, CulqiLink y CulqiFull."
            : "Carga los cuatro archivos: Xafiro, Mifact, CulqiLink y CulqiFull.";
        showMessage(consolidadoMessage, messageText, "error");
        missing[0].focus();
        return false;
    }
    return true;
};
consolidadoForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearMessage(consolidadoMessage);
    renderSummary(null);
    resetConsolidadoDownload();
    if (!validateConsolidadoFiles())
        return;
    const formData = new FormData(consolidadoForm);
    setConsolidadoLoading(true);
    try {
        const response = await fetch("/api/consolidado/procesar", { method: "POST", headers: { "X-CSRFToken": csrfToken }, body: formData });
        if (!response.ok) {
            const errorBody = await response.text();
            let data = {};
            try {
                data = JSON.parse(errorBody);
            }
            catch (_error) {
            }
            throw new Error(data.error ?? `No se pudo procesar la consolidacion. (HTTP ${response.status})`);
        }
        consolidadoBlob = await response.blob();
        consolidadoFilename = getFilename(response, `consolidado_${getConsolidadoType()}.xlsx`);
        consolidadoDownloadButton.disabled = false;
        renderSummary(parseSummary(response));
        showMessage(consolidadoMessage, "Consolidado procesado correctamente.", "success");
    }
    catch (error) {
        showMessage(consolidadoMessage, error instanceof Error ? error.message : "Ocurri\u00f3 un error inesperado.", "error");
    }
    finally {
        setConsolidadoLoading(false);
    }
});
consolidadoDownloadButton.addEventListener("click", () => {
    if (consolidadoBlob)
        downloadBlob(consolidadoBlob, consolidadoFilename);
});
startInput.addEventListener("change", () => {
    endInput.min = startInput.value;
    clearMessage(message);
});
endInput.addEventListener("change", () => clearMessage(message));
sourceInputs.forEach((input) => input.addEventListener("change", updateSelectedSource));
consolidadoFileInputs.forEach((input) => {
    input.addEventListener("change", () => {
        clearMessage(consolidadoMessage);
        resetConsolidadoDownload();
        updateFileLabels();
    });
});
consolidadoTypeInputs.forEach((input) => input.addEventListener("change", updateConsolidadoType));
updateSelectedSource();
updateFileLabels();
updateConsolidadoType();
export {};
