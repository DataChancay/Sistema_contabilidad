const form = document.querySelector("#export-form");
const startInput = document.querySelector("#start-date");
const endInput = document.querySelector("#end-date");
const button = document.querySelector("#export-button");
const buttonLabel = document.querySelector(".button-label");
const message = document.querySelector("#form-message");
const animationPanel = document.querySelector("#export-animation-panel");
const animationContainer = document.querySelector("#export-animation");
const animationText = document.querySelector("#export-animation-text");
const sourceInputs = Array.from(document.querySelectorAll("input[name='source']"));
if (!form ||
    !startInput ||
    !endInput ||
    !button ||
    !buttonLabel ||
    !message ||
    !animationPanel ||
    !animationContainer ||
    !animationText ||
    sourceInputs.length === 0) {
    throw new Error("No se pudo inicializar el formulario de exportación.");
}
let currentAnimation = null;
const toLocalIsoDate = (value) => {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
};
const today = new Date();
const thirtyDaysAgo = new Date(today);
thirtyDaysAgo.setDate(today.getDate() - 30);
endInput.value = toLocalIsoDate(today);
startInput.value = toLocalIsoDate(thirtyDaysAgo);
endInput.max = toLocalIsoDate(today);
const showMessage = (text, kind) => {
    message.textContent = text;
    message.className = `form-message ${kind}`;
    message.hidden = false;
};
const clearMessage = () => {
    message.hidden = true;
    message.textContent = "";
    message.className = "form-message";
};
const playExportAnimation = (state, text) => {
    const path = state === "loading"
        ? animationPanel.dataset.loadingAnimation
        : animationPanel.dataset.successAnimation;
    animationPanel.hidden = false;
    animationText.textContent = text;
    currentAnimation?.destroy();
    animationContainer.replaceChildren();
    if (!path || !window.lottie) {
        return;
    }
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
const setLoading = (loading) => {
    button.disabled = loading;
    button.classList.toggle("loading", loading);
    buttonLabel.textContent = loading ? "Generando reportes…" : "Exportar reportes";
};
const getSelectedSources = () => {
    return sourceInputs
        .filter((input) => input.checked)
        .map((input) => (input.value === "neo" ? "neo" : "xafiro"));
};
const getExportEndpoint = (source) => {
    return source === "neo" ? "/api/neo/export" : "/api/xafiro/export";
};
const getSourceLabel = (source) => {
    return source === "neo" ? "Neo" : "Xafiro";
};
const updateSelectedSource = () => {
    sourceInputs.forEach((input) => {
        input.closest(".source-card")?.classList.toggle("selected", input.checked);
    });
    clearMessage();
};
const getFilename = (response, source) => {
    const disposition = response.headers.get("Content-Disposition") ?? "";
    const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
    const basicMatch = disposition.match(/filename=["']?([^"';]+)["']?/i);
    if (utf8Match?.[1]) {
        return decodeURIComponent(utf8Match[1]);
    }
    if (basicMatch?.[1]) {
        return basicMatch[1].trim();
    }
    return `facturacion_${source}_${startInput.value}_${endInput.value}.xlsx`;
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
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            fecha_inicio: startInput.value,
            fecha_fin: endInput.value,
        }),
    });
    if (!response.ok) {
        const data = (await response.json().catch(() => ({})));
        throw new Error(`${getSourceLabel(source)}: ${data.error ?? "No se pudo generar el reporte."}`);
    }
    const blob = await response.blob();
    downloadBlob(blob, getFilename(response, source));
    return source;
};
form.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearMessage();
    if (!startInput.value || !endInput.value) {
        showMessage("Selecciona una fecha inicial y una fecha final.", "error");
        return;
    }
    if (startInput.value > endInput.value) {
        showMessage("La fecha inicial no puede ser posterior a la fecha final.", "error");
        startInput.focus();
        return;
    }
    const selectedSources = getSelectedSources();
    if (selectedSources.length === 0) {
        showMessage("Selecciona al menos un sistema para exportar.", "error");
        return;
    }
    setLoading(true);
    playExportAnimation("loading", selectedSources.length > 1 ? "Generando exportables seleccionados..." : "Generando exportable...");
    try {
        const results = await Promise.allSettled(selectedSources.map(exportSource));
        const exported = results
            .filter((result) => result.status === "fulfilled")
            .map((result) => getSourceLabel(result.value));
        const failed = results
            .filter((result) => result.status === "rejected")
            .map((result) => (result.reason instanceof Error ? result.reason.message : "Ocurrió un error inesperado."));
        if (failed.length > 0) {
            hideExportAnimation();
            throw new Error(exported.length > 0
                ? `Se descargó ${exported.join(" y ")}. Falló: ${failed.join(" ")}`
                : failed.join(" "));
        }
        playExportAnimation("success", `Exportables listos: ${exported.join(" y ")}.`);
        showMessage(`Reportes descargados correctamente: ${exported.join(" y ")}.`, "success");
    }
    catch (error) {
        const detail = error instanceof Error ? error.message : "Ocurrió un error inesperado.";
        showMessage(detail, "error");
    }
    finally {
        setLoading(false);
    }
});
startInput.addEventListener("change", () => {
    endInput.min = startInput.value;
    clearMessage();
});
endInput.addEventListener("change", clearMessage);
sourceInputs.forEach((input) => input.addEventListener("change", updateSelectedSource));
updateSelectedSource();
export {};
