(function () {
  "use strict";

  const STORAGE_KEY = "mit-personalidad-juridica-v1";
  const AGREEMENT_SIGNING_DATE = "2025-11-08";
  const form = document.getElementById("case-form");
  const panels = [...document.querySelectorAll("[data-step-panel]")];
  const navButtons = [...document.querySelectorAll("[data-step]")];
  const membersBody = document.getElementById("members-body");
  const memberTemplate = document.getElementById("member-row-template");
  const checklistRoot = document.getElementById("checklist");
  const toast = document.getElementById("toast");

  const checklistDefinitions = [
    { id: "request", title: "Solicitud dirigida a la autoridad competente", detail: "La petición identifica la autoridad del MIT competente en la jurisdicción." },
    { id: "act_signed", title: "Acta constitutiva suscrita por todos los fundadores", detail: "Consta el acta de asamblea general constitutiva y las firmas de todos los miembros fundadores." },
    { id: "act_name", title: "Nombre de la organización", detail: "El acta contiene la denominación completa y esta es distinguible en los sistemas de registro." },
    { id: "act_founders", title: "Identificación de los miembros fundadores", detail: "Constan nombres y apellidos, nacionalidad y documento de identidad de cada fundador." },
    { id: "act_will", title: "Voluntad de constituir la organización", detail: "El acta expresa de forma clara la voluntad de los fundadores." },
    { id: "act_purposes", title: "Fines y objetivos generales", detail: "El acta describe los fines y objetivos generales propuestos." },
    { id: "act_board", title: "Directiva provisional", detail: "El acta identifica la nómina de la directiva provisional." },
    { id: "act_responsible", title: "Responsable del trámite y datos de contacto", detail: "Constan nombres, documento de identidad, teléfono, correo y domicilio para notificaciones." },
    { id: "act_statutes", title: "Estatutos aprobados por la asamblea", detail: "El acta deja constancia expresa de la aprobación de los estatutos." },
    { id: "act_domicile", title: "Domicilio de la organización", detail: "Constan calle o dirección, parroquia, cantón, provincia y un medio de contacto." },
    { id: "higher_degree", title: "Decisión de asociarse de organizaciones integrantes", detail: "Para segundo o tercer grado, se adjuntan actas certificadas de los máximos órganos de cada persona jurídica.", allowNA: true, higherDegreeOnly: true },
    { id: "statute_identity", title: "Estatuto con denominación, ámbito y domicilio", detail: "El estatuto define la denominación, el ámbito de acción y el domicilio." },
    { id: "statute_territory", title: "Alcance territorial", detail: "El estatuto determina el alcance territorial de la organización." },
    { id: "statute_purposes", title: "Fines, objetivos y voluntariado", detail: "Se detallan fines y objetivos y se precisa si realizará actividades o programas de voluntariado." },
    { id: "statute_structure", title: "Órgano de gobierno y estructura organizacional", detail: "El estatuto regula los órganos de gobierno, administración y representación." },
    { id: "statute_rights", title: "Derechos y obligaciones de los miembros", detail: "El estatuto establece derechos y obligaciones de sus integrantes." },
    { id: "statute_democracy", title: "Régimen democrático interno", detail: "Se regulan elecciones, participación, alternabilidad y rendición de cuentas." },
    { id: "statute_assemblies", title: "Convocatorias, instalación y decisiones", detail: "Constan las formas y épocas de convocatoria y los quórums de instalación y decisión." },
    { id: "statute_membership", title: "Inclusión y exclusión con debido proceso", detail: "El estatuto establece mecanismos de admisión y exclusión, defensa y decisión motivada." },
    { id: "statute_reform", title: "Procedimiento de reforma de estatutos", detail: "Se identifica el órgano, convocatoria, mayoría y procedimiento para reformar el estatuto." },
    { id: "statute_disputes", title: "Solución de controversias", detail: "El estatuto regula mecanismos internos y otras vías permitidas para resolver conflictos." },
    { id: "statute_assets", title: "Patrimonio", detail: "Se regula la conformación, administración y destino del patrimonio." },
    { id: "statute_dissolution", title: "Disolución y liquidación", detail: "Constan causales, procedimiento, liquidador y destino del remanente patrimonial." },
    { id: "board_list", title: "Nómina de la directiva provisional", detail: "Se adjunta la nómina como documento del expediente." },
    { id: "legal_domicile", title: "Domicilio legal", detail: "Se presenta el domicilio legal completo de la organización." },
    { id: "assets_declaration", title: "Declaración juramentada de patrimonio inicial", detail: "Se adjunta la declaración juramentada con el patrimonio inicial." },
  ];

  let currentStep = 1;
  let members = [];
  let checklistState = Object.fromEntries(checklistDefinitions.map((item) => [item.id, { status: "pendiente", observation: "" }]));
  let saveTimer;

  function emptyData() {
    return { members: [], checklist: checklistState };
  }

  function showToast(message) {
    toast.textContent = message;
    toast.classList.add("show");
    clearTimeout(showToast.timer);
    showToast.timer = setTimeout(() => toast.classList.remove("show"), 2400);
  }

  function fieldData() {
    return Object.fromEntries(new FormData(form).entries());
  }

  function fullData() {
    const data = fieldData();
    data.members = members.filter((member) => member.name || member.id);
    data.checklist = checklistDefinitions.map((definition) => ({
      ...definition,
      ...checklistState[definition.id],
    }));
    return data;
  }

  function applyData(data) {
    Object.entries(data || {}).forEach(([name, value]) => {
      if (name === "members" || name === "checklist") return;
      const field = form.elements.namedItem(name);
      if (field && typeof value === "string") field.value = value;
    });
    members = Array.isArray(data?.members) ? data.members : [];
    if (data?.checklist) {
      checklistDefinitions.forEach((item) => {
        const saved = Array.isArray(data.checklist) ? data.checklist.find((entry) => entry.id === item.id) : data.checklist[item.id];
        if (saved) checklistState[item.id] = { status: saved.status || "pendiente", observation: saved.observation || "" };
      });
    }
    renderMembers();
    syncHigherDegreeRequirement();
    renderChecklist();
    updateSummary();
    evaluateCase();
  }

  function scheduleSave() {
    document.getElementById("save-status").textContent = "Guardando…";
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(fullData()));
      document.getElementById("save-status").textContent = "Borrador guardado";
    }, 250);
  }

  function showStep(step) {
    currentStep = Math.max(1, Math.min(6, Number(step)));
    panels.forEach((panel) => panel.classList.toggle("active", Number(panel.dataset.stepPanel) === currentStep));
    navButtons.forEach((button) => button.classList.toggle("active", Number(button.dataset.step) === currentStep));
    document.getElementById("previous-step").disabled = currentStep === 1;
    const next = document.getElementById("next-step");
    next.textContent = currentStep === 6 ? "Volver al inicio" : "Continuar";
    if (currentStep === 6) evaluateCase();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function addMember(member = { name: "", id: "", nationality: "ECUATORIANA" }) {
    members.push(member);
    renderMembers();
    scheduleSave();
  }

  function renderMembers() {
    membersBody.textContent = "";
    members.forEach((member, index) => {
      const row = memberTemplate.content.firstElementChild.cloneNode(true);
      row.dataset.index = String(index);
      row.querySelector(".row-number").textContent = String(index + 1);
      row.querySelectorAll("[data-member-field]").forEach((input) => { input.value = member[input.dataset.memberField] || ""; });
      membersBody.appendChild(row);
    });
    document.getElementById("members-empty").hidden = members.length > 0;
    document.getElementById("member-count").textContent = `${members.length} ${members.length === 1 ? "miembro" : "miembros"}`;
  }

  function syncHigherDegreeRequirement() {
    const degree = form.elements.organizationDegree.value;
    if (degree === "primer") {
      checklistState.higher_degree = { status: "no_aplica", observation: "Organización de primer grado." };
    } else if (checklistState.higher_degree.status === "no_aplica") {
      checklistState.higher_degree = { status: "pendiente", observation: "" };
    }
  }

  function renderChecklist() {
    checklistRoot.textContent = "";
    checklistDefinitions.forEach((item, index) => {
      const state = checklistState[item.id];
      const wrapper = document.createElement("article");
      wrapper.className = `check-item${state.status === "no_cumple" ? " needs-observation" : ""}`;
      wrapper.dataset.checkId = item.id;
      const statuses = [
        ["cumple", "Cumple"],
        ["no_cumple", "No cumple"],
        ...(item.allowNA ? [["no_aplica", "No aplica"]] : []),
      ];
      wrapper.innerHTML = `
        <span class="check-index">${index + 1}</span>
        <div class="check-copy"><strong>${item.title}</strong><p>${item.detail}</p></div>
        <div class="check-controls">${statuses.map(([value, label]) => `<button type="button" class="status-button${state.status === value ? " selected" : ""}" data-status="${value}">${label}</button>`).join("")}</div>
        <div class="check-observation"><textarea aria-label="Observación para ${item.title}" placeholder="Indique qué falta, qué debe corregirse y cómo subsanarlo">${state.observation || ""}</textarea></div>`;
      checklistRoot.appendChild(wrapper);
    });
    updateChecklistProgress();
  }

  function updateChecklistProgress() {
    const applicable = checklistDefinitions.filter((item) => checklistState[item.id].status !== "no_aplica");
    const complete = applicable.filter((item) => checklistState[item.id].status === "cumple").length;
    document.getElementById("check-complete").textContent = `${complete} de ${applicable.length}`;
    document.getElementById("progress-bar").style.width = `${applicable.length ? (complete / applicable.length) * 100 : 0}%`;
  }

  function requiredMissing() {
    const missing = [];
    form.querySelectorAll("[required]").forEach((field) => {
      field.classList.remove("invalid");
      if (!String(field.value || "").trim()) {
        missing.push(field.closest("label")?.querySelector("span")?.textContent?.replace(" *", "") || field.name);
      }
    });
    if (!members.length || members.some((member) => !member.name || !member.id || !member.nationality)) missing.push("Nómina completa de miembros fundadores");
    return missing;
  }

  function evaluateCase() {
    const missing = requiredMissing();
    const failed = checklistDefinitions.filter((item) => checklistState[item.id].status === "no_cumple");
    const pending = checklistDefinitions.filter((item) => checklistState[item.id].status === "pendiente");
    const resultCard = document.getElementById("result-card");
    const resultTitle = document.getElementById("result-title");
    const resultDescription = document.getElementById("result-description");
    const summary = document.getElementById("missing-summary");
    resultCard.className = "result-card";

    let result = "pending";
    if (!missing.length && !pending.length && !failed.length) result = "approved";
    else if (!missing.length && failed.length) result = "observed";

    if (result === "approved") {
      resultCard.classList.add("approved");
      resultCard.querySelector(".result-icon").textContent = "✓";
      resultTitle.textContent = "Expediente conforme";
      resultDescription.textContent = "El checklist está completo. Se pueden generar los tres documentos de aprobación.";
    } else if (result === "observed") {
      resultCard.classList.add("observed");
      resultCard.querySelector(".result-icon").textContent = "!";
      resultTitle.textContent = "Expediente con observaciones";
      resultDescription.textContent = `${failed.length} ${failed.length === 1 ? "requisito requiere" : "requisitos requieren"} subsanación.`;
    } else {
      resultCard.classList.add("pending");
      resultCard.querySelector(".result-icon").textContent = "…";
      resultTitle.textContent = "Revisión pendiente";
      resultDescription.textContent = "Completa la información obligatoria y asigna un estado a cada requisito.";
    }

    const messages = [];
    if (missing.length) messages.push(`<strong>Datos por completar:</strong> ${missing.join(", ")}.`);
    if (pending.length) messages.push(`<strong>Requisitos pendientes:</strong> ${pending.map((item) => item.title).join(", ")}.`);
    summary.innerHTML = messages.join("<br>");
    summary.classList.toggle("show", messages.length > 0);
    document.querySelectorAll(".generate-approval").forEach((button) => { button.disabled = result !== "approved"; });
    document.getElementById("generate-observation").disabled = result !== "observed";
    return result;
  }

  function updateSummary() {
    const data = fieldData();
    document.getElementById("case-name").textContent = data.organizationName || "Sin nombre";
    document.getElementById("case-number").textContent = data.caseNumber || "Sin número de trámite";
  }

  function validateOriginalCaseDate() {
    const field = form.elements.originalCaseDate;
    const help = document.getElementById("original-date-help");
    const isAfterAgreement = Boolean(field.value && field.value > AGREEMENT_SIGNING_DATE);
    field.classList.toggle("invalid", isAfterAgreement);
    help.textContent = isAfterAgreement
      ? "Esta fecha es posterior al 8 de noviembre de 2025; no se aplicará la Disposición General Tercera."
      : "Úsala únicamente si el expediente fue presentado hasta el 8 de noviembre de 2025.";
    help.classList.toggle("date-warning", isAfterAgreement);
    return !isAfterAgreement;
  }

  async function generate(kind) {
    const result = evaluateCase();
    if (kind === "observation" && result !== "observed") return showToast("Registra al menos un requisito como no cumplido y completa los datos obligatorios.");
    if (kind !== "observation" && result !== "approved") return showToast("Completa el expediente antes de generar documentos de aprobación.");
    const button = kind === "observation" ? document.getElementById("generate-observation") : document.querySelector(`[data-document="${kind}"]`);
    const originalText = button.textContent;
    button.disabled = true;
    button.textContent = "Generando…";
    try {
      await window.MITDocuments.downloadDocument(kind, fullData());
      showToast("Documento Word generado correctamente.");
    } catch (error) {
      console.error(error);
      showToast("No se pudo generar el documento. Revise los datos e intente otra vez.");
    } finally {
      button.textContent = originalText;
      evaluateCase();
    }
  }

  form.addEventListener("input", (event) => {
    if (event.target.matches("[data-member-field]")) return;
    if (event.target.name === "originalCaseDate") validateOriginalCaseDate();
    updateSummary();
    scheduleSave();
  });

  form.elements.organizationDegree.addEventListener("change", () => {
    syncHigherDegreeRequirement();
    renderChecklist();
    scheduleSave();
  });

  membersBody.addEventListener("input", (event) => {
    const row = event.target.closest("tr");
    if (!row || !event.target.dataset.memberField) return;
    members[Number(row.dataset.index)][event.target.dataset.memberField] = event.target.value;
    scheduleSave();
  });

  membersBody.addEventListener("click", (event) => {
    const button = event.target.closest(".remove-member");
    if (!button) return;
    members.splice(Number(button.closest("tr").dataset.index), 1);
    renderMembers();
    scheduleSave();
  });

  checklistRoot.addEventListener("click", (event) => {
    const button = event.target.closest("[data-status]");
    if (!button) return;
    const wrapper = button.closest("[data-check-id]");
    const id = wrapper.dataset.checkId;
    checklistState[id].status = button.dataset.status;
    if (button.dataset.status === "cumple") checklistState[id].observation = "";
    renderChecklist();
    scheduleSave();
  });

  checklistRoot.addEventListener("input", (event) => {
    const wrapper = event.target.closest("[data-check-id]");
    if (!wrapper || event.target.tagName !== "TEXTAREA") return;
    checklistState[wrapper.dataset.checkId].observation = event.target.value;
    scheduleSave();
  });

  document.getElementById("add-member").addEventListener("click", () => addMember());
  navButtons.forEach((button) => button.addEventListener("click", () => showStep(button.dataset.step)));
  document.getElementById("previous-step").addEventListener("click", () => showStep(currentStep - 1));
  document.getElementById("next-step").addEventListener("click", () => showStep(currentStep === 6 ? 1 : currentStep + 1));
  document.querySelectorAll(".generate-approval").forEach((button) => button.addEventListener("click", () => generate(button.dataset.document)));
  document.getElementById("generate-observation").addEventListener("click", () => generate("observation"));

  document.getElementById("new-case").addEventListener("click", () => {
    if (!window.confirm("¿Desea iniciar un nuevo trámite? El borrador actual se eliminará de este navegador.")) return;
    localStorage.removeItem(STORAGE_KEY);
    form.reset();
    members = [];
    checklistState = Object.fromEntries(checklistDefinitions.map((item) => [item.id, { status: "pendiente", observation: "" }]));
    applyData(emptyData());
    addMember();
    showStep(1);
    showToast("Nuevo trámite iniciado.");
  });

  document.getElementById("export-case").addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(fullData(), null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `tramite_${(fieldData().caseNumber || "sin_numero").replace(/[^a-zA-Z0-9_-]+/g, "_")}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  });

  document.getElementById("import-case").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      form.reset();
      checklistState = Object.fromEntries(checklistDefinitions.map((item) => [item.id, { status: "pendiente", observation: "" }]));
      applyData(data);
      scheduleSave();
      showStep(1);
      showToast("Trámite importado correctamente.");
    } catch {
      showToast("El archivo seleccionado no contiene un trámite válido.");
    } finally {
      event.target.value = "";
    }
  });

  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    try { applyData(JSON.parse(saved)); } catch { applyData(emptyData()); }
  } else {
    applyData(emptyData());
    addMember();
  }
  validateOriginalCaseDate();
  showStep(1);
})();
