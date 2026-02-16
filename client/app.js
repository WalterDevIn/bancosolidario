function getAPI() {
  return document.getElementById("apiSelect")?.value
    || "https://bancosolidario.onrender.com";
}

let suggestedPlanNumber = null;

function setInterest(value) {
  document.getElementById("tasaMensual").value = (value * 100) + "%";
}

function setCuotas(value) {
  document.getElementById("cuotas").value = value;
}

function parseInterest(value) {
  if (!value) return 0.08;
  value = value.replace("%", "");
  let num = Number(value);
  if (num > 1) return num / 100;
  return num;
}

function getSecondSaturdayOfCurrentMonth() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth();
  const firstDay = new Date(y, m, 1);
  const offset = (6 - firstDay.getDay() + 7) % 7;
  return new Date(y, m, 1 + offset + 7).toISOString().split("T")[0];
}

function setMonto(value) {
  document.getElementById("monto").value = parseMonto(value);
}

// Convierte "2M", "2.5M", "2000000" → número real
function parseMonto(value) {
  value = value.toString().trim().toUpperCase();

  if (value.endsWith("M")) {
    return parseFloat(value.replace("M", "")) * 1000000;
  }

  return Number(value);
}

async function suggestNextPlan() {
  const res = await fetch(`${getAPI()}/plans`);
  const plans = await res.json();

  if (!plans.length) {
    document.getElementById("planNumero").value = "1";
    return;
  }

  const max = Math.max(
    ...plans.map(p => Number(p.planNumero) || 0)
  );

  document.getElementById("planNumero").value = max + 1;
}


async function updateSuggestedPlan() {
  const res = await fetch(`${getAPI()}/plans`);
  const plans = await res.json();

  let next = 1;

  if (plans.length) {
    const max = Math.max(
      ...plans.map(p => Number(p.planNumero) || 0)
    );
    next = max + 1;
  }

  suggestedPlanNumber = next;

  const btn = document.getElementById("suggestBtn");
  btn.innerText = `${next}`;
}

function applySuggestedPlan() {
  if (suggestedPlanNumber !== null) {
    document.getElementById("planNumero").value = suggestedPlanNumber;
  }
}

document.getElementById("gestion").addEventListener("input", function () {
  this.value = this.value.toUpperCase();
});

document.getElementById("fechaDesembolso").value =
  getSecondSaturdayOfCurrentMonth();

document.getElementById("planForm").addEventListener("submit", async (e) => {
  e.preventDefault();

  const data = {
    gestion: document.getElementById("gestion").value.trim(),
    planNumero: document.getElementById("planNumero").value.trim(),

    nombre: nombre.value,
    dni: dni.value,
    monto: parseMonto(document.getElementById("monto").value),
    fechaDesembolso: fechaDesembolso.value,
    tasaMensual: parseInterest(tasaMensual.value),
    cuotas: Number(cuotas.value) || 24,
  };


  const res = await fetch(`${getAPI()}/plans`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });

  if (res.ok) {
    loadPlans();
    e.target.reset();
  }
});

async function loadPlans() {
  const searchInput = document.getElementById("searchInput");
  const search = searchInput ? searchInput.value.trim() : "";

  const url = search
    ? `${getAPI()}/plans?search=${encodeURIComponent(search)}`
    : `${getAPI()}/plans`;

  const res = await fetch(url);
  const plans = await res.json();

  plans.reverse();

  const container = document.getElementById("plansList");
  container.innerHTML = "";

  let totalAmount = 0;

  plans.forEach(plan => {
    totalAmount += Number(plan.monto);

    const card = document.createElement("div");
    card.className = "loan-card";

    card.innerHTML = `
      <div class="loan-top">
        <div class="loan-client">
          <strong>${plan.nombre}</strong>
            <div class="loan-meta">
              DNI: ${plan.dni}
            </div>

            <div style="margin-top:6px; display:flex; gap:8px;">
              <span class="badge badge-blue">
                Gestión ${plan.gestion}
              </span>
              <span class="badge badge-green">
                Plan ${plan.planNumero}
              </span>
            </div>

        </div>

        <div class="loan-actions">
          <button class="btn-view" onclick="viewPlan('${plan.id}')">Ver</button>
          <button class="btn-excel" onclick="downloadExcel('${plan.id}')">Excel</button>
          <button class="btn-delete" onclick="deletePlan('${plan.id}')">Eliminar</button>
        </div>
      </div>

      <div class="loan-financial">
        <div>
          <span>Monto</span>
          <strong>$${Number(plan.monto).toLocaleString()}</strong>
        </div>

        <div>
          <span>Cuotas</span>
          <strong>${plan.cuotas}</strong>
        </div>

        <div>
          <span>Interés Total</span>
          <strong>$${Number(plan.schedule.sumInteres).toLocaleString()}</strong>
        </div>

        <div>
          <span>Total a Pagar</span>
          <strong>$${Number(plan.schedule.sumTotal).toLocaleString()}</strong>
        </div>
      </div>
    `;

    container.appendChild(card);
  });

  document.getElementById("totalCount").innerText = plans.length;
  document.getElementById("totalAmount").innerText =
    "$" + totalAmount.toLocaleString();

  updateSuggestedPlan();
}



function downloadExcel(id) {
  window.open(`${getAPI()}/plans/${id}/excel`);
}

async function deletePlan(id) {
  await fetch(`${getAPI()}/plans/${id}`, { method: "DELETE" });
  loadPlans();
}

async function viewPlan(id) {
  const res = await fetch(`${getAPI()}/plans/${id}`);
  const plan = await res.json();

  document.getElementById("modalTitle").innerText = `Plan de ${plan.nombre}`;
  document.getElementById("modalBody").innerHTML = `
    <p><strong>Monto:</strong> $${plan.monto}</p>
    <p><strong>Cuotas:</strong> ${plan.cuotas}</p>
    <p><strong>Interés total:</strong> $${plan.schedule.sumInteres}</p>
    <p><strong>Total a pagar:</strong> $${plan.schedule.sumTotal}</p>
  `;

  document.getElementById("modal").classList.remove("hidden");
}

function closeModal() {
  document.getElementById("modal").classList.add("hidden");
}

loadPlans();
