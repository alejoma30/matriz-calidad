const asesoresPorProceso = {
  "PQRSD": ["Avila, Marcela Paola", "Beltran, Nicole Juliana", "Castro, Diana Angelica", "Fuentes, Sandra Milena", "Lopez, Angela Marcela", "Roa, Cindy Paola", "Sanchez, Hellen Viviana", "Nadya Eliscet Bernal Escobar"],
  "NOTIFICACIONES - PQRSD": ["Insuasty, Daniel Ismael"],
  "NOTIFICACIONES": ["Gomez, Natalia", "Gutierrez, Valentina", "Alvarez, Carlos William", "Garavito, Gabriela Alexandra", "Mahecha, Diego Andres", "Peña, Jairo Esteban", "Rincon, Nathaly Dayana", "Sandoval, Diego Mauricio", "Santamaria, Edinson Yesid", "Hernandez, Diego Andres", "John Edwar Olarte"],
  "LEGALIZACIONES": ["Castiblanco, Jonathan Javier", "Saavedra, Jenny Alexandra", "Ojeda, Maria Alejandra", "Rodriguez, Andrés Eduardo", "Ruiz, Daissy Katerine"],
  "ANTENCION PRESENCIAL": ["Alvarez, Katherine"]
};

const dimensiones = [
  { peso: 50, tipo: "ENC", precision: "Usuario Final", dimension: "Uso de plantillas", id: "c1" },
  { peso: 30, tipo: "ENC", precision: "Usuario Final", dimension: "Claridad del lenguaje", id: "c2" },
  { peso: 20, tipo: "ENC", precision: "Usuario Final", dimension: "Redacción – puntuación", id: "c3" },
  { peso: 100, tipo: "EC", precision: "Usuario Final", dimension: "Redacción – ortografía", id: "c4" },
  { peso: 100, tipo: "EC", precision: "Usuario Final", dimension: "Interpretación de la solicitud", id: "c5" },
  { peso: 100, tipo: "EC", precision: "Negocio / Entidad", dimension: "Oportunidad en la respuesta", id: "c6" },
  { peso: 100, tipo: "EC", precision: "Negocio / Entidad", dimension: "Pertinencia de la respuesta", id: "c7" },
  { peso: 100, tipo: "EC", precision: "Negocio / Entidad", dimension: "Desempeño", id: "c8" }
];

const tabla = document.getElementById("tablaDimensiones");

dimensiones.forEach(d => {
  const fila = document.createElement("tr");
  fila.innerHTML = `
    <td>${d.peso}%</td>
    <td>${d.tipo}</td>
    <td>${d.precision}</td>
    <td>${d.dimension}</td>
    <td>
      <select id="${d.id}">
        <option value="SI" selected>SI</option>
        <option value="NO">NO</option>
        <option value="N/A">N/A</option>
      </select>
    </td>
  `;
  tabla.appendChild(fila);
});

document.getElementById("proceso").addEventListener("change", function () {
  const asesores = asesoresPorProceso[this.value] || [];
  const select = document.getElementById("asesor");
  select.innerHTML = '<option value="">Seleccione asesor</option>';
  asesores.forEach(nombre => {
    const option = document.createElement("option");
    option.value = nombre;
    option.textContent = nombre;
    select.appendChild(option);
  });
});

function calcularNota() {
  let nota = 100;
  for (const d of dimensiones) {
    const valor = document.getElementById(d.id).value;
    if (valor === "NO") {
      if (d.tipo === "EC") return 0;
      nota -= d.peso;
    }
  }
  return Math.max(0, nota);
}

function mostrarNota() {
  const nota = calcularNota();
  document.getElementById("notaFinal").textContent = `Nota: ${nota}%`;
}

document.querySelectorAll("select").forEach(select => {
  select.addEventListener("change", mostrarNota);
});

document.getElementById("formulario").addEventListener("submit", function (e) {
  e.preventDefault();

  const nota = calcularNota();
  const evaluador = document.getElementById("evaluador").value;

  const registro = {
    fechaAuditoria: document.getElementById("fechaAuditoria").value,
    fechaGestion: document.getElementById("fechaGestion").value,
    proceso: document.getElementById("proceso").value,
    asesor: document.getElementById("asesor").value,
    evaluador: evaluador,
    radicado: document.getElementById("radicado").value,
    observaciones: document.getElementById("observaciones").value,
    feedback: document.getElementById("feedback").value,
    nota,
    c1: document.getElementById("c1").value,
    c2: document.getElementById("c2").value,
    c3: document.getElementById("c3").value,
    c4: document.getElementById("c4").value,
    c5: document.getElementById("c5").value,
    c6: document.getElementById("c6").value,
    c7: document.getElementById("c7").value,
    c8: document.getElementById("c8").value
  };

  const registros = JSON.parse(localStorage.getItem("registros")) || [];
  registros.push(registro);
  localStorage.setItem("registros", JSON.stringify(registros));

  alert(`${evaluador}, la auditoría se ha guardado con éxito.\nNota: ${nota}%`);
  this.reset();
  mostrarNota();
});

document.getElementById("btnExportarExcel").addEventListener("click", function () {
  const registros = JSON.parse(localStorage.getItem("registros")) || [];
  if (registros.length === 0) {
    alert("No hay evaluaciones guardadas.");
    return;
  }

  const headers = [
    "Fecha Auditoría", "Fecha Gestión", "Proceso", "Asesor", "Evaluador", "Radicado",
    "Uso de plantillas", "Claridad del lenguaje", "Redacción – puntuación",
    "Redacción – ortografía", "Interpretación de la solicitud", "Oportunidad en la respuesta",
    "Pertinencia de la respuesta", "Desempeño", "Nota", "Observaciones", "Retroalimentación"
  ];

  const data = registros.map(r => [
    r.fechaAuditoria, r.fechaGestion, r.proceso, r.asesor, r.evaluador, r.radicado,
    r.c1, r.c2, r.c3, r.c4, r.c5, r.c6, r.c7, r.c8,
    `${r.nota}%`, r.observaciones, r.feedback
  ]);

  const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Evaluaciones");

  XLSX.writeFile(wb, "monitoreos_calidad.xlsx");
});

mostrarNota();
