// app.js

document.addEventListener('DOMContentLoaded', function () {
  const asesoresPorProceso = {
    "PQRSD": ["Avila, Marcela Paola", "Beltran, Nicole Juliana", "Castro, Diana Angelica", "Fuentes, Sandra Milena", "Lopez, Angela Marcela", "Roa, Cindy Paola", "Sanchez, Hellen Viviana", "Nadya Eliscet Bernal Escobar"],
    "NOTIFICACIONES - PQRSD": ["Insuasty, Daniel Ismael"],
    "NOTIFICACIONES": ["Gomez, Natalia", "Gutierrez, Valentina", "Alvarez, Carlos William", "Garavito, Gabriela Alexandra", "Mahecha, Diego Andres", "Peña, Jairo Esteban", "Rincon, Nathaly Dayana", "Sandoval, Diego Mauricio", "Santamaria, Edinson Yesid", "Hernandez, Diego Andres", "John Edwar Olarte"],
    "LEGALIZACIONES": ["Castiblanco, Jonathan Javier", "Saavedra, Jenny Alexandra", "Ojeda, Maria Alejandra", "Rodriguez, Andrés Eduardo", "Ruiz, Daissy Katerine"],
    "ANTENCION PRESENCIAL": ["Alvarez, Katherine"],
  };

  const lideresCalidad = ["Rene Alejandro Mayorga", "Andrea Guzman Botache"];

  const evaluadorSelect = document.getElementById('evaluador');
  lideresCalidad.forEach(nombre => {
    const option = document.createElement('option');
    option.value = nombre;
    option.textContent = nombre;
    evaluadorSelect.appendChild(option);
  });

  const procesoSelect = document.getElementById('proceso');
  const asesorSelect = document.getElementById('asesor');

  procesoSelect.addEventListener('change', () => {
    const proceso = procesoSelect.value;
    const asesores = asesoresPorProceso[proceso] || [];
    asesorSelect.innerHTML = '<option value="">-- Selecciona --</option>';
    asesores.forEach(nombre => {
      const option = document.createElement('option');
      option.value = nombre;
      option.textContent = nombre;
      asesorSelect.appendChild(option);
    });
  });

  const cumplimientoSelects = document.querySelectorAll('.cumplimiento');
  const notaSpan = document.getElementById('nota');
  const pesos = [50, 30, 20];

  const calcularNota = () => {
    let nota = 100;
    for (let i = 0; i < cumplimientoSelects.length; i++) {
      const val = cumplimientoSelects[i].value;
      if (val === 'NO') {
        if (i < 3) nota -= pesos[i];
        else return 0;
      }
    }
    return nota;
  };

  const updateNota = () => {
    const nota = calcularNota();
    notaSpan.textContent = `${nota}%`;
  };

  cumplimientoSelects.forEach(select => {
    select.addEventListener('change', updateNota);
  });

  document.getElementById('fecha-auditoria').valueAsDate = new Date();

  const guardarEnLocalStorage = (registro) => {
    const registros = JSON.parse(localStorage.getItem("monitoreos")) || [];
    registros.push(registro);
    localStorage.setItem("monitoreos", JSON.stringify(registros));
  };

  document.getElementById('formulario').addEventListener('submit', function (e) {
    e.preventDefault();
    const nota = calcularNota();
    const semaforo = nota >= 90 ? '🟢 Excelente' : nota >= 80 ? '🟡 Aceptable' : '🔴 Debe mejorar';
    const evaluador = document.getElementById('evaluador').value;

    const data = {
      "Fecha Auditoría": document.getElementById('fecha-auditoria').value,
      "Fecha Gestión": document.getElementById('fecha-gestion').value,
      "Proceso": document.getElementById('proceso').value,
      "Asesor": document.getElementById('asesor').value,
      "Evaluador": evaluador,
      "Radicado": document.getElementById('radicado').value,
      "Uso de plantillas": document.getElementById('c1').value,
      "Claridad del lenguaje": document.getElementById('c2').value,
      "Redacción – puntuación": document.getElementById('c3').value,
      "Redacción – ortografía": document.getElementById('c4').value,
      "Interpretación de la solicitud": document.getElementById('c5').value,
      "Oportunidad en la respuesta": document.getElementById('c6').value,
      "Pertinencia de la respuesta": document.getElementById('c7').value,
      "Desempeño": document.getElementById('c8').value,
      "Observaciones": document.getElementById('observaciones').value,
      "Retroalimentación": document.getElementById('feedback').value,
      "Nota": `${nota}%`,
      "Semáforo": semaforo
    };

    guardarEnLocalStorage(data);

    alert(`${evaluador}, la auditoría se ha guardado con éxito.\nNota: ${nota}%\n${semaforo}`);

    this.reset();
    notaSpan.textContent = '100%';
    document.getElementById('fecha-auditoria').valueAsDate = new Date();
  });

  document.getElementById('btnExportarExcel').addEventListener('click', function () {
    const registros = JSON.parse(localStorage.getItem("monitoreos")) || [];

    if (registros.length === 0) {
      alert("No hay monitoreos guardados para exportar.");
      return;
    }

    const headers = Object.keys(registros[0]);
    const filas = registros.map(obj => headers.map(header => obj[header]));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...filas]);
    XLSX.utils.book_append_sheet(wb, ws, "Monitoreos");

    XLSX.writeFile(wb, "monitoreos_calidad.xlsx");
  });
});
