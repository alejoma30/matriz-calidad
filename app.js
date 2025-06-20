document.addEventListener('DOMContentLoaded', function () {
  const asesoresPorProceso = {
    "PQRSD": ["Avila, Marcela Paola", "Beltran, Nicole Juliana", "Castro, Diana Angelica", "Fuentes, Sandra Milena", "Lopez, Angela Marcela", "Roa, Cindy Paola", "Sanchez, Hellen Viviana", "Nadya Eliscet Bernal Escobar"],
    "NOTIFICACIONES - PQRSD": ["Insuasty, Daniel Ismael"],
    "NOTIFICACIONES": ["Gomez, Natalia", "Gutierrez, Valentina", "Alvarez, Carlos William", "Garavito, Gabriela Alexandra", "Mahecha, Diego Andres", "Peña, Jairo Esteban", "Rincon, Nathaly Dayana", "Sandoval, Diego Mauricio", "Santamaria, Edinson Yesid", "Hernandez, Diego Andres", "John Edwar Olarte"],
    "LEGALIZACIONES": ["Castiblanco, Jonathan Javier", "Saavedra, Jenny Alexandra", "Ojeda, Maria Alejandra", "Rodriguez, Andrés Eduardo", "Ruiz, Daissy Katerine"],
    "ANTENCION PRESENCIAL": ["Alvarez, Katherine"]
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

  const radicadoInput = document.getElementById('radicado');
  radicadoInput.addEventListener('input', function () {
    this.value = this.value.replace(/\D/g, '');
  });

  document.getElementById('formulario').addEventListener('submit', function (e) {
    e.preventDefault();
    const nota = calcularNota();
    const semaforo = nota >= 90 ? '🟢 Excelente' : nota >= 80 ? '🟡 Aceptable' : '🔴 Debe mejorar';
    const evaluador = document.getElementById('evaluador').value;

    alert(`${evaluador}, la auditoría se ha guardado con éxito.\nNota: ${nota}%\n${semaforo}`);

    const data = {
      fechaAuditoria: document.getElementById('fecha-auditoria').value,
      fechaGestion: document.getElementById('fecha-gestion').value,
      proceso: document.getElementById('proceso').value,
      asesor: document.getElementById('asesor').value,
      evaluador: evaluador,
      radicado: document.getElementById('radicado').value,
      c1: document.getElementById('c1').value,
      c2: document.getElementById('c2').value,
      c3: document.getElementById('c3').value,
      c4: document.getElementById('c4').value,
      c5: document.getElementById('c5').value,
      c6: document.getElementById('c6').value,
      c7: document.getElementById('c7').value,
      c8: document.getElementById('c8').value,
      observaciones: document.getElementById('observaciones').value,
      retroalimentacion: document.getElementById('feedback').value,
      nota: nota,
      semaforo: semaforo
    };

    // ✅ Enviar a Google Sheets - URL ACTUALIZADA:
    fetch("https://script.google.com/macros/s/AKfycbx_NtKOyxjSeAeuvaEvYN-aGEB6d7luqi6RcdOaqN7J_245K3wDa2JVNfGHX1QaYoDZ/exec", {
      method: 'POST',
      body: JSON.stringify(data),
      headers: { 'Content-Type': 'application/json' }
    }).then(res => res.text())
      .then(resp => console.log("Enviado a Sheets:", resp))
      .catch(err => console.error("Error al enviar a Sheets:", err));

    this.reset();
    notaSpan.textContent = '100%';
  });

  document.getElementById('btnExportarExcel').addEventListener('click', function () {
    const headers = ["Fecha Auditoría", "Fecha Gestión", "Proceso", "Asesor", "Evaluador", "Radicado",
      "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8",
      "Observaciones", "Retroalimentación", "Nota", "Semáforo"];
    const fila = [
      document.getElementById('fecha-auditoria').value,
      document.getElementById('fecha-gestion').value,
      document.getElementById('proceso').value,
      document.getElementById('asesor').value,
      document.getElementById('evaluador').value,
      document.getElementById('radicado').value,
      document.getElementById('c1').value,
      document.getElementById('c2').value,
      document.getElementById('c3').value,
      document.getElementById('c4').value,
      document.getElementById('c5').value,
      document.getElementById('c6').value,
      document.getElementById('c7').value,
      document.getElementById('c8').value,
      document.getElementById('observaciones').value,
      document.getElementById('feedback').value,
      document.getElementById('nota').textContent,
      calcularNota() >= 90 ? '🟢 Excelente' : calcularNota() >= 80 ? '🟡 Aceptable' : '🔴 Debe mejorar'
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, fila]);
    XLSX.utils.book_append_sheet(wb, ws, "Monitoreo");
    XLSX.writeFile(wb, "monitoreo.xlsx");
  });
});
