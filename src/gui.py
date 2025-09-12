<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Procesador de Facturas</title>
  <link rel="icon" href="{{ url_for('static', filename='logo.ico') }}">
  <style>
    html, body {
      height: 100%;
      margin: 0;
    }
    body {
      display: flex;
      justify-content: center;
      align-items: center;
      background: #f7f9fc;
      font-family: Arial, sans-serif;
      color: #333;
    }
    .container {
      text-align: center;
      background: #fff;
      padding: 2rem;
      border-radius: 8px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.1);
      max-width: 500px;
      width: 100%;
    }
    h1 {
      color: #2c3e50;
    }
    label {
      font-weight: bold;
    }
    input, button {
      margin-top: 0.5rem;
      margin-bottom: 1rem;
      width: 100%;
      padding: 0.6rem;
      border-radius: 4px;
      border: 1px solid #ccc;
    }
    button {
      background: #3498db;
      color: white;
      border: none;
      cursor: pointer;
    }
    button:hover {
      background: #2980b9;
    }
    #resetBtn {
      background: #e74c3c;
    }
    #resetBtn:hover {
      background: #c0392b;
    }
    #status {
      margin-top: 1rem;
      font-style: italic;
    }
    #counters {
      margin-top: 1.5rem;
      background: #ecf0f1;
      padding: 1rem;
      border-radius: 6px;
      text-align: left;
      white-space: pre-line; /* respeta saltos de línea */
      font-family: monospace;
    }
    progress {
      width: 100%;
      height: 20px;
      margin-top: 1rem;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>Procesador de Facturas</h1>
    <form id="form">
      <div>
        <label>Selecciona la carpeta con tus archivos:</label><br>
        <input type="file" name="folder" webkitdirectory multiple required>
      </div>
      <div>
        <label>Nombre del archivo de salida:</label><br>
        <input type="text" name="output_name" placeholder="Excel_final">
      </div>
      <button type="submit">Procesar y descargar</button>
    </form>

    <button id="resetBtn">Limpiar y subir otro</button>

    <progress id="progress" value="0" max="100" style="display:none;"></progress>
    <p id="status"></p>
    <div id="counters"></div>
  </div>

  <script>
    const form = document.getElementById('form');
    const status = document.getElementById('status');
    const countersDiv = document.getElementById('counters');
    const progress = document.getElementById('progress');
    const resetBtn = document.getElementById('resetBtn');

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      status.textContent = 'Procesando...';
      countersDiv.innerHTML = '';
      progress.style.display = 'block';
      progress.value = 10;

      // Simula avance
      let interval = setInterval(() => {
        if (progress.value < 90) progress.value += 10;
      }, 400);

      const formData = new FormData(e.target);
      const res = await fetch('/process-folder', { method: 'POST', body: formData });

      clearInterval(interval);

      if (!res.ok) {
        status.textContent = 'Error al procesar';
        progress.style.display = 'none';
        return;
      }

      const blob = await res.blob();
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = res.headers.get('Content-Disposition').split('filename=')[1].replace(/"/g, '');
      link.click();

      progress.value = 100;
      status.textContent = '¡Procesado con éxito!';

      // === Construcción del resumen estilo PyQt ===
      const total = res.headers.get('X-Counter-Total') || 0;
      const ie = res.headers.get('X-Counter-I/E') || 0;
      const p = res.headers.get('X-Counter-P') || 0;
      const n = res.headers.get('X-Counter-N') || 0;
      const desconocido = res.headers.get('X-Counter-Desconocido') || 0;

      const resumen = 
        "Resumen de procesamiento:\n" +
        `Total XML procesados: ${total}\n` +
        `I/E: ${ie}, P: ${p}, N: ${n}, Desconocidos: ${desconocido}`;

      countersDiv.textContent = resumen; // lo muestra en la interfaz
      alert("Éxito: El procesamiento se completó con éxito.\n\n" + resumen); // popup

      progress.style.display = 'none';
    });

    // Limpiar todo
    resetBtn.addEventListener('click', () => {
      form.reset();
      status.textContent = '';
      countersDiv.innerHTML = '';
      progress.style.display = 'none';
      progress.value = 0;
    });
  </script>
</body>
</html>
