importScripts(
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
  );
  
  self.onmessage = async function(e) {
  
    const { action, buffer, sheets } = e.data;
  
    try {
  
      if (action === 'readGram') {
  
        self.postMessage({
          type: 'progress',
          percent: 12,
          message: 'Leyendo GRAM EAM',
          detail: 'Procesando el archivo de Solicitudes de Trabajo…'
        });
  
        const wb = XLSX.read(buffer, {
          type: 'array',
          cellDates: false,
          raw: true
        });
  
        const sheetName = wb.SheetNames[0];
  
        const ws = wb.Sheets[sheetName];
  
        const aoa = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          raw: true,
          defval: null
        });
  
        self.postMessage({
          type: 'progress',
          percent: 20,
          message: 'GRAM EAM cargado',
          detail: aoa.length.toLocaleString('es-PE') + ' filas encontradas.'
        });
  
        self.postMessage({
          type: 'gram',
          aoa: aoa
        });
  
        return;
      }
  
  
      if (action === 'readBase') {
  
        self.postMessage({
          type: 'progress',
          percent: 42,
          message: 'Procesando.....',
          detail: 'Procesando BD 2022 – BD 2026…aguanta un momento mi king'
        });
  
        const wb = XLSX.read(buffer, {
          type: 'array',
          cellDates: true,
          sheets: sheets
        });
  
        const result = {};
  
        sheets.forEach(sheetName => {
  
          if (!wb.Sheets[sheetName]) return;
  
          const ws = wb.Sheets[sheetName];
  
          result[sheetName] = XLSX.utils.sheet_to_json(ws, {
            header: 1,
            raw: true,
            defval: null
          });
  
        });
  
        self.postMessage({
          type: 'progress',
          percent: 55,
          message: 'Base de Seguimiento cargada',
          detail: 'Hojas procesadas correctamente.'
        });
  
        self.postMessage({
          type: 'base',
          sheets: result
        });
  
        return;
      }
  
    } catch (err) {
  
      self.postMessage({
        type: 'error',
        message: err.message || String(err)
      });
  
    }
  
  };