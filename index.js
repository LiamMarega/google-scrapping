const axios = require('axios');
const fs = require('fs');
const ExcelJS = require('exceljs');

const apiUrl = `https://api.app.outscraper.com/maps/search`;

require('dotenv').config();

// Realiza la solicitud HTTP a la API de Places
axios
  .get(apiUrl, {
    params: {
      query: 'Dentistas,Entre Rios, Argentina',
      limit: 2,
      async: false,
      fields: 'query,name,full_address,phone,site,type,photo',
    },
    headers: {
      'X-API-KEY': process.env.API_KEY,
    },
  })
  .then((response) => {
    // Leer los datos del archivo JSON
    fs.readFile('data.json', 'utf8', (err, jsonString) => {
      if (err) {
        console.log('Error al leer el archivo desde el disco:', err);
        return;
      }
      try {
        const data = JSON.parse(jsonString);

        // Crear un nuevo libro de trabajo
        var wb = new ExcelJS.Workbook();

        // Crear una nueva hoja de trabajo
        var ws = wb.addWorksheet('Datos');

        // Definir el estilo
        var style = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFF00' }, // Color de fondo amarillo
        };

        // Definir el orden de las claves
        var keyOrder = ['name', 'full_address', 'phone', 'site', 'type', 'photo'];

        // Reordenar los objetos en el array de datos
        var reorderedData = data.data[0].map((obj) => {
          var newObj = {};
          keyOrder.forEach((key) => {
            if (obj.hasOwnProperty(key)) {
              // Cambiar la primera letra a mayúscula
              var newKey = key.charAt(0).toUpperCase() + key.slice(1);
              newObj[newKey] = obj[key];
            }
          });
          return newObj;
        });

        // Convertir los datos a formato de hoja de trabajo y aplicar el estilo
        reorderedData.forEach((obj, rowIndex) => {
          Object.keys(obj).forEach((key, colIndex) => {
            var cell = ws.getCell(rowIndex + 2, colIndex + 1);
            cell.value = obj[key];
          });
        });

        // Agregar los nombres de las claves en la primera fila y aplicar el estilo
        Object.keys(reorderedData[0]).forEach((key, colIndex) => {
          var cell = ws.getCell(1, colIndex + 1);
          cell.value = key;
          cell.style.fill = style;
        });

        // Escribir el libro de trabajo en un archivo .xlsx
        wb.xlsx.writeFile('data.xlsx').then(() => {
          console.log('Archivo creado exitosamente.');
        });
      } catch (err) {
        console.log('Error al analizar el JSON:', err);
      }
    });
  });
