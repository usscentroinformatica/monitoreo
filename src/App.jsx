import React, { useState, useMemo } from "react";
import ExcelJS from "exceljs";
import ControlPanel from "./components/ControlPanel";
import DataTable from "./components/DataTable";

function App() {
  const [data, setData] = useState([]);
  const [zoomData, setZoomData] = useState([]);
  const [selectedDocente, setSelectedDocente] = useState("");
  const [numFilas, setNumFilas] = useState("");
  const [isLoading, setIsLoading] = useState(false);

  // ===== FUNCIONES DE UTILIDAD =====
  const normalizeDocenteName = (name) => {
    if (!name) return "";
    return name.toUpperCase().trim().split(/\s+/).sort().join(" ");
  };

  const normalizeCursoName = (name) => {
  if (!name) return "";
  
  // Primero convertir números romanos a arábigos
  let normalized = convertRomanToArabic(name);
  
  return normalized
    .toUpperCase()
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ")
    .split(/\s+/)
    .filter(word => word.length > 1)
    .join(" ");
};

  const convertRomanToArabic = (text) => {
  if (!text) return text;
  
  const romanToArabic = {
    'II': '2',
    'III': '3',
    'IV': '4',
    'V': '5'
  };
  
  let result = text;
  Object.keys(romanToArabic).forEach(roman => {
    // Reemplazar números romanos que estén como palabra completa
    const regex = new RegExp(`\\b${roman}\\b`, 'gi');
    result = result.replace(regex, romanToArabic[roman]);
  });
  
  return result;
};

  const matchDocente = (docenteExcel, docenteZoom) => {
    const normalizedExcel = normalizeDocenteName(docenteExcel);
    const normalizedZoom = normalizeDocenteName(docenteZoom);
    
    if (normalizedExcel === normalizedZoom) return true;
    
    const wordsExcel = normalizedExcel.split(" ");
    const wordsZoom = normalizedZoom.split(" ");
    const commonWords = wordsExcel.filter(word => wordsZoom.includes(word));
    
    return commonWords.length >= 2;
  };

  const extractDate = (dateTimeStr) => {
    if (!dateTimeStr) return "";
    const match = dateTimeStr.match(/^([A-Za-z]+\s+\d{1,2},\s+\d{4})/);
    return match ? match[1] : "";
  };

  const extractTime = (dateTimeStr) => {
    if (!dateTimeStr) return "";
    
    let match = dateTimeStr.match(/(\d{1,2}:\d{2}:\d{2}\s*[AP]M)/i);
    if (match) return match[1];
    
    match = dateTimeStr.match(/(\d{1,2}:\d{2}:\d{2}\s*[ap]\.\s*m\.)/i);
    if (match) return match[1];
    
    match = dateTimeStr.match(/^(\d{1,2}:\d{2}:\d{2})/);
    if (match) return match[1];
    
    return dateTimeStr;
  };

  const extractCursoFromTema = (tema) => {
    if (!tema) return "";
    const match = tema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
    return match ? match[1].trim() : tema;
  };

  // ===== HANDLERS =====
  const handleZoomCsvUpload = async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  if (!selectedDocente) {
    alert("⚠️ Por favor, primero selecciona un docente en el filtro antes de cargar el CSV de Zoom");
    event.target.value = "";
    return;
  }

  setIsLoading(true);
  try {
    const text = await file.text();
    
    const lines = text.split('\n').filter(line => line.trim());
    let delimiter = ';';
    
    if (!lines[0].includes(';')) {
      delimiter = lines[0].includes('\t') ? '\t' : ',';
    }
    
    const headers = lines[0].split(delimiter).map(h => h.trim().replace(/^"|"$/g, ''));
    
    console.log("Delimitador detectado:", delimiter);
    console.log("Headers detectados:", headers);
    
    const parsedZoomData = [];
    for (let i = 1; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      
      const values = lines[i].split(delimiter).map(v => v.trim().replace(/^"|"$/g, ''));
      const row = {};
      headers.forEach((header, index) => {
        row[header] = values[index] || "";
      });
      parsedZoomData.push(row);
    }

    console.log("Total de registros Zoom:", parsedZoomData.length);
    console.log("Primer registro de Zoom:", parsedZoomData[0]);
    console.log("Docente seleccionado:", selectedDocente);

    setZoomData(parsedZoomData);

    let updatedCount = 0;
    let createdCount = 0;
    const newData = [...data];

    // Buscar fila template del docente
    let templateRow = data.find(row => row.DOCENTE === selectedDocente);
    
    // Si NO existe template, crear uno básico desde el primer registro de Zoom
    const needsTemplate = !templateRow;
    if (needsTemplate) {
      const firstZoomForDocente = parsedZoomData.find(zoomRow => {
        const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
        return matchDocente(selectedDocente, zoomDocente);
      });

      if (firstZoomForDocente) {
        console.log("⚠️ No hay filas base para el docente. Creando template desde Zoom...");
        
        templateRow = {
          PERIODO: "",
          MODELO: "PROTECH XP",
          MODALIDAD: "VIRTUAL",
          CURSO: "",
          SECCION: "",
          "AULA USS": "",
          DOCENTE: selectedDocente,
          TURNO: "",
          DIAS: "",
          "HORA INICIO": "",
          "HORA FIN": "",
          SESION: "",
          "Columna 13": "",
          inicio: "",
          fin: "",
          "Columna 16": "",
          "Columna 17": "",
          TOTAL: ""
        };
      } else {
        alert("⚠️ No se encontraron registros de Zoom para este docente");
        return;
      }
    }

    const unmatchedZoom = [];
    parsedZoomData.forEach((zoomRow, idx) => {
      const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
      const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
      const fechaInicio = zoomRow['Hora de inicio'] || zoomRow['Start Time'] || "";
      const fechaFin = zoomRow['Hora de finalización'] || zoomRow['End Time'] || "";

      if (!zoomDocente || !zoomTema) return;

      const docenteCoincide = matchDocente(selectedDocente, zoomDocente);
      if (!docenteCoincide) return;

      console.log(`\n✓ Registro #${idx + 1} - Docente: ${zoomDocente}`);
      console.log(`  Tema: ${zoomTema}`);

      let temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
      
      if (!temaMatch) {
        console.log("  ✗ No se pudo extraer info del tema");
        return;
      }

      const [, cursoParte, seccionZoom, sesionNumeroStr] = temaMatch;
      const cursoZoom = cursoParte.trim();
      const sesionNumero = sesionNumeroStr || "";

      console.log(`  📝 Extraído del tema: CURSO="${cursoZoom}", SECCION="${seccionZoom}", SESION="${sesionNumero}"`);

      let matched = false;

      // Solo buscar match si HAY filas existentes del docente
      if (!needsTemplate) {
        newData.forEach((row, index) => {
          if (row.DOCENTE !== selectedDocente) return;

          let cursoMatch = false;
          if (row.CURSO) {
            const cursoExcelNorm = normalizeCursoName(row.CURSO);
            const cursoZoomNorm = normalizeCursoName(cursoZoom);
            
            if (cursoExcelNorm === cursoZoomNorm) {
              cursoMatch = true;
            } else {
              const wordsExcel = cursoExcelNorm.split(" ");
              const wordsZoom = cursoZoomNorm.split(" ");
              const commonWords = wordsExcel.filter(word => wordsZoom.includes(word));
              const similarity = commonWords.length / Math.max(wordsExcel.length, wordsZoom.length);
              
              cursoMatch = commonWords.length >= 2 || similarity >= 0.5;
            }
          }

          const seccionMatch = row.SECCION && row.SECCION.toUpperCase() === seccionZoom.toUpperCase();
          const sesionMatch = sesionNumero ? parseInt(String(row.SESION)) === parseInt(sesionNumero) : false;

          if (seccionMatch && sesionMatch) {
            const fechaExtraida = extractDate(fechaInicio);
            const horaInicioExtraida = extractTime(fechaInicio);
            const horaFinExtraida = extractTime(fechaFin);
            
            console.log(`    ✓✓✓ MATCH ENCONTRADO - Actualizando fila ${index}`);
            
            let updateCurso = false;
            if (row.CURSO !== cursoZoom) {
              updateCurso = true;
              console.log(`      Actualizando CURSO de "${row.CURSO}" a "${cursoZoom}"`);
            }
            
            newData[index] = {
              ...newData[index],
              CURSO: updateCurso ? cursoZoom : newData[index].CURSO,
              "Columna 13": fechaExtraida,
              inicio: horaInicioExtraida,
              fin: horaFinExtraida
            };
            updatedCount++;
            matched = true;
          }
        });
      }

      if (!matched) {
        unmatchedZoom.push({ zoomRow, seccionZoom, sesionNumero, cursoZoom, fechaInicio, fechaFin });
      }
    });

    // Crear filas para los unmatched
    unmatchedZoom.forEach(({ zoomRow, seccionZoom, sesionNumero, cursoZoom, fechaInicio, fechaFin }) => {
      const fechaExtraida = extractDate(fechaInicio);
      const horaInicioExtraida = extractTime(fechaInicio);
      const horaFinExtraida = extractTime(fechaFin);

      const newRow = {
        PERIODO: templateRow.PERIODO,
        MODELO: templateRow.MODELO,
        MODALIDAD: templateRow.MODALIDAD,
        CURSO: cursoZoom,  // ✅ Usar el curso extraído del Tema
        SECCION: seccionZoom,  // ✅ Usar la sección extraída del Tema
        "AULA USS": templateRow["AULA USS"],
        DOCENTE: templateRow.DOCENTE,
        TURNO: templateRow.TURNO,
        DIAS: templateRow.DIAS,
        "HORA INICIO": templateRow["HORA INICIO"],
        "HORA FIN": templateRow["HORA FIN"],
        SESION: sesionNumero,  // ✅ Usar el número de sesión extraído del Tema
        "Columna 13": fechaExtraida,
        inicio: horaInicioExtraida,
        fin: horaFinExtraida,
        "Columna 16": "",
        "Columna 17": "",
        TOTAL: ""
      };

      const alreadyExists = newData.some(row => 
  row.DOCENTE === selectedDocente &&
  row.SECCION.toUpperCase() === seccionZoom.toUpperCase() &&
  String(row.SESION) === sesionNumero &&
  row["Columna 13"] === fechaExtraida &&
  row.inicio === horaInicioExtraida
);

      if (!alreadyExists) {
        newData.push(newRow);
        createdCount++;
        console.log(`✓ Nueva fila creada: ${cursoZoom} - ${seccionZoom} - SESIÓN ${sesionNumero}`);
      }
    });

    setData(newData);
    
    let message = "";
    if (updatedCount > 0 || createdCount > 0) {
      message = needsTemplate 
        ? `✅ Se crearon ${createdCount} registros nuevos para ${selectedDocente} (sin filas base previas)`
        : `✅ Se actualizaron ${updatedCount} registros y se crearon ${createdCount} nuevos para ${selectedDocente}`;
    } else {
      message = `⚠️ No se encontraron coincidencias. Verifica la consola (F12) para más detalles.`;
    }
    alert(message);
  } catch (error) {
    alert("❌ Error al procesar el archivo CSV: " + error.message);
    console.error(error);
  } finally {
    setIsLoading(false);
    event.target.value = "";
  }
};

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    setIsLoading(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      
      const worksheet = workbook.worksheets[0];
      const loadedData = [];

      worksheet.eachRow((row, rowIndex) => {
        if (rowIndex === 1) return;
        
        const getCellValue = (cell) => {
          if (!cell || !cell.value) return "";
          if (cell.value.hyperlink) return cell.value.hyperlink;
          if (typeof cell.value === 'object' && cell.value.text) return cell.value.text;
          return cell.value;
        };
        
        const rowData = {
          PERIODO: getCellValue(row.getCell(1)),
          MODELO: getCellValue(row.getCell(2)),
          MODALIDAD: getCellValue(row.getCell(3)),
          CURSO: getCellValue(row.getCell(4)),
          SECCION: getCellValue(row.getCell(5)),
          "AULA USS": getCellValue(row.getCell(6)),
          DOCENTE: getCellValue(row.getCell(7)),
          TURNO: getCellValue(row.getCell(8)),
          DIAS: getCellValue(row.getCell(9)),
          "HORA INICIO": getCellValue(row.getCell(10)),
          "HORA FIN": getCellValue(row.getCell(11)),
          SESION: getCellValue(row.getCell(12)),
          "Columna 13": getCellValue(row.getCell(13)),
          inicio: getCellValue(row.getCell(14)),
          fin: getCellValue(row.getCell(15)),
          "Columna 16": getCellValue(row.getCell(16)),
          "Columna 17": getCellValue(row.getCell(17)),
          TOTAL: getCellValue(row.getCell(18))
        };

        const notHeader = rowData.PERIODO !== "PERIODO" && rowData.MODELO !== "MODELO";
        if ((rowData.DOCENTE || rowData.CURSO) && notHeader) {
          loadedData.push(rowData);
        }
      });

      setData(loadedData);
      alert(`✅ Se cargaron ${loadedData.length} registros correctamente`);
    } catch (error) {
      alert("❌ Error al cargar el archivo: " + error.message);
      console.error(error);
    } finally {
      setIsLoading(false);
      event.target.value = "";
    }
  };

  const createRowsForDocente = () => {
  if (!selectedDocente || !numFilas || numFilas <= 0) {
    alert("Por favor selecciona un docente y especifica el número de filas");
    return;
  }

  const docenteRow = data.find(row => row.DOCENTE === selectedDocente);
  
  if (!docenteRow) {
    alert("No se encontró información del docente seleccionado");
    return;
  }

  const teacherZoom = zoomData.filter(zoomRow => {
    const zoomDocente = zoomRow['Anfitrión'] || zoomRow['Host'] || "";
    return matchDocente(selectedDocente, zoomDocente);
  });

  const sectionsMap = new Map();
  teacherZoom.forEach(zoomRow => {
    const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
    let temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
    if (temaMatch) {
      const [, , seccionZoom] = temaMatch;
      if (!sectionsMap.has(seccionZoom)) {
        sectionsMap.set(seccionZoom, []);
      }
      sectionsMap.get(seccionZoom).push(zoomRow);
    }
  });

  let uniqueSections = Array.from(sectionsMap.keys());

  if (uniqueSections.length === 0 && docenteRow.SECCION) {
    uniqueSections = [docenteRow.SECCION];
  }

  const numSections = uniqueSections.length;
  if (numSections === 0) {
    alert("No se encontró sección para crear filas");
    return;
  }

  // CAMBIO IMPORTANTE: Contar filas existentes por sección
  const existingRowsBySection = {};
  data.forEach(row => {
    if (row.DOCENTE === selectedDocente) {
      const seccion = row.SECCION;
      if (!existingRowsBySection[seccion]) {
        existingRowsBySection[seccion] = 0;
      }
      existingRowsBySection[seccion]++;
    }
  });

  const rowsPerSection = Math.ceil(parseInt(numFilas) / numSections);
  const allNewRows = [];
  let totalAutoCompleted = 0;

  uniqueSections.forEach(seccion => {
    let sectionZoom = sectionsMap.get(seccion) || [];
    
    // NUEVO: Calcular cuántas filas faltan para esta sección
    const existingCount = existingRowsBySection[seccion] || 0;
    const rowsToCreate = rowsPerSection - existingCount;
    
    console.log(`Sección ${seccion}: Existen ${existingCount}, se crearán ${rowsToCreate} (total deseado: ${rowsPerSection})`);
    
    if (rowsToCreate <= 0) {
      console.log(`⚠️ Sección ${seccion} ya tiene suficientes filas`);
      return;
    }

    // Empezar desde la siguiente sesión disponible
    const startSession = existingCount + 1;

    for (let i = 0; i < rowsToCreate; i++) {
      const sesionActual = startSession + i;
      
      const matchingZoom = sectionZoom.find(zoomRow => {
        const zoomTema = zoomRow['Tema'] || zoomRow['Topic'] || "";
        let temaMatch = zoomTema.match(/(.+?)(?:(?:–|-|\/|:)\s*)(PEAD-[a-zA-Z]+)(?:\s*(?:SESION|SESIÓN|Session|Sesión)\s*(\d+)?)?/i);
        if (!temaMatch) return false;

        const [, cursoParte, seccionZoomMatch, sesionNumeroStr] = temaMatch;
        const sesionNum = parseInt(sesionNumeroStr || "");

        if (isNaN(sesionNum) || sesionNum !== sesionActual) return false;
        if (seccionZoomMatch.toUpperCase() !== seccion.toUpperCase()) return false;

        if (docenteRow.CURSO) {
          const cursoExcelNorm = normalizeCursoName(docenteRow.CURSO);
          const cursoZoomNorm = normalizeCursoName(cursoParte.trim());
          
          const wordsExcel = cursoExcelNorm.split(" ");
          const wordsZoom = cursoZoomNorm.split(" ");
          const commonWords = wordsExcel.filter(word => wordsZoom.includes(word));
          const similarity = commonWords.length / Math.max(wordsExcel.length, wordsZoom.length);
          
          return commonWords.length >= 1 || similarity >= 0.4;
        }

        return true;
      });

      const thisCurso = matchingZoom ? extractCursoFromTema(matchingZoom['Tema'] || matchingZoom['Topic'] || "") : docenteRow.CURSO;

      const thisRow = {
        PERIODO: docenteRow.PERIODO,
        MODELO: docenteRow.MODELO,
        MODALIDAD: docenteRow.MODALIDAD,
        CURSO: thisCurso,
        SECCION: seccion,
        "AULA USS": docenteRow["AULA USS"],
        DOCENTE: docenteRow.DOCENTE,
        TURNO: docenteRow.TURNO,
        DIAS: docenteRow.DIAS,
        "HORA INICIO": docenteRow["HORA INICIO"],
        "HORA FIN": docenteRow["HORA FIN"],
        SESION: sesionActual,
        "Columna 13": "",
        inicio: "",
        fin: "",
        "Columna 16": "",
        "Columna 17": "",
        TOTAL: ""
      };

      if (matchingZoom) {
        const fechaInicio = matchingZoom['Hora de inicio'] || matchingZoom['Start Time'] || "";
        const fechaFin = matchingZoom['Hora de finalización'] || matchingZoom['End Time'] || "";
        const fechaExtraida = extractDate(fechaInicio);
        const horaInicioExtraida = extractTime(fechaInicio);
        const horaFinExtraida = extractTime(fechaFin);
        
        thisRow["Columna 13"] = fechaExtraida;
        thisRow.inicio = horaInicioExtraida;
        thisRow.fin = horaFinExtraida;
        
        console.log(`✓ Autocompletado SESIÓN ${sesionActual} de ${seccion} (${thisCurso})`);
        totalAutoCompleted++;
      }

      allNewRows.push(thisRow);
    }
  });

  // NO eliminar las filas existentes del docente, solo agregar las nuevas
  setData([...data, ...allNewRows]);
  setNumFilas("");

  const sectionsList = uniqueSections.join(', ');
  const totalExistentes = Object.values(existingRowsBySection).reduce((a, b) => a + b, 0);
  const message = totalAutoCompleted > 0 
    ? `Se crearon ${allNewRows.length} filas nuevas para ${selectedDocente} (ya existían ${totalExistentes}). Total: ${totalExistentes + allNewRows.length} filas. ${totalAutoCompleted} autocompletadas con Zoom`
    : `Se crearon ${allNewRows.length} filas nuevas para ${selectedDocente} (ya existían ${totalExistentes}). Total: ${totalExistentes + allNewRows.length} filas`;
  alert(`✅ ${message}`);
};

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Monitoreo");

    worksheet.addRow(headers);

    const headerRow = worksheet.getRow(1);
    headerRow.height = 40;
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF203864" }
      };
      cell.font = {
        color: { argb: "FFFFFFFF" },
        bold: true,
        size: 11,
        name: "Arial"
      };
      cell.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFFFFFFF" } },
        left: { style: "thin", color: { argb: "FFFFFFFF" } },
        bottom: { style: "thin", color: { argb: "FFFFFFFF" } },
        right: { style: "thin", color: { argb: "FFFFFFFF" } }
      };
    });

    data.forEach((row, index) => {
      const rowData = headers.map(h => row[h] !== undefined ? row[h] : "");
      const excelRow = worksheet.addRow(rowData);
      excelRow.height = 25;

      const isEven = (index + 2) % 2 === 0;
      excelRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: isEven ? "FFE8F4F8" : "FFFFFFFF" }
        };
        cell.font = {
          size: 10,
          name: "Arial",
          color: { argb: "FF000000" }
        };
        cell.alignment = {
          vertical: "middle",
          horizontal: "center"
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE0E0E0" } },
          left: { style: "thin", color: { argb: "FFE0E0E0" } },
          bottom: { style: "thin", color: { argb: "FFE0E0E0" } },
          right: { style: "thin", color: { argb: "FFE0E0E0" } }
        };
      });
    });

    const columnWidths = [13, 14, 13, 18, 12, 45, 25, 11, 13, 14, 12, 10, 13, 10, 10, 13, 13, 13];
    worksheet.columns = columnWidths.map((width, index) => ({
      key: headers[index],
      width: width
    }));

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = "Monitoreo_USS.xlsx";
    link.click();
    window.URL.revokeObjectURL(url);
  };

  const addRow = () => {
    const newRow = {};
    headers.forEach(header => {
      newRow[header] = "";
    });
    setData([...data, newRow]);
  };

  const deleteRow = (index) => {
    if (selectedDocente) {
      const realIndex = data.findIndex((row) => row === filteredData[index]);
      const newData = data.filter((_, i) => i !== realIndex);
      setData(newData);
    } else {
      const newData = data.filter((_, i) => i !== index);
      setData(newData);
    }
  };

  const handleCellChange = (rowIndex, columnName, value) => {
    if (selectedDocente) {
      const realIndex = data.findIndex((row, idx) => 
        row === filteredData[rowIndex]
      );
      const newData = [...data];
      newData[realIndex][columnName] = value;
      setData(newData);
    } else {
      const newData = [...data];
      newData[rowIndex][columnName] = value;
      setData(newData);
    }
  };

  // ===== CONSTANTES =====
  const headers = [
    "PERIODO", "MODELO", "MODALIDAD", "CURSO", "SECCION", "AULA USS",
    "DOCENTE", "TURNO", "DIAS", "HORA INICIO", "HORA FIN", "SESION",
    "Columna 13", "inicio", "fin", "Columna 16", "Columna 17", "TOTAL"
  ];

  const uniqueDocentes = [
    "CARRASCO CHEVEZ HENRY", "LARA PERLECHE LOURDES", "CALLACNA SENCIO IVAN",
    "MEJIA ZELADA CARLOS", "MOGOLLON GALECIO POLO", "DIAZ ESPINO MIGUEL",
    "GARCIA CABRERA MARTIN", "DIAZ MUSAYON CESAR", "CASTAÑEDA BALCAZAR EDWARD",
    "QUINTANA DAVILA TONY", "SANCHEZ JAEGER CRISTHIAM", "TULLUME DIAZ JOSE",
    "GUERRERO AGURTO GINO", "CRIOLLO VALVERDE STEPHANY", "TICONA TAPIA ESTRELLA",
    "DELGADO MARIELLA", "SANCHEZ PEREZ DIANA", "QUESADA QUIROZ JENNIE",
    "BURGOS SUERO LUCIANA", "SALAZAR LLUEN IVONNE", "SANCHEZ GUEVARA OMAR",
    "BRUNO SARMIENTO JOSE", "GONZALES ÑIQUE PERCY", "OLIVOS KATHERIN",
    "SANDOVAL HORNA CARMEN", "SALAZAR LLUEN JAIRO", "MECHAN FRANCISCO",
    "CISNEROS LEANDRO", "MURO EFRAIN", "QUESADA JENNIE", "NIETO NELSON",
    "SALAZAR DANIEL", "DIAZ CESAR"
  ];

  const dropdownOptions = {
    MODELO: ["PROTECH XP", "TRADICIONAL"],
    MODALIDAD: ["PRESENCIAL", "VIRTUAL"],
    CURSO: [
      "AUTOCAD 2D", "AUTOCAD 3D", "COMPUTACION 2", "COMPUTACION 2 ARQ",
      "COMPUTACION 2 CIV", "COMPUTACIÓN 2 MEC", "COMPUTACIÓN 2 SIST",
      "COMPUTACION 3", "COMPUTACION 3 ARQ", "COMPUTACION 3 CIV",
      "COMPUTACIÓN 3 MEC", "COMPUTACIÓN 3 SIST", "WORD 365", "WORD ASOCIADO",
      "EXCEL 365", "EXCEL ASOCIADO", "DISEÑO CON CANVA", "DISEÑO WEB",
      "POWER BI", "TALLER WORD 2019/365", "TALLER EXCEL 2019/365"
    ],
    SECCION: [
      "A", "PEAD-a", "PEAD-b", "PEAD-C", "PEAD-d", "PEAD-e", "PEAD-f",
      "PEAD-g", "PEAD-h", "PEAD-i", "PEAD-j", "PEAD-k", "PEAD-l", "PEAD-m",
      "PEAD-n", "PEAD-o", "PEAD-p", "PEAD-q", "PEAD-A"
    ],
    TURNO: ["MAÑANA", "TARDE", "NOCHE"],
    DIAS: [
      "LUN", "MAR", "MIE", "JUE", "VIE", "SAB", "LUN y MIE", "LUN-MIE",
      "LUN-VIE", "MAR Y JUE", "MAR y VIE", "MAR-JUE", "MAR-VIE", "MIE y JUE",
      "MIE y VIE", "MIE-VIE", "JUE-SAB", "VIE y SAB", "VIE-SAB"
    ],
    DOCENTE: uniqueDocentes
  };

  // ===== DATOS COMPUTADOS =====
  const filteredData = useMemo(() => {
    if (!selectedDocente) return data;
    return data.filter(row => row.DOCENTE === selectedDocente);
  }, [data, selectedDocente]);

  const displayData = useMemo(() => {
    if (!selectedDocente) return data;
    
    return filteredData.map((row, index) => ({
      ...row,
      SESION: index + 1
    }));
  }, [filteredData, selectedDocente, data]);

  // ===== RENDER =====
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 p-4">
      <div className="max-w-full mx-auto">
        <ControlPanel
          selectedDocente={selectedDocente}
          setSelectedDocente={setSelectedDocente}
          numFilas={numFilas}
          setNumFilas={setNumFilas}
          uniqueDocentes={uniqueDocentes}
          onCreateRows={createRowsForDocente}
          onAddRow={addRow}
          onExport={exportToExcel}
          onLoadExcel={handleFileUpload}
          onLoadZoomCsv={handleZoomCsvUpload}
          isLoading={isLoading}
          displayDataLength={displayData.length}
        />

        <DataTable
          data={displayData}
          headers={headers}
          dropdownOptions={dropdownOptions}
          onCellChange={handleCellChange}
          onDeleteRow={deleteRow}
        />
      </div>
    </div>
  );
}

export default App;