function doPost(e) {
  var params = JSON.parse(e.postData.contents);
  var action = params.action;
  var result;

  try {
    if (action === "login") {
      result = login(params.usuario, params.password);
    } else if (action === "getClassroomData") {
      result = getClassroomData(params.aula);
    } else if (action === "updateGrades") {
      result = updateGrades(params.aula, params.legajo, params.notas);
    } else if (action === "updateStudentStatus") {
      result = updateStudentStatus(params.aula, params.legajo, params.nuevoEstado);
    } else if (action === "adminAddStudent") {
      result = adminAddStudent(params.aula, params.legajo, params.apellido, params.nombre, params.genero);
    } else if (action === "adminGetTeachers") {
      result = adminGetTeachers();
    } else if (action === "adminUpdateAssignments") {
      result = adminUpdateAssignments(params.usuario, params.nuevasAulas);
    } else if (action === "adminGenerateReport") {
      result = adminGenerateReport();
    } else {
      result = { success: false, message: "Acción no válida" };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function login(usuario, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Usuarios");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] == usuario && row[1] == password) {
      return {
        success: true,
        usuario: row[0],
        rol: row[2],
        aulas: row[3] ? row[3].toString().split(",") : []
      };
    }
  }
  return { success: false, message: "Usuario o contraseña incorrectos" };
}

function getClassroomData(aula) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(aula.toString());
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var alumnos = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {
      alumnos.push({
        legajo: row[0],
        apellido: row[1],
        nombre: row[2],
        genero: row[3],
        estado: row[4],
        n1_p1: row[5],
        n2_p1: row[6],
        prom_p1: row[7],
        n1_p2: row[8],
        n2_p2: row[9],
        prom_p2: row[10],
        prom_anual: row[11]
      });
    }
  }
  return alumnos;
}

function updateGrades(aula, legajo, notas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(aula.toString());
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == legajo) {
      var rowNum = i + 1;
      if (notas.n1_p1 !== undefined) sheet.getRange(rowNum, 6).setValue(notas.n1_p1);
      if (notas.n2_p1 !== undefined) sheet.getRange(rowNum, 7).setValue(notas.n2_p1);
      if (notas.n1_p2 !== undefined) sheet.getRange(rowNum, 9).setValue(notas.n1_p2);
      if (notas.n2_p2 !== undefined) sheet.getRange(rowNum, 10).setValue(notas.n2_p2);
      return { success: true, message: "Calificaciones actualizadas correctamente" };
    }
  }
  return { success: false, message: "Alumno no encontrado" };
}

function updateStudentStatus(aula, legajo, nuevoEstado) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(aula.toString());
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == legajo) {
      var rowNum = i + 1;
      sheet.getRange(rowNum, 5).setValue(nuevoEstado);
      return { success: true, message: "Estado actualizado a " + nuevoEstado };
    }
  }
  return { success: false, message: "Alumno no encontrado" };
}

function adminAddStudent(aula, legajo, apellido, nombre, genero) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(aula.toString());
  if (!sheet) return { success: false, message: "El aula no existe" };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == legajo) {
      return { success: false, message: "El legajo ya existe en esta aula" };
    }
  }
  
  var nextRow = sheet.getLastRow() + 1;
  
  sheet.getRange(nextRow, 1).setValue(legajo);
  sheet.getRange(nextRow, 2).setValue(apellido);
  sheet.getRange(nextRow, 3).setValue(nombre);
  sheet.getRange(nextRow, 4).setValue(genero);
  sheet.getRange(nextRow, 5).setValue("ALTA");
  
  sheet.getRange(nextRow, 8).setFormula('=IF(OR(ISBLANK(F' + nextRow + '), ISBLANK(G' + nextRow + ')), "Falta nota", AVERAGE(F' + nextRow + ':G' + nextRow + '))');
  sheet.getRange(nextRow, 11).setFormula('=IF(OR(ISBLANK(I' + nextRow + '), ISBLANK(J' + nextRow + ')), "Falta nota", AVERAGE(I' + nextRow + ':J' + nextRow + '))');
  sheet.getRange(nextRow, 12).setFormula('=IF(OR(H' + nextRow + '="Falta nota", K' + nextRow + '="Falta nota"), "Falta nota", AVERAGE(H' + nextRow + ', K' + nextRow + '))');
  
  return { success: true, message: "Alumno agregado exitosamente" };
}

function adminUpdateAssignments(usuario, nuevasAulas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Usuarios");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == usuario) {
      sheet.getRange(i + 1, 4).setValue(nuevasAulas);
      return { success: true, message: "Aulas reasignadas correctamente" };
    }
  }
  return { success: false, message: "Usuario no encontrado" };
}

function adminGetTeachers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Usuarios");
  var data = sheet.getDataRange().getValues();
  var teachers = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] == "profesor") {
      teachers.push({
        usuario: data[i][0],
        aulas: data[i][3]
      });
    }
  }
  return teachers;
}

function adminGenerateReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aulas = ["101", "102", "103"];
  var sinNota = [];
  var promedioBajo = [];
  
  aulas.forEach(function(aula) {
    var sheet = ss.getSheetByName(aula);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var legajo = row[0];
      var apellido = row[1];
      var nombre = row[2];
      var estado = row[4];
      
      if (legajo && estado === "ALTA") {
        var n1_p1 = row[5];
        var n2_p1 = row[6];
        var n1_p2 = row[8];
        var n2_p2 = row[9];
        var promAnual = row[11];
        
        if (n1_p1 === "" || n2_p1 === "" || n1_p2 === "" || n2_p2 === "") {
          sinNota.push({
            aula: aula,
            legajo: legajo,
            apellido: apellido,
            nombre: nombre,
            detalle: "Faltan notas en alguna instancia"
          });
        }
        
        if (typeof promAnual === "number" && promAnual < 4) {
          promedioBajo.push({
            aula: aula,
            legajo: legajo,
            apellido: apellido,
            nombre: nombre,
            promedio: promAnual
          });
        }
      }
    }
  });
  
  return {
    sinCalificacion: sinNota,
    promedioMenorA4: promedioBajo
  };
}
