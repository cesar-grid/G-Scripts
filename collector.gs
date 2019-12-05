function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Generar CSR")
  .addItem('Eventos generados', 'eventosOTRS')
  .addItem('Eventos por Areas', 'eventosAreas')
  .addItem('Top 25', 'topAlertProducers')
  .addItem('Eventos por turno', 'diaNoche')
  .addItem('Ayuda', 'showHelp')
  .addToUi();
}

function showHelp() {
 var ss=SpreadsheetApp.getActiveSpreadsheet();
 var ui = SpreadsheetApp.getUi();
 var Alert = ui.alert("En caso de necesitar ayuda con este documento contacte a: César Granados.");
}

function eventosOTRS() {
  var ss = SpreadsheetApp.getActive();
  var sheetConfig = ss.getSheetByName('Config');
  var sheetCollector = ss.getSheetByName('Collector');
  var host = sheetConfig.getRange("B1").getValue();
  var database = sheetConfig.getRange("B2").getValue();
  var user = sheetConfig.getRange("B3").getValue();
  var password = sheetConfig.getRange("B4").getValue();
  var port = sheetConfig.getRange("B5").getValue();
  var FechaInicio = sheetCollector.getRange("L4").getValue();
  var FechaFin = sheetCollector.getRange("L5").getValue();
  var Cliente = sheetConfig.getRange("B6").getValue();  
  var url = 'jdbc:mysql://'+host+':'+port+'/'+database;
  var EventosGenerados = 'SELECT customer_company.customer_id ID, customer_company.name, COUNT(1) Total, SUM(CASE WHEN ticket.user_id = 1 THEN 1 ELSE 0 END) SinAnalisis, SUM(CASE WHEN ticket.ticket_state_id = 11 THEN 1 ELSE 0 END) Escalados, SUM(CASE WHEN ticket.ticket_state_id = 14 THEN 1 ELSE 0 END) Recuperados, SUM(CASE WHEN ticket.ticket_state_id IN (2,3) THEN 1 ELSE 0 END) SatisfechosInsatisfechos, SUM(CASE WHEN ticket.ticket_state_id = 9 THEN 1 ELSE 0 END) Fusionados FROM ticket, customer_company WHERE ticket.customer_id = customer_company.customer_id AND ticket.queue_id IN(8,9,10) AND customer_company.customer_id  = "'+Cliente+'" AND ticket.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59")';

  try{
    var connection = Jdbc.getConnection(url, user, password);
    var result = connection.createStatement().executeQuery(EventosGenerados);
    var metaData = result.getMetaData();
    var columns = metaData.getColumnCount();  
    var values = [];
    var value = [];
    var element = '';

    for (i = 1; i <= columns; i ++){
      element = metaData.getColumnLabel(i);
      value.push(element);
    }
    values.push(value);
  
    while(result.next()){
      value = [];
      for (i = 1; i <= columns; i ++){
        element = result.getString(i);
        value.push(element);
      }
        values.push(value);
    }
  //Cierra conexion
    result.close();
  //Escribe datos en las celdas
    sheetCollector.getRange(1,1, values.length, value.length).setValues(values);
    SpreadsheetApp.getActive().toast('Datos actualizado correctamente en [Tab: Collector]!');
  }catch(err){
    SpreadsheetApp.getActive().toast(err.message);
  } 
}

function topAlertProducers() {
  var ss = SpreadsheetApp.getActive();
  var sheetConfig = ss.getSheetByName('Config');
  var sheetCollector = ss.getSheetByName('Collector');  
  var host = sheetConfig.getRange("B1").getValue();
  var database = sheetConfig.getRange("B2").getValue();
  var user = sheetConfig.getRange("B3").getValue();
  var password = sheetConfig.getRange("B4").getValue();
  var port = sheetConfig.getRange("B5").getValue();
  var FechaInicio = sheetCollector.getRange("L4").getValue();
  var FechaFin = sheetCollector.getRange("L5").getValue();
  var Cliente = sheetConfig.getRange("B6").getValue();  
  var url = 'jdbc:mysql://'+host+':'+port+'/'+database;
  var Top25 = 'SELECT ticket.title, COUNT(1) AS Total FROM customer_company, ticket WHERE ticket.customer_id = customer_company.customer_id AND ticket.archive_flag IN (0,1) AND ticket.queue_id IN(8,9,10) AND customer_company.customer_id  = "'+Cliente+'" AND ticket.create_time BETWEEN CONCAT(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") GROUP BY ticket.title ORDER BY Total desc limit 0, 25';

  try{
    var connection = Jdbc.getConnection(url, user, password); 
    var result = connection.createStatement().executeQuery(Top25);
    var metaData = result.getMetaData();
    var columns = metaData.getColumnCount();
    var values = [];
    var value = [];
    var element = '';

    for (i = 1; i <= columns; i ++){
      element = metaData.getColumnLabel(i);
      value.push(element);
    }
    values.push(value);
  
    while(result.next()){
      value = [];
      for (i = 1; i <= columns; i ++){
        element = result.getString(i);
        value.push(element);
      }
      values.push(value);
    }
  //Cierra conexion
    result.close(); 
    sheetCollector.getRange('A6:B30').clearContent();
  //Escribe datos en las celdas
    sheetCollector.getRange(5,1, values.length, value.length).setValues(values);
    SpreadsheetApp.getActive().toast('Datos actualizado correctamente en [Tab: Collector]');
  }catch(err){
    SpreadsheetApp.getActive().toast(err.message);
  } 
}

function diaNoche() {
  var ss = SpreadsheetApp.getActive();
  var sheetConfig = ss.getSheetByName('Config');
  var sheetCollector = ss.getSheetByName('Collector');  
  var host = sheetConfig.getRange("B1").getValue();
  var database = sheetConfig.getRange("B2").getValue();
  var user = sheetConfig.getRange("B3").getValue();
  var password = sheetConfig.getRange("B4").getValue();
  var port = sheetConfig.getRange("B5").getValue();
  var FechaInicio = sheetCollector.getRange("L4").getValue();
  var FechaFin = sheetCollector.getRange("L5").getValue();
  var Cliente = sheetConfig.getRange("B6").getValue();
  var url = 'jdbc:mysql://'+host+':'+port+'/'+database;
  var EnventosDiaNoche = 'SELECT DAYNAME(ticket.create_time) AS "Day of Week", SUM(CASE WHEN date_format(ticket.create_time, "%H:%i:%s") BETWEEN "07:00:00" AND "18:59:59" THEN 1 ELSE 0 END ) AS Diurnal, SUM(CASE WHEN date_format(ticket.create_time, "%H:%i:%s") BETWEEN "19:00:00" AND "23:59:59"  THEN 1 ELSE 0 END ) AS Nightly1, SUM(CASE WHEN date_format(ticket.create_time, "%H:%i:%s") BETWEEN  "00:00:00" and "06:59:59" THEN 1 ELSE 0 END ) AS Nightly2, COUNT(*) AS Total FROM customer_company, ticket WHERE ticket.customer_id = customer_company.customer_id AND customer_company.customer_id  =  "'+Cliente+'" AND date_format(ticket.create_time, "%Y-%m-%d") BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d"),"23:59:59") GROUP BY dayofweek(ticket.create_time) ORDER BY dayofweek(ticket.create_time)';
  var EnventosDiaNocheEscalados = 'SELECT DAYNAME(ticket.create_time) AS "Day of Week", SUM(CASE WHEN date_format(ticket.create_time, "%H:%i:%s") BETWEEN "07:00:00" AND "18:59:59" THEN 1 ELSE 0 END ) AS Diurnal, SUM(CASE WHEN date_format(ticket.create_time, "%H:%i:%s") BETWEEN "19:00:00" AND "23:59:59"  THEN 1 ELSE 0 END ) AS Nightly1, SUM(CASE WHEN date_format(ticket.create_time, "%H:%i:%s") BETWEEN  "00:00:00" and "06:59:59" THEN 1 ELSE 0 END ) AS Nightly2, COUNT(*) AS Total FROM customer_company, ticket WHERE ticket.customer_id = customer_company.customer_id AND ticket.ticket_state_id = 11 AND customer_company.customer_id  =  "'+Cliente+'" AND date_format(ticket.create_time, "%Y-%m-%d") BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d"),"23:59:59") GROUP BY dayofweek(ticket.create_time) ORDER BY dayofweek(ticket.create_time)';

  try{
    var connection = Jdbc.getConnection(url, user, password); 
    var result = connection.createStatement().executeQuery(EnventosDiaNoche);
    var metaData = result.getMetaData();
    var columns = metaData.getColumnCount();  
    var values = [];
    var value = [];
    var element = '';

    for (i = 1; i <= columns; i ++){
      element = metaData.getColumnLabel(i);
      value.push(element);
    }
    values.push(value);
  
    while(result.next()){
      value = [];
      for (i = 1; i <= columns; i ++){
        element = result.getString(i);
        value.push(element);
      }
      values.push(value);
    }
  //Cierra conexion
  result.close();
  sheetCollector.getRange('D33:H40').clearContent();    
  //Escribe datos en las celdas
  sheetCollector.getRange(33,4, values.length, value.length).setValues(values);
  
  var result = connection.createStatement().executeQuery(EnventosDiaNocheEscalados);
  var metaData = result.getMetaData();
  var columns = metaData.getColumnCount();  
  var values = [];
  var value = [];
  var element = '';

  for (i = 1; i <= columns; i ++){
    element = metaData.getColumnLabel(i);
    value.push(element);
  }
  values.push(value);
  
  while(result.next()){
    value = [];
    for (i = 1; i <= columns; i ++){
      element = result.getString(i);
       value.push(element);
    }
    values.push(value);
  }
  //Cierra conexion
  result.close(); 
  sheetCollector.getRange('D42:H49').clearContent();
  //Escribe datos en las celdas
  sheetCollector.getRange(42,4, values.length, value.length).setValues(values);
  SpreadsheetApp.getActive().toast('Datos actualizado correctamente en [Tab: Collector]!');   
  }catch(err){
    SpreadsheetApp.getActive().toast(err.message);
  } 
}

function moveCols() {
  var ss = SpreadsheetApp.getActive();
  var sourceSheet = ss.getSheetByName('Datos');
  var destSheet = ss.getSheetByName('Datos');
  
  //Eventos escaldos													
  sourceSheet.getRange('C3:C6').copyTo(destSheet.getRange('B3:B6'))
  sourceSheet.getRange('D3:D6').copyTo(destSheet.getRange('C3:C6'))
  sourceSheet.getRange('E3:E6').copyTo(destSheet.getRange('D3:D6'))
  sourceSheet.getRange('F3:F6').copyTo(destSheet.getRange('E3:E6'))
  sourceSheet.getRange('G3:G6').copyTo(destSheet.getRange('F3:F6'))
  sourceSheet.getRange('H3:H6').copyTo(destSheet.getRange('G3:G6'))
  sourceSheet.getRange('I3:I6').copyTo(destSheet.getRange('H3:H6'))
  sourceSheet.getRange('J3:J6').copyTo(destSheet.getRange('I3:I6'))
  sourceSheet.getRange('K3:K6').copyTo(destSheet.getRange('J3:J6'))
  sourceSheet.getRange('L3:L6').copyTo(destSheet.getRange('K3:K6'))
  sourceSheet.getRange('M3:M6').copyTo(destSheet.getRange('L3:L6'))
  sourceSheet.getRange('N3:N6').copyTo(destSheet.getRange('M3:M6'),{contentsOnly:true})
  sourceSheet.getRange('B3').copyTo(destSheet.getRange('N3'))
  //Porcentaje de eventos escaldos
  sourceSheet.getRange('C8:C11').copyTo(destSheet.getRange('B8:B11'))
  sourceSheet.getRange('D8:D11').copyTo(destSheet.getRange('C8:C11'))
  sourceSheet.getRange('E8:E11').copyTo(destSheet.getRange('D8:D11'))
  sourceSheet.getRange('F8:F11').copyTo(destSheet.getRange('E8:E11'))
  sourceSheet.getRange('G8:G11').copyTo(destSheet.getRange('F8:F11'))
  sourceSheet.getRange('H8:H11').copyTo(destSheet.getRange('G8:G11'))
  sourceSheet.getRange('I8:I11').copyTo(destSheet.getRange('H8:H11'))
  sourceSheet.getRange('J8:J11').copyTo(destSheet.getRange('I8:I11'))
  sourceSheet.getRange('K8:K11').copyTo(destSheet.getRange('J8:J11'))
  sourceSheet.getRange('L8:L11').copyTo(destSheet.getRange('K8:K11'))
  sourceSheet.getRange('M8:M11').copyTo(destSheet.getRange('L8:L11'))
  sourceSheet.getRange('N8:N11').copyTo(destSheet.getRange('M8:M11'),{contentsOnly:true})
  sourceSheet.getRange('B8').copyTo(destSheet.getRange('N8'))
  //Tiempo Promedio de Atención (min)
  sourceSheet.getRange('C15:C18').copyTo(destSheet.getRange('B15:B18'))
  sourceSheet.getRange('D15:D18').copyTo(destSheet.getRange('C15:C18'))
  sourceSheet.getRange('E15:E18').copyTo(destSheet.getRange('D15:D18'))
  sourceSheet.getRange('F15:F18').copyTo(destSheet.getRange('E15:E18'))
  sourceSheet.getRange('G15:G18').copyTo(destSheet.getRange('F15:F18'))
  sourceSheet.getRange('H15:H18').copyTo(destSheet.getRange('G15:G18'))
  sourceSheet.getRange('I15:I18').copyTo(destSheet.getRange('H15:H18'))
  sourceSheet.getRange('J15:J18').copyTo(destSheet.getRange('I15:I18'))
  sourceSheet.getRange('K15:K18').copyTo(destSheet.getRange('J15:J18'))
  sourceSheet.getRange('L15:L18').copyTo(destSheet.getRange('K15:K18'))
  sourceSheet.getRange('M15:M18').copyTo(destSheet.getRange('L15:L18'))
  sourceSheet.getRange('N15:N18').copyTo(destSheet.getRange('M15:M18'),{contentsOnly:true})
  sourceSheet.getRange('B15').copyTo(destSheet.getRange('N15'))
  //Cumplimiento de SLA (%)
  sourceSheet.getRange('C20:C23').copyTo(destSheet.getRange('B20:B23'))
  sourceSheet.getRange('D20:D23').copyTo(destSheet.getRange('C20:C23'))
  sourceSheet.getRange('E20:E23').copyTo(destSheet.getRange('D20:D23'))
  sourceSheet.getRange('F20:F23').copyTo(destSheet.getRange('E20:E23'))
  sourceSheet.getRange('G20:G23').copyTo(destSheet.getRange('F20:F23'))
  sourceSheet.getRange('H20:H23').copyTo(destSheet.getRange('G20:G23'))
  sourceSheet.getRange('I20:I23').copyTo(destSheet.getRange('H20:H23'))
  sourceSheet.getRange('J20:J23').copyTo(destSheet.getRange('I20:I23'))
  sourceSheet.getRange('K20:K23').copyTo(destSheet.getRange('J20:J23'))
  sourceSheet.getRange('L20:L23').copyTo(destSheet.getRange('K20:K23'))
  sourceSheet.getRange('M20:M23').copyTo(destSheet.getRange('L20:L23'))
  sourceSheet.getRange('N20:N23').copyTo(destSheet.getRange('M20:M23'),{contentsOnly:true})
  sourceSheet.getRange('B20').copyTo(destSheet.getRange('N20'))
  //Disponibilidad de Servicios
  sourceSheet.getRange('C26:C29').copyTo(destSheet.getRange('B26:B29'))
  sourceSheet.getRange('D26:D29').copyTo(destSheet.getRange('C26:C29'))
  sourceSheet.getRange('E26:E29').copyTo(destSheet.getRange('D26:D29'))
  sourceSheet.getRange('F26:F29').copyTo(destSheet.getRange('E26:E29'))
  sourceSheet.getRange('G26:G29').copyTo(destSheet.getRange('F26:F29'))
  sourceSheet.getRange('H26:H29').copyTo(destSheet.getRange('G26:G29'))
  sourceSheet.getRange('I26:I29').copyTo(destSheet.getRange('H26:H29'))
  sourceSheet.getRange('J26:J29').copyTo(destSheet.getRange('I26:I29'))
  sourceSheet.getRange('K26:K29').copyTo(destSheet.getRange('J26:J29'))
  sourceSheet.getRange('L26:L29').copyTo(destSheet.getRange('K26:K29'))
  sourceSheet.getRange('M26:M29').copyTo(destSheet.getRange('L26:L29'))
  sourceSheet.getRange('N26:N29').copyTo(destSheet.getRange('M26:M29'),{contentsOnly:true})
  sourceSheet.getRange('B26').copyTo(destSheet.getRange('N26'))
  //Clasificacion monitoreo - Indidacores
  sourceSheet.getRange('C33:C35').copyTo(destSheet.getRange('B33:B35'))
  sourceSheet.getRange('D33:D35').copyTo(destSheet.getRange('C33:C35'))
  sourceSheet.getRange('E33:E35').copyTo(destSheet.getRange('D33:D35'))
  sourceSheet.getRange('F33:F35').copyTo(destSheet.getRange('E33:E35'))
  sourceSheet.getRange('G33:G35').copyTo(destSheet.getRange('F33:F35'))
  sourceSheet.getRange('H33:H35').copyTo(destSheet.getRange('G33:G35'))
  sourceSheet.getRange('I33:I35').copyTo(destSheet.getRange('H33:H35'))
  sourceSheet.getRange('J33:J35').copyTo(destSheet.getRange('I33:I35'))
  sourceSheet.getRange('K33:K35').copyTo(destSheet.getRange('J33:J35'))
  sourceSheet.getRange('L33:L35').copyTo(destSheet.getRange('K33:K35'))
  sourceSheet.getRange('M33:M35').copyTo(destSheet.getRange('L33:L35'))
  sourceSheet.getRange('N33:N35').copyTo(destSheet.getRange('M33:M35'),{contentsOnly:true})
  sourceSheet.getRange('B33').copyTo(destSheet.getRange('N33'))
  //Clasificacion monitoreo - Bandas
  sourceSheet.getRange('C37:C41').copyTo(destSheet.getRange('B37:B41'))
  sourceSheet.getRange('D37:D41').copyTo(destSheet.getRange('C37:C41'))
  sourceSheet.getRange('E37:E41').copyTo(destSheet.getRange('D37:D41'))
  sourceSheet.getRange('F37:F41').copyTo(destSheet.getRange('E37:E41'))
  sourceSheet.getRange('G37:G41').copyTo(destSheet.getRange('F37:F41'))
  sourceSheet.getRange('H37:H41').copyTo(destSheet.getRange('G37:G41'))
  sourceSheet.getRange('I37:I41').copyTo(destSheet.getRange('H37:H41'))
  sourceSheet.getRange('J37:J41').copyTo(destSheet.getRange('I37:I41'))
  sourceSheet.getRange('K37:K41').copyTo(destSheet.getRange('J37:J41'))
  sourceSheet.getRange('L37:L41').copyTo(destSheet.getRange('K37:K41'))
  sourceSheet.getRange('M37:M41').copyTo(destSheet.getRange('L37:L41'))
  sourceSheet.getRange('N37:N41').copyTo(destSheet.getRange('M37:M41'),{contentsOnly:true})
  sourceSheet.getRange('B37').copyTo(destSheet.getRange('N37'))
  //Otras Metricas - Tickets escalados por banda													
  sourceSheet.getRange('C57:C60').copyTo(destSheet.getRange('B57:B60'))
  sourceSheet.getRange('D57:D60').copyTo(destSheet.getRange('C57:C60'))
  sourceSheet.getRange('E57:E60').copyTo(destSheet.getRange('D57:D60'))
  sourceSheet.getRange('F57:F60').copyTo(destSheet.getRange('E57:E60'))
  sourceSheet.getRange('G57:G60').copyTo(destSheet.getRange('F57:F60'))
  sourceSheet.getRange('H57:H60').copyTo(destSheet.getRange('G57:G60'))
  sourceSheet.getRange('I57:I60').copyTo(destSheet.getRange('H57:H60'))
  sourceSheet.getRange('J57:J60').copyTo(destSheet.getRange('I57:I60'))
  sourceSheet.getRange('K57:K60').copyTo(destSheet.getRange('J57:J60'))
  sourceSheet.getRange('L57:L60').copyTo(destSheet.getRange('K57:K60'))
  sourceSheet.getRange('M57:M60').copyTo(destSheet.getRange('L57:L60'))
  sourceSheet.getRange('N57:N60').copyTo(destSheet.getRange('M57:M60'),{contentsOnly:true})
  sourceSheet.getRange('B57').copyTo(destSheet.getRange('N57'))
  //Otras Metricas - Eventos sin atencion fuera de SLA
  sourceSheet.getRange('C62:C66').copyTo(destSheet.getRange('B62:B66'))
  sourceSheet.getRange('D62:D66').copyTo(destSheet.getRange('C62:C66'))
  sourceSheet.getRange('E62:E66').copyTo(destSheet.getRange('D62:D66'))
  sourceSheet.getRange('F62:F66').copyTo(destSheet.getRange('E62:E66'))
  sourceSheet.getRange('G62:G66').copyTo(destSheet.getRange('F62:F66'))
  sourceSheet.getRange('H62:H66').copyTo(destSheet.getRange('G62:G66'))
  sourceSheet.getRange('I62:I66').copyTo(destSheet.getRange('H62:H66'))
  sourceSheet.getRange('J62:J66').copyTo(destSheet.getRange('I62:I66'))
  sourceSheet.getRange('K62:K66').copyTo(destSheet.getRange('J62:J66'))
  sourceSheet.getRange('L62:L66').copyTo(destSheet.getRange('K62:K66'))
  sourceSheet.getRange('M62:M66').copyTo(destSheet.getRange('L62:L66'))
  sourceSheet.getRange('N62:N66').copyTo(destSheet.getRange('M62:M66'),{contentsOnly:true})
  sourceSheet.getRange('B62').copyTo(destSheet.getRange('N62'))
  // Colocar valores en 0  
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange('N10').setValue(0);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange('N27').setValue(0);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange('N28').setValue(0);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange('N29').setValue(0);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange('N34').setValue(0);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange('N35').setValue(0);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange('N38').setValue(0);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange('N39').setValue(0);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos').getRange('N40').setValue(0);
}

function eventosAreas() {
  var ss = SpreadsheetApp.getActive();
  var sheetConfig = ss.getSheetByName('Config');
  var sheetCollector = ss.getSheetByName('Collector');
  var host = sheetConfig.getRange("B1").getValue();
  var database = sheetConfig.getRange("B2").getValue();
  var user = sheetConfig.getRange("B3").getValue();
  var password = sheetConfig.getRange("B4").getValue();
  var port = sheetConfig.getRange("B5").getValue();
  var FechaInicio = sheetCollector.getRange("L4").getValue();
  var FechaFin = sheetCollector.getRange("L5").getValue();
  var Cliente = sheetConfig.getRange("B6").getValue();  
  var url = 'jdbc:mysql://'+host+':'+port+'/'+database;
  var EventosGeneradosAreas = 'SELECT D.value_text AS Area, COUNT(1) Total, SUM(CASE WHEN T.ticket_state_id = 11 THEN 1 ELSE 0 END) Escalados, SUM(CASE WHEN T.user_id = 1 THEN 1 ELSE 0 END) SinAnalisis FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.customer_id = "'+Cliente+'" AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.value_text  in ("Windows","Conectividad","Unix","Oracle","Telefonia","AS400","database","Fortigate","Seguridad") group by D.value_text order by D.value_text  asc';

  try{
    var connection = Jdbc.getConnection(url, user, password);
    var result = connection.createStatement().executeQuery(EventosGeneradosAreas);
    var metaData = result.getMetaData();
    var columns = metaData.getColumnCount();  
    var values = [];
    var value = [];
    var element = '';

    for (i = 1; i <= columns; i ++){
      element = metaData.getColumnLabel(i);
      value.push(element);
    }
    values.push(value);
  
    while(result.next()){
      value = [];
      for (i = 1; i <= columns; i ++){
        element = result.getString(i);
        value.push(element);
      }
        values.push(value);
    }
  //Cierra conexion
    result.close();
   sheetCollector.getRange('D4:G13').clearContent();
  //Escribe datos en las celdas
    sheetCollector.getRange(4,4, values.length, value.length).setValues(values);
    SpreadsheetApp.getActive().toast('Datos actualizado correctamente en [Tab: Collector]!');
  }catch(err){
    SpreadsheetApp.getActive().toast(err.message);
  } 
}

function topAreas() {
 var ss = SpreadsheetApp.getActive();
 var sheetConfig = ss.getSheetByName('Config');
 var sheetCollector = ss.getSheetByName('Collector');
 var host = sheetConfig.getRange("B1").getValue();
 var database = sheetConfig.getRange("B2").getValue();
 var user = sheetConfig.getRange("B3").getValue();
 var password = sheetConfig.getRange("B4").getValue();
 var port = sheetConfig.getRange("B5").getValue();
 var FechaInicio = sheetCollector.getRange("L4").getValue();
 var FechaFin = sheetCollector.getRange("L5").getValue();
 var Cliente = sheetConfig.getRange("B6").getValue();
 var url = 'jdbc:mysql://' + host + ':' + port + '/' + database;
 var topWindows = 'SELECT T.title AS Titulo, count(2) AS Total FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.field_id = 15 AND D.value_text = "Windows" GROUP BY T.title ORDER BY Total DESC limit 0, 25';
 var topConectividad = 'SELECT T.title AS Titulo, count(2) AS Total FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.field_id = 15 AND D.value_text = "Conectividad" GROUP BY T.title ORDER BY Total DESC limit 0, 25';
 var topUnix = 'SELECT T.title AS Titulo, count(2) AS Total FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.field_id = 15 AND D.value_text = "Unix" GROUP BY T.title ORDER BY Total DESC limit 0, 25';
 var topOracle = 'SELECT T.title AS Titulo, count(2) AS Total FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.field_id = 15 AND D.value_text = "Oracle" GROUP BY T.title ORDER BY Total DESC limit 0, 25';
 var topTelefonia = 'SELECT T.title AS Titulo, count(2) AS Total FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.field_id = 15 AND D.value_text = "telefonia" GROUP BY T.title ORDER BY Total DESC limit 0, 25';
 var topAs400 = 'SELECT T.title AS Titulo, count(2) AS Total FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.field_id = 15 AND D.value_text = "as400" GROUP BY T.title ORDER BY Total DESC limit 0, 25';
 var topMssql = 'SELECT T.title AS Titulo, count(2) AS Total FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.field_id = 15 AND D.value_text = "database" GROUP BY T.title ORDER BY Total DESC limit 0, 25';
 var topFortigate = 'SELECT T.title AS Titulo, count(2) AS Total FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.field_id = 15 AND D.value_text = "Fortigate" GROUP BY T.title ORDER BY Total DESC limit 0, 25';
 var topSeguridad = 'SELECT T.title AS Titulo, count(2) AS Total FROM ticket T INNER JOIN dynamic_field_value D ON T.id = D.object_id WHERE T.queue_id IN (8, 9, 10) AND T.create_time BETWEEN concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-"),"01 00:00:00") AND concat(date_format(LAST_DAY(now() - interval 1 month),"%Y-%m-%d")," 23:59:59") AND D.field_id = 15 AND D.value_text = "Seguridad" GROUP BY T.title ORDER BY Total DESC limit 0, 25';

 try {
  var connection = Jdbc.getConnection(url, user, password);
  var result = connection.createStatement().executeQuery(topWindows);
  var metaData = result.getMetaData();
  var columns = metaData.getColumnCount();
  var values = [];
  var value = [];
  var element = '';

  for (i = 1; i <= columns; i++) {
   element = metaData.getColumnLabel(i);
   value.push(element);
  }
  values.push(value);

  while (result.next()) {
   value = [];
   for (i = 1; i <= columns; i++) {
    element = result.getString(i);
    value.push(element);
   }
   values.push(value);
  }
  //Cierra conexion
  result.close();
  sheetCollector.getRange('A33:B58').clearContent();
  //Escribe datos en las celdas
  sheetCollector.getRange(33, 1, values.length, value.length).setValues(values);
  SpreadsheetApp.getActive().toast('Top Windows Actualizado Correctamente');

  var connection = Jdbc.getConnection(url, user, password);
  var result = connection.createStatement().executeQuery(topConectividad);
  var metaData = result.getMetaData();
  var columns = metaData.getColumnCount();
  var values = [];
  var value = [];
  var element = '';

  for (i = 1; i <= columns; i++) {
   element = metaData.getColumnLabel(i);
   value.push(element);
  }
  values.push(value);

  while (result.next()) {
   value = [];
   for (i = 1; i <= columns; i++) {
    element = result.getString(i);
    value.push(element);
   }
   values.push(value);
  }
  //Cierra conexion
  result.close();
  sheetCollector.getRange('A61:B86').clearContent();
  //Escribe datos en las celdas
  sheetCollector.getRange(61, 1, values.length, value.length).setValues(values);
  SpreadsheetApp.getActive().toast('Top Conectividad Actualizado Correctamente');

  var connection = Jdbc.getConnection(url, user, password);
  var result = connection.createStatement().executeQuery(topUnix);
  var metaData = result.getMetaData();
  var columns = metaData.getColumnCount();
  var values = [];
  var value = [];
  var element = '';

  for (i = 1; i <= columns; i++) {
   element = metaData.getColumnLabel(i);
   value.push(element);
  }
  values.push(value);

  while (result.next()) {
   value = [];
   for (i = 1; i <= columns; i++) {
    element = result.getString(i);
    value.push(element);
   }
   values.push(value);
  }
  //Cierra conexion
  result.close();
  sheetCollector.getRange('A89:B114').clearContent();
  //Escribe datos en las celdas
  sheetCollector.getRange(89, 1, values.length, value.length).setValues(values);
  SpreadsheetApp.getActive().toast('Top Conectividad Unix Correctamente');

  var connection = Jdbc.getConnection(url, user, password);
  var result = connection.createStatement().executeQuery(topOracle);
  var metaData = result.getMetaData();
  var columns = metaData.getColumnCount();
  var values = [];
  var value = [];
  var element = '';

  for (i = 1; i <= columns; i++) {
   element = metaData.getColumnLabel(i);
   value.push(element);
  }
  values.push(value);

  while (result.next()) {
   value = [];
   for (i = 1; i <= columns; i++) {
    element = result.getString(i);
    value.push(element);
   }
   values.push(value);
  }
  //Cierra conexion
  result.close();
  sheetCollector.getRange('A117:B142').clearContent();
  //Escribe datos en las celdas
  sheetCollector.getRange(117, 1, values.length, value.length).setValues(values);
  SpreadsheetApp.getActive().toast('Top Oracle Actualizado Correctamente');
   
  var connection = Jdbc.getConnection(url, user, password);
  var result = connection.createStatement().executeQuery(topTelefonia);
  var metaData = result.getMetaData();
  var columns = metaData.getColumnCount();
  var values = [];
  var value = [];
  var element = '';

  for (i = 1; i <= columns; i++) {
   element = metaData.getColumnLabel(i);
   value.push(element);
  }
  values.push(value);

  while (result.next()) {
   value = [];
   for (i = 1; i <= columns; i++) {
    element = result.getString(i);
    value.push(element);
   }
   values.push(value);
  }
  //Cierra conexion
  result.close();
  sheetCollector.getRange('A145:B170').clearContent();
  //Escribe datos en las celdas
  sheetCollector.getRange(145, 1, values.length, value.length).setValues(values);
  SpreadsheetApp.getActive().toast('Top Telefonia Actualizado Correctamente');   
   
  var connection = Jdbc.getConnection(url, user, password);
  var result = connection.createStatement().executeQuery(topAs400);
  var metaData = result.getMetaData();
  var columns = metaData.getColumnCount();
  var values = [];
  var value = [];
  var element = '';

  for (i = 1; i <= columns; i++) {
   element = metaData.getColumnLabel(i);
   value.push(element);
  }
  values.push(value);

  while (result.next()) {
   value = [];
   for (i = 1; i <= columns; i++) {
    element = result.getString(i);
    value.push(element);
   }
   values.push(value);
  }
  //Cierra conexion
  result.close();
  sheetCollector.getRange('A173:B198').clearContent();
  //Escribe datos en las celdas
  sheetCollector.getRange(173, 1, values.length, value.length).setValues(values);
  SpreadsheetApp.getActive().toast('Top AS400 Actualizado Correctamente.');   

   
 } catch (err) {
  SpreadsheetApp.getActive().toast(err.message);
 }
}