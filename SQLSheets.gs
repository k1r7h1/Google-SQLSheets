/*
Author: Kirthi Banothu
Project Details: Allows users to interact with an SQL database using SQL queries while leveraging google spreadsheet capabilities

  Functions:                      Example use:
  initializeTable                 initializeTable("likes")
  queryDB                         queryDB("SELECT * FROM beers")
  insert                          insert("likes", B16:C16)
  tablesInDB                      tablesInDB()
  update                          update("UPDATE likes SET beers='Budweiser' WHERE drinkers='Peter' ")
  deleteRow                       deleteRow("drinkers",B44:C44)
  createTable                     createTable("Persons", "PersonID int, LastName varchar(255)");
  dropTable                       dropTable("Persons")

*/

// Global Variables 
  var conn; 
  var stmt; 
  var rs; 
  var start;



function incrementRefresh() {
  // Increments the values in all the cells in the active range (i.e., selected cells).
  // Numbers increase by one, text strings get a "1" appended.
  // Cells that contain a formula are ignored.
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeRange = ss.getActiveRange();
  var cell, cellValue, cellFormula;
  
  // iterate through all cells in the active range
  for (var cellRow = 1; cellRow <= activeRange.getHeight(); cellRow++) {
    for (var cellColumn = 1; cellColumn <= activeRange.getWidth(); cellColumn++) {
      cell = activeRange.getCell(cellRow, cellColumn);
      //data[2][3];
      cellFormula = cell.getFormula();
      
      // if not a formula, increment numbers by one, or add "1" to text strings
      // if the leftmost character is "=", it contains a formula and is ignored
      // otherwise, the cell contains a constant and is safe to increment
      // does not work correctly with cells that start with '=
      if (cellFormula[0] != "=") {
        cellValue = cell.getValue();
        cell.setValue(cellValue + 1);
      }
    }
  }
}


/*
  Establishes connection to the database. Using JDBC driver to connect to MySQL database. 
  Retrieves parameters hostname, instance, username, and password from 2nd column of spreadsheet.
  data[0][1]: Database Host
  data[1][1]: Database Instance
  data[2][1]: Database User
  data[3][1]: Databse Password
*/
function connectToDatabase(){
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  // JDBC connection
    conn = Jdbc.getConnection('jdbc:mysql://' + data[0][1] + ':3306/' + data[1][1], data[2][1], data[3][1]);
    stmt = conn.createStatement();
    stmt.setMaxRows(data[4][1]);
};

/*
  Triggered after every query to ensure all connections are closed. 
*/
function closeConnection(){
  if(rs != undefined) {
    rs.close();
  };
  stmt.close();
  conn.close();
};

/*
  Used to return all tables within the instance specified.
  No parameters - Defaults on DB connection. 
*/
function tablesInDB(){
  connectToDatabase();
 
  rs = stmt.executeQuery("SHOW tables");
  
  var row = 0;
  var out=[];
  var rowAdd=[];
   
  var header=[];

  for(var i = 0; i < rs.getMetaData().getColumnCount(); i++){
    header.push(rs.getMetaData().getColumnName(i+1));
  }
  out.push([ , header]);
  
  while (rs.next()) {
    for (var col = 0; col < rs.getMetaData().getColumnCount(); col++) {
      //cell.offset(row, col).setValue(rs.getString(col + 1));
      rowAdd.push(rs.getString(col+1));
    }
    out.push([row + 1, rowAdd]);
    rowAdd = []
    row++;
  }
  
  closeConnection();
  
  return out; 
};

/*
  Quick way to retrieve entire table from database
  Parameter: Table Name 
  Returns: Entire Table
*/
function initializeTable(table_name){
  connectToDatabase();

  var sqlQuery = "SELECT * FROM " + table_name;
  rs = stmt.executeQuery(String(sqlQuery));
  
  var row = 0;
  var out=[];
  var rowAdd=[];
   
  var header=[];
  //out.push([0, rs.getMetaData().getTableName(1)]);
  for(var i = 0; i < rs.getMetaData().getColumnCount(); i++){
    header.push(rs.getMetaData().getColumnName(i+1));
  }
  out.push(["Table Name: ", table_name]);
  
  out.push(["Headers:", header]);
  while (rs.next()){
    for (var col = 0; col < rs.getMetaData().getColumnCount(); col++) {
      rowAdd.push(rs.getString(col+1));
    }
    out.push([row, rowAdd]);
    rowAdd = []
    row++;
  }
  
  closeConnection();
  return out;
};

/*
  Can use for general case queries to database. 
  Parameter: query to be executes. Example: SELECT name FROM persons WHERE age = 15;
  Restriction: Cannot use for queries that are designed to modify database. Only for retrival of data. 
*/
function queryDB(stringQuery){
  connectToDatabase();
 
  rs = stmt.executeQuery(stringQuery);
  
  var row = 0;
  var out=[];
  var rowAdd=[];
  var header=[];

  for(var i = 0; i < rs.getMetaData().getColumnCount(); i++){
    header.push(rs.getMetaData().getColumnName(i+1));
  }
  out.push([ , header]);
  
  while (rs.next()) {
    for (var col = 0; col < rs.getMetaData().getColumnCount(); col++) {
      //cell.offset(row, col).setValue(rs.getString(col + 1));
      rowAdd.push(rs.getString(col+1));
    }
    out.push([row, rowAdd]);
    rowAdd = []
    row++;
  }

  closeConnection();
  return out; 
};

/*
  Used to insert data into MySQL database. 
  Parameters: Table Name & Values. Example: insert("likes", B16:C16)
*/
function insert(table_name, values){
  connectToDatabase();
  
  /* 
    INSERT INTO table_name
    VALUES (value1,value2,value3,...);
  */
  
  var questionValues = "?";
  for(j = 1; j < values[0].length; j++){
    questionValues = questionValues + ", ?";
  }
  
  var stringQuery = "INSERT INTO " + table_name + " VALUES (" + questionValues+ ")";
  
  var stmt = conn.prepareStatement(String(stringQuery));
  
  for(i = 0; i < values[0].length; i++){
    stmt.setObject(i+1, values[0][i]);
  }
    stmt.addBatch();
    
    var res = stmt.executeBatch();
  
    
     //rs = stmt.executeQuery("INSERT INTO drinkers VALUES ('apple', 'banana')");
  
  var addedRow = [];
  // var rowCount = stmt.executeQuery("SELECT COUNT(*) FROM " + table_name);
  // addedRow.push([rowCount , values]);
 
  //return SpreadsheetApp.getActiveSheet().getActiveCell().getRow(); 
  closeConnection();
  return "Added!"; 
};

/*
  General case method to execute queries that made changes to database. 
  Parameters: SQL query
*/
function update(query){
  connectToDatabase();
  /* 
    INSERT INTO table_name
    VALUES (value1,value2,value3,...);
  */
  
  //var stringQuery = "INSERT INTO " + table_name + " VALUES (?, ?)";
  var updateQuery = query;

  var stmt = conn.prepareStatement(String(updateQuery));
  //stmt.setObject(1, values[0][0]);
  //stmt.setObject(2, values[0][1]);
  stmt.addBatch();
  var res = stmt.executeBatch();

  //rs = stmt.executeQuery("INSERT INTO drinkers VALUES ('apple', 'banana')");

  var addedRow = [];
  // var rowCount = stmt.executeQuery("SELECT COUNT(*) FROM " + table_name);
  // addedRow.push([rowCount , values]);
 
  //return SpreadsheetApp.getActiveSheet().getActiveCell().getRow(); 
  closeConnection();
  return "updated!";
};

/*
  Used to delete an entire row for a table. 
  Requirements: Give entire row in values parameter
  Parameters: Table Name & Row to be deleted.
*/
function deleteRow(table_name, values){
  
  var header = getHeader(table_name);
  
  var stringWhereClause;
  stringWhereClause = header[0] + "='" + values[0][0] + "' ";

  for(var j = 1; j < header.length; j++){
    stringWhereClause = stringWhereClause + "AND " +header[j] + "='" + values[0][j] + "' ";
  }
  
  var stringDeleteQuery = "DELETE FROM " + table_name + " WHERE " + stringWhereClause;

  connectToDatabase();
  /* 
    INSERT INTO table_name
    VALUES (value1,value2,value3,...);
  */
  
  //var stringQuery = "INSERT INTO " + table_name + " VALUES (?, ?)";
  
  //var deleteQuery = query;

  var stmt = conn.prepareStatement(String(stringDeleteQuery));
  //stmt.setObject(1, values[0][0]);
  //stmt.setObject(2, values[0][1]);
  stmt.addBatch();
  
  var res = stmt.executeBatch();

  
  //rs = stmt.executeQuery("INSERT INTO drinkers VALUES ('apple', 'banana')");
  
  var addedRow = [];
  // var rowCount = stmt.executeQuery("SELECT COUNT(*) FROM " + table_name);
  // addedRow.push([rowCount , values]);
 
  //return SpreadsheetApp.getActiveSheet().getActiveCell().getRow(); 
  closeConnection();
  return "Deleted!";
};

/*
  Can be used to create a table.
  Parameters: Table name & table column properties
    Example: "Persons", "PersonID int, LastName varchar(255)"
*/
function createTable(table_name, parameters){
  connectToDatabase();

  /*
    CREATE TABLE Persons
    (
      PersonID int,
      LastName varchar(255),
      FirstName varchar(255),
      Address varchar(255),
      City varchar(255)
    );
  */

  stringCreateTable = "CREATE TABLE " + table_name + " ( " + parameters + " );";
  var stmt = conn.prepareStatement(String(stringCreateTable));
  stmt.addBatch();
  var rs = stmt.executeBatch();

  closeConnection();

  return "Table Created";

};

/*
  Used to delete/drop a table.
  Parameters: Table name
*/
function deleteTable(table_name){
    connectToDatabase();
  
    stringDropTable = "DROP TABLE  " + table_name;
    var stmt = conn.prepareStatement(String(stringDropTable));
    stmt.addBatch();
    var rs = stmt.executeBatch();
  
    closeConnection();
  
    return "Table " + table_name + " Dropped";
};

/*
  function onEdit(e) {
    if (e.range.getValue() == "_now") {
      e.range.setValue(new Date());
    }
  }
*/

/*
  Gets headers of a given table
  Parameter: table_name 
*/
function getHeader(table_name){
  connectToDatabase();

  // gets header values 
  var keysString = "SELECT * FROM " + table_name + " LIMIT 1";
  rs = stmt.executeQuery(String(keysString));

  var header=[];
  for(var i = 0; i < rs.getMetaData().getColumnCount(); i++){
    header.push(rs.getMetaData().getColumnName(i+1));
  }

  closeConnection();
  return header;
};

/*
  Sample code structure. Shows how to push data fundamentally to google spreadsheet from MySQL
*/
function sample() {
  connectToDatabase();
  rs = stmt.executeQuery('select * from beers');

  var doc = SpreadsheetApp.getActiveSpreadsheet();
  //var cell = doc.getRange('a8');
 
  var row = 0;
  var out=[];
  var rowAdd=[];  
  var header=[];
  //out.push([0, rs.getMetaData().getTableName(1)]);
  
  for(var i = 0; i < rs.getMetaData().getColumnCount(); i++){
    header.push(rs.getMetaData().getColumnName(i+1));
  }
  out.push([ , header]);
  while (rs.next()) {
    for (var col = 0; col < rs.getMetaData().getColumnCount(); col++) {
      //cell.offset(row, col).setValue(rs.getString(col + 1));
      rowAdd.push(rs.getString(col+1));
    }
    out.push([row, rowAdd]);
    rowAdd = []
    row++;
  }
  
  closeConnection();
  //var end = new Date();
  //Logger.log('Time elapsed: ' + (end.getTime() - start.getTime()));
  
  return out;
};

/* 
  Given parameters in the data array, you can retrieve data from any cell.
  Use this to initialize database 
*/
function getValueOfCell(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  return data[3][1];
};

/**
* Find an element in array.
*
* @param {array} array where element need to be searched
* @param {object} element that will be searched in the array.
* @return {number} index of the element, -1 if not found.
*/
function findInArray_(arr, ele) {
  for (var i=0; i < arr.length; i++)
    if (arr[i]==ele)
      return i;
  return -1;
};

/**
* Check if the input row is empty.
*
* @param {array} row to be checked.
* @param {number} number of columns in the row.
* @return {boolean} if range is empty return true.
*/
function emptyRow_(arr, columnCount) {
  if (typeof arr=="undefined")
    return true;
  
  for (var i=0; i < columnCount; i++)
    if (typeof arr[i]=="string" && arr[i].length>0)
      return false;
  return true;
};

/**
* Get data from a range.
*
* @param {range} range containing a table.
* @return {range} table in form of a 2d array.
*/
function getDataRange_(inputRange){
  var out=[];
  if (inputRange.length < 2)
      throw "Range should atleast have a header";
  
  /* Copy the table Name */
  var tableName=[];
  tableName.push(inputRange[0][0])
  out.push(tableName);
  
  var column_count = -1;
  var column_hdr = [];
  /* Copy attribute name, and get column count */
  while (typeof inputRange[1][++column_count]=="string" && inputRange[1][column_count].length > 0)
    column_hdr.push(inputRange[1][column_count]);

  out.push(column_hdr);
 
  var row_count=-1;
  while(!emptyRow_(inputRange[++row_count+2], column_count))
  {
    var row=[];
    for (var i=0;i<column_count;i++) 
      row.push(inputRange[row_count+2][i]);
    
    out.push(row);
  } 
  return out;
};
