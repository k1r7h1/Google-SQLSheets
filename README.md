Google-SQLSheets
================

Google spreadsheets implemented to connect seemlessly with an SQL database for interaction


Link to Google spreadsheet implmentation: https://docs.google.com/spreadsheet/ccc?key=0AqWxps8PkYg2dGROTHZraWx0N3hSR2FqZG9wd3h0OWc&usp=sharing

Code available in SQLSheets.gs file.

/*
Author: Kirthi Banothu
Project Details: Allows users to interact with an SQL database using SQL queries while leveraging google spreadsheet capabilities

  Functions:
  
  initializeTable | initializeTable("likes")
  
  queryDB | queryDB("SELECT * FROM beers")
  
  insert | insert("likes", B16:C16)
  
  tablesInDB | tablesInDB()
  
  update | update("UPDATE likes SET beers='Budweiser' WHERE drinkers='Peter' ")
  
  deleteRow | deleteRow("drinkers",B44:C44)
  
  createTable | createTable("Persons", "PersonID int, LastName varchar(255)");
  
  dropTable | dropTable("Persons")
  
*/










