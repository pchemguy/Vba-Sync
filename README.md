# VBA-Sync

In VBA, asynchronous queries require a class module, a unique ADODB.Command object, a unique ADODB.Connection object, and a unique user-defined event linked to the ADODB.Conection object. This class module abstracts away the complicated and boilerplate code of executing asynchronous queries to a database in VBA. 

Although this class is mainly designed for executing asynchronous queries, it is capable of executing synchronous queries as well. This functionality can be useful if you need to do something like run a SQL package to populate various temp tables. That query must be executed synchronously because the temp tables must be populated before any select queries can be executed on them. Once these temp tables are populated, any select queries that query them can be run asynchronously.

## Getting started

I've named the class module cQuerable, but you can name it whatever you want. If you don't name it as cQueryable, you have to create an instance of whatever you decide to name it in order to use it. All of my examples will assume a cQueryable class module name.

As I noted in the code, the class modules requires a reference to the Microsoft ActiveX Data Objects 6.1 library. It will not work without a reference to that library or a similar one (I've only tested on the 6.1 library.)

NOTE: cQueryable variables must be declared with module-level scope. While synchronous queries may work with local scope, asynchronous queries will not. So in the normal module that you write your executable code, ensure that any cQueryable variables you create have module-level scope.

For a detailed overview of the properties and methods in the class, please see the [wiki](https://github.com/beyphy/VBA-SQL-Async/wiki).

## Current limitations / future features

cQueryable does not have extensive error handling. I do not currently check whether the connection string you've provided is valid, or whether you've provided a query to the SQL property for example. If these things are not provided, the program will crash, but no detailed, custom error will be provided. I will likely implement custom errors in a future release.

## A note on support

Be aware that I have limited time for feature requests or bug fixes. And even if I have some time to do those things now, I may not in the future. Also be aware that while this code works for me, it has not been extensively tested. I would recommend extensive testing if you plan on using this code in a production environment to ensure it works and fits your needs.

## Example

    Option Explicit
    
    'this code is in a normal module
    
    Private QueryableArr(2) As cQueryable
    
    Sub AsyncQueryExample()
        Dim ConnectionString As String
        Dim i As Long
        
        ConnectionString = "Dsn=MyDsn"
        
        For i = LBound(QueryableArr) To UBound(QueryableArr)
            Set QueryableArr(i) = New cQueryable
            QueryableArr(i).ConnectionString = ConnectionString
        Next i
        
        QueryableArr(0).Sql = "select pg_sleep(10)"
        QueryableArr(1).Sql = "Select * from sales.invoices"
        QueryableArr(2).Sql = "select * from sales.rates"
        
        QueryableArr(1).procedureAfterQuery = "updateSheet1"
        QueryableArr(2).procedureAfterQuery = "updateSheet2"
        
        QueryableArr(0).AsyncExecute
        QueryableArr(1).AsyncExecute
        QueryableArr(2).AsyncExecute
    End Sub

    Private Sub updateSheet1(rs As ADODB.Recordset)
        With Sheet1.Range("A1")
            .CurrentRegion.ClearContents
            .CopyFromRecordset rs
        End With
    End Sub
    
    Private Sub updateSheet2(rs As ADODB.Recordset)
        With Sheet2.Range("A1")
            .ClearContents
            .CopyFromRecordset rs
        End With
    End Sub
