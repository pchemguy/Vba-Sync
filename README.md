# VBA-Sync

In VBA, to use asynchronous queries you need an object module or a class module, an ADODB.Connection object, and a user-defined event linked to the ADODB.Conection object. While this is not that much code in itself, it is limited. It is not capable of supporting queries that utilize ordinal parameters. For that, you also need an ADODB.Command object and an ADODB.Parameter object for every ordinal argument you have. So this class module abstracts away the complicated and boilerplate code of executing asynchronous queries to a database in VBA. 

Although this class is mainly designed for executing asynchronous queries, it is capable of executing synchronous queries as well. This functionality can be useful if you need to do something like run a SQL package to populate various temp tables. That query must be executed synchronously because the temp tables must be populated before any select queries can be executed on them. Once these temp tables are populated, any select queries that query them can be run asynchronously.

## Getting started

I've named the class module cQuerable, but you can name it whatever you want. If you don't name it as cQueryable, you have to create an instance of whatever you decide to name it in order to use it. All of my examples will assume a cQueryable class module name.

As I noted in the code, **the class modules requires a reference to the Microsoft ActiveX Data Objects 6.1 library.** It will not work without a reference to that library or a similar one (I've only tested on the 6.1 library.)

**NOTE:** cQueryable variables must be declared with **module-level scope**. While synchronous queries may work with local scope, asynchronous queries will not. So in the normal module that you write your executable code, ensure that any cQueryable variables you create have module-level scope.

## Untested usage

This code was developed to be utilized on SQL queries to a database. Under the hood though, it just utilizes the objects in the ADODB library. Since ADODB can connect to a variety of data sources provided that a driver is supplied, it should be able to be utilized in situations other than SQL queries to a database. However, this is out of the scope of the project and not something I have tested. Part of the reasons for this is that finding the connection strings for the different data sources can be a pain. Not only do you need to find the correct one, but the one you need to use may vary depending on whether you're in a 32 bit or 64 bit environment.

## A note on support

At this point, I believe the code for this module is essentially complete. While the code utilizes some limited error handling, be aware that I have limited time for feature requests or bug fixes. And even if I have some time to do those things now, I may not in the future. Also be aware that while this code works for me, it has not been extensively tested. I would recommend extensive testing if you plan on using this code in a production environment to ensure it works and fits your needs.

# Examples

## Synchronous and asynchronous queries

    Option Explicit
    
    'this code is in a normal module
    
    Private QueryableArr(2) As cQueryable
    
    Sub AsyncQueryExample1()
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
        
        QueryableArr(1).AsyncProcedure = "updateSheet1"
        QueryableArr(2).AsyncProcedure = "updateSheet2"
        
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

## Ordinal parameters query

    Private queryable As cQueryable
    
    'this is a normal code module
    
    Sub AsyncQueryExample2()
        
        Set queryable = New cQueryable
        
        With queryable
            .ConnectionString = "Dsn=MyDsn"
            .Sql = "select * from company.customers where first_name = ? and age > ?"
            .createParam "firstName", adVarChar, "John", pSize:=50
            .createParam "age", adInteger, 30
            .AsyncProcedure = "updateSheet1"
            .AsyncExecute
        End With
    End Sub
