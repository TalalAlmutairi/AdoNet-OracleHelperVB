Imports Oracle.ManagedDataAccess.Client

Partial Public Class Examples
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
    End Sub

    ' Getting all employees data
    Protected Sub btnLoadData_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sql As String = "SELECT * FROM Employees"
        GridView1.DataSource = OracleHelper.ExecuteQuery(sql, CommandType.Text, Nothing)
        GridView1.DataBind()
    End Sub

    ' Selecting only the employee with id=2
    Protected Sub btnSqlWhere_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sql As String = "SELECT * FROM Employees WHERE EmpID=:pID"

        'SqlParameter uses for security to prevent SQL injection
        '@ID in the SQL above must match the parameter in SqlParameter
        Dim parametersList As OracleParameter() = New OracleParameter() {
            New OracleParameter(":pID", "2")
        }
        GridView1.DataSource = OracleHelper.ExecuteQuery(sql, CommandType.Text, parametersList)
        GridView1.DataBind()
    End Sub

    ' Getting the maximum salary of all employees as only one value
    Protected Sub btnExecuteScalar_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sql As String = "SELECT MAX(Age) FROM Employees"
        lbMsg.Text = OracleHelper.ExecuteScalar(sql, CommandType.Text, Nothing)
    End Sub



    'Oracle PACKAGE

    'CREATE Or REPLACE PACKAGE EMP_TS.EMPLOYEES .GET_ALL_EMPLOYEES AS
    'TYPE refcur Is REF CURSOR;
    'PROCEDURE GET_EMPLOYEES_INFO(CurEmp OUT GET_ALL_EMPLOYEES.refcur);
    'END GET_ALL_EMPLOYEES;
    '/

    'CREATE Or REPLACE PACKAGE BODY EMP_TS.EMPLOYEES .GET_ALL_EMPLOYEES Is
    'PROCEDURE GET_EMPLOYEES_INFO(CurEmp OUT GET_ALL_EMPLOYEES.refcur) Is
    'BEGIN
    'OPEN CurEmp FOR SELECT * FROM Employees;
    'END GET_EMPLOYEES_INFO;
    'END;
    '/

    ' Execute stored procedure
    Protected Sub btnSP_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim PKG_SP_Name As String = "GET_ALL_EMPLOYEES.GET_EMPLOYEES_INFO"

        Dim parametersList As OracleParameter() = New OracleParameter() {
            New OracleParameter("CurEmp", OracleDbType.RefCursor, ParameterDirection.Output)
        }
        GridView1.DataSource = OracleHelper.ExecuteQuery(PKG_SP_Name, CommandType.StoredProcedure, parametersList)
        GridView1.DataBind()
    End Sub


    Protected Sub btnInsert_Click(ByVal sender As Object, ByVal e As EventArgs)

        Dim sql As String = "INSERT INTO Employees VALUES(SEQ_EMPID.NEXTVAL,:pFName,:pLName,:pAge,:pCountryID)"

        Dim parametersList As OracleParameter() = New OracleParameter() {
            New OracleParameter(":pFName", txtFName.Value),
            New OracleParameter(":pLName", txtLName.Value),
            New OracleParameter(":pAge", txtAge.Value),
            New OracleParameter(":pCountryID", ddlCouontries.Value)
        }

        If OracleHelper.ExecuteNonQuery(sql, CommandType.Text, parametersList) Then
            lbMsg.Text = "Inserted successfully"
        Else
            lbMsg.Text = "Error"
        End If
    End Sub

    Protected Sub btnUpdate_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sql As String = "UPDATE Employees SET FirstName=:pFName,LastName=:pLName,Age=:pAge,CountryID=:pCountryID WHERE EmpID =:pID"
        Dim parametersList As OracleParameter() = New OracleParameter() {
            New OracleParameter(":pID", txtEmpID.Value),
            New OracleParameter(":pFName", txtFName.Value),
            New OracleParameter(":pLName", txtLName.Value),
            New OracleParameter(":pAge", txtAge.Value),
            New OracleParameter(":pCountryID", ddlCouontries.Value)
        }

        If OracleHelper.ExecuteNonQuery(sql, CommandType.Text, parametersList) Then
            lbMsg.Text = "Updated successfully"
        Else
            lbMsg.Text = "Error"
        End If
    End Sub

    Protected Sub btnDelete_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim sql As String = "DELETE FROM Employees WHERE EmpID =:pID"
        Dim parametersList As OracleParameter() = New OracleParameter() {New OracleParameter(":pID", txtEmpID.Value)}

        If OracleHelper.ExecuteNonQuery(sql, CommandType.Text, parametersList) Then
            lbMsg.Text = "Deleted successfully"
        Else
            lbMsg.Text = "Error"
        End If
    End Sub

    'Execute two SQL statements Insert and update, which all SQL statements in a single transaction, rolling back if an error has occurred
    Protected Sub btnExecuteTransaction_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim listOfSQLs As ArrayList = New ArrayList()
        Dim listOfParamerters As List(Of OracleParameter()) = New List(Of OracleParameter())()

        Dim sql1 As String = "INSERT INTO Employees VALUES(SEQ_EMPID.NEXTVAL,:pFName,:pLName,:pAge,:pCountryID)"
        Dim parameters1 As OracleParameter() = New OracleParameter() {
            New OracleParameter(":pFName", "Test F Name"),
            New OracleParameter(":pLName", "Test L Name"),
            New OracleParameter(":pAge", 25),
            New OracleParameter(":pCountryID", 1)
        }
        listOfSQLs.Add(sql1)
        listOfParamerters.Add(parameters1)

        Dim sql2 As String = "UPDATE Employees SET FirstName=:pFName,LastName=:pLName,Age=:pAge,CountryID=:pCountryID WHERE EmpID =:pID"
        Dim parameters2 As OracleParameter() = New OracleParameter() {
            New OracleParameter(":pFName", "New F Name"),
            New OracleParameter(":pLName", "New L Name"),
            New OracleParameter(":pAge", 30),
            New OracleParameter(":pCountryID", 2),
            New OracleParameter(":pID", 4) 'This number for testing, make sure a record with id=4 is exist
        }
        listOfSQLs.Add(sql2)
        listOfParamerters.Add(parameters2)

        If OracleHelper.ExecuteTransaction(listOfSQLs, listOfParamerters) Then
            lbMsg.Text = "All SQL statements executed successfully"
        Else
            lbMsg.Text = "Error"
        End If
    End Sub


    'Execute two select queries, and returns employees and country tables
    Protected Sub btnReturnDS_Click(ByVal sender As Object, ByVal e As EventArgs)

        'make sure parameter name of Cursor (CurEmp and CurCountry ) same as the name in the oracle PACKAGE 
        Dim PKG_SP_Name As String = "GET_MULTIPLE_TABLES.GET_TABLES"
        Dim parametersList As OracleParameter() = New OracleParameter() {
            New OracleParameter("CurEmp", OracleDbType.RefCursor, ParameterDirection.Output),
            New OracleParameter("CurCountry", OracleDbType.RefCursor, ParameterDirection.Output)
        }
        Dim ds As DataSet = OracleHelper.ExecuteQueryDS(PKG_SP_Name, CommandType.StoredProcedure, parametersList)

        GridViewEmp.DataSource = ds.Tables(0)
        GridViewEmp.DataBind()
        GridViewCountry.DataSource = ds.Tables(1)
        GridViewCountry.DataBind()
    End Sub
End Class
