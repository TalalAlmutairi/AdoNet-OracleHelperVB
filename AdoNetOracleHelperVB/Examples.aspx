<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Examples.aspx.vb" Inherits="AdoNetOracleHelperVB.Examples" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:GridView ID="GridView1" runat="server"></asp:GridView> <br />
            <asp:Label ID="lbMsg" runat="server" Text="Label"></asp:Label><br />
            <asp:Button ID="btnLoadData" runat="server" Text="Load Data" OnClick="btnLoadData_Click" />
            <asp:Button ID="btnSqlWhere" runat="server" Text="Sql Where" OnClick="btnSqlWhere_Click" />
            <asp:Button ID="btnExecuteScalar" runat="server" Text="Execute Scalar" OnClick="btnExecuteScalar_Click" />
            <asp:Button ID="btnSP" runat="server" Text="Stored Procedure" OnClick="btnSP_Click" />
             <asp:Button ID="btnExecuteTransaction" runat="server" Text="Execute Transaction" OnClick="btnExecuteTransaction_Click" />
            <asp:Button ID="btnReturnDS" runat="server" Text="DataSet" OnClick="btnReturnDS_Click" />
            <asp:GridView ID="GridViewEmp" runat="server"></asp:GridView> <br />
            <br />
            <asp:GridView ID="GridViewCountry" runat="server"></asp:GridView> <br />
            For Insert, update, and delete <br />
            <input id="txtEmpID" type="text" runat="server" placeholder="Employee ID" />
            <input id="txtFName" type="text" runat="server" placeholder="First Name" />
            <input id="txtLName" type="text" runat="server" placeholder="Last Name" />
            <input id="txtAge" type="number" min="18" max="50" step="1" runat="server" placeholder="Age" />
            <select id="ddlCouontries" runat="server">
                <option value="1">Saudi Arabia</option>
                <option value="2">Kuwait</option>
                 <option value="3">United Arab Emirates</option>
            </select>
            <br />
            <asp:Button ID="btnInsert" runat="server" Text="Insert" OnClick="btnInsert_Click" />
            <asp:Button ID="btnUpdate" runat="server" Text="Update" OnClick="btnUpdate_Click" /> 
            <asp:Button ID="btnDelete" runat="server" Text="Delete" OnClick="btnDelete_Click" /> <br />
           
        </div>
    </form>
</body>
</html>
