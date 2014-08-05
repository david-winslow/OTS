<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="wwwroot._default" EnableEventValidation="false" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
  <h3>We suggest the following:</h3>

<ol class="round">
    <li class="one">
        <h5>Getting Started</h5>
        First Fill out the template data in excel with your client. The current template can be found <a href="../Web/templates/testData.xlsx">here</a>
    </li>

    <li class="two">
        <h5>Upload your excel data file </h5>
        Click browse and select your filled in Excel data file              
        <form method="POST" enctype="multipart/form-data">
            @Html.ValidationSummary()
            <ol>
                
                <li><p>Upload your <b>Excel file</b> here</p>  
                    <input type="file" id="excelFile" name="file" runat="server"/>
                </li>
                <li>
                    <p>Optionally change the <b>Word template</b> here</p>
                    <input type="file" id="wordTemplateFile" name="file" runat="server" />
                </li>
            </ol>
            <input type="submit" id="btnSubmit" value="Generate FCE report" runat="server"  />
        </form>     
    </li>

    <li class="three">
        <h5>Download your report</h5>
        <br/>
        <asp:HyperLink Visible="False" ID="HyperLink1" runat="server">Generated report</asp:HyperLink>
        
    </li>
</ol>
    </form>
</body>
</html>
