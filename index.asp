
<style>
.thead{
  background-color: rgba(128, 128, 128, 0.458);
  
  }
</style>
<table class="table">
  <thead class="thead">
    <tr class="theadTr">
      <th class="theadTh">name
      </th>
       <th class="theadTh">id
      </th>
       <th class="theadTh">description
      </th>
       <th class="theadTh">created
      </th>
       <th class="theadTh">action
      </th>
    </tr>
  </thead>
  <tbody>
           <% Dim conn 
    Dim fainul
    fainul = "khong thanh cong"
    Set conn = Server.CreateObject("ADODB.Connection")
    On Error Resume Next
    
    conn.Open "Provider=SQLOLEDB;Data Source=LAPTOP-UHRJ0SA1;Initial Catalog=test;User Id=sa;Password=123;"
    
     Dim rs
      Set rs=Server.CreateObject("ADODB.Recordset") 
      rs.Open "SELECT * FROM st" , conn 
      If Err.Number <> 0 Then
        Response.Write (fainul )
        Else

        If Not rs.EOF Then
        Do While Not rs.EOF
       
    Response.Write "<tr class='names'><td class='name'>"&rs("name")&"</td></tr>"
    
        rs.MoveNext
        Loop
        End If
        End If
        conn.Close
        %>
  </tbody>
</table>
