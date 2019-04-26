Imports System.Data.SqlClient
Imports System.Data
''' <summary>
'''  集成最常用的对SQL Server数据库的添加修改删除操作和对数据表的填充
''' </summary>
''' <remarks></remarks>
Public Class SqlServerOperaion
    '连接字符串
    Shared conn As New SqlConnection("Data Source=192.168.1.100;Initial Catalog=hbposv7;uid=sa; password=zaqxswcde")
    'Shared conn As New SqlConnection("Data Source=localhost;Initial Catalog=hbv7test;uid=sa; password=zaqxswcde")

    ''' <summary>
    ''' 根据SQL语句返回值
    ''' 返回true时表示有结果集
    ''' 返回false时表示没有结果集
    ''' </summary>
    ''' <param name="sqlstr">查询语句</param>
    ''' <returns>返回true时表示值存在，false表示值不存在</returns>
    ''' <remarks></remarks>
    Shared Function SqlJudge(ByVal sqlstr As String) As Boolean
        Dim judeg As Boolean = False
        conn.Open()
        Dim da As SqlDataAdapter = New SqlDataAdapter(sqlstr, conn)
        Dim ds As New System.Data.DataSet()
        da.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            judeg = True
        End If
        conn.Close()
        Return judeg
    End Function

    ''' <summary>
    '''  根据输入的SQL查询语句返回数据填充好的DataTable
    ''' </summary>
    ''' <param name="sqlStr">SQL查询语句</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Shared Function SqlInquiresDataTable(ByVal sqlStr As String) As DataTable
        conn.Open()
        Dim da As SqlDataAdapter = New SqlDataAdapter(sqlStr, conn)
        Dim ds As New System.Data.DataSet()
        da.Fill(ds)
        Dim t As DataTable = ds.Tables(0)
        conn.Close()
        Return t
    End Function


    ''' <summary>
    ''' 根据SQL语句对SQL Server数据库进行添加、修改、删除操作
    ''' </summary>
    ''' <param name="sqlstr">需要执行的SQL语句</param>
    ''' <remarks></remarks>
    Shared Sub SqlOperaion(ByVal sqlstr As String)
        conn.Open()
        Dim comm As New SqlCommand(sqlstr, conn)
        comm.ExecuteNonQuery()
        conn.Close()
    End Sub
    ''' <summary>
    ''' 事务封装，输入字符串数组执行
    ''' </summary>
    ''' <param name="sqlstr">需要执行的SQL语句</param>
    ''' <returns>Boolean：成功执行为True</returns>
    ''' <remarks></remarks>
    Shared Function SqlTransactions(ByVal sqlstr() As String) As Boolean
        Dim j As Boolean = True
        conn.Open()
        Dim myTrans As SqlTransaction = conn.BeginTransaction
        Dim comm As New SqlCommand()
        comm.Connection = conn
        comm.Transaction = myTrans
        Try
            Dim i As Integer = 0
            While i < sqlstr.Length
                comm.CommandText = sqlstr(i)
                comm.ExecuteNonQuery()
                i = i + 1
            End While
            myTrans.Commit()
        Catch ex As Exception
            myTrans.Rollback()
            j = False
        Finally
            conn.Close()
        End Try
        Return j
    End Function
    'Shared Sub del(ByVal sqlstr As String)
    '    conn.Open()
    '    Dim comm As New SqlCommand(sqlstr, conn)
    '    comm.ExecuteNonQuery()
    '    conn.Close()
    'End Sub
    'Shared Sub insert(ByVal sqlstr As String)
    '    conn.Open()
    '    Dim comm As New SqlCommand(sqlstr, conn)
    '    comm.ExecuteNonQuery()
    '    conn.Close()
    'End Sub
End Class
