''' <summary>
''' 常用方法封装汇总
''' </summary>
''' <remarks></remarks>
Public Class pubClass

#Region "SQL语句拼接"
    ''' <summary>
    '''<para>SQL语句拼接(查询修改删除均可)</para>
    '''<para>两者以";"为分隔符进行分隔，再进行拼接</para>
    '''<para>例子：</para>
    '''<para>sqlQuery=select * from ATabel where aID =';' and bID=';'</para>
    '''<para>Condition=条件变量A+";"+条件变量B</para>
    ''' </summary>
    ''' <param name="SqlQuery">SQL查询语句中固定的部分</param>
    ''' <param name="Condition">SQL查询语句中的变量</param>
    ''' <returns>String类型的SQL语句</returns>
    ''' <remarks>
    ''' </remarks>
    Shared Function SqlStatement_combination(ByVal SqlQuery As String, ByVal Condition As String)
        Dim combination_str As String
        Dim Condition_strs As String() = Condition.Split(";")
        Dim sqlQuery_strs As String() = SqlQuery.Split(";")
        '判断SQL语句所需要的变量是否与输入的一致
        If sqlQuery_strs.Length <> Condition_strs.Length + 1 Then
            combination_str = ""
        Else
            combination_str = sqlQuery_strs(0)
            For i As Integer = 0 To Condition_strs.Length - 1
                combination_str = combination_str + Condition_strs(i) + sqlQuery_strs(i + 1)
            Next
        End If
        Return combination_str
    End Function


    ''' <summary>
    ''' SQL查询语句拼接
    ''' <para>功能：可根据条件是否为空自动注释掉为空的条件，适用于字段不为null的表格查询</para>
    ''' <para>-----SqlQuery例子----------------------------------------------------</para>
    ''' <para>原SQL语句</para>
    ''' <para>select ATabel.aID , ATabel.aName ,ATable.acount from ATabel</para>
    ''' <para>where ATabel.aID like '01%' </para>
    ''' <para>and ATabel.aName like '张三%' </para>
    ''' <para>and ATabel.acount like '90%' </para>
    ''' <para>修改后的SQL语句</para>
    ''' <para>select ATabel.aID , ATabel.aName ,ATable.acount from ATabel </para>
    ''' <para>/* ATabel.aID like ';%' */</para>
    ''' <para>/* ATabel.aName like ';%' */</para>
    ''' <para>/* ATabel.acount like ';%' */</para>
    ''' <para>-----Condition例子---------------------------------------------------</para>
    ''' <para>Condition="01"+";"+"张三"+";"+"90"</para>
    ''' </summary>
    ''' <param name="SqlQuery">查询语句中固定的部分【注意：需要对SQL语句进行改写】</param>
    ''' <param name="Condition">查询语句中的变量,变量之间以";"为分隔符进行进行分隔，再进行拼接</param>
    ''' <returns>String类型的SQL语句</returns>
    ''' <remarks></remarks>
    Shared Function SqlQueryStatement_combination_condition(ByVal SqlQuery As String, ByVal Condition As String)
        Dim return_str As String
        Dim Condition_strs As String() = Condition.Split(";") '条件变量数组
        Dim sqlQuery_strs As String() = SqlQuery.Split(";") 'SQL语句主句子
        Dim isfristCondition As Boolean = True '为true用于标记是否为第一个条件
        '判断SQL语句所需要的变量是否与输入的一致
        If sqlQuery_strs.Length <> Condition_strs.Length + 1 Then
            return_str = ""
        Else
            return_str = ""
            For i As Integer = 0 To Condition_strs.Length - 1
                If Condition_strs(i).ToString <> "" Then '条件不为空时
                    If isfristCondition Then
                        sqlQuery_strs(i) = sqlQuery_strs(i).Replace("/*", " where ")
                        isfristCondition = False
                    Else
                        sqlQuery_strs(i) = sqlQuery_strs(i).Replace("/*", " and ")
                    End If
                    sqlQuery_strs(i + 1) = sqlQuery_strs(i + 1).Replace("*/", "  ")
                End If
                return_str = return_str + sqlQuery_strs(i) + Condition_strs(i)
            Next
            return_str = return_str + sqlQuery_strs(Condition_strs.Length) '为SQL查询语句加上最后面一部分，是否需要处理注释符号已经在循环中解决
        End If
        Return return_str
    End Function
#End Region

#Region "格式控制"
    ''' <summary>
    ''' 为数字添加千分号，并保留小数点后两位
    ''' </summary>
    ''' <param name="Num">需要进行格式控制的数字</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Shared Function NumFormat(ByVal Num As String)
        Dim integerPart As String = Num
        Dim decimalPart As String = "00"
        Dim flatLenght As Integer
        Dim isNegative As Boolean = True
        Dim i As Integer = 3
        If Num >= 0 Then
            isNegative = False
        End If

        If Num.Contains(".") Then '去掉小数位
            Dim NumPart As String() = Num.Split(".")
            integerPart = NumPart(0)
            decimalPart = NumPart(1)
        End If
        '======小数位位数控制==================================
        If decimalPart.Length = 1 Then
            decimalPart = decimalPart & "0"
        ElseIf decimalPart.Length > 2 Then
            If decimalPart.Substring(2) >= 5 Then '五入
                decimalPart = decimalPart.Substring(0, 2)
                decimalPart = decimalPart + 1
            Else '四舍
                decimalPart = decimalPart.Substring(0, 2)
            End If
        End If

        '=================================================
        '======排除不需要添加千分号的数字==================

        If isNegative Then '去掉负数号
            integerPart = -integerPart
        End If
        flatLenght = integerPart.Count
        If flatLenght <= 3 Then
            If isNegative Then '是负数
                Num = "-" & integerPart & "." & decimalPart
            Else
                Num = integerPart & "." & decimalPart
            End If
            Return Num
            Exit Function
        End If
        '=======排除完成================================
        '>>>>>>>>>>>>>千分号添加>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        While flatLenght > 3
            integerPart = integerPart.Insert(integerPart.Length - i, ",")
            flatLenght = flatLenght - 3
            i = i + 4
        End While
        '>>>>>>>>>>>>>>>>千分号>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        If isNegative Then '是负数
            Num = "-" & integerPart & "." & decimalPart
        Else
            Num = integerPart & "." & decimalPart
        End If

        Return Num
    End Function

    '=========================待用================================
    ''日期格式控制
    'Shared Function DateFormat(ByVal dt As String)
    '    Dim yearStr, MonthStr, dayStr As String
    '    Dim dateStr() As String = dt.Split("/")
    '    yearStr = dateStr(0)
    '    MonthStr = dateStr(1)
    '    dayStr = dateStr(2)
    '    If MonthStr.Length = 1 Then
    '        MonthStr = "0" & MonthStr
    '    End If
    '    If dayStr.Length = 1 Then
    '        dayStr = "0" & dayStr
    '    End If
    '    Dim reDt As String = yearStr & "-" & MonthStr & "-" & dayStr
    '    Return reDt
    'End Function
#End Region

#Region "DataGridView相关操作方法"
    ''' <summary>
    '''  用于增加Datagridview表格的行头数字
    ''' </summary>
    ''' <param name="dgv">需要添加行头数字的DataGridView</param>
    ''' <remarks></remarks>
    Shared Sub DGV_rowHeaderNameAdd(ByVal dgv As Windows.Forms.DataGridView)
        Dim rowindex As Int16 = 1
        While rowindex <= dgv.Rows.Count
            dgv.Rows(rowindex - 1).HeaderCell.Value = rowindex.ToString
            rowindex = rowindex + 1
        End While
    End Sub
#End Region

#Region "读取TXT文件"

#End Region

  
End Class
