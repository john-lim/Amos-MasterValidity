#Region "Header"
Imports System.Xml
Imports AmosGraphics
Imports AmosEngineLib
Imports AmosDebug
Imports MiscAmosTypes
Imports PBayes
#End Region

Public Class ValidityMaster
    Implements AmosGraphics.IPlugin
    'This plugin was written by John Lim Fall 2016 for James Gaskin

#Region "Name Plugin"
    Public Function Name() As String Implements AmosGraphics.IPlugin.Name
        Return "Master Validity"
    End Function

    Public Function Description() As String Implements AmosGraphics.IPlugin.Description
        Return "Calculates CR, AVE, ASV, MSV and outputs them in a table. See statwiki.kolobkreations.com for more information."
    End Function
#End Region

    Public Function Mainsub() As Integer Implements AmosGraphics.IPlugin.MainSub
#Region "Table Input"
        'Fits the specified model.
        AmosGraphics.pd.AnalyzeCalculateEstimates()

        'Get  from xpath expression for the output tables.
        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(AmosGraphics.pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim eRoot As Xml.XmlElement = doc.DocumentElement
        Dim tableCorrelation As XmlElement 'Table of correlations.
        Dim tableRegression As XmlElement 'Table of Standardized Regression Weights
        Dim tableVariance As XmlElement 'Table of Variances
        Dim tableCovariance As XmlElement 'Table of covariances
        'Selects the entire table node and stores it in a table object.
        tableRegression = eRoot.SelectSingleNode("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody", nsmgr)
        tableCorrelation = eRoot.SelectSingleNode("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody", nsmgr)
        tableVariance = eRoot.SelectSingleNode("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Variances:']/table/tbody", nsmgr)
        tableCovariance = eRoot.SelectSingleNode("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Covariances:']/table/tbody", nsmgr)
#End Region

#Region "Counters"
        'Loop to count the number of latent variables in the model. Used to size the array and later print the array.
        Dim iLatent As Integer
        For Each e As PDElement In pd.PDElements
            If e.IsLatentVariable And Not e.IsEndogenousVariable Then 'Checks if the variable is latent
                iLatent += 1 'Will return the number of variables in the model.
            End If
        Next
        'Variables for the count of the number of nodes in the output table.
        Dim iCorrelations As Integer 'Number of correlations
        Dim iRegression As Integer = tableRegression.ChildNodes.Count 'Number of regressions
        If IsNothing(tableCorrelation) Then
            iCorrelations = 0
        Else
            iCorrelations = tableCorrelation.ChildNodes.Count
        End If
#End Region

#Region "One Indicator Exception"
        'Exception handling for if the user tries to get CFA with a variable that only has one indicator.
        For j = 1 To iLatent + 1
            Try
                MatrixName(tableVariance, j, 3) 'Checks if there is no variance for an unobserved variable.
            Catch ex As Exception
                For k = 1 To iLatent + 1
                    Try
                        If Not MatrixName(tableVariance, k, 3) = Nothing Then 'Checks if the variable is "unidentified"
                            MsgBox(MatrixName(tableVariance, k, 0) + " is causing an error. Either:
1. It only has one indicator and is therefore, not latent. The CFA is only for latent variables, so don’t include " + MatrixName(tableVariance, k, 0) + " in the CFA.
2. You are missing a constraint on an indicator.")
                            Exit Function
                        End If
                    Catch exc As NullReferenceException
                        Continue For
                    End Try
                Next
            End Try
        Next
#End Region

#Region "Arrays"
        'Set up the arrays that will hold all the names and values in the matrix]
        Dim rowIndex As Integer = 0
        Dim columnIndex As Integer = 0
        Dim aColumn(iLatent + 3) As String
        Dim aRow(iLatent - 1) As String
        Dim valueArray((aRow.Length - 1), (aColumn.Length - 1)) As Double 'Two dimensional array to hold all the values in the table.
        aColumn(columnIndex) = "CR"
        columnIndex += 1
        aColumn(columnIndex) = "AVE"
        columnIndex += 1
        If iCorrelations > 1 Then
            aColumn(columnIndex) = "MSV"
            columnIndex += 1
        End If
        aColumn(columnIndex) = "MaxR(H)"
        columnIndex += 1

        For Each e As PDElement In pd.PDElements
            If e.IsLatentVariable And Not e.IsEndogenousVariable Then
                aColumn(columnIndex) = e.NameOrCaption
                aRow(rowIndex) = e.NameOrCaption
                columnIndex += 1
                rowIndex += 1
            End If
        Next
#End Region


#Region "Calculate CR-AVE-MSV-MR"
        Dim sColumn As String
        Dim dValue As Double
        For Each sRow As String In aRow
            Dim CR As Double
            Dim AVE As Double
            Dim MR As Double
            Dim MSV As Double = 0
            Dim SSI As Double 'Sum of the square SRW values for a varialbe.
            Dim SL2 As Double 'The squared value of each SRW value.
            Dim SL3 As Double 'The sum of the SRW values for a variable.
            Dim SQRT As Double 'Square root of the AVE
            Dim sumErr As Double
            Dim MaxR As Double
            Dim numItems As Integer
            Dim testVal As Double = 0
            Dim testVal2 As Double 'Holds the SRW value to use and compare.

            rowIndex = Array.IndexOf(aRow, sRow)

            For v = 1 To iCorrelations
                'If the variable name matches a node from the table, square the correlation.
                If sRow = MatrixName(tableCorrelation, v, 0) Or sRow = MatrixName(tableCorrelation, v, 2) Then
                    testVal = Math.Pow(MatrixElement(tableCorrelation, v, 3), 2)
                    'Compare all the correlations to find the largest one.
                    If testVal > MSV Then
                        MSV = testVal
                    End If
                End If
            Next
            For i = 1 To iRegression
                'If the variable name matches a node from the table, perform these calculations.
                If sRow = MatrixName(tableRegression, i, 2) Then
                    testVal2 = MatrixElement(tableRegression, i, 3)
                    SL2 = Math.Pow(testVal2, 2)
                    SL3 = SL3 + testVal2
                    SSI = SSI + SL2
                    sumErr = sumErr + (1 - Math.Pow(testVal2, 2))
                    MaxR = MaxR + (Math.Pow(testVal2, 2) / (1 - Math.Pow(testVal2, 2)))
                    numItems = numItems + 1
                End If
            Next
            SL3 = SL3 * SL3
            CR = SL3 / (SL3 + sumErr)
            AVE = SSI / numItems
            MR = 1 / (1 + (1 / MaxR))
            SQRT = Math.Sqrt(AVE)

            columnIndex = 0
            valueArray(rowIndex, columnIndex) = CR
            columnIndex += 1
            valueArray(rowIndex, columnIndex) = AVE
            columnIndex += 1
            If iCorrelations > 1 Then
                valueArray(rowIndex, columnIndex) = MSV
                columnIndex += 1
            End If
            valueArray(rowIndex, columnIndex) = MR
            columnIndex += 1

            For a = 1 To iCorrelations
                If sRow = MatrixName(tableCorrelation, a, 2) Then
                    columnIndex = -1
                    rowIndex = -1
                    sColumn = MatrixName(tableCorrelation, a, 0)
                    dValue = MatrixElement(tableCorrelation, a, 3)
                    For index = 0 To aColumn.Length - 1
                        If aColumn(index) = sColumn Then
                            columnIndex = index
                        End If
                    Next
                    For index = 0 To aRow.Length - 1
                        If aRow(index) = sRow Then
                            rowIndex = index
                        End If
                    Next
                    If columnIndex <> -1 And rowIndex <> -1 Then
                        valueArray(rowIndex, columnIndex) = dValue
                    End If
                End If
            Next

            columnIndex = -1
            rowIndex = -1
            For index = 0 To aColumn.Length - 1
                If sRow = aColumn(index) Then
                    columnIndex = index
                    rowIndex = index - 4
                End If

            Next
            If columnIndex <> -1 And rowIndex <> -1 Then
                valueArray(rowIndex, columnIndex) = SQRT
            End If

            MSV = 0
                CR = 0
                AVE = 0
                MR = 0
                SSI = 0
                SL2 = 0
                SL3 = 0
                MaxR = 0
                sumErr = 0
                numItems = 0
            Next

#End Region

#Region "Create html"
            'Remove the old table files
            If (System.IO.File.Exists("MasterValidity.html")) Then
            System.IO.File.Delete("MasterValidity.html")
        End If

        'Start the Amos debugger to print the table
        Dim debug As New AmosDebug.AmosDebug

        'Set up the listener To output the debugs
        Dim resultWriter As New TextWriterTraceListener("MasterValidity.html")
        Trace.Listeners.Add(resultWriter)

        'Queues used to hold messages for validity concerns.
        Dim msgDiscriminant As New Queue()
        Dim msgCR As New Queue()
        Dim msgCRAVE As New Queue()
        Dim msgAVE As New Queue()
        Dim msgAVEMSV As New Queue()
        Dim red As Boolean = False 'Boolean for displaying if correlations are higher than the squareroot of AVE
        Dim bMalhotra As Boolean = True
        Dim iMalhotra As Integer = 0
        Dim sMessage As String = ""
        Dim significance As String = ""

        'Write the beginning Of the document and the table header
        debug.PrintX("<html><body><h1>Model Validity Measures</h1><hr/>")
#End Region

#Region "Table matrix"
        debug.PrintX("<table><tr><td></td>")
        For Each sColumn In aColumn
            debug.PrintX("<th>" + sColumn + "</th>")
        Next
        debug.PrintX("</tr>")
        For y = 0 To aRow.Length - 1 'Loops through all the rows in the matrix
            debug.PrintX("<tr><td><span style='font-weight:bold;'>" + aRow(y) + "</span></td>")
            For x = 0 To aColumn.Length - 1 'Loops through all the columns in the matrix
                If x = 0 Then
                    If valueArray(y, x) < 0.7 Then 'Check if CR is low.
                        bMalhotra = False
                        debug.PrintX("<td style='color:red'>")
                        'Make a recommendation if there are enough indicators.
                        Dim iCount As Integer = 0
                        Dim dCheck As Integer = 0
                        Dim sCheck As String = vbNull
                        For i = 1 To iRegression
                            If aRow(y) = MatrixName(tableRegression, i, 2) Then
                                iCount += 1
                            End If
                        Next
                        If iCount > 2 Then 'Find the lowest indicator
                            For i = 1 To iRegression
                                If aRow(y) = MatrixName(tableRegression, i, 2) Then
                                    If MatrixElement(tableRegression, i, 3) > dCheck Then
                                        sCheck = MatrixName(tableRegression, i, 0)
                                        dCheck = MatrixElement(tableRegression, i, 3)
                                    End If
                                End If
                            Next
                            msgCR.Enqueue("Reliability: the CR for " & aRow(y) & " is less than 0.70. Try removing " & sCheck & " to improve CR.")
                        Else
                            msgCR.Enqueue("Reliability: the CR for " & aRow(y) & " is less than 0.70. No way to improve CR because you only have two indicators for that variable. Removing one indicator will make this not latent.")
                        End If
                    ElseIf valueArray(y, x) < valueArray(y, (x + 1)) Then 'Check if CR is less than AVE.
                        debug.PrintX("<td style='color:red'>")
                        msgCRAVE.Enqueue("Convergent Validity: the CR for " & aRow(y) & " is less than the AVE.")
                    Else
                        debug.PrintX("<td>")
                    End If
                ElseIf x = 1 Then
                    If valueArray(y, x) < 0.5 Then 'Check if AVE is low.
                        debug.PrintX("<td style='color:red'>")
                        If bMalhotra = True Then 'If there is no problem with CR, the message won't reference Malhotra. Conversely, a reference to Malhotra.
                            sMessage = "<sup>1 </sup>"
                            iMalhotra += 1 'Variable to show Malhotra reference later.
                        Else sMessage = ""
                        End If
                        Dim iCount As Integer = 0
                        Dim dCheck As Double = 0
                        Dim sCheck As String = vbNull
                        For e = 1 To iRegression 'Check if there are enough indicators
                            If aRow(y) = MatrixName(tableRegression, e, 2) Then
                                iCount += 1
                            End If
                        Next
                        If iCount > 2 Then 'Find the lowest indicator
                            For i = 1 To iRegression
                                If aRow(y) = MatrixName(tableRegression, i, 2) Then
                                    If MatrixElement(tableRegression, i, 3) > dCheck Then
                                        sCheck = MatrixName(tableRegression, i, 0)
                                        dCheck = MatrixElement(tableRegression, i, 3)
                                    End If
                                End If
                            Next
                            msgAVE.Enqueue(sMessage & "Convergent Validity: the AVE for " & aRow(y) & " is less than 0.50. Try removing " & sCheck & " to improve AVE.")
                        Else
                            msgAVE.Enqueue(sMessage & "Convergent Validity: the AVE for " & aRow(y) & " is less than 0.50. No way to improve AVE because you only have two indicators for that variable. Removing one indicator will make this not latent.")
                        End If
                    ElseIf valueArray(y, x) < valueArray(y, x + 1) And iCorrelations > 1 Then 'Check if AVE is less than MSV.
                        debug.PrintX("<td style='color:red'>")
                        msgAVEMSV.Enqueue("Discriminant Validity: the AVE for " & aRow(y) & " is less than the MSV.")
                    Else
                        debug.PrintX("<td>")
                    End If
                    For i = 1 To iCorrelations
                        If aRow(y) = MatrixName(tableCorrelation, i, 0) Then
                            If (Math.Sqrt(valueArray(y, x)) < (MatrixElement(tableCorrelation, i, 3))) Then 'Check if the square root of AVE is less than the correlation.
                                msgDiscriminant.Enqueue("Discriminant Validity: the square root of the AVE for " & aRow(y) & " is less than its correlation with " & MatrixName(tableCorrelation, i, 2) & ".") 'Message if any cases are found.
                                red = True
                            End If
                        ElseIf aRow(y) = MatrixName(tableCorrelation, i, 2) Then 'If the correlation matches the variable we are checking.
                            If (Math.Sqrt(valueArray(y, x)).ToString("#0.000")) < (MatrixElement(tableCorrelation, i, 3)) Then 'Check if the square root of AVE is less than the correlation.
                                msgDiscriminant.Enqueue("Discriminant Validity: the square root of the AVE for " & aRow(y) & " is less than its correlation with " & MatrixName(tableCorrelation, i, 0) & ".") 'Message if any cases are found.
                                red = True
                            End If
                        End If
                    Next

                ElseIf aColumn(x) = aRow(y) Then
                    If red = True Then
                        debug.PrintX("<td><span style='font-weight:bold; color:red;'>")
                    Else
                        debug.PrintX("<td><span style='font-weight:bold;'>")
                    End If
                Else
                    debug.PrintX("<td>")
                End If

                If x > 3 And aColumn(x) <> aRow(y) Then
                    For a = 1 To iCorrelations
                        If MatrixName(tableCovariance, a, 0) = aColumn(x) And MatrixName(tableCovariance, a, 2) = aRow(y) Then
                            If MatrixName(tableCovariance, a, 6) = "***" Then
                                significance = "***"
                            ElseIf MatrixElement(tableCovariance, a, 6) < 0.01 Then
                                significance = "**"
                            ElseIf MatrixElement(tableCovariance, a, 6) < 0.05 Then
                                significance = "*"
                            ElseIf MatrixElement(tableCovariance, a, 6) < 0.1 Then
                                significance = "&#8224;"
                            Else
                                significance = ""
                            End If
                        End If
                    Next
                End If

                If valueArray(y, x) <> 0 Then 'Prints the value if the value is not 0
                    If aColumn(x) = aRow(y) Then
                        debug.PrintX(valueArray(y, x).ToString("#0.000") + "</span></td>")
                    Else
                        debug.PrintX(valueArray(y, x).ToString("#0.000") + significance + "</td>")
                    End If
                End If
            Next
            bMalhotra = True
            significance = ""
            debug.PrintX("</tr>")
        Next

        debug.PrintX("</table><hr/><h3>Validity Concerns</h3>") 'Next section displaying validity concerns.
        Dim max As Integer = 0 'Variable to count the number of validity concerns.
        If msgDiscriminant.Count > max Then
            max = msgDiscriminant.Count
        ElseIf msgCR.Count > max Then
            max = msgCR.Count
        ElseIf msgCRAVE.Count > max Then
            max = msgCRAVE.Count
        ElseIf msgAVE.Count > max Then
            max = msgAVE.Count
        ElseIf msgAVEMSV.Count > max Then
            max = msgAVEMSV.Count
        End If

        While msgDiscriminant.Count > 0 'Prints out the queues.
            debug.PrintX(msgDiscriminant.Dequeue() & "<br>")
            msgDiscriminant.Dequeue()
        End While
        While msgCR.Count > 0 'Prints out the queues.
            debug.PrintX(msgCR.Dequeue() & "<br>")
        End While
        While msgCRAVE.Count > 0 'Prints out the queues.
            debug.PrintX(msgCRAVE.Dequeue() & "<br>")
        End While
        While msgAVE.Count > 0 'Prints out the queues.
            debug.PrintX(msgAVE.Dequeue() & "<br>")
        End While
        While msgAVEMSV.Count > 0 'Prints out the queues.
            debug.PrintX(msgAVEMSV.Dequeue() & "<br>")
        End While

        If iCorrelations = 0 Then 'Prints message if there was no correlation matrix
            debug.PrintX("You only had one latent variable so there is no correlation matrix or MSV.<br>")
        End If
        If max = 0 Then 'Prints if there are no validity concerns.
            debug.PrintX("No validity concerns here.<br>")
        End If

        'Print references. Malhotra only prints if the described condition exists.
        debug.PrintX("<h3>References</h3>Significance of Correlations:<br>&#8224; p < 0.100<br>* p < 0.050<br>** p < 0.010<br>*** p < 0.001<br>")
        debug.PrintX("<br>Thresholds From:<br>Hu, L., Bentler, P.M. (1999), ""Cutoff Criteria for Fit Indexes in Covariance Structure Analysis: Conventional Criteria Versus New Alternatives"" SEM vol. 6(1), pp. 1-55.")
        If iMalhotra > 0 Then
            debug.PrintX("<br><p><sup>1 </sup>Malhotra N. K., Dash S. argue that AVE is often too strict, and reliability can be established through CR alone.<br>Malhotra N. K., Dash S. (2011). Marketing Research an Applied Orientation. London: Pearson Publishing.")
        End If
        debug.PrintX("<p>--If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J. & Lim, J. (2016), ""Master Validity Tool"", AMOS Plugin. <a href=\""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write Style And close
        debug.PrintX("<style>table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("MasterValidity.html")
#End Region

        Return 0
    End Function

#Region "Matrix functions"
    'This function gets a value of type double from the xml matrix
    Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double
        Dim e As XmlElement
        e = eTableBody.ChildNodes(row - 1).ChildNodes(column)
        MatrixElement = CDbl(e.GetAttribute("x"))
    End Function
    'This function gets a value of type string from the matrix
    Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String
        Dim e As XmlElement
        'The row is offset one.
        e = eTableBody.ChildNodes(row - 1).ChildNodes(column)
        MatrixName = e.InnerText
    End Function
#End Region
End Class
