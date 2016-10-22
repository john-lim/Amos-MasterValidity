
Imports System.Xml
Imports AmosGraphics
Imports AmosEngineLib
Imports AmosDebug
Imports MiscAmosTypes
Imports PBayes


Namespace Gaskination
    Public Class MVClass
        Implements AmosGraphics.IPlugin
        'This plugin was written by John Lim September 2016 for James Gaskin

        Public Function Name() As String Implements AmosGraphics.IPlugin.Name
            Return "Master Validity"
        End Function

        Public Function Description() As String Implements AmosGraphics.IPlugin.Description
            Return "Calculates CR, AVE, ASV, MSV and outputs them in a table. See statwiki.kolobkreations.com for more information."
        End Function

        Public Function Mainsub() As Integer Implements AmosGraphics.IPlugin.MainSub
            'Remove the old table files
            If (System.IO.File.Exists("MasterValidity.html")) Then
                System.IO.File.Delete("MasterValidity.html")
            End If

            'Loop to count the number of latent variables in the model. Used to size the array and later print the array.
            Dim iCount As Integer
            Dim iCount2 As Integer
            Dim iCount3 As Integer
            For Each e As PDElement In pd.PDElements
                If e.IsLatentVariable Then
                    iCount += 1 'Will return the number of variables in the model.
                End If
                If e.IsCovariance Then 'Will return the number of covariances in the model
                    iCount2 += 1
                End If
            Next
            If iCount > 1 Then 'Doesn't run this check if there is only one latent variable.
                iCount3 = iCount - 1
                Dim iFactorial As Integer = (iCount - 1)
                Do 'Loop to count the number of factorials there should be.
                    iCount3 -= 1
                    iFactorial += (iCount3)
                Loop Until iCount3 = 0
                If iCount2 < iFactorial Then
                    MsgBox("You are missing at least one correlation arrow between latent factors. In order for the Plugin to work, all latent factors must be correlated.")
                    Exit Function
                End If
            End If

            'Fits the specified model.
            AmosGraphics.pd.AnalyzeCalculateEstimates()

            'Get  from xpath expression for the output tables.
            Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
            doc.Load(AmosGraphics.pd.ProjectName & ".AmosOutput")
            Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
            Dim eRoot As Xml.XmlElement = doc.DocumentElement
            Dim CorrelationsTable As XmlElement 'Table of correlations.
            Dim SRWTable As XmlElement 'Table of Standardized Regression Weights
            Dim Variances As XmlElement 'Table of Standardized Regression Weights
            'Selects the entire table node and stores it in a table object.
            SRWTable = eRoot.SelectSingleNode("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody", nsmgr)
            CorrelationsTable = eRoot.SelectSingleNode("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Correlations:']/table/tbody", nsmgr)
            Variances = eRoot.SelectSingleNode("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Variances:']/table/tbody", nsmgr)

            'Exception handling for if the user tries to get CFA with a variable that only has one indicator.
            For j = 1 To iCount + 1
                Try
                    MatrixName(Variances, j, 3) 'Checks if there is no variance for an unobserved variable.
                Catch ex As Exception
                    For k = 1 To iCount + 1
                        Try
                            If Not MatrixName(Variances, k, 3) = Nothing Then 'Checks if the variable is "unidentified"
                                MsgBox(MatrixName(Variances, k, 0) + "  is missing an indicator, and is therefore, not latent. The CFA is only for latent variables, so don’t include " + MatrixName(Variances, k, 0) + " in the CFA.")
                                Exit Function
                            End If
                        Catch exc As NullReferenceException
                            Continue For
                        End Try
                    Next
                End Try
            Next

            'Variables for the count of the number of nodes in the output table.
            Dim NumberOfRows As Integer 'Number of correlations
            Dim SRWNumberOfRows As Integer = SRWTable.ChildNodes.Count 'Number of regressions
            If IsNothing(CorrelationsTable) Then
                NumberOfRows = 0
            Else
                NumberOfRows = CorrelationsTable.ChildNodes.Count
            End If

            'Arrays to hold values for each latent variable.
            Dim nameArray(iCount) As String
            Dim MSVarray(iCount) As Double
            Dim CRarray(iCount) As Double
            Dim AVEarray(iCount) As Double
            Dim MRarray(iCount) As Double
            Dim a As Integer = 0 'Variable to iterate through locations of the arrays.

            'Compute MSV. Loops through all variables in the model.
            For Each e As PDElement In pd.PDElements
                'Finds the latent variables (ovals with more than one arrow).
                If e.IsLatentVariable Then
                    Dim testVal As Double
                    Dim MSV As Double
                    Dim i As Integer
                    For i = 1 To NumberOfRows
                        'If the variable name matches a node from the table, square the correlation.
                        If e.NameOrCaption = MatrixName(CorrelationsTable, i, 0) Or e.NameOrCaption = MatrixName(CorrelationsTable, i, 2) Then
                            testVal = Math.Pow(MatrixElement(CorrelationsTable, i, 3), 2)
                            'Compare all the correlations to find the largest one.
                            If testVal > MSV Then
                                MSV = testVal
                            End If
                        End If
                    Next
                    MSVarray(a) = MSV 'Store the largest squared correlation to an array.
                    nameArray(a) = e.NameOrCaption
                    a += 1
                    MSV = 0
                End If
            Next

            'Variables used for computing CR, AVE, and MaxR
            Dim CR As Double
            Dim AVE As Double
            Dim MR As Double
            Dim SSI As Double 'Sum of the square SRW values for a varialbe.
            Dim SL2 As Double 'The squared value of each SRW value.
            Dim SL3 As Double 'The sum of the SRW values for a variable.
            Dim sumErr As Double
            Dim MaxR As Double
            Dim numItems As Integer
            a = 0

            'Compute CR & AVE. 
            For Each e As PDElement In pd.PDElements 'Loops through all variables in the model.
                If e.IsLatentVariable Then 'Finds the latent variables (ovals with more than one arrow).
                    Dim testVal As Double 'Holds the SRW value to use and compare.
                    For i = 1 To SRWNumberOfRows
                        'If the variable name matches a node from the table, perform these calculations.
                        If e.NameOrCaption = MatrixName(SRWTable, i, 2) Then
                            testVal = MatrixElement(SRWTable, i, 3)
                            SL2 = Math.Pow(testVal, 2)
                            SL3 = SL3 + testVal
                            SSI = SSI + SL2
                            sumErr = sumErr + (1 - Math.Pow(testVal, 2))
                            MaxR = MaxR + (Math.Pow(testVal, 2) / (1 - Math.Pow(testVal, 2)))
                            numItems = numItems + 1
                        End If
                    Next

                    SL3 = SL3 * SL3
                    CR = SL3 / (SL3 + sumErr)
                    AVE = SSI / numItems
                    MR = 1 / (1 + (1 / MaxR))

                    'Load the calculations into the arrays.
                    CRarray(a) = CR
                    AVEarray(a) = AVE
                    MRarray(a) = MR
                    'Increase increment for loading arrays, reset values to 0 to compute next variable.
                    a += 1
                    CR = 0
                    AVE = 0
                    MR = 0
                    SSI = 0
                    SL2 = 0
                    SL3 = 0
                    MaxR = 0
                    sumErr = 0
                    numItems = 0
                End If
            Next

            'Start the Amos debugger to print the table
            Dim debug As New AmosDebug.AmosDebug

            'Set up the listener To output the debugs
            Dim resultWriter As New TextWriterTraceListener("MasterValidity.html")
            Trace.Listeners.Add(resultWriter)

            'Write the beginning Of the document and the teable header
            debug.PrintX("<html><body><h1>Model Validity Measures</h1><hr/>")
            debug.PrintX("<table><tr><th></th><th>CR</th><th>AVE</th>")
            If NumberOfRows > 0 Then
                debug.PrintX("<th>MSV</th>")
            End If
            debug.PrintX("<th>MaxR(H)</th>")
            'Loop to print all the latent variable names in the header.
            If NumberOfRows > 0 Then
                For b = 0 To iCount - 1
                    debug.PrintX("<th>")
                    debug.PrintX(nameArray(b))
                    debug.PrintX("</th>")
                Next
            End If

            debug.PrintX("</tr>")

            'Variables used to dynamically create a table according to the number of variables.
            Dim t As Integer = 1
            Dim u As Integer = iCount - 3
            Dim v As Integer = iCount - 3
            Dim w As Integer = iCount - 2
            Dim x As Integer = 1
            Dim y As Integer = iCount
            Dim z As Integer = 1
            Dim pass As Boolean = False
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

            'Make the table of validity measures while inspecting each element.
            For c = 0 To iCount - 1 'Unload the arrays holding each calculated measure.
                debug.PrintX("<tr><td><span style='font-weight:bold;'>")
                debug.PrintX(nameArray(c))
                debug.PrintX("</span></td>")
                If CRarray(c) < 0.7 Then 'Check if CR is low.
                    bMalhotra = False
                    debug.PrintX("<td style='color:red'>")

                    'Make a recommendation if there are enough indicators.
                    Dim iCount4 As Integer = 0
                    Dim dCheck As Double = 0
                    Dim sCheck As String = vbNull
                    For e = 1 To SRWNumberOfRows 'Check if there are enough indicators
                        If nameArray(c) = MatrixName(SRWTable, e, 2) Then
                            iCount4 += 1
                        End If
                    Next
                    If iCount4 > 2 Then 'Find the lowest indicator
                        For f = 1 To SRWNumberOfRows
                            If nameArray(c) = MatrixName(SRWTable, f, 2) Then
                                If MatrixElement(SRWTable, f, 3) > dCheck Then
                                    sCheck = MatrixName(SRWTable, f, 0)
                                    dCheck = MatrixElement(SRWTable, f, 3)
                                End If
                            End If
                        Next
                        msgCR.Enqueue("Reliability: the CR for " & nameArray(c) & " is less than 0.70. Try removing " & sCheck & " to improve CR.")
                    Else
                        msgCR.Enqueue("Reliability: the CR for " & nameArray(c) & " is less than 0.70. No way to improve CR because you only have two indicators for that variable. Removing one indicator will make this not latent.")
                    End If
                ElseIf CRarray(c) < AVEarray(c) Then 'Check if CR is less than AVE.
                    debug.PrintX("<td style='color:red'>")
                    msgCRAVE.Enqueue("Convergent Validity: the CR for " & nameArray(c) & " is less than the AVE.")
                Else
                    debug.PrintX("<td>")
                End If
                debug.PrintX(CRarray(c).ToString("#0.000"))
                debug.PrintX("</td>")
                If AVEarray(c) < 0.5 Then 'Check if AVE is low.
                    debug.PrintX("<td style='color:red'>")
                    If bMalhotra = True Then 'If there is no problem with CR, the message won't reference Malhotra. Conversely, a reference to Malhotra.
                        sMessage = "*"
                        iMalhotra += 1 'Variable to show Malhotra reference later.
                    Else sMessage = ""
                    End If
                    Dim iCount4 As Integer = 0
                    Dim dCheck As Double = 0
                    Dim sCheck As String = vbNull
                    For e = 1 To SRWNumberOfRows 'Check if there are enough indicators
                        If nameArray(c) = MatrixName(SRWTable, e, 2) Then
                            iCount4 += 1
                        End If
                    Next
                    If iCount4 > 2 Then 'Find the lowest indicator
                        For f = 1 To SRWNumberOfRows
                            If nameArray(c) = MatrixName(SRWTable, f, 2) Then
                                If MatrixElement(SRWTable, f, 3) > dCheck Then
                                    sCheck = MatrixName(SRWTable, f, 0)
                                    dCheck = MatrixElement(SRWTable, f, 3)
                                End If
                            End If
                        Next
                        msgAVE.Enqueue(sMessage & "Convergent Validity: the AVE for " & nameArray(c) & " is less than 0.50. Try removing " & sCheck & " to improve AVE.")
                    Else
                        msgAVE.Enqueue(sMessage & "Convergent Validity: the AVE for " & nameArray(c) & " is less than 0.50. No way to improve AVE because you only have two indicators for that variable. Removing one indicator will make this not latent.")
                    End If
                ElseIf AVEarray(c) < MSVarray(c) Then 'Check if AVE is less than MSV.
                    debug.PrintX("<td style='color:red'>")
                    msgAVEMSV.Enqueue("Discriminant Validity: the AVE for " & nameArray(c) & " is less than the MSV.")
                Else
                    debug.PrintX("<td>")
                End If
                debug.PrintX(AVEarray(c).ToString("#0.000"))
                debug.PrintX("</td>")
                If NumberOfRows > 0 Then 'Does not print the MSV if there are no correlations.
                    debug.PrintX("<td>")
                    debug.PrintX(MSVarray(c).ToString("#0.000"))
                    debug.PrintX("</td>")
                End If
                debug.PrintX("<td>")
                debug.PrintX(MRarray(c).ToString("#0.000"))
                debug.PrintX("</td>")

                'Inspect for discriminant validity.
                For f = 1 To iCount - y
                    For g = 1 To NumberOfRows 'Iterate through the correlations table.
                        If nameArray(c) = MatrixName(CorrelationsTable, g, 0) Then
                            If (Math.Sqrt(AVEarray(c)).ToString("#0.000")) < (MatrixElement(CorrelationsTable, g, 3)) Then 'Check if the square root of AVE is less than the correlation.
                                msgDiscriminant.Enqueue("Discriminant Validity: the square root of the AVE for " & nameArray(c) & " is less than its correlation with " & MatrixName(CorrelationsTable, g, 2) & ".") 'Message if any cases are found.
                                red = True
                            End If
                        ElseIf nameArray(c) = MatrixName(CorrelationsTable, g, 2) Then 'If the correlation matches the variable we are checking.
                            If (Math.Sqrt(AVEarray(c)).ToString("#0.000")) < (MatrixElement(CorrelationsTable, g, 3)) Then 'Check if the square root of AVE is less than the correlation.
                                msgDiscriminant.Enqueue("Discriminant Validity: the square root of the AVE for " & nameArray(c) & " is less than its correlation with " & MatrixName(CorrelationsTable, g, 0) & ".") 'Message if any cases are found.
                                red = True
                            End If
                        End If
                    Next

                    'Print out the correlations for that variable.
                    debug.PrintX("<td>")
                    debug.PrintX(MatrixElement(CorrelationsTable, z, 3).ToString("#0.000"))
                    debug.PrintX("</td>")

                    'A set of conditionals that formats the correlation matrix correctly.
                    If z > 1 And t < iCount - y Then
                        z = z + w
                        w -= 1
                        pass = True
                    End If
                    If z = 1 Then
                        z += 1
                    End If
                    t += 1
                Next
                If pass = True Then
                    z = z - u
                    u += v
                    v -= 1
                End If

                If NumberOfRows > 0 Then
                    If red = True Then 'If there is discriminant validity, mark it as red.
                        debug.PrintX("<td><span style='font-weight:bold; color:red;'>")
                        debug.PrintX(Math.Sqrt(AVEarray(c)).ToString("#0.000"))
                        debug.PrintX("</span></td>")
                    Else
                        debug.PrintX("<td><span style='font-weight:bold;'>")
                        debug.PrintX(Math.Sqrt(AVEarray(c)).ToString("#0.000"))
                        debug.PrintX("</span></td>")
                    End If
                End If

                For e = 1 To iCount - x 'Empty cells in the matrix.
                    debug.PrintX("<td></td>")
                Next

                w = iCount - 2
                t = 1
                x += 1 'Decrement the number of empty cells.
                y -= 1 'Increment the number of correlations printed.
                red = False
                bMalhotra = True
                debug.PrintX("</tr>") 'Close row to move on to the next.
            Next

            debug.PrintX("</table><hr/>") 'Close table
            debug.PrintX("<h3>Validity Concerns</h3>") 'Next section displaying validity concerns.
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

            If NumberOfRows = 0 Then 'Prints message if there was no correlation matrix
                debug.PrintX("You only had one latent variable so there is no correlation matrix or MSV.<br>")
            End If
            If max = 0 Then 'Prints if there are no validity concerns.
                debug.PrintX("No validity concerns here.<br>")
            End If


            'Print references. Malhotra only prints if the described condition exists.
            debug.PrintX("<h3>References</h3>Thresholds From:<br>Hu, L., Bentler, P.M. (1999), ""Cutoff Criteria for Fit Indexes in Covariance Structure Analysis: Conventional Criteria Versus New Alternatives"" SEM vol. 6(1), pp. 1-55.")
            If iMalhotra > 0 Then
                debug.PrintX("<br>*Malhotra N. K., Dash S. argue that AVE is often too strict, and reliability can be established through CR alone.<br>Malhotra N. K., Dash S. (2011). Marketing Research an Applied Orientation. London: Pearson Publishing.")
            End If
            debug.PrintX("<p>**If you would like to cite this tool directly, please use the following:")
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
            Return 0
        End Function
        Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double
            Dim e As XmlElement
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column)
            MatrixElement = CDbl(e.GetAttribute("x"))
        End Function
        Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String
            Dim e As XmlElement
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column)
            MatrixName = e.InnerText
        End Function
    End Class
End Namespace