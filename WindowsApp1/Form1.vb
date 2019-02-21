﻿Imports System.Data.OleDb

Public Class lblMathsforIT
    Dim adapter As New OleDbDataAdapter
    Dim dataSet As DataSet = New DataSet
    Dim connection As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source= ..\..\resources\SkillsDemo#2.accdb;Persist Security Info=True")
    Dim id As Single
    Dim CAO As Double
    Dim operatingResult As Short
    Dim networkingResult As Short
    Dim computerResult As Short
    Dim virtualisationResult As Short
    Dim programmingResult As Short
    Dim mathResult As Short
    Dim databaseResult As Short
    Dim communicationResult As Short
    Dim workExperienceResult As Short
    Dim totalResult As Double
    Dim rowNumber As Short

    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        dataSet.Clear()

        adapter.SelectCommand = New OleDbCommand("Select PPSN, FirstName + ' ' + Surname As [Name], OpeartingSystems, NetworkingEssentials, ComputerSystemsHardware, VirtualisationSupport, 
                                                ProgrammingandDesignPrinciples, MathsforIT, DatabaseMethods, Communications, WorkExperience
                                                from [5M0536 Module Results]", connection)
        adapter.Fill(dataSet, "[5M0536 Module Results]")
        DataGridView1.DataSource = dataSet.Tables("[5M0536 Module Results]")
        DataGridView1.Columns.Add("Total Cao", "Total Cao")


        rowNumber = DataGridView1.RowCount
        MsgBox(rowNumber)

        For i = 0 To rowNumber - 2
            operatingResult = dataSet.Tables("[5M0536 Module Results]").Rows(i).Item(2)
            networkingResult = dataSet.Tables("[5M0536 Module Results]").Rows(i).Item(3)
            computerResult = dataSet.Tables("[5M0536 Module Results]").Rows(i).Item(4)
            virtualisationResult = dataSet.Tables("[5M0536 Module Results]").Rows(i).Item(5)
            programmingResult = dataSet.Tables("[5M0536 Module Results]").Rows(i).Item(6)
            mathResult = dataSet.Tables("[5M0536 Module Results]").Rows(i).Item(7)
            databaseResult = dataSet.Tables("[5M0536 Module Results]").Rows(i).Item(8)
            communicationResult = dataSet.Tables("[5M0536 Module Results]").Rows(i).Item(9)
            workExperienceResult = dataSet.Tables("[5M0536 Module Results]").Rows(i).Item(10)
            totalResult = calculateCAO(operatingResult) + calculateCAO(networkingResult) + calculateCAO(computerResult) + calculateCAO(virtualisationResult) + calculateCAO(programmingResult) +
            calculateCAO(mathResult) + calculateCAO(databaseResult) + calculateCAO(communicationResult) + calculateCAO(workExperienceResult)
            DataGridView1.Rows(i).Cells(11).Value = totalResult
        Next

    End Sub

    Private Sub btnOneStudent_Click(sender As Object, e As EventArgs) Handles btnOneStudent.Click
        Dim text As String = txtOneStudent.Text
        dataSet.Clear()
        adapter.SelectCommand = New OleDbCommand("Select * from [5M0536 Module Results] where PPSN = '" & text & "'", connection)
        adapter.Fill(dataSet, "Search Result")
        DataGridView1.DataSource = dataSet.Tables("Search Result")

    End Sub

    Private Sub btnSearchLike_Click(sender As Object, e As EventArgs) Handles btnSearchLike.Click
        dataSet.Clear()
        Dim text As String = txtSerchNSP.Text
        Dim quert As String = "Select * from [5M0536 Module Results] where FirstName like '%" & text & "%' or Surname like '%" & text & "%' or PPSN like '%" & text & "%'"
        adapter.SelectCommand = New OleDbCommand(quert, connection)
        adapter.Fill(dataSet, "Search Result")
        DataGridView1.DataSource = dataSet.Tables("Search Result")
        DataGridView1.Columns.Add("Total Cao", "Total Cao")
        rowNumber = DataGridView1.RowCount

        For i = 0 To rowNumber - 2
            operatingResult = dataSet.Tables("Search Result").Rows(i).Item(3)
            networkingResult = dataSet.Tables("Search Result").Rows(i).Item(4)
            computerResult = dataSet.Tables("Search Result").Rows(i).Item(5)
            virtualisationResult = dataSet.Tables("Search Result").Rows(i).Item(6)
            programmingResult = dataSet.Tables("Search Result").Rows(i).Item(7)
            mathResult = dataSet.Tables("Search Result").Rows(i).Item(8)
            databaseResult = dataSet.Tables("Search Result").Rows(i).Item(9)
            communicationResult = dataSet.Tables("Search Result").Rows(i).Item(10)
            workExperienceResult = dataSet.Tables("Search Result").Rows(i).Item(11)
            totalResult = calculateCAO(operatingResult) + calculateCAO(networkingResult) + calculateCAO(computerResult) + calculateCAO(virtualisationResult) + calculateCAO(programmingResult) +
            calculateCAO(mathResult) + calculateCAO(databaseResult) + calculateCAO(communicationResult) + calculateCAO(workExperienceResult)
            DataGridView1.Rows(i).Cells(12).Value = totalResult
        Next
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        connection.Open()
        adapter.InsertCommand = New OleDbCommand("INSERT INTO [5M0536 Module Results] 
                                                (PPSN, FirstName, Surname, OpeartingSystems, NetworkingEssentials, ComputerSystemsHardware, VirtualisationSupport, ProgrammingandDesignPrinciples, 
                                                MathsforIT, DatabaseMethods, Communications, WorkExperience) 
                                                Values ('" & txtPPSN.Text & "', '" & txtFirstName.Text & "', '" & txtLastName.Text & "', '" & txtSystem.Text & "', '" & txtNetworking.Text & "',
                                                '" & txtComputerSystem.Text & "', '" & txtVirtualisation.Text & "', '" & txtProgramming.Text & "', '" & txtMath.Text & "', '" & txtDatabase.Text & "',
                                                '" & txtCommunication.Text & "', '" & txtWorkExperience.Text & "')", connection)
        adapter.InsertCommand.ExecuteNonQuery()
        connection.Close()
    End Sub

    Public Function calculateCAO(ByVal value As Short) As Double
        Dim CAO As Double
        If value < 50 Then
            CAO = 0
        ElseIf value >= 50 And value < 65 Then
            CAO = 16.25
        ElseIf value >= 65 And value < 80 Then
            CAO = 32.5
        Else
            CAO = 48.75
        End If
        Return CAO
    End Function

End Class
