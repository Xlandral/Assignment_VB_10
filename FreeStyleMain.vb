
Imports System.Diagnostics.Eventing
Imports Microsoft.Data.SqlClient
Imports MySql.Data.MySqlClient

Public Class FreeStyleMain
    'Global Variables'
    Dim Size As Integer
    Dim Product(3) As String
    Dim CCORF As Integer
    Dim CCCF As Integer
    Dim CCCVF As Integer
    Dim CCLEF As Integer
    Dim CCLIF As Integer
    Dim CCOF As Integer
    Dim CCOVF As Integer
    Dim CCRF As Integer
    Dim DateString As String
    Dim CCVF As Integer
    Dim CCCSF As Integer
    Dim CO2 As Integer
    Dim number As Integer
    Dim ID As Integer
    Dim IDLINE As Integer
    Dim ORDERDATE As String
    Dim DESC As String
    Dim FID As Integer
    Dim FAMOUNT As Integer
    Dim ORDERAMOUNT As Integer
    Dim MOSTFREQUENT As String
    Dim LEASTFREQUENT As String
    Dim RandomNumber As New Random()
    Dim _rnd As Integer
    Dim conn As New SqlConnection


    'On Form Load Make sure Reaction Colors are reset and values are laoded'
    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim connectionString As String = ("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\cscha\OneDrive\Desktop\VB_Final_Project_Submission\Coca-Cola_Freestyle\Coca-Cola\COCACOLA.mdf;Integrated Security=True;Connect Timeout=30")

        ' Open database connection
        conn = New SqlConnection(connectionString)
        conn.Open()

        updateValues()
        CocaCola.BackColor = Color.White
        CokeCherry.BackColor = Color.White
        CokeCherryVanilla.BackColor = Color.White
        CokeLime.BackColor = Color.White
        CokeLemon.BackColor = Color.White
        CokeRaspberry.BackColor = Color.White
        CokeOrange.BackColor = Color.White
        CokeOrangeVanilla.BackColor = Color.White
        CokeCremeSoda.BackColor = Color.White
        CokeVanilla.BackColor = Color.White
        Small.BackColor = Color.White
        Medium.BackColor = Color.White
        Large.BackColor = Color.White
        ExtraLarge.BackColor = Color.White
    End Sub

    'Reset Button'
    'Sets all Reaction Colors back to normal and resets things needed to be reset.'
    Private Sub btnMainMenu_Click(sender As Object, e As EventArgs) Handles btnMainMenu.Click
        CocaCola.BackColor = Color.White
        CokeCherry.BackColor = Color.White
        CokeCherryVanilla.BackColor = Color.White
        CokeLime.BackColor = Color.White
        CokeLemon.BackColor = Color.White
        CokeRaspberry.BackColor = Color.White
        CokeOrange.BackColor = Color.White
        CokeOrangeVanilla.BackColor = Color.White
        CokeCremeSoda.BackColor = Color.White
        CokeVanilla.BackColor = Color.White

        Small.BackColor = Color.White
        Medium.BackColor = Color.White
        Large.BackColor = Color.White
        ExtraLarge.BackColor = Color.White

        number = 0
        Array.Clear(Product, 0, Product.Length)
        Size = 0
    End Sub

    'Coka Cola Original'
    Private Sub CocaCola_Click(sender As Object, e As EventArgs) Handles CocaCola.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Coca-Cola Original"
            number = number + 1
        End If

        CocaCola.BackColor = Color.SlateGray
    End Sub

    'Cherry Coke'
    Private Sub CokeCherry_Click(sender As Object, e As EventArgs) Handles CokeCherry.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Cherry Coke"
            number = number + 1
        End If

        CokeCherry.BackColor = Color.SlateGray
    End Sub

    'Cherry Vanilla Coke'
    Private Sub CokeCherryVanilla_Click(sender As Object, e As EventArgs) Handles CokeCherryVanilla.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Cherry Vanilla Coke"
            number = number + 1
        End If

        CokeCherryVanilla.BackColor = Color.SlateGray
    End Sub

    'Lemon Coke'
    Private Sub CokeLemon_Click(sender As Object, e As EventArgs) Handles CokeLemon.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Lemon Coke"
            number = number + 1
        End If

        CokeLemon.BackColor = Color.SlateGray
    End Sub

    'Lime Coke'
    Private Sub CokeLime_Click(sender As Object, e As EventArgs) Handles CokeLime.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Lime Coke"
            number = number + 1
        End If

        CokeLime.BackColor = Color.SlateGray
    End Sub

    'Orange Coke'
    Private Sub CokeOrange_Click(sender As Object, e As EventArgs) Handles CokeOrange.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Orange Coke"
            number = number + 1
        End If

        CokeOrange.BackColor = Color.SlateGray
    End Sub

    'Orange Vanilla Coke'
    Private Sub CokeOrangeVanilla_Click(sender As Object, e As EventArgs) Handles CokeOrangeVanilla.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Orange Vanilla Coke"
            number = number + 1
        End If

        CokeOrangeVanilla.BackColor = Color.SlateGray
    End Sub

    'Raspberry Coke'
    Private Sub CokeRaspberry_Click(sender As Object, e As EventArgs) Handles CokeRaspberry.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Raspberry Coke"
            number = number + 1
        End If

        CokeRaspberry.BackColor = Color.SlateGray
    End Sub

    'Coke Vanilla'
    Private Sub CokeVanilla_Click(sender As Object, e As EventArgs) Handles CokeVanilla.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Vanilla Coke"
            number = number + 1
        End If

        CokeVanilla.BackColor = Color.SlateGray
    End Sub

    'CokeCremeSoda'
    Private Sub CokeCremeSoda_Click(sender As Object, e As EventArgs) Handles CokeCremeSoda.Click
        If number = 3 Then
            MsgBox("Cannot have more than 3 flavors",, " ")
            Return
        End If
        If number < 3 Then
            Product(number) = "Creme Soda Coke"
            number = number + 1
        End If

        CokeCremeSoda.BackColor = Color.SlateGray
    End Sub

    'Small Size Option/Button'
    Private Sub Small_Click(sender As Object, e As EventArgs) Handles Small.Click
        Size = 8

        Small.BackColor = Color.SlateGray
        Medium.BackColor = Color.White
        Large.BackColor = Color.White
        ExtraLarge.BackColor = Color.White
    End Sub

    'Medium Size Option/Button'
    Private Sub Medium_Click(sender As Object, e As EventArgs) Handles Medium.Click
        Size = 16

        Small.BackColor = Color.White
        Medium.BackColor = Color.SlateGray
        Large.BackColor = Color.White
        ExtraLarge.BackColor = Color.White
    End Sub

    'Large Size Option/Button'
    Private Sub Large_Click(sender As Object, e As EventArgs) Handles Large.Click
        Size = 24

        Small.BackColor = Color.White
        Medium.BackColor = Color.White
        Large.BackColor = Color.SlateGray
        ExtraLarge.BackColor = Color.White
    End Sub

    'Extra Size Large Option/Button'
    Private Sub ExtraLarge_Click(sender As Object, e As EventArgs) Handles ExtraLarge.Click
        Size = 32

        Small.BackColor = Color.White
        Medium.BackColor = Color.White
        Large.BackColor = Color.White
        ExtraLarge.BackColor = Color.SlateGray
    End Sub

    'Mixing/Dispensing Button Action'
    Private Sub btnMixDis_Click(sender As Object, e As EventArgs) Handles btnMixDis.Click
        Dim FluidLvls(11) As Integer
        Dim answer As Integer
        Dim number2 As Integer = 0
        'Calls subfunction to update values'
        updateValues()

        'If someone does not pick a size it stops the mixing from happening'
        If Size = 0 Then
            MsgBox("You must pick a size.",, " ")
            Return
        End If

        'If user picked one flavor then mix'
        If number = 1 Then
            answer = MsgBox("Do you want to dispense " & Size & "oz" & " " & Product(0) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "")
            If answer = vbYes Then
                MsgBox("Dispensing: " & Size & " " & Product(0),, " ")
            Else
                MsgBox("Canceling",, " ")
                Return
            End If
        End If
        'If user picked two flavor then mix'
        If number = 2 Then
            answer = MsgBox("Do you want to dispense " & Size & "oz" & " " & Product(0) & ", " & Product(1) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "")
            If answer = vbYes Then
                MsgBox("Dispensing: " & Size & " " & Product(0) & ", " & Product(1),, " ")
            Else
                MsgBox("Canceling",, " ")
                Return
            End If
        End If
        'If user picked three flavor then mix'
        If number = 3 Then
            answer = MsgBox("Do you want to dispense " & Size & "oz" & " " & Product(0) & ", " & Product(1) & ", " & Product(2) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "")
            If answer = vbYes Then
                MsgBox("Dispensing: " & Size & "oz" & " " & Product(0) & ", " & Product(1) & ", " & Product(2),, " ")
            Else
                MsgBox("Canceling",, " ")
                Return
            End If
        End If

        'Calculates the number of flavors used and then sets the division amount for those flavors based on it'
        If number = 1 Then
            number2 = 2
        End If
        If number = 2 Then
            number2 = 4
        End If
        If number = 3 Then
            number2 = 8
        End If

        'Calculates how much fluid is used up on mixing'
        Dim pro As Integer
        For pro = 0 To 3
            If Product(pro) = "Coca-Cola Original" Then
                If 0 <= (CCORF - (Size / number2)) Then
                    CCORF = CCORF - (Size / number2)
                    'If syrup gets too low'
                    If CCORF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
            If Product(pro) = "Cherry Coke" Then
                If 0 <= (CCCF - (Size / number2)) Then
                    CCCF = CCCF - (Size / number2)
                    'If syrup gets too low'
                    If CCCF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
            If Product(pro) = "Cherry Vanilla Coke" Then
                If 0 <= (CCCVF - (Size / number2)) Then
                    CCCVF = CCCVF - (Size / number2)
                    'If syrup gets too low'
                    If CCCVF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
            If Product(pro) = "Lemon Coke" Then
                If 0 <= (CCLEF - (Size / number2)) Then
                    CCLEF = CCLEF - (Size / number2)
                    'If syrup gets too low'
                    If CCLEF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
            If Product(pro) = "Lime Coke" Then
                If 0 <= (CCLIF - (Size / number2)) Then
                    CCLIF = CCLIF - (Size / number2)
                    'If syrup gets too low'
                    If CCLIF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
            If Product(pro) = "Orange Coke" Then
                If 0 <= (CCOF - (Size / number2)) Then
                    CCOF = CCOF - (Size / number2)
                    'If syrup gets too low'
                    If CCOF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
            If Product(pro) = "Orange Vanilla Coke" Then
                If 0 <= (CCOVF - (Size / number2)) Then
                    CCOVF = CCOVF - (Size / number2)
                    'If syrup gets too low'
                    If CCOVF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
            If Product(pro) = "Raspberry Coke" Then
                If 0 <= (CCRF - (Size / number2)) Then
                    CCRF = CCRF - (Size / number2)
                    'If syrup gets too low'
                    If CCRF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
            If Product(pro) = "Vanilla Coke" Then
                If 0 <= (CCVF - (Size / number2)) Then
                    CCVF = CCVF - (Size / number2)
                    'If syrup gets too low'
                    If CCVF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
            If Product(pro) = "Creme Soda Coke" Then
                If 0 <= (CCCSF - (Size / number2)) Then
                    CCCSF = CCCSF - (Size / number2)
                    'If syrup gets too low'
                    If CCCSF < 12 Then
                        MsgBox("Syrup is getting low. Order more!",, " ")
                    End If
                    'Stops syrup from going under 0 in the calculations'
                Else
                    MsgBox("Not Enough Syrup",, " ")
                End If
            End If
        Next

        'Calculates amount of CO2 used while mixing'
        If 0 <= (CO2 - (Size / 2)) Then
            CO2 = CO2 - (Size / 2)
        Else
            'Stops syrup from going under 0 in the calculations'
            MsgBox("Not Enough C02",, " ")
        End If

        'If CO2 gets too low'
        If CO2 < 54 Then
            MsgBox("CO2 is getting low. Order more!",, " ")
        End If

        'Loads new variables into array to insert into database'
        Dim I As Integer
        Dim H As Integer
        FluidLvls(0) = CCORF
        FluidLvls(1) = CCCF
        FluidLvls(2) = CCCVF
        FluidLvls(3) = CCLIF
        FluidLvls(4) = CCLEF
        FluidLvls(5) = CCRF
        FluidLvls(6) = CCOF
        FluidLvls(7) = CCOVF
        FluidLvls(8) = CCCSF
        FluidLvls(9) = CCVF
        FluidLvls(10) = CO2

        'Inserts data into database'
        H = 0
        For I = 1 To 11
            Dim strSql As String
            'SQL query'
            strSql = ("UPDATE FLUID SET CURR_AMOUNT = @lvl WHERE FLUID_ID =@ID")
            Dim cmd As New SqlCommand(strSql, conn)
            cmd.Parameters.AddWithValue("@lvl", FluidLvls(H))
            cmd.Parameters.AddWithValue("@ID", I)
            cmd.ExecuteNonQuery()
            H = H + 1
        Next
    End Sub

    'Syrup Level Report'
    Private Sub btnSyrupLvl_Click(sender As Object, e As EventArgs) Handles btnSyrupLvl.Click
        'Updates values'
        updateValues()

        'Report Contents'
        MsgBox(
        "Fluid Levels:" & vbCrLf &
        "----------------------------" & vbCrLf &
        "Coka-Cola Original: " & CCORF & vbCrLf &
        "Cherry Coke: " & CCCF & vbCrLf &
        "Cherry Vanilla Coke: " & CCCVF & vbCrLf &
        "Lemon Coke: " & CCLEF & vbCrLf &
        "Lime Coke: " & CCLIF & vbCrLf &
        "Orange Coke: " & CCOF & vbCrLf &
        "Orange Vanilla Coke: " & CCOVF & vbCrLf &
        "Raspberry Coke: " & CCRF & vbCrLf &
        "Vanilla Coke: " & CCVF & vbCrLf &
        "Creme Soda Coke: " & CCCSF & vbCrLf &
        vbCrLf &
        "CO2: " & CO2,, " ")
    End Sub

    'Exit Button'
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Closes application'
        conn.Close()
        Me.Close()
    End Sub

    'Maintnance Report'
    Private Sub btnMaintRep_Click(sender As Object, e As EventArgs) Handles btnMaintRep.Click
        'Updates Values'
        updateValues()

        'Local Variable Declarations'
        Dim CAPACITY As Integer
        Dim MAINDATE As String
        Dim strSql As String
        Dim myReader As SqlDataReader

        'Calls to see what is the most requested flavor'
        mostRequested()

        'Calls to see what is the least requested flavor'
        leastRequested()

        'SQL query'
        strSql = ("SELECT FLAVOR_NAME, EXP_DATE, MAIN_DATE, CAPACITY, CURR_AMOUNT FROM FLUID")
        Dim cmd As New SqlCommand(strSql, conn)

        'SQL reader'
        myReader = cmd.ExecuteReader()
        Dim I As Integer = 0
        Dim TotalAmount As Integer = 0
        Dim STRINGLIST(11) As String

        'Reads values from database'
        If myReader.HasRows Then
            Do While myReader.Read
                DESC = myReader.Item("FLAVOR_NAME")
                ORDERDATE = myReader.Item("EXP_DATE")
                MAINDATE = myReader.Item("MAIN_DATE")
                CAPACITY = myReader.Item("CAPACITY")
                FAMOUNT = myReader.Item("CURR_AMOUNT")
                STRINGLIST(I) = ("Flavor Name: " + DESC.ToString() + vbCrLf + "Expiration Date: " + ORDERDATE.ToString() + "     " +
                "Last Fill: " + MAINDATE.ToString + vbCrLf + "Capacity: " + CAPACITY.ToString + "                            " + "Current Amount: " + FAMOUNT.ToString() + vbCrLf +
                "------------------------------------------------------------")

                I = I + 1
            Loop
        End If

        'Report Contents'
        MsgBox("Maintnance Report:" + vbCrLf +
               "------------------------------------------------------------" + vbCrLf +
               Join(STRINGLIST, vbCrLf))
        myReader.Close()
    End Sub

    'Management Report'
    Private Sub btnManRep_Click(sender As Object, e As EventArgs) Handles btnManRep.Click
        'Updates Values'
        updateValues()

        'Calls to see what is the most requested flavor'
        mostRequested()

        'Calls to see what is the least requested flavor'
        leastRequested()

        'Local Variable Declarations'
        Dim strSql As String
        Dim myReader As SqlDataReader
        Dim I As Integer = 0
        Dim TotalAmount As Integer = 0
        Dim TotalIncome As Integer = 0
        Dim TotalNumber As Integer = 0

        'SQL query'
        strSql = ("SELECT ORDER_AMOUNT FROM ORDERTABLE WHERE ORDER_AMOUNT <> 100")
        Dim cmd As New SqlCommand(strSql, conn)
        myReader = cmd.ExecuteReader()

        'Reads data from orders'
        If myReader.HasRows Then
            Do While myReader.Read
                TotalAmount = TotalAmount + Convert.ToInt32(myReader.Item("ORDER_AMOUNT"))
                TotalIncome = TotalIncome + Convert.ToInt32(myReader.Item("ORDER_AMOUNT"))
            Loop
        End If

        'Calculates average income'
        TotalIncome = TotalIncome * 4
        TotalNumber = TotalIncome / 4

        'Report content'
        MsgBox("Management Report:" + vbCrLf +
               "---------------------------------------------------------------------------------" + vbCrLf +
               "Total Number of Orders: " + TotalNumber.ToString() + vbCrLf +
               "Total Income of Orders: $" + TotalIncome.ToString() + vbCrLf +
               "Average Size and Revenue: Extra Large $4.00" + vbCrLf +
               "Most Common Size: " + "32oz" + vbCrLf +
               "Most Requested Flavor: " + MOSTFREQUENT + vbCrLf +
               "Least Requested FLavor: " + LEASTFREQUENT + vbCrLf +
               "---------------------------------------------------------------------------------")
        myReader.Close()
    End Sub

    'Order Report'
    Private Sub btnOrders_Click(sender As Object, e As EventArgs) Handles btnOrders.Click
        'Calls to see what is the most requested flavor'
        mostRequested()

        'Calls to see what is the least requested flavor'
        leastRequested()

        'Local Variable Declarations'
        Dim StartDate As Date
        Dim EndDate As Date
        Dim TotalIncome As Integer = 0
        Dim TotalAmount As Integer = 0
        Dim TotalNumber As Integer = 0
        Dim I As Integer = 0
        Dim STRINGLIST(25) As String
        Dim strSql As String
        Dim myReader As SqlDataReader

        'Calculates dates that the orders must be between'
        Try
            StartDate = Convert.ToDateTime(txtbxStrtDate.Text)
            EndDate = Convert.ToDateTime(txtbxEndDate.Text)
        Catch ex As Exception
            MsgBox("Please insert valid dates")
            Return
        End Try


        'SQL Query'
        strSql = ("SELECT ORDER_ID, ORDER_ID_LINE_NO, ORDER_DATE, ORDER_DESC, FLUID_ID, FLUID_AMOUNT, ORDER_AMOUNT FROM ORDERTABLE WHERE @STRTDATE <= ORDER_DATE AND @ENDDATE >= ORDER_DATE ")
        Dim cmd As New SqlCommand(strSql, conn)
        cmd.Parameters.AddWithValue("@STRTDATE", StartDate)
        cmd.Parameters.AddWithValue("@ENDDATE", EndDate)
        myReader = cmd.ExecuteReader()

        'Reads data from database'
        If myReader.HasRows Then
            Do While myReader.Read
                'Reads total income and total number of orders'
                TotalAmount = TotalAmount + Convert.ToInt32(myReader.Item("ORDER_AMOUNT"))
                If Convert.ToInt32(myReader.Item("ORDER_AMOUNT")) <> 100 Then
                    TotalIncome = TotalIncome + Convert.ToInt32(myReader.Item("ORDER_AMOUNT"))
                End If
                ID = myReader.Item("ORDER_ID")
                IDLINE = myReader.Item("ORDER_ID_LINE_NO")
                ORDERDATE = myReader.Item("ORDER_DATE")
                DESC = myReader.Item("ORDER_DESC")
                FID = myReader.Item("FLUID_ID")
                FAMOUNT = myReader.Item("FLUID_AMOUNT")
                ORDERAMOUNT = myReader.Item("ORDER_AMOUNT")
                STRINGLIST(I) = ("|   " + ID.ToString() + "   |   " + IDLINE.ToString() + "   |   " + ORDERDATE + "   |   " + DESC + "   |   " +
                    FID.ToString() + "   |   " + FAMOUNT.ToString() + "oz   |   $" + ORDERAMOUNT.ToString() + "   |")

                I = I + 1
            Loop
        End If

        'Calculates total income and and total number of orders'
        TotalIncome = TotalIncome * 4
        TotalNumber = TotalIncome / 4

        'Report contents'
        MsgBox("ORDERS:" + vbCrLf +
               "---------------------------------------------------------------------------------" + vbCrLf +
               "ID, Line Number, Date Ordered, Description, Fluid ID, Fluid Amount, $" + vbCrLf +
               "---------------------------------------------------------------------------------" + vbCrLf +
               Join(STRINGLIST, vbCrLf) +
               "---------------------------------------------------------------------------------" + vbCrLf +
               "Total Number of Orders: " + TotalNumber.ToString() + vbCrLf +
               "Total Income of Orders: $" + TotalIncome.ToString() + vbCrLf +
               "Average Size and Revenue: Extra Large $4.00" + vbCrLf +
               "Most Common Size: " + "32oz" + vbCrLf +
               "Most Requested Flavor: " + MOSTFREQUENT + vbCrLf +
               "Least Requested FLavor: " + LEASTFREQUENT + vbCrLf +
               "---------------------------------------------------------------------------------")
        myReader.Close()
    End Sub

    'Reorder Button'
    Private Sub btnReOrder_Click(sender As Object, e As EventArgs) Handles btnReOrder.Click
        Dim strSql As String
        Dim strSql2 As String
        Dim datee As Date
        datee = Today
        Try
            'Arrays of the values to check'
            Dim fluidIds() As Integer = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}
            Dim fluidNames() As String = {"CokaCola", "CokeCherry", "CherryVanilla", "LimeCoke", "LemonCoke", "RaspberryCoke", "OrangeCoke", "OrangeVanilla", "CremeSodaCoke", "VanillaCoke"}
            Dim fluidLevels() As Integer = {CCORF, CCCF, CCCVF, CCLIF, CCLEF, CCRF, CCOF, CCOVF, CCCSF, CCVF}

            'Checks each value to see if it needs to be refilled and adds to order which ones do'
            For i As Integer = 0 To fluidIds.Length - 1
                If fluidLevels(i) < 12 Then
                    Dim strSql3 As String = "UPDATE FLUID SET CURR_AMOUNT = @lvl WHERE FLUID_ID = @ID"
                    Dim cmd As New SqlCommand(strSql3, conn)
                    cmd.Parameters.AddWithValue("@lvl", 34)
                    cmd.Parameters.AddWithValue("@ID", fluidIds(i))
                    cmd.ExecuteNonQuery()

                    Dim strSql4 As String = "INSERT INTO ORDERTABLE ([ORDER_ID], [ORDER_ID_LINE_NO], [ORDER_DATE], [ORDER_DESC], [FLUID_ID], [FLUID_AMOUNT], [ORDER_AMOUNT]) 
                       VALUES (@OID, @LINEID, @DATE, @DESC, @FID, @FAMOUNT, @OAMOUNT)"
                    Dim fnd As New SqlCommand(strSql4, conn)
                    RandomNumbers()
                    fnd.Parameters.AddWithValue("@OID", _rnd)
                    fnd.Parameters.AddWithValue("@LINEID", i + 1)
                    RandomDate()
                    fnd.Parameters.AddWithValue("@DATE", datee)
                    fnd.Parameters.AddWithValue("@DESC", fluidNames(i))
                    fnd.Parameters.AddWithValue("@FID", fluidIds(i))
                    fnd.Parameters.AddWithValue("@FAMOUNT", 34)
                    fnd.Parameters.AddWithValue("@OAMOUNT", 50)
                    fnd.ExecuteNonQuery()
                End If
            Next

            'Updates CO2 and inserts order when it is low'
            If CO2 < 54 Then
                strSql = ("UPDATE FLUID SET CURR_AMOUNT = @lvl WHERE FLUID_ID =@ID")
                Dim cmd As New SqlCommand(strSql, conn)
                cmd.Parameters.AddWithValue("@lvl", 170)
                cmd.Parameters.AddWithValue("@ID", 11)
                cmd.ExecuteNonQuery()

                strSql2 = ("INSERT INTO ORDERTABLE ([ORDER_ID], [ORDER_ID_LINE_NO], [ORDER_DATE], [ORDER_DESC], [FLUID_ID], [FLUID_AMOUNT], [ORDER_AMOUNT]) 
                       VALUES (@OID, @LINEID, @DATE, @DESC, @FID, @FAMOUNT, @OAMOUNT)")
                Dim fnd As New SqlCommand(strSql2, conn)
                fnd.Parameters.AddWithValue("@OID", _rnd)
                fnd.Parameters.AddWithValue("@LINEID", 11)
                RandomDate()
                fnd.Parameters.AddWithValue("@DATE", datee)
                fnd.Parameters.AddWithValue("@DESC", "CO2")
                fnd.Parameters.AddWithValue("@FID", 11)
                fnd.Parameters.AddWithValue("@FAMOUNT", 170)
                fnd.Parameters.AddWithValue("@OAMOUNT", 100)
                fnd.ExecuteNonQuery()
            End If
        Catch ex As Exception
            MsgBox("Error: Please try again.")
        End Try
    End Sub

    'Sub to update global variables according to database'
    Sub updateValues()
        'Local Variable declarations'
        Dim strSql As String
        Dim myReader As SqlDataReader
        Dim FluidLvls(11) As Integer
        Dim I As Integer = 0

        'Database Query'
        strSql = "SELECT CURR_AMOUNT FROM FLUID"
        Dim cmd As New SqlCommand(strSql, conn)
        myReader = cmd.ExecuteReader()

        'Reads each fluid level'
        If myReader.HasRows Then
            Do While myReader.Read
                FluidLvls(I) = myReader.GetInt32(0)
                I += 1
            Loop
        End If

        'Updates global array variable based on fluid levels'
        CCORF = FluidLvls(0)
        CCCF = FluidLvls(1)
        CCCVF = FluidLvls(2)
        CCLIF = FluidLvls(3)
        CCLEF = FluidLvls(4)
        CCRF = FluidLvls(5)
        CCOF = FluidLvls(6)
        CCOVF = FluidLvls(7)
        CCCSF = FluidLvls(8)
        CCVF = FluidLvls(9)
        CO2 = FluidLvls(10)
        myReader.Close()
    End Sub

    'Sub to see what is the most requested flavor'
    Sub mostRequested()
        'Local Variable Declarations'
        Dim strSql As String

        'Database Query'
        strSql = "SELECT TOP 1 FLUID_ID, COUNT(*) AS CNT FROM ORDERTABLE GROUP BY FLUID_ID ORDER BY CNT DESC"

        Dim cmd As New SqlCommand(strSql, conn)

        'Execute SQL command and retrieve the result'
        Dim reader As SqlDataReader = cmd.ExecuteReader()

        If reader.HasRows Then
            'Read the first row, which should contain the most frequent fluid ID'
            reader.Read()
            Dim mostFrequentFluidId As Integer = reader.GetInt32(0)

            'Determine the most frequent flavor based on the fluid ID'
            Select Case mostFrequentFluidId
                Case 1
                    MOSTFREQUENT = "CocaColaOriginal"
                Case 2
                    MOSTFREQUENT = "CokeCherry"
                Case 3
                    MOSTFREQUENT = "CokeCherryVanilla"
                Case 4
                    MOSTFREQUENT = "CokeLime"
                Case 5
                    MOSTFREQUENT = "CokeLemon"
                Case 6
                    MOSTFREQUENT = "CokeRaspberry"
                Case 7
                    MOSTFREQUENT = "OrangeCoke"
                Case 8
                    MOSTFREQUENT = "OrangeVanillaCoke"
                Case 9
                    MOSTFREQUENT = "CremeSodaCoke"
                Case 10
                    MOSTFREQUENT = "CokeVanilla"
                Case 11
                    MOSTFREQUENT = "CO2"
            End Select
        End If
        reader.Close()
    End Sub

    Sub leastRequested()
        'Local Variable Declarations'
        Dim strSql As String

        'Database Query'
        strSql = "SELECT TOP 1 FLUID_ID, COUNT(*) AS CNT FROM ORDERTABLE GROUP BY FLUID_ID ORDER BY CNT ASC"

        Dim cmd As New SqlCommand(strSql, conn)

        'Execute SQL command and retrieve the result'
        Dim reader As SqlDataReader = cmd.ExecuteReader()

        If reader.HasRows Then
            'Read the first row, which should contain the most frequent fluid ID'
            reader.Read()
            Dim mostFrequentFluidId As Integer = reader.GetInt32(0)

            'Determine the most frequent flavor based on the fluid ID'
            Select Case mostFrequentFluidId
                Case 1
                    LEASTFREQUENT = "CocaColaOriginal"
                Case 2
                    LEASTFREQUENT = "CokeCherry"
                Case 3
                    LEASTFREQUENT = "CokeCherryVanilla"
                Case 4
                    LEASTFREQUENT = "CokeLime"
                Case 5
                    LEASTFREQUENT = "CokeLemon"
                Case 6
                    LEASTFREQUENT = "CokeRaspberry"
                Case 7
                    LEASTFREQUENT = "OrangeCoke"
                Case 8
                    LEASTFREQUENT = "OrangeVanillaCoke"
                Case 9
                    LEASTFREQUENT = "CremeSodaCoke"
                Case 10
                    LEASTFREQUENT = "CokeVanilla"
                Case 11
                    LEASTFREQUENT = "CO2"
            End Select
        End If
        reader.Close()
    End Sub

    'Random Number Generator'
    Sub RandomNumbers()
        _rnd = RandomNumber.Next(100000, 300000)
    End Sub

    'Random Date Generator'
    Sub RandomDate()
        'Define a start date and the range of possible offsets'
        Dim startDate As New DateTime(2010, 1, 1)
        Dim yearOffset As Integer = 10
        Dim monthOffset As Integer = 12
        Dim dayOffset As Integer = 20

        'Generate random year, month, and day offsets'
        Dim rand As New Random()
        Dim year As Integer = rand.Next(-yearOffset, yearOffset + 1)
        Dim month As Integer = rand.Next(1, monthOffset + 1)
        Dim day As Integer = rand.Next(1, dayOffset + 1)

        'Calculate the random date by adding the offsets to the start date'
        Dim randomDate As DateTime = startDate.AddYears(year).AddMonths(month).AddDays(day)

        'Convert the date to a string using a specific format'
        DateString = randomDate.ToString("yyyy-MM-dd")
    End Sub

    'Order Class'
    Public Class Order
        Public Property ID As Integer
        Public Property IDLINE As Integer
        Public Property ORDERDATE As String
        Public Property DESC As String
        Public Property FID As Integer
        Public Property FAMOUNT As Integer

        Public Sub New(ByVal id As Integer)
            'Open Database Connection'
            Dim conn As New SqlConnection
            conn.ConnectionString = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\cscha\OneDrive\Desktop\Coca-Cola_Freestyle\Coca-Cola\COCACOLA.mdf;Integrated Security=True;Connect Timeout=30"
            conn.Open()

            'Database Query'
            Dim strSql As String = "SELECT ID, IDLINE, ORDERDATE, [DESC], FID, FAMOUNT FROM ORDERTABLE WHERE ID = @id"
            Dim cmd As New SqlCommand(strSql, conn)
            cmd.Parameters.AddWithValue("@id", id)

            'Execute SQL command and retrieve the result'
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            If reader.HasRows Then
                'Read the first row, which should contain the order details'
                reader.Read()
                Me.ID = reader.GetInt32(0)
                Me.IDLINE = reader.GetInt32(1)
                Me.ORDERDATE = reader.GetDateTime(2).ToString("yyyy-MM-dd HH:mm:ss")
                Me.DESC = reader.GetString(3)
                Me.FID = reader.GetInt32(4)
                Me.FAMOUNT = reader.GetInt32(5)
            End If

            'Close the database connection'
            conn.Close()
            reader.Close()
        End Sub
    End Class

    'Flavor Class'
    Public Class Flavor
        Public Property Name As String
        Public Property ExpirationDate As DateTime
        Public Property MaintenanceDate As DateTime
        Public Property Capacity As Integer
        Public Property CurrentAmount As Integer

        Public Sub New(ByVal name As String)
            Dim connectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\cscha\OneDrive\Desktop\Coca-Cola_Freestyle\Coca-Cola\COCACOLA.mdf;Integrated Security=True;Connect Timeout=30"
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand("SELECT FLAVOR_NAME, EXP_DATE, MAIN_DATE, CAPACITY, CURR_AMOUNT FROM FLAVOR_TABLE WHERE FLAVOR_NAME=@Name", connection)
                    command.Parameters.AddWithValue("@Name", name)
                    Using reader As SqlDataReader = command.ExecuteReader()
                        If reader.Read() Then
                            Me.Name = reader.GetString(0)
                            Me.ExpirationDate = reader.GetDateTime(1)
                            Me.MaintenanceDate = reader.GetDateTime(2)
                            Me.Capacity = reader.GetInt32(3)
                            Me.CurrentAmount = reader.GetInt32(4)
                        End If
                        reader.Close()
                    End Using
                End Using
                connection.Close()
            End Using
        End Sub
    End Class
End Class

