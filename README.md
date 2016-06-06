'  Author: Andrew Jo
'  Date: 3 April 2002
'  Purpose: Check line data for missing records
'  Parameters:
'              CheckType - Flight, Car, Hotel
'               Value - Row number
'
Function CheckMissing( CheckType , Value )
	Dim ws As New notesuiworkspace
	Dim doc As notesdocument
	Dim uidoc As notesuidocument
	Dim counter As Integer
	
	Set uidoc = ws.currentdocument
	Set doc = uidoc.document
	
	counter = 0
	Select Case CheckType
	Case "Flight"
		If doc.getitemvalue("Departing" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If doc.getitemvalue("DepartureDate" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If doc.getitemvalue("DepartureTime" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If doc.getitemvalue("Destination" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If counter = 0 Or counter = 4 Then
			CheckMissing = True
			Exit Function
		Else
			Messagebox "Missing value on flight details line " & Cstr(Value) , 0 + 64 , "Missing Value"
			Call uidoc.GotoField( "Departing" & Cstr(Value) )
			CheckMissing = False
			Exit Function
		End If
		
	Case "Car"
		If doc.getitemvalue("CarHireDate" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If doc.getitemvalue("CarPickUp" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If doc.getitemvalue("CarHireDate" & Cstr(Value) & "a")(0) <> "" Then
			counter = counter + 1	
		End If
		If doc.getitemvalue("CarType" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If counter = 0 Or counter = 4 Then
			CheckMissing = True
			Exit Function
		Else
			Messagebox "Missing value on Car Hire details line " & Cstr(Value) , 0 + 64 , "Missing Value"
			Call uidoc.GotoField( "CarHireDate" & Cstr(Value) )
			CheckMissing = False
			Exit Function
		End If
		
	Case "Hotel"
		If doc.getitemvalue("Hotel" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If doc.getitemvalue("AccDateIn" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If doc.getitemvalue("AccDateOut" & Cstr(Value))(0) <> "" Then
			counter = counter + 1	
		End If
		If counter = 0 Or counter = 3 Then
			CheckMissing = True
			Exit Function
		Else
			Messagebox "Missing value on hotel details line " & Cstr(Value) , 0 + 64 , "Missing Value"
			Call uidoc.GotoField( "Hotel" & Cstr(Value) )
			CheckMissing = False
			Exit Function
		End If		
	End Select
	
End Function
