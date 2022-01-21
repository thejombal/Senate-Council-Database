
'----------------------------------
Insert  6/30
'----------------------------------

CurrentDb.Execute "INSERT INTO tbl_Council(CouncABR, CouncName, Type, TotalCouncSeat)" & "VALUES('" & Me.cbo_CouncABR & "','" & Me.cbo_CouncName & "','" & Me.cbo_Type & "','" & Me.cbo_TotalCouncSeat & "')"


CurrentDb.Execute "INSERT INTO tbl_Member(FirstName, LastName, Campus, Senator, Active, Office, Phone, Email, LastUpdated)" & "VALUES('" & Me.cbo_FirstName & "','" & Me.cbo_LastName & "','" & Me.cbo_Campus & "','" & Me.cbo_Senator & "','" & Me.cbo_Active & "','" & Me.cbo_Office & "','" & Me.cbo_Phone & "','" & Me.cbo_Email & "','" & Me.txt_LastUpdated & "')"

MsgBox "Entries have been added.", , "Success!"


'Member AND Council Lookup Query with the Assignment Table Insert (Subform)

insert rows for testing

fix bug on member insert

yellow tail carbaent sauvigonon


CurrentDb.Execute "INSERT INTO tbl_AssignmentID(CouncABR, MemberID, TermEnding, TermSequence, Role, UnitABR, CurrentlyServing)" & "VALUES('" & Me.cbo_CouncABR & "','" & Me.cbo_MemberID & "','" & Me.cbo_TermEnding & "','" & Me.cbo_TermSequence & "','" & Me.cbo_Role & "','" & Me.cbo_UnitABR & "','" & Me.cbo_CurrentlyServing & "')"



'------------------------
7/7 2nd Half-Insert 
'------------------------

CurrentDb.Execute "INSERT INTO tbl_CouncilUnitSeat(CouncABR, UnitABR, UnitSeatAlloc)" & "VALUES('" & Me.cbo_CouncABR & "','" & Me.cbo_UnitABR & "','" & Me.cbo_UnitSeatAlloc & "')"

MsgBox "Entries have been added.", , "Success!"

CurrentDb.Execute "INSERT INTO tbl_MemberUnit(UnitABR, MemberID)" & "VALUES('" & Me.cbo_UnitABR & "','" & Me.cbo_MemberID & "')"

MsgBox "Entries have been added.", , "Success!"

CurrentDb.Execute "INSERT INTO tbl_Unit(UnitABR, UnitName)" & "VALUES('" & Me.cbo_UnitABR & "','" & Me.cbo_UnitName & "')"

MsgBox "Entries have been added.", , "Success!"


'------------------------
7/7 Update Statements
'------------------------



'First Name

If cbo_FirstName.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_FirstName.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Member " & " SET [FirstName] = " & "'" & Me.cbo_FirstName & "'" & "" & " WHERE [MemberID] = " & Me.cbo_MemberID & ""
End If
End If

'Last Name

If cbo_LastName.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_LastName.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Member " & " SET [Last Name] = " & "'" & Me.cbo_LastName & "'" & "" & " WHERE [MemberID] = " & Me.cbo_MemberID & ""
End If
End If

'Campus

If cbo_Campus.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_Campus.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Member " & " SET [Campus] = " & "'" & Me.cbo_Campus & "'" & "" & " WHERE [MemberID] = " & Me.cbo_MemberID & ""
End If
End If


'Senator

If cbo_Senator.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_Senator.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Member " & " SET [Senator] = " & "'" & Me.cbo_Senator & "'" & "" & " WHERE [MemberID] = " & Me.cbo_MemberID & ""
End If
End If

'Active


If cbo_Active.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_Active.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Member " & " SET [Active] = " & "'" & Me.cbo_Active & "'" & "" & " WHERE [MemberID] = " & Me.cbo_MemberID & ""
End If
End If

'Office

If cbo_Office.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_Office.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Member " & " SET [Office] = " & "'" & Me.cbo_Office & "'" & "" & " WHERE [MemberID] = " & Me.cbo_MemberID & ""
End If
End If

'Phone

If cbo_Phone.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_Phone.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Member " & " SET [Phone] = " & "'" & Me.cbo_Phone & "'" & "" & " WHERE [MemberID] = " & Me.cbo_MemberID & ""
End If
End If

'Email

If cbo_Email.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_Email.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Member " & " SET [Email] = " & "'" & Me.cbo_Email & "'" & "" & " WHERE [MemberID] = " & Me.cbo_MemberID & ""
End If
End If


'Last Updated

If txt_LastUpdated.Value = "" Then
    Exit Sub
Else   
If IsNull(txt_LastUpdated.Value) = False Then
CurrentDb.Execute "UPDATE tbl_Member " & " SET [LastUpdated] = " & CDbl(Me.txt_LastUpdated) & "" & " WHERE [MemberID] = " & Me.cbo_MemberID & ""
End If
End If


'Update All Buttons

Call cmd_FirstName_Click
Call cmd_LastName_Click
Call cmd_Campus_Click
Call cmd_Senator_Click
Call cmd_Active_Click
Call cmd_Office_Click
Call cmd_Phone_Click
Call cmd_Email_Click
Call cmd_LastUpdated_Click


MsgBox "Entry updated.", , "Success!"

End Sub


'------------------------
7/8 Council Update
'------------------------

'CouncName
If cbo_CouncName.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_CouncName.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Council " & " SET [CouncName] = " & "'" & Me.cbo_CouncName & "'" & "" & " WHERE [CouncABR] = " & "'" & Me.cbo_CouncABR & "'" & ""

End If
End If

'Type
If cbo_Type.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_Type.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Council " & " SET [Type] = " & "'" & Me.cbo_Type & "'" & "" & " WHERE [CouncABR] = " & "'" & Me.cbo_CouncABR & "'" & ""

End If
End If


'Total Council Seat
If cbo_TotalCouncSeat.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_TotalCouncSeat.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Council " & " SET [TotalCouncSeat] = " &  Me.cbo_TotalCouncSeat & "" & " WHERE [CouncABR] = " & "'" & Me.cbo_CouncABR & "'" & ""

End If
End If

'Update All Buttons

Call cmd_CouncName_Click
Call cmd_Type_Click
Call cmd_TotalCouncSeat_Click


MsgBox "Entries have been added.", , "Success!"

End Sub


'------------------------
7/8 Assignment Update
'------------------------

'MemberID

If cbo_MemberID.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_MemberID.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Assignment " & " SET [MemberID] = " & Me.cbo_MemberID & "" & " WHERE [AssignmentID] = " & Me.cbo_AssignmentID & ""
End If
End If

End Sub

'Council ABR

If cbo_CouncABR.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_CouncABR.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Assignment " & " SET [CouncABR] = " & "'" & Me.cbo_CouncABR & "'" & "" & " WHERE [AssignmentID] = " & Me.cbo_AssignmentID & ""
End If
End If


'Term Ending

If cbo_TermEnding.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_TermEnding.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Assignment " & " SET [TermEnding] = " & Me.cbo_TermEnding & "" & " WHERE [AssignmentID] = " & Me.cbo_AssignmentID & ""
End If
End If


'Term Sequence

If cbo_TermSequence.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_TermSequence.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Assignment " & " SET [TermSequence] = " & "'" & Me.cbo_TermSequence & "'" & "" & " WHERE [AssignmentID] = " & Me.cbo_AssignmentID & ""
End If
End If


'Role

If cbo_Role.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_Role.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Assignment " & " SET [Role] = " & "'" & Me.cbo_Role & "'" & "" & " WHERE [AssignmentID] = " & Me.cbo_AssignmentID & ""
End If
End If

'Unit ABR

If cbo_UnitABR.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_UnitABR.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Assignment " & " SET [UnitABR] = " & "'" & Me.cbo_UnitABR & "'" & "" & " WHERE [AssignmentID] = " & Me.cbo_AssignmentID & ""
End If
End If


'Currently Serving

If cbo_CurrentlyServing.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_CurrentlyServing.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_Assignment " & " SET [CurrentlyServing] = " & "'" & Me.cbo_CurrentlyServing & "'" & "" & " WHERE [AssignmentID] = " & Me.cbo_AssignmentID & ""
End If
End If



Call cmd_MemberID_Click
Call cmd_CouncABR_Click
Call cmd_TermEnding_Click
Call cmd_TermSequence_Click
Call cmd_Role_Click
Call cmd_UnitABR_Click
Call cmd_CurrentlyServing_Click

MsgBox "Entry updated.", , "Success!"

Me.Refresh



'------------------------
7/11 Last Updates
'------------------------

'Counc Abreviation
If cbo_CouncABR.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_CouncABR.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_CouncilUnitSeat " & " SET [CouncABR] = " & "'" & Me.cbo_CouncABR & "'" & "" & " WHERE [CouncUnitSeatID] = " & Me.cbo_CouncUnitSeatID & ""

End If
End If

'Unit Abreviation
If cbo_UnitABR.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_UnitABR.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_CouncilUnitSeat " & " SET [UnitABR] = " & "'" & Me.cbo_UnitABR & "'" & "" & " WHERE [CouncUnitSeatID] = " & Me.cbo_CouncUnitSeatID & ""

End If
End If


'Unit Seat Allocation
If cbo_UnitSeatAlloc.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_UnitSeatAlloc.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_CouncilUnitSeat " & " SET [UnitSeatAlloc] = " &  Me.cbo_UnitSeatAlloc &  "" & " WHERE [CouncUnitSeatID] = " & Me.cbo_CouncUnitSeatID & ""

End If
End If



'Call
Call cmd_CouncABR_Click
Call cmd_UnitABR_Click
Call cmd_UnitSeatAlloc_Click

MsgBox "Entry updated.", , "Success!"

Me.Refresh

'--------------
Member Unit Update
'--------------

If cbo_MemberID.Value = "" Then
    Exit Sub
Else
    If IsNull(cbo_MemberID.Value) = False Then
        CurrentDb.Execute "UPDATE tbl_CouncilUnitSeat " & " SET [UnitSeatAlloc] = " &  Me.cbo_MemberID &  "" & " WHERE [UnitABR] = " & "'" & Me.cbo_UnitABR & "'" & ""

End If
End If








 --council Roster of second termers
 select a.CouncABR, m.LastName, a.termending 'Term Ending', a.TermSequence, a.unitABR, m.office, m.Phone, m.Email
 from assignment a 
 join member m
 on a.MemberID = m.MemberID
 where currentlyserving = 'Y' and CouncABR is not null and a.TermSequence = 'Second'
 order by a.termending 



 --list of senators grouped by council
 select a.CouncABR, m.LastName, m.Senator, m.Active
 from assignment a
 join member m
 on a.MemberID = m.MemberID
 where currentlyserving = 'Y'and m.Senator = 'Y'
 group by a.CouncABR, m.LastName, m.Senator, m.Active
 order by CouncABR desc



--Unit seats filled on a council
 select a.CouncABR, unitABR 'Unit', count(unitABR) 'Seats filled'
 from assignment a
 where currentlyserving = 'Y'  and a.CouncABR = 'GEN' --and unitABR = 'CAS'
 group by a.CouncABR, unitABR



--Vacant seats on council by unit
 select a.CouncABR,a.unitABR,UnitSeatAlloc 'Authorized',count(a.unitABR) 'Filled',unitseatalloc - count(a.unitABR) 'Vacant'
 from Assignment a
 join CouncilUnitSeat cu
 on a.CouncABR = a.CouncABR and a.unitABR = cu.UnitABR
 where CurrentlyServing = 'y'
 group by a.CouncABR, a.unitABR, UnitSeatAlloc


