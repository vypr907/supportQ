Sub userForm_Initialize()
   setUsersRange
   'load data
   Dim item As Variant
   For Each item In usersRng 'dataSht.Range("users")
      With Me.userCboBx
         .AddItem item.Value
      End With
   Next item
End Sub

Sub submitBtn_Click()
   'authorizer(Me.userCboBx.ListIndex, Me.pinEntryBx)
   Hide Me
End Sub
