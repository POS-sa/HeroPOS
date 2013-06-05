Attribute VB_Name = "modDatabaseUpdate"
Public Sub updatedatabase()
'2.1012.02
ActiveUpdateServer "update users set owner_transfer = 'False' where owner_transfer IS NULL"
ActiveUpdateServer "update products set Receipe_charge_item = 'False'  where Receipe_charge_item IS NULL"
ActiveUpdateServer "update users set reprint = 'False' where reprint IS NULL"


End Sub
