Attribute VB_Name = "Demo2_CodingABC"
Option Explicit
Option Private Module


Private Sub DemoModelObjects_ROUND_1()
    ' Demonstrate:
    ' 1. how to obtain a DB "instance"
    ' 2. how to do things with DB, like record_count
    ' 3. basic looping of MultifamilyAmenity
    ' 4. how to delete a record row

    ' How to obtain a DB instance
    Dim DB As clsDb
    Set DB = GetDb_MultifamilyAmenity
    'Debug.Print "DB.RecordCount: " & DB.RecordCount

    ' How to LOOP through DB records
    Dim amenity As clsMultifamilyAmenity
    
    For Each amenity In DB.GetAllRecords
        'Debug.Print amenity.name
    Next amenity

    ' How to get a SPECIFIC DB record
    Dim library_amenity As clsMultifamilyAmenity
    
    Set library_amenity = DB.GetRecordById(15)
    Debug.Print "library_amenity.name: " & library_amenity.name
    
    MsgBox library_amenity.nlg_alias
End Sub



Private Sub DemoModelObjects_ROUND_2()
    ' Demonstrate
    ' 1. data validation enforcement with property.status
    ' 2. data validation enforcement with REMOTE field
    ' 3. all data validation mirrors Generic EditView

    Dim DB As clsDb
    Set DB = GetDb_MultifamilyRentComp

    ' Get an arbitrary record to work with
    Dim prop As clsMultifamilyRentComp
    Set prop = DB.GetFirstRecordOrNothing
    Set prop = DB.GetRecordByNumber(4)
    
    'Debug.Print "prop.name: " & prop.name, prop.id, prop.street_address

    ' REMOTE fields cannot be changed with the dot operator
    ' the only way to change a remote field is through DB.IngestExternalRecords()
    'prop.name = "SOME OTHER NAME"

    ' Demonstrate how to WRITE a LEGAL value
    'prop.status = "Subject"

    ' Demonstrate writing ILLEGAL value
    'prop.status = "FDFDFDF"

    ' Demonstrate LOOP assignment (Exclude all)
    For Each prop In DB.GetAllRecords
        prop.status = "Excluded"
        Debug.Print prop.name
    Next prop
End Sub



Private Sub DemoModelObjects_ROUND_3()
    ' Demonstrate
    ' 1. ManyToManyField relation

    Dim DB As clsDb
    Set DB = GetDb_MultifamilyRentComp

    Dim prop As clsMultifamilyRentComp
    Set prop = DB.GetRecordByNumber(1)

    Debug.Print "------- AMENITIES @ " & prop.name & " ------------"
    
    ' Demonstration ManyToMany, notice prop.amenities.all()!
    Dim amenity As clsMultifamilyAmenity

    For Each amenity In prop.amenities.all
        Debug.Print amenity.name
    Next amenity
End Sub




Private Sub DemoModelObjects_ROUND_4()
    ' Demonstrate
    ' 1. ForeignKey relation

    Dim DB As clsDb
    Set DB = GetDb_MultifamilyRentCompUnit
    
    Dim unit As clsMultifamilyRentCompUnit
    
    For Each unit In DB.GetAllRecords
        'Debug.Print unit.id, unit.parent_property.name
    Next unit
    
    Dim prop As clsMultifamilyRentComp
    Set prop = GetDb_MultifamilyRentComp.GetFirstRecordOrNothing
    
    Debug.Print "------- UNITS @ " & prop.name & " ------------"
    
    For Each unit In prop.associated_units
        Debug.Print unit.beds & "BR / " & unit.baths & "BA"
    Next unit
End Sub



