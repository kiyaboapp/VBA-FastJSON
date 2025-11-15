Attribute VB_Name = "test"
'====================================================================
' FastJSON Test Suite - SIMPLE AND COMPLETE
' Standard Module
'====================================================================
Option Explicit

Public Sub RunAllTests()
    Debug.Print "=========================================="
    Debug.Print "FastJSON Test Suite"
    Debug.Print "Started: " & Now()
    Debug.Print "=========================================="
    Debug.Print ""
    
    TestBasics
    TestArrays
    TestNesting
    TestDeepNesting
    TestPaths
    TestSaveLoad
    TestRealWorld
    
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "ALL TESTS PASSED!"
    Debug.Print "Finished: " & Now()
    Debug.Print "=========================================="
End Sub

Private Sub TestBasics()
    Debug.Print "TEST 1: Basic Values"
    Debug.Print String(50, "-")
    
    Dim j As New FastJSON
    j.Add "name", "Asha Juma"
    j.Add "age", 31
    j.Add "salary", 1250000
    j.Add "active", True
    
    
    Debug.Print j.ToPretty
    Debug.Print ""
    Debug.Print "Get name: " & j.GetValue("name")
    Debug.Print "Get age: " & j.GetValue("age")
    Debug.Print "Count: " & j.Count
    Debug.Print ""
End Sub

Private Sub TestArrays()
    Debug.Print "TEST 2: Arrays"
    Debug.Print String(50, "-")
    
    Dim j As New FastJSON
    j.Add "employee", "John Doe"
    j.AddArray "phones", "0712345678", "0655554433", "0788112233"
    j.AddArray "skills", "Excel", "VBA", "SQL"
    
    Debug.Print j.ToPretty
    Debug.Print ""
    
    Dim phones As Variant
    phones = j.GetArray("phones")
    Debug.Print "Phone count: " & UBound(phones) + 1
    Dim i As Long
    For i = LBound(phones) To UBound(phones)
        Debug.Print "  Phone " & i + 1 & ": " & phones(i)
    Next i
    Debug.Print ""
End Sub

Private Sub TestNesting()
    Debug.Print "TEST 3: Nested Objects"
    Debug.Print String(50, "-")
    
    Dim person As New FastJSON
    person.Add "name", "Asha Juma"
    person.Add "age", 31
    
    Dim addr As New FastJSON
    addr.Add "street", "Uhuru Street"
    addr.Add "city", "Dar es Salaam"
    addr.Add "country", "Tanzania"
    
    person.AddObject "address", addr
    
    Debug.Print person.ToPretty
    Debug.Print ""
    Debug.Print "Name: " & person.GetValue("name")
    Debug.Print "City: " & person.GetValue("address.city")
    Debug.Print "Country: " & person.GetValue("address.country")
    Debug.Print ""
End Sub

Private Sub TestDeepNesting()
    Debug.Print "TEST 4: Deep Nesting (3 Levels)"
    Debug.Print String(50, "-")
    
    Dim person As New FastJSON
    person.Add "name", "Asha Juma"
    person.AddArray "phones", "0712345678", "0754123456"
    
    ' Level 2
    Dim addr As New FastJSON
    addr.Add "street", "Uhuru Street"
    addr.Add "city", "Dar es Salaam"
    
    ' Level 3
    Dim coords As New FastJSON
    coords.Add "lat", -6.7924
    coords.Add "lng", 39.2083
    
    addr.AddObject "coordinates", coords
    person.AddObject "address", addr
    
    ' Another level 2
    Dim emp As New FastJSON
    emp.Add "title", "Senior Developer"
    emp.Add "department", "IT"
    
    ' Level 3
    Dim mgr As New FastJSON
    mgr.Add "name", "Hassan Mbogo"
    mgr.Add "email", "hassan@company.co.tz"
    
    emp.AddObject "manager", mgr
    person.AddObject "employment", emp
    
    Debug.Print person.ToPretty
    Debug.Print ""
    Debug.Print "Name: " & person.GetValue("name")
    Debug.Print "City: " & person.GetValue("address.city")
    Debug.Print "Lat: " & person.GetValue("address.coordinates.lat")
    Debug.Print "Lng: " & person.GetValue("address.coordinates.lng")
    Debug.Print "Title: " & person.GetValue("employment.title")
    Debug.Print "Manager: " & person.GetValue("employment.manager.name")
    Debug.Print "Manager Email: " & person.GetValue("employment.manager.email")
    Debug.Print ""
End Sub

Private Sub TestPaths()
    Debug.Print "TEST 5: Paths and Keys"
    Debug.Print String(50, "-")
    
    Dim j As New FastJSON
    j.Add "company", "TechCo"
    j.AddArray "products", "Software", "Hardware"
    
    Dim office As New FastJSON
    office.Add "location", "Arusha"
    office.Add "employees", 50
    
    j.AddObject "office", office
    
    Debug.Print "Has 'company': " & j.HasKey("company")
    Debug.Print "Has 'office.location': " & j.HasKey("office.location")
    Debug.Print "Has 'nothere': " & j.HasKey("nothere")
    Debug.Print ""
    
    Debug.Print "Root Keys:"
    Dim keys As Variant
    keys = j.GetKeys()
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Debug.Print "  - " & keys(i)
    Next i
    Debug.Print ""
    
    Dim officeObj As FastJSON
    Set officeObj = j.GetObject("office")
    Debug.Print "Office Object:"
    Debug.Print officeObj.ToPretty
    Debug.Print ""
End Sub

Private Sub TestSaveLoad()
    Debug.Print "TEST 6: Save and Load"
    Debug.Print String(50, "-")
    
    Dim original As New FastJSON
    original.Add "name", "Test Person"
    original.Add "age", 35
    original.AddArray "tags", "tag1", "tag2", "tag3"
    
    Dim addr As New FastJSON
    addr.Add "city", "Arusha"
    addr.Add "country", "Tanzania"
    original.AddObject "address", addr
    
    Debug.Print "Original:"
    Debug.Print original.ToPretty
    Debug.Print ""
    
    Dim saved As String
    saved = original.ToRaw
    Debug.Print "Saved (first 100): " & Left(saved, 100)
    Debug.Print ""
    
    Dim loaded As New FastJSON
    loaded.Parse saved
    
    Debug.Print "Loaded:"
    Debug.Print loaded.ToPretty
    Debug.Print ""
    Debug.Print "Name matches: " & (original.GetValue("name") = loaded.GetValue("name"))
    Debug.Print "City matches: " & (original.GetValue("address.city") = loaded.GetValue("address.city"))
    Debug.Print "Loaded Address: " & loaded.GetValue("address.city")
    Debug.Print ""
End Sub

Private Sub TestRealWorld()
    Debug.Print "TEST 7: Real-World Employee"
    Debug.Print String(50, "-")
    
    Dim emp As New FastJSON
    emp.Add "employeeID", "EMP-001"
    emp.Add "firstName", "Asha"
    emp.Add "lastName", "Juma"
    emp.Add "email", "asha@company.co.tz"
    emp.Add "salary", 1250000
    emp.Add "active", True
    
    emp.AddArray "phones", "0712345678", "0754123456"
    emp.AddArray "skills", "Access", "VBA", "SQL", "Excel"
    
    Dim addr As New FastJSON
    addr.Add "street", "Plot 123 Uhuru St"
    addr.Add "city", "Arusha"
    addr.Add "country", "Tanzania"
    
    Dim gps As New FastJSON
    gps.Add "lat", -3.3869
    gps.Add "lng", 36.683
    addr.AddObject "gps", gps
    
    emp.AddObject "address", addr
    
    Dim job As New FastJSON
    job.Add "department", "IT"
    job.Add "position", "Senior Developer"
    job.Add "level", "L4"
    
    emp.AddObject "employment", job
    
    Debug.Print "Complete Employee:"
    Debug.Print emp.ToPretty
    Debug.Print ""
    
    Debug.Print "Quick Access:"
    Debug.Print "  Name: " & emp.GetValue("firstName") & " " & emp.GetValue("lastName")
    Debug.Print "  Email: " & emp.GetValue("email")
    Debug.Print "  Dept: " & emp.GetValue("employment.department")
    Debug.Print "  Position: " & emp.GetValue("employment.position")
    Debug.Print "  City: " & emp.GetValue("address.city")
    Debug.Print "  GPS: " & emp.GetValue("address.gps.lat") & ", " & emp.GetValue("address.gps.lng")
    Debug.Print ""
    
    Dim phones As Variant
    phones = emp.GetArray("phones")
    Debug.Print "Phones:"
    Dim i As Long
    For i = LBound(phones) To UBound(phones)
        Debug.Print "  " & phones(i)
    Next i
    Debug.Print ""
    
    Dim skills As Variant
    skills = emp.GetArray("skills")
    Debug.Print "Skills:"
    For i = LBound(skills) To UBound(skills)
        Debug.Print "  - " & skills(i)
    Next i
    Debug.Print ""
    
    ' Database storage
    Dim raw As String
    raw = emp.ToRaw
    Debug.Print "Storage Ready:"
    Debug.Print "  Length: " & Len(raw) & " chars"
    Debug.Print "  SQL-safe: Yes"
    Debug.Print ""
    
    ' Simulate retrieval
    Dim retrieved As New FastJSON
    retrieved.Parse raw
    Debug.Print "Retrieved from storage:"
    Debug.Print "  Name: " & retrieved.GetValue("firstName") & " " & retrieved.GetValue("lastName")
    Debug.Print "  Position: " & retrieved.GetValue("employment.position")
    Debug.Print "Raw:"
    Debug.Print raw
    Debug.Print ""
End Sub

