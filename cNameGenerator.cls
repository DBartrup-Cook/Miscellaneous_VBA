VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNameGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

'This class generates two arrays of the same size. These arrays contain randomly selected sets of first names & surnames.

'After initialising the class use these Public procedures and properties:
'GenerateNames(NumberToGenerate As Long) - Creates a random set of names. The number of sets created is controlled by
'                                          the NumberToGenerate parameter.
'ClearLists()                            - Clears the currently selected array of names.
'FirstNames()                            - Property returns the first name array.
'Surnames()                              - Property returns the surname array.

'200 first names x 200 surnames.
'An example of use in a normal module:
'-------------------------------------
'Public Sub ToUse()
'    Dim Test As cNameGenerator
'    Set Test = New cNameGenerator
'
'    With Test
'        .ClearLists
'        .GenerateNames NumberToGenerate:=4
'        Dim x As Long
'        For x = 0 To 3
'            Debug.Print x & ": "; .FirstNames(x) & " " & .Surnames(x)
'        Next x
'    End With
'
'End Sub

Private Type ClassVariables
    FirstNames As Variant
    Surnames As Variant
    SelectedFirstNames As Variant
    SelectedSurnames As Variant
End Type
Private This As ClassVariables

'When the class is created initialise the two variables containing 200 first and last names.
Private Sub Class_Initialize()
    This.FirstNames = Array("Marie", "Sophie", "Charlotte", "Alexander", "Noah", "David", "Emilia", "Leon", "Maria", "Maximilian", _
                            "Valentin", "Ella", "Lea", "Theodor", "Karl", "Emma", "Elias", "Paul", "Sophia", "Adam", "Anton", "Emil", _
                            "Jonas", "Anna", "Johanna", "Felix", "Liam", "Elisabeth", "Clara", "Mathilda", "Jakob", "Leo", "Lukas", _
                            "Mia", "Henry", "Frida", "Hannah", "Olivia", "Levi", "Luise", "Ben", "Gabriel", "Leonard", "Hanna", "Marie", _
                            "Mila", "Nora", "Paula", "Benjamin", "Alma", "Lara", "Leni", "Leonie", "Lina", "Max", "Helena", "Julius", _
                            "Lena", "Pauline", "Ali", "Johann", "Finn", "Louis", "Amelie", "Luca", "Maya", "Malik", "Milan", "Michael", _
                            "Emily", "Frieda", "Jonathan", "Linus", "Sofia", "Luisa", "Aaron", "Luis", "Ava", "Mateo", "Isabella", _
                            "Raphael", "Juna", "Samuel", "Liya", "Tim", "Lotta", "Marlene", "Mina", "Friedrich", "Moritz", "Viktor", _
                            "Vincent", "Ada", "Amira", "Flora", "Ida", "Lia", "Lilly", "Mara", "Thea", "Tilda", "Antonia", "Carlotta", _
                            "Helene", "Magdalena", "Rosa", "Victoria", "Arthur", "August", "Bruno", "Caspar", "Erik", "Joshua", "Kian", _
                            "Matteo", "Mika", "Nicolas", "Oskar", "Theo", "Peter", "Theresa", "Alina", "Amalia", "Carla", "Carl", "Cleo", _
                            "Daniel", "Frederik", "Julia", "Jan", "Karla", "Leopold", "Nisa", "Ferdinand", "Johannes", "Katharina", "Otto", _
                            "Lou", "Wilhelm", "Arian", "Aliya", "Eymen", "Elisa", "Gustav", "Klara", "Hamza", "Lana", "Hugo", "Luna", _
                            "Jasper", "Joris", "Matilda", "Julian", "Medina", "Justus", "Mira", "Lian", "Zoe", "Mert", "Niklas", "Noel", _
                            "Oliver", "Philipp", "Louisa", "Joel", "Viktoria", "Nael", "Robert", "Thomas", "Aaliyah", "Alva", "Amelia", _
                            "Anouk", "Elif", "Elina", "Ayaz", "Laura", "Elio", "Louise", "Fynn", "Lucy", "Henri", "Maja", "Ibrahim", _
                            "Martha", "Jona", "Merle", "Lasse", "Romy", "Luka", "Ruby", "Mattis", "Sarah", "Milo", "Simon", "Aurelia", _
                            "Georg", "Paulina", "Josef", "Can")
                       
    This.Surnames = Array("Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", _
                          "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson", _
                          "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson", "Walker", "Young", "Allen", "King", "Wright", _
                          "Scott", "Torres", "Nguyen", "Hill", "Flores", "Green", "Adams", "Nelson", "Baker", "Hall", "Rivera", "Campbell", _
                          "Mitchell", "Carter", "Roberts", "Gomez", "Phillips", "Evans", "Turner", "Diaz", "Parker", "Cruz", "Edwards", "Collins", _
                          "Reyes", "Stewart", "Morris", "Morales", "Murphy", "Cook", "Rogers", "Gutierrez", "Ortiz", "Morgan", "Cooper", _
                          "Peterson", "Bailey", "Reed", "Kelly", "Howard", "Ramos", "Kim", "Cox", "Ward", "Richardson", "Watson", "Brooks", _
                          "Chavez", "Wood", "James", "Bennett", "Gray", "Mendoza", "Ruiz", "Hughes", "Price", "Alvarez", "Castillo", "Sanders", _
                          "Patel", "Myers", "Long", "Ross", "Foster", "Jimenez", "Powell", "Jenkins", "Perry", "Russell", "Sullivan", "Bell", _
                          "Coleman", "Butler", "Henderson", "Barnes", "Gonzales", "Fisher", "Vasquez", "Simmons", "Romero", "Jordan", "Patterson", _
                          "Alexander", "Hamilton", "Graham", "Reynolds", "Griffin", "Wallace", "Moreno", "West", "Cole", "Hayes", "Bryant", _
                          "Herrera", "Gibson", "Ellis", "Tran", "Medina", "Aguilar", "Stevens", "Murray", "Ford", "Castro", "Marshall", "Owens", _
                          "Harrison", "Fernandez", "Mcdonald", "Woods", "Washington", "Kennedy", "Wells", "Vargas", "Henry", "Chen", "Freeman", _
                          "Webb", "Tucker", "Guzman", "Burns", "Crawford", "Olson", "Simpson", "Porter", "Hunter", "Gordon", "Mendez", "Silva", _
                          "Shaw", "Snyder", "Mason", "Dixon", "Munoz", "Hunt", "Hicks", "Holmes", "Palmer", "Wagner", "Black", "Robertson", _
                          "Boyd", "Rose", "Stone", "Salazar", "Fox", "Warren", "Mills", "Meyer", "Rice", "Schmidt", "Garza", "Daniels", "Ferguson", _
                          "Nichols", "Stephens", "Soto", "Weaver", "Ryan", "Gardner", "Payne", "Grant", "Dunn", "Kelley", "Spencer", "Hawkins")
End Sub

'The NumberToGenerate parameter indicates how many first name/surname combinations to generate.
'This number is passed to the SelectNames function along with either the This.FirstNames or the
'This.Surname array of names to create an array of the indicated part of the name.
'These arrays are then stored within the Private ClassVariables Type.
Public Sub GenerateNames(NumberToGenerate As Long)
    If NumberToGenerate > 0 Then
        This.SelectedFirstNames = SelectNames(NumberToGenerate, This.FirstNames)
        This.SelectedSurnames = SelectNames(NumberToGenerate, This.Surnames)
    End If
End Sub

'Clears the currently selected names allowing the class to generate a new set.
Public Sub ClearLists()
    This.SelectedFirstNames = Empty
    This.SelectedSurnames = Empty
End Sub

'These properties allow the generated first and surnames to be passed out of the class.
Public Property Get FirstNames() As Variant
    FirstNames = This.SelectedFirstNames
End Property

Public Property Get Surnames() As Variant
    Surnames = This.SelectedSurnames
End Property

'The NumberToGenerate paramater is passed to the GenerateRandomNumbers function to return an array of numbers.
'These numbers are then used to return items from either This.FirstNames or This.Surnames which is passed to
'the function through the UseList parameter.
'A comma delimited string is built from the items in the UseList parameter before being split into an array
'and returned by the function.
Private Function SelectNames(NumberToGenerate As Long, UseList As Variant)
    Dim itm As Variant
    Dim ReturnString As Variant
    For Each itm In GenerateRandomNumbers(NumberToGenerate, LBound(UseList), UBound(UseList))
        ReturnString = ReturnString & UseList(itm) & ","
    Next itm
    SelectNames = Split(Left(ReturnString, Len(ReturnString) - 1), ",")
End Function

'Generates an array of randomly generated numbers between the MinNum and MaxNum limits.
'The total numbers generated is controlled by the NumbersToGenerate parameter.
'An example of a string returned by the function is: "140,106,115"
Private Function GenerateRandomNumbers(Optional NumbersToGenerate As Long = 20, _
                                      Optional MinNum As Long = 1, _
                                      Optional MaxNum As Long = 200) As Variant
                                      
    Dim NumbersGenerated As Long
    Dim ThisNum As Long
    Dim ReturnString As String
    'Create random numbers in a loop.
    'Only add the number to the return string if it hasn't already been generated.
    'The final number will be followed by a comma so this is removed before the
    'text string is split into an array and returned by the function.
    Do
        ThisNum = Rnd * (MaxNum - MinNum) + MinNum
        If InStr(ReturnString, ThisNum & ",") = 0 Then
            ReturnString = ReturnString & ThisNum & ","
            NumbersGenerated = NumbersGenerated + 1
        End If
    Loop Until NumbersGenerated = NumbersToGenerate
    GenerateRandomNumbers = Split(Left(ReturnString, Len(ReturnString) - 1), ",")
End Function





