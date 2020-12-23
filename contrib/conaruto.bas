REM  *****  BASIC  *****

sub Main
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------

Dim arraySize
arraySize = 10
Dim wordsArray(arraySize,2) As String
wordsArray(0,0) = "Voie du ninja medecin"
wordsArray(0,1) = "Voie du ninja mdecin"
wordsArray(1,0) = "Voie du clan Hyuga"
wordsArray(1,1) = "Voie du clan Hyga"
wordsArray(2,0) = "Voie du jinchuriki"
wordsArray(2,1) = "Voie du jinchriki"
wordsArray(3,0) = "Voie des huit portes celestes"
wordsArray(3,1) = "Voie des huit portes clestes"
wordsArray(4,0) = "Voies et capacites"
wordsArray(4,1) = "Voies et capacits"
wordsArray(5,0) = "Competences"
wordsArray(5,1) = "Comptences"
wordsArray(6,0) = "Des de vie"
wordsArray(6,1) = "Ds de vie"
wordsArray(7,0) = "Systeme monetaire"
wordsArray(7,1) = "Systme montaire"
wordsArray(8,0) = "Acheter de l equipement"
wordsArray(8,1) = "Acheter de l'quipement"
wordsArray(9,0) = "Materiels"
wordsArray(9,1) = "Matriels"
wordsArray(10,0) = "Comment se proteger"
wordsArray(10,1) = "Comment se protger"

dim args1(21) as new com.sun.star.beans.PropertyValue
dim i
For i = 0 To arraySize 
args1(0).Name = "SearchItem.StyleFamily"
args1(0).Value = 2
args1(1).Name = "SearchItem.CellType"
args1(1).Value = 0
args1(2).Name = "SearchItem.RowDirection"
args1(2).Value = true
args1(3).Name = "SearchItem.AllTables"
args1(3).Value = false
args1(4).Name = "SearchItem.SearchFiltered"
args1(4).Value = false
args1(5).Name = "SearchItem.Backward"
args1(5).Value = false
args1(6).Name = "SearchItem.Pattern"
args1(6).Value = false
args1(7).Name = "SearchItem.Content"
args1(7).Value = false
args1(8).Name = "SearchItem.AsianOptions"
args1(8).Value = false
args1(9).Name = "SearchItem.AlgorithmType"
args1(9).Value = 0
args1(10).Name = "SearchItem.SearchFlags"
args1(10).Value = 65552
args1(11).Name = "SearchItem.SearchString"
args1(11).Value = wordsArray(i,0)
args1(12).Name = "SearchItem.ReplaceString"
args1(12).Value = wordsArray(i,1)
args1(13).Name = "SearchItem.Locale"
args1(13).Value = 255
args1(14).Name = "SearchItem.ChangedChars"
args1(14).Value = 2
args1(15).Name = "SearchItem.DeletedChars"
args1(15).Value = 2
args1(16).Name = "SearchItem.InsertedChars"
args1(16).Value = 2
args1(17).Name = "SearchItem.TransliterateFlags"
args1(17).Value = 1280
args1(18).Name = "SearchItem.Command"
args1(18).Value = 3
args1(19).Name = "SearchItem.SearchFormatted"
args1(19).Value = false
args1(20).Name = "SearchItem.AlgorithmType2"
args1(20).Value = 1
args1(21).Name = "Quiet"
args1(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())
Next i
End Sub
