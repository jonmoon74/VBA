VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dataform 
   Caption         =   "Data Administration"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
   OleObjectBlob   =   "dataform2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dataform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancelbutton_Click()
Unload Me
End Sub


Private Sub okbutton_Click()
Dim datee As String, datef As String, dateg As String, dateh As String, datei As String, datej As String
Dim datek As String, datel As String, datem As String, daten As String, dateo As String, datep As String
Dim dateq As String, dater As String, dates As String, datet As String, dateu As String, datev As String
Dim datew As String, datex As String, datey As String, datez As String, dateaa As String, dateab As String
Dim dateac As String, datead As String, dateae As String, dateaf As String, dateag As String, dateah As String
Dim dateai As String, dateaj As String, dateak As String, dateal As String, dateam As String, datean As String
Dim dateao As String, dateap As String, a As Integer, fname As String, sname As String, role As String, sdate As Date

a = ActiveCell.Row

Cells(a, 1).Value = dataform.fname.Value
Cells(a, 2).Value = dataform.sname.Value
Cells(a, 3).Value = dataform.role.Value
Cells(a, 4).Value = dataform.sdate.Value
Cells(a, 4) = Format(Cells(a, 4), "dd/mm/yy")
Cells(a, 5).Value = dataform.datee.Value
Cells(a, 5) = Format(Cells(a, 5), "dd/mm/yy")
Cells(a, 6).Value = dataform.datef.Value
Cells(a, 6) = Format(Cells(a, 6), "dd/mm/yy")
Cells(a, 7).Value = dataform.dateg.Value
Cells(a, 7) = Format(Cells(a, 7), "dd/mm/yy")
Cells(a, 8).Value = dataform.dateh.Value
Cells(a, 8) = Format(Cells(a, 8), "dd/mm/yy")
Cells(a, 9).Value = dataform.datei.Value
Cells(a, 9) = Format(Cells(a, 9), "dd/mm/yy")
Cells(a, 10).Value = dataform.datej.Value
Cells(a, 10) = Format(Cells(a, 10), "dd/mm/yy")
Cells(a, 11).Value = dataform.datek.Value
Cells(a, 11) = Format(Cells(a, 11), "dd/mm/yy")
Cells(a, 12).Value = dataform.datel.Value
Cells(a, 12) = Format(Cells(a, 12), "dd/mm/yy")
Cells(a, 13).Value = dataform.datem.Value
Cells(a, 13) = Format(Cells(a, 13), "dd/mm/yy")
Cells(a, 14).Value = dataform.daten.Value
Cells(a, 14) = Format(Cells(a, 14), "dd/mm/yy")
Cells(a, 15).Value = dataform.dateo.Value
Cells(a, 15) = Format(Cells(a, 15), "dd/mm/yy")
Cells(a, 16).Value = dataform.datep.Value
Cells(a, 16) = Format(Cells(a, 16), "dd/mm/yy")
Cells(a, 17).Value = dataform.dateq.Value
Cells(a, 17) = Format(Cells(a, 17), "dd/mm/yy")
Cells(a, 18).Value = dataform.dater.Value
Cells(a, 18) = Format(Cells(a, 18), "dd/mm/yy")
Cells(a, 19).Value = dataform.dates.Value
Cells(a, 19) = Format(Cells(a, 19), "dd/mm/yy")
Cells(a, 20).Value = dataform.datet.Value
Cells(a, 20) = Format(Cells(a, 20), "dd/mm/yy")
Cells(a, 21).Value = dataform.dateu.Value
Cells(a, 21) = Format(Cells(a, 21), "dd/mm/yy")
Cells(a, 22).Value = dataform.datev.Value
Cells(a, 22) = Format(Cells(a, 22), "dd/mm/yy")
Cells(a, 23).Value = dataform.datew.Value
Cells(a, 23) = Format(Cells(a, 23), "dd/mm/yy")
Cells(a, 24).Value = dataform.datex.Value
Cells(a, 24) = Format(Cells(a, 24), "dd/mm/yy")
Cells(a, 25).Value = dataform.datey.Value
Cells(a, 25) = Format(Cells(a, 25), "dd/mm/yy")
Cells(a, 26).Value = dataform.datez.Value
Cells(a, 26) = Format(Cells(a, 26), "dd/mm/yy")
Cells(a, 27).Value = dataform.dateaa.Value
Cells(a, 27) = Format(Cells(a, 27), "dd/mm/yy")
Cells(a, 28).Value = dataform.dateab.Value
Cells(a, 28) = Format(Cells(a, 28), "dd/mm/yy")
Cells(a, 29).Value = dataform.dateac.Value
Cells(a, 29) = Format(Cells(a, 29), "dd/mm/yy")
Cells(a, 30).Value = dataform.datead.Value
Cells(a, 30) = Format(Cells(a, 30), "dd/mm/yy")
Cells(a, 31).Value = dataform.dateae.Value
Cells(a, 31) = Format(Cells(a, 31), "dd/mm/yy")
Cells(a, 32).Value = dataform.dateaf.Value
Cells(a, 32) = Format(Cells(a, 32), "dd/mm/yy")
Cells(a, 33).Value = dataform.dateag.Value
Cells(a, 33) = Format(Cells(a, 33), "dd/mm/yy")
Cells(a, 34).Value = dataform.dateah.Value
Cells(a, 34) = Format(Cells(a, 34), "dd/mm/yy")
Cells(a, 35).Value = dataform.dateai.Value
Cells(a, 35) = Format(Cells(a, 35), "dd/mm/yy")
Cells(a, 36).Value = dataform.dateaj.Value
Cells(a, 36) = Format(Cells(a, 36), "dd/mm/yy")
Cells(a, 37).Value = dataform.dateak.Value
Cells(a, 37) = Format(Cells(a, 37), "dd/mm/yy")
Cells(a, 38).Value = dataform.dateal.Value
Cells(a, 38) = Format(Cells(a, 38), "dd/mm/yy")
Cells(a, 39).Value = dataform.dateam.Value
Cells(a, 39) = Format(Cells(a, 39), "dd/mm/yy")
Cells(a, 40).Value = dataform.datean.Value
Cells(a, 40) = Format(Cells(a, 40), "dd/mm/yy")
Cells(a, 41).Value = dataform.dateao.Value
Cells(a, 41) = Format(Cells(a, 41), "dd/mm/yy")
Cells(a, 42).Value = dataform.dateap.Value
Cells(a, 42) = Format(Cells(a, 42), "dd/mm/yy")

Unload Me

End Sub

Private Sub searchbutton_Click()
FormatDateTime (Date = [vbShortDate])

Dim datee As String, datef As String, dateg As String, dateh As String, datei As String, datej As String
Dim datek As String, datel As String, datem As String, daten As String, dateo As String, datep As String
Dim dateq As String, dater As String, dates As String, datet As String, dateu As String, datev As String
Dim datew As String, datex As String, datey As String, datez As String, dateaa As String, dateab As String
Dim dateac As String, datead As String, dateae As String, dateaf As String, dateag As String, dateah As String
Dim dateai As String, dateaj As String, dateak As String, dateal As String, dateam As String, datean As String
Dim dateao As String, dateap As String, a As Integer, fname As String, sname As String, role As String, sdate As Date

On Error Resume Next

ActiveSheet.Cells.Find(what:=dataform.searchbox.Value, after:=ActiveCell, LookIn:=xlValues, lookat:=xlPart, _
searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False, searchformat:=False).Activate

a = ActiveCell.Row

dataform.fname.Value = Cells(a, 1).Value
dataform.sname.Value = Cells(a, 2).Value
dataform.role.Value = Cells(a, 3).Value
dataform.sdate.Value = Cells(a, 4).Value
dataform.sdate = Format(dataform.sdate, "dd/mm/yy")
dataform.datee.Value = Cells(a, 5).Value
dataform.datee = Format(dataform.datee, "dd/mm/yy")
dataform.datef.Value = Cells(a, 6).Value
dataform.datef = Format(dataform.datef, "dd/mm/yy")
dataform.dateg.Value = Cells(a, 7).Value
dataform.dateg = Format(dataform.dateg, "dd/mm/yy")
dataform.dateh.Value = Cells(a, 8).Value
dataform.dateh = Format(dataform.dateh, "dd/mm/yy")
dataform.datei.Value = Cells(a, 9).Value
dataform.datei = Format(dataform.datei, "dd/mm/yy")
dataform.datej.Value = Cells(a, 10).Value
dataform.datej = Format(dataform.datej, "dd/mm/yy")
dataform.datek.Value = Cells(a, 11).Value
dataform.datek = Format(dataform.datek, "dd/mm/yy")
dataform.datel.Value = Cells(a, 12).Value
dataform.datel = Format(dataform.datel, "dd/mm/yy")
dataform.datem.Value = Cells(a, 13).Value
dataform.datem = Format(dataform.datem, "dd/mm/yy")
dataform.daten.Value = Cells(a, 14).Value
dataform.daten = Format(dataform.daten, "dd/mm/yy")
dataform.dateo.Value = Cells(a, 15).Value
dataform.dateo = Format(dataform.dateo, "dd/mm/yy")
dataform.datep.Value = Cells(a, 16).Value
dataform.datep = Format(dataform.datep, "dd/mm/yy")
dataform.dateq.Value = Cells(a, 17).Value
dataform.dateq = Format(dataform.dateq, "dd/mm/yy")
dataform.dater.Value = Cells(a, 18).Value
dataform.dater = Format(dataform.dater, "dd/mm/yy")
dataform.dates.Value = Cells(a, 19).Value
dataform.dates = Format(dataform.dates, "dd/mm/yy")
dataform.datet.Value = Cells(a, 20).Value
dataform.datet = Format(dataform.datet, "dd/mm/yy")
dataform.dateu.Value = Cells(a, 21).Value
dataform.dateu = Format(dataform.dateu, "dd/mm/yy")
dataform.datev.Value = Cells(a, 22).Value
dataform.datev = Format(dataform.datev, "dd/mm/yy")
dataform.datew.Value = Cells(a, 23).Value
dataform.datew = Format(dataform.datew, "dd/mm/yy")
dataform.datex.Value = Cells(a, 24).Value
dataform.datex = Format(dataform.datex, "dd/mm/yy")
dataform.datey.Value = Cells(a, 25).Value
dataform.datey = Format(dataform.datey, "dd/mm/yy")
dataform.datez.Value = Cells(a, 26).Value
dataform.datez = Format(dataform.datez, "dd/mm/yy")
dataform.dateaa.Value = Cells(a, 27).Value
dataform.dateaa = Format(dataform.dateaa, "dd/mm/yy")
dataform.dateab.Value = Cells(a, 28).Value
dataform.dateab = Format(dataform.dateab, "dd/mm/yy")
dataform.dateac.Value = Cells(a, 29).Value
dataform.dateac = Format(dataform.dateac, "dd/mm/yy")
dataform.datead.Value = Cells(a, 30).Value
dataform.datead = Format(dataform.datead, "dd/mm/yy")
dataform.dateae.Value = Cells(a, 31).Value
dataform.dateae = Format(dataform.dateae, "dd/mm/yy")
dataform.dateaf.Value = Cells(a, 32).Value
dataform.dateaf = Format(dataform.dateaf, "dd/mm/yy")
dataform.dateag.Value = Cells(a, 33).Value
dataform.dateag = Format(dataform.dateag, "dd/mm/yy")
dataform.dateah.Value = Cells(a, 34).Value
dataform.dateah = Format(dataform.dateah, "dd/mm/yy")
dataform.dateai.Value = Cells(a, 35).Value
dataform.dateai = Format(dataform.dateai, "dd/mm/yy")
dataform.dateaj.Value = Cells(a, 36).Value
dataform.dateaj = Format(dataform.dateaj, "dd/mm/yy")
dataform.dateak.Value = Cells(a, 37).Value
dataform.dateak = Format(dataform.dateak, "dd/mm/yy")
dataform.dateal.Value = Cells(a, 38).Value
dataform.dateal = Format(dataform.dateal, "dd/mm/yy")
dataform.dateam.Value = Cells(a, 39).Value
dataform.dateam = Format(dataform.dateam, "dd/mm/yy")
dataform.datean.Value = Cells(a, 40).Value
dataform.datean = Format(dataform.datean, "dd/mm/yy")
dataform.dateao.Value = Cells(a, 41).Value
dataform.dateao = Format(dataform.dateao, "dd/mm/yy")
dataform.dateap.Value = Cells(a, 42).Value
dataform.dateap = Format(dataform.dateap, "dd/mm/yy")

End Sub

Private Sub submitbutton_Click()
Dim datee As String, datef As String, dateg As String, dateh As String, datei As String, datej As String
Dim datek As String, datel As String, datem As String, daten As String, dateo As String, datep As String
Dim dateq As String, dater As String, dates As String, datet As String, dateu As String, datev As String
Dim datew As String, datex As String, datey As String, datez As String, dateaa As String, dateab As String
Dim dateac As String, datead As String, dateae As String, dateaf As String, dateag As String, dateah As String
Dim dateai As String, dateaj As String, dateak As String, dateal As String, dateam As String, datean As String
Dim dateao As String, dateap As String, a As Integer, fname As String, sname As String, role As String, sdate As Date

a = ActiveCell.Row

Cells(a, 1).Value = dataform.fname.Value
Cells(a, 2).Value = dataform.sname.Value
Cells(a, 3).Value = dataform.role.Value
Cells(a, 4).Value = dataform.sdate.Value
Cells(a, 4) = Format(Cells(a, 4), "dd/mm/yy")
Cells(a, 5).Value = dataform.datee.Value
Cells(a, 5) = Format(Cells(a, 5), "dd/mm/yy")
Cells(a, 6).Value = dataform.datef.Value
Cells(a, 6) = Format(Cells(a, 6), "dd/mm/yy")
Cells(a, 7).Value = dataform.dateg.Value
Cells(a, 7) = Format(Cells(a, 7), "dd/mm/yy")
Cells(a, 8).Value = dataform.dateh.Value
Cells(a, 8) = Format(Cells(a, 8), "dd/mm/yy")
Cells(a, 9).Value = dataform.datei.Value
Cells(a, 9) = Format(Cells(a, 9), "dd/mm/yy")
Cells(a, 10).Value = dataform.datej.Value
Cells(a, 10) = Format(Cells(a, 10), "dd/mm/yy")
Cells(a, 11).Value = dataform.datek.Value
Cells(a, 11) = Format(Cells(a, 11), "dd/mm/yy")
Cells(a, 12).Value = dataform.datel.Value
Cells(a, 12) = Format(Cells(a, 12), "dd/mm/yy")
Cells(a, 13).Value = dataform.datem.Value
Cells(a, 13) = Format(Cells(a, 13), "dd/mm/yy")
Cells(a, 14).Value = dataform.daten.Value
Cells(a, 14) = Format(Cells(a, 14), "dd/mm/yy")
Cells(a, 15).Value = dataform.dateo.Value
Cells(a, 15) = Format(Cells(a, 15), "dd/mm/yy")
Cells(a, 16).Value = dataform.datep.Value
Cells(a, 16) = Format(Cells(a, 16), "dd/mm/yy")
Cells(a, 17).Value = dataform.dateq.Value
Cells(a, 17) = Format(Cells(a, 17), "dd/mm/yy")
Cells(a, 18).Value = dataform.dater.Value
Cells(a, 18) = Format(Cells(a, 18), "dd/mm/yy")
Cells(a, 19).Value = dataform.dates.Value
Cells(a, 19) = Format(Cells(a, 19), "dd/mm/yy")
Cells(a, 20).Value = dataform.datet.Value
Cells(a, 20) = Format(Cells(a, 20), "dd/mm/yy")
Cells(a, 21).Value = dataform.dateu.Value
Cells(a, 21) = Format(Cells(a, 21), "dd/mm/yy")
Cells(a, 22).Value = dataform.datev.Value
Cells(a, 22) = Format(Cells(a, 22), "dd/mm/yy")
Cells(a, 23).Value = dataform.datew.Value
Cells(a, 23) = Format(Cells(a, 23), "dd/mm/yy")
Cells(a, 24).Value = dataform.datex.Value
Cells(a, 24) = Format(Cells(a, 24), "dd/mm/yy")
Cells(a, 25).Value = dataform.datey.Value
Cells(a, 25) = Format(Cells(a, 25), "dd/mm/yy")
Cells(a, 26).Value = dataform.datez.Value
Cells(a, 26) = Format(Cells(a, 26), "dd/mm/yy")
Cells(a, 27).Value = dataform.dateaa.Value
Cells(a, 27) = Format(Cells(a, 27), "dd/mm/yy")
Cells(a, 28).Value = dataform.dateab.Value
Cells(a, 28) = Format(Cells(a, 28), "dd/mm/yy")
Cells(a, 29).Value = dataform.dateac.Value
Cells(a, 29) = Format(Cells(a, 29), "dd/mm/yy")
Cells(a, 30).Value = dataform.datead.Value
Cells(a, 30) = Format(Cells(a, 30), "dd/mm/yy")
Cells(a, 31).Value = dataform.dateae.Value
Cells(a, 31) = Format(Cells(a, 31), "dd/mm/yy")
Cells(a, 32).Value = dataform.dateaf.Value
Cells(a, 32) = Format(Cells(a, 32), "dd/mm/yy")
Cells(a, 33).Value = dataform.dateag.Value
Cells(a, 33) = Format(Cells(a, 33), "dd/mm/yy")
Cells(a, 34).Value = dataform.dateah.Value
Cells(a, 34) = Format(Cells(a, 34), "dd/mm/yy")
Cells(a, 35).Value = dataform.dateai.Value
Cells(a, 35) = Format(Cells(a, 35), "dd/mm/yy")
Cells(a, 36).Value = dataform.dateaj.Value
Cells(a, 36) = Format(Cells(a, 36), "dd/mm/yy")
Cells(a, 37).Value = dataform.dateak.Value
Cells(a, 37) = Format(Cells(a, 37), "dd/mm/yy")
Cells(a, 38).Value = dataform.dateal.Value
Cells(a, 38) = Format(Cells(a, 38), "dd/mm/yy")
Cells(a, 39).Value = dataform.dateam.Value
Cells(a, 39) = Format(Cells(a, 39), "dd/mm/yy")
Cells(a, 40).Value = dataform.datean.Value
Cells(a, 40) = Format(Cells(a, 40), "dd/mm/yy")
Cells(a, 41).Value = dataform.dateao.Value
Cells(a, 41) = Format(Cells(a, 41), "dd/mm/yy")
Cells(a, 42).Value = dataform.dateap.Value
Cells(a, 42) = Format(Cells(a, 42), "dd/mm/yy")

End Sub

Private Sub userform_Initialize()
Dim namee As String, namef As String, nameg As String, nameh As String, namei As String, namej As String
Dim namek As String, namel As String, namem As String, namen As String, nameo As String, namep As String
Dim nameq As String, namer As String, names As String, namet As String, nameu As String, namev As String
Dim namew As String, namex As String, namey As String, namez As String, nameaa As String, nameab As String
Dim nameac As String, namead As String, nameae As String, nameaf As String, nameag As String, nameah As String
Dim nameai As String, nameaj As String, nameak As String, nameal As String, nameam As String, namean As String
Dim nameao As String, nameap As String

dataform.namee.Value = Cells(2, 5).Value
dataform.namef.Value = Cells(2, 6).Value
dataform.nameg.Value = Cells(2, 7).Value
dataform.nameh.Value = Cells(2, 8).Value
dataform.namei.Value = Cells(2, 9).Value
dataform.namej.Value = Cells(2, 10).Value
dataform.namek.Value = Cells(2, 11).Value
dataform.namel.Value = Cells(2, 12).Value
dataform.namem.Value = Cells(2, 13).Value
dataform.namen.Value = Cells(2, 14).Value
dataform.nameo.Value = Cells(2, 15).Value
dataform.namep.Value = Cells(2, 16).Value
dataform.nameq.Value = Cells(2, 17).Value
dataform.namer.Value = Cells(2, 18).Value
dataform.names.Value = Cells(2, 19).Value
dataform.namet.Value = Cells(2, 20).Value
dataform.nameu.Value = Cells(2, 21).Value
dataform.namev.Value = Cells(2, 22).Value
dataform.namew.Value = Cells(2, 23).Value
dataform.namex.Value = Cells(2, 24).Value
dataform.namey.Value = Cells(2, 25).Value
dataform.namez.Value = Cells(2, 26).Value
dataform.nameaa.Value = Cells(2, 27).Value
dataform.nameab.Value = Cells(2, 28).Value
dataform.nameac.Value = Cells(2, 29).Value
dataform.namead.Value = Cells(2, 30).Value
dataform.nameae.Value = Cells(2, 31).Value
dataform.nameaf.Value = Cells(2, 32).Value
dataform.nameag.Value = Cells(2, 33).Value
dataform.nameah.Value = Cells(2, 34).Value
dataform.nameai.Value = Cells(2, 35).Value
dataform.nameaj.Value = Cells(2, 36).Value
dataform.nameak.Value = Cells(2, 37).Value
dataform.nameal.Value = Cells(2, 38).Value
dataform.nameam.Value = Cells(2, 39).Value
dataform.namean.Value = Cells(2, 40).Value
dataform.nameao.Value = Cells(2, 41).Value
dataform.nameap.Value = Cells(2, 42).Value

End Sub


