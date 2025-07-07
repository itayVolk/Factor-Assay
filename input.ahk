#Requires AutoHotkey v2.0

win := Gui(, "Factor Assay Input")
win.SetFont("S10")
win.AddText(,"Nom")
win.AddEdit("yp vLast")
win.AddText("yp","Prénom")
win.AddEdit("yp vFirst")
win.AddText("xm0","Age")
win.AddEdit("yp vAge")
win.AddText("yp","No. de dossier")
win.AddEdit("yp vDossier")
win.AddText("yp","Service")
win.AddEdit("yp vService")
win.AddText("yp","ID Sample")
win.AddEdit("yp vID")
win.AddText("xm0","Technicien")
win.AddEdit("yp vTechnician")
win.AddText("yp","Date")
win.AddDateTime("yp vDate")

win.AddText("yp+100 xm0","Actin FSL No. Lot")
win.AddEdit("yp vFSL")
win.AddText("yp x400","Date d'exp.")
win.AddDateTime("yp vEDFSL")
win.AddText("xm0","Standard Plasma No. Lot")
win.AddEdit("yp vSP")
win.AddText("yp x400","Date d'exp.")
win.AddDateTime("yp vEDSP")
win.AddText("xm0","Factor Def. Plasma No. Lot")
win.AddEdit("yp vFDP")
win.AddText("yp x400","Date d'exp.")
win.AddDateTime("yp vEDFDP")
win.AddText("xm0","Valeur standard du facteur")
win.AddEdit("yp vFAV")
win.AddText("yp", "%")

win.SetFont("S10 bold")
win.AddText("xm0 yp+100", "Factor Select")
win.AddDropDownList("yp vFactor", ["VIII", "IX"])

win.SetFont("S10 bold")
win.AddText("yp+100 xm0 ","Données plasma standards")
win.SetFont("S10 norm")
win.AddText("xm0", "Rep. 1")
win.AddEdit("yp x250 w50 vSP1_10")
win.AddEdit("yp x350 w50 vSP1_20")
win.AddEdit("yp x450 w50 vSP1_40")
win.AddText("xm0", "Rep. 2")
win.AddEdit("yp x250 w50 vSP2_10")
win.AddEdit("yp x350 w50 vSP2_20")
win.AddEdit("yp x450 w50 vSP2_40")

win.SetFont("S10 bold")
win.AddText("yp+100 xm0 ","Données Patient/QC")
win.SetFont("S10 norm")
win.AddText("yp x250", "1:10")
win.AddText("yp x350", "1:20")
win.AddText("yp x450", "1:40")
win.AddText("xm0", "Rep. 1")
win.AddEdit("yp x250 w50 vP1_10")
win.AddEdit("yp x350 w50 vP1_20")
win.AddEdit("yp x450 w50 vP1_40")
win.AddText("xm0", "Rep. 2")
win.AddEdit("yp x250 w50 vP2_10")
win.AddEdit("yp x350 w50 vP2_20")
win.AddEdit("yp x450 w50 vP2_40")

win.SetFont("S10 bold")
win.AddButton("xm0 yp+100", "Sauvegarder").OnEvent("Click", store)
win.AddButton("yp", "Rapport du labo").OnEvent("Click", lab)
win.AddButton("yp", "Rapport du patient").OnEvent("Click", patient)
win.AddButton("yp", "Effacer Patient").OnEvent("Click", pClear)
win.AddButton("yp", "Tout Effacer").OnEvent("Click", clear)

win.Show("Center AutoSize")

pClear(control, *) {
    static fields := ["Last", "First", "Age", "Dossier", "Service", "ID"]
    for field in fields {
        control.Gui[field].Value := ""
    }
}

clear(control, *) {
    for field in control.Gui {
        if field.Type == "DateTime" {
            field.Value := A_Now
        } else if field.Name {
            field.Value := ""
        }
    }
}

calculate(control) {
    static default := {Last: "Test", First: "Test", Age: 55, Dossier: "Test", Service: "Test", ID: "SAC2", Technician: "Bob", FSL: 3333, SP: 4444, FDP: 5555, FAV: 0.26,
                        Factor: 2, SP1_10: 80.8, SP1_20: 97.5, SP1_40: 117.3, SP2_10: 82.8, SP2_20: 102.3, SP2_40: 118.3,
                        P1_10: 130.2, P1_20:123.5, P1_40: 139.7, P2_10: 110.5, P2_20: 128.5, P2_40: 139.6}
    static dates := ["EDFSL", "EDSP", "EDFDP", "Date"]
    static seconds := ["P_avg_10", "P_avg_20", "P_avg_40", "SP_avg_10", "SP_avg_20", "SP_avg_40"]
    static percentages := ["res_10", "res_20", "res_40", "avg", "sd", "cv", "FAV"]
    
    sub := control.Gui.Submit(false)

    for key, val in default.OwnProps() {
        if !sub.%key% {
            if val is Number {
                if val == Round(val) {
                    control.Gui[key].Value :=  Format("{:.0f}", val)
                } else if val == Round(val, 1){
                    control.Gui[key].Value :=  Format("{:.1f}", val)
                } else {
                    control.Gui[key].Value :=  Format("{:.2f}", val)
                }
            } else {
                control.Gui[key].Value := val
            }

            if control.Gui[key].Type == "DropDownList" {
                sub.%key% := control.Gui[key].Text
            } else {
               sub.%key% := val
            }
        }
    }

    for field in dates{
        sub.%field% := FormatTime(sub.%field%, "MM dd yyyy")
    }

    sub.P_avg_10 := (sub.P1_10+sub.P2_10)/2
    sub.P_avg_20 := (sub.P1_20+sub.P2_20)/2
    sub.P_avg_40 := (sub.P1_40+sub.P2_40)/2

    sub.SP_avg_10 := (sub.SP1_10+sub.SP2_10)/2
    sub.SP_avg_20 := (sub.SP1_20+sub.SP2_20)/2
    sub.SP_avg_40 := (sub.SP1_40+sub.SP2_40)/2

    m := (3*(sub.SP_avg_10*Ln(100)+sub.SP_avg_20*Ln(50)+sub.SP_avg_40*Ln(25)) - (sub.SP_avg_10+sub.SP_avg_20+sub.SP_avg_40)*(Ln(100)+Ln(50)+Ln(25)))
        /(3*(Ln(100)**2+Ln(50)**2+Ln(25)**2)-(Ln(100)+Ln(50)+Ln(25))**2)
    b := (sub.SP_avg_10+sub.SP_avg_20+sub.SP_avg_40-m*(Ln(100)+Ln(50)+Ln(25)))/3

    sub.res_10 := Exp((sub.P_avg_10-b)/m)*sub.FAV
    sub.res_20 := Exp((sub.P_avg_20-b)/m)*sub.FAV
    sub.res_40 := Exp((sub.P_avg_40-b)/m)*sub.FAV

    sub.avg := (sub.res_10+sub.res_20+sub.res_40)/3
    sub.sd := Sqrt(((sub.res_10-sub.avg)**2+(sub.res_20-sub.avg)**2+(sub.res_40-sub.avg)**2)/3)
    sub.cv := sub.sd/sub.avg*100

    for field in seconds {
        sub.%field% := Format("{:.1f}", sub.%field%)
    }

    for field in percentages {
        sub.%field% := Format("{:.0f}", sub.%field%)
    }
    return sub
}

store(control, *) {
    sub := calculate(control)
    name := FileSelect(8,,,"CSV File (*.csv)")
    if !name {
        return
    }
    if SubStr(name, -4) != ".csv" {
        name .= ".csv"
    }
    csv := FileOpen(name, "rw")
    if(csv.AtEOF) {
        csv.Write("Nom,Prénom,No. de dossier,Actin FSL No. Lot,Date d'exp.,Standard Plasma No. Lot,Date d'exp.,Factor Def. Plasma No. Lot,Date d'exp.,Résultat,Date,Technicien`n")
    }
    csv.Seek(0,2)
    csv.Write(sub.Last "," sub.First "," sub.Dossier "," sub.FSL "," sub.EDFSL "," sub.SP "," sub.EDSP "," sub.FDP "," sub.EDFDP "," sub.avg "," sub.Date "," sub.Technician "`n")
    csv.Close()
    MsgBox("Data stored",,"T5")
}

lab(control, *) {
    sub := calculate(control)
    name := FileSelect(16,,,"PDF File (*.pdf)")
    if !name {
        return
    }
    if SubStr(name, -4) != ".pdf" {
        name .= ".pdf"
    }
    file := FileOpen(name, "w")
    template := FileOpen("Lab-template.html", "r").Read()
    static vars := ["Factor", "Last", "First", "Age", "Dossier", "Service", "P_avg_10", "P_avg_20", "P_avg_40", "SP_avg_10", "SP_avg_20", "SP_avg_40", "res_10", "res_20", "res_40", "avg", "sd", "cv", "Technician", "Date", "ID", "SP", "EDSP", "FAV"]
    for var in vars {
        template := StrReplace(template, "<" var ">", sub.%var%)
    }
    temp := FileOpen("temp.html", "w")
    temp.Write(template)
    temp.Close()
    RunWait(A_ComSpec ' /c ".\wkhtmltopdf.exe temp.html "' name '""',, "Hide")
    FileDelete("temp.html")
}

patient(control, *) {
    sub := calculate(control)
    name := FileSelect(16,,,"PDF File (*.pdf)")
    if !name {
        return
    }
    if SubStr(name, -4) != ".pdf" {
        name .= ".pdf"
    }
    file := FileOpen(name, "w")
    template := FileOpen("Patient-template.html", "r").Read()
    static vars := ["Factor", "Last", "First", "Age", "Dossier", "Service", "avg", "Technician", "Date"]
    for var in vars {
        template := StrReplace(template, "<" var ">", sub.%var%)
    }
    temp := FileOpen("temp.html", "w")
    temp.Write(template)
    temp.Close()
    RunWait(A_ComSpec ' /c ".\wkhtmltopdf.exe temp.html "' name '""',, "Hide")
    FileDelete("temp.html")
}