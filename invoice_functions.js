function setTotalsNew(rowChange) {
    /* 24.8.2016 Merging the previous separate functions setTotals and setTotalsRowChange.
    B uses net amounts on rows. Get gross total and tax amount from vendor line. Calculate item totals.
    B does not show domestic currency.
    Balance = item total + tax total - gross total
    What we get from ABAP:
    normal invoices, header GROSS_AMOUNT > 0. For credit memos, GROSS_AMOUNT < 0.
    Item WRBTR > 0 always, but KDIND is D or K. Normally invoices debit costs and taxes, credit vendor.
    For display, let's just put everything > 0 for invoices, < 0 for credit memos.
    */
    "use strict";
    var i, x, txt, wrbtr, totDoc, valIt, totIt, kd, valHTax, valHVal, valBal;

    i = tblItems.getMaxItemsCount();
    txt = "";
    totIt = 0;
    valHVal = 0;
    valHTax = 0;

    // Calculate item total in document currency (B only uses document currency)
    for (x = 0; x < i; x++) {
        wrbtr = modeltblItems.getData()[x].WRBTR;
        valIt = txtToVal(wrbtr);
            if (txtDocCur.getValue() === txtDomCur.getValue()) {
                modeltblItems.getData()[x].DMBTR = modeltblItems.getData()[x].WRBTR;
            }
            kd = tblItems.getItems()[x].getBindingContext().getProperty("KDIND");
            if (kd === "K") {
                valIt = valIt * -1;
            }
        totIt = totIt + valIt;
    }

    //Invoice total (=vendor total)
    txt = txtHeadValue.getValue();
    valHVal = txtToVal(txt);

    //Total tax amount (from vendor item)
    txt = txtHeadTax.getValue();
    valHTax = txtToVal(txt);
    if (valHVal < 0){
        //Credit memo
        valHTax = valHTax * -1;
    }

    //Calculate balance
    valBal = 0;
    valBal = totIt + valHTax - valHVal;
    if (valBal === 0) {
        txtBalance.removeStyleClass("balanceAlert");
        txtBalanceCur.removeStyleClass("balanceAlert");
    } else {
        txtBalance.addStyleClass("balanceAlert");
        txtBalanceCur.addStyleClass("balanceAlert");
    }
    if (rowChange) {
        modeltblItems.setData(modeltblItems.getData());
    }

    //Format amounts for display
    txt = valToTxt(valHVal);
    txtDocTot.setValue(txt);

    txt = valToTxt(totIt);
    txtItemValue.setValue(txt);

    txt = valToTxt(valHTax);
    txtHeadTax.setValue(txt);

    txt = valToTxt(valBal);
    txtBalance.setValue(txt);
}

function valToTxt(val) {
    var txt = "";
    if (typeof(val) == "string") {
        txt = val;
    } else {
        txt = txt + val.toFixed(2);
        txt = txt.replace(".", ",");
    }
    return txt;
}

function txtToVal(txt) {
    var val = 0;
    if (typeof(txt) == "string") {
        txt = txt.replace(",", ".");
        txt = txt.replace(/\s+/g, '');
        val = parseFloat(txt);
        if (isNaN(val)) {
            //    val = 0; Let's display the error!
        }
    } else {
        val = txt;
    }
    return val;
}

// 9.9.2016 Explicit conversion functions for Excel import/export etc.
function commaToPoint(val) {
    var txt = String(val);
    txt = txt.replace(",", ".");
    return txt;
}

function pointToComma(val) {
    var txt = String(val);
    txt = txt.replace(".", ",");
    return txt;
}

function exportToExcel() {
    var i = tblItems.getMaxItemsCount();
// Build header
    var str = "Account" + "\t";
    str = str + "Profit center" + "\t";
    str = str + "Cost center" + "\t";
    str = str + "Order" + "\t";
    str = str + "Project" + "\t"; 
    str = str + "Func.area" + "\t";
    str = str + "Tax" + "\t";
    str = str + "Value (doc.)" + "\t";
    str = str + "Value (dom.)" + "\t";
    str = str + "Quantity" + "\t";
    str = str + "Unit" + "\t";
    str = str + "Deb/Cre" + "\t";
    str = str + "Row text" + "\t";
    str = str + "Reviewer name" + "\t";
    str = str + "Approver" + "\t";
    str = str + "Approver name";
    str = str + "\n";
    for (x = 0; x < i; x++) {
// 1st-PKTILI, 2nd-PRCTR, ... 7th-WRBTR, 8-DMBTR, 9-MENGE, 10-MEINS, 11th-KDIND
        str = str + modeltblItems.getData()[x].PKTILI + "\t";
        str = str + modeltblItems.getData()[x].PRCTR + "\t";
        str = str + modeltblItems.getData()[x].KUSTPAIKKA + "\t";
        str = str + modeltblItems.getData()[x].ORDER + "\t";
        str = str + modeltblItems.getData()[x].FUNC_AREA + "\t";
        str = str + modeltblItems.getData()[x].VERO + "\t";
        str = str + pointToComma(modeltblItems.getData()[x].WRBTR) + "\t";
        str = str + pointToComma(modeltblItems.getData()[x].DMBTR) + "\t";
        str = str + pointToComma(modeltblItems.getData()[x].MENGE) + "\t";
        str = str + modeltblItems.getData()[x].MEINS + "\t";
        str = str + modeltblItems.getData()[x].KDIND + "\t";
        str = str + modeltblItems.getData()[x].RIVITXT + "\t";
        str = str + modeltblItems.getData()[x].REV_TEXT + "\t";
        str = str + modeltblItems.getData()[x].APPROVER + "\t";
        str = str + modeltblItems.getData()[x].APPR_TEXT;
        str = str + "\n";
    }
    oTxtPaste.setValue(str);
    oDialogExportToExcel.open();
    oTxtPaste.focus();
    var txtLen = oTxtPaste.getValue().length;
    oTxtPaste.selectText(0, txtLen);
}


function importFromExcel() {
    Busy.open();
    var data = oTxtFromExcel.getValue();
    var rows = data.match(/[^\r\n]+/g);
    var i = tblItems.getMaxItemsCount();
    i--;
    var i2 = 0;
    if (i > -1) {
        i2 = parseInt(modeltblItems.getData()[i].ITEM);
    } else {
        i2 = 0;
    }
    tblItems.setGrowing(true);
    if (typeof(modeltblItems.getData()) == 'undefined') {
        var arr = new Array();
    } else {
        if (modeltblItems.getData().length) {
            var arr = modeltblItems.getData();
        } else {
            var arr = new Array();
        }
    }
    for (var y in rows) {
        var cells = rows[y].split("\t");
        // if (cells[0] != "Account") {
        if (isNaN(parseInt(cells[0], 10))) {
            //Assume this a header line
        } else {
// 1st-PKTILI, 2nd-PRCTR, ... 7th-WRBTR, 8-DMBTR, 9-MENGE, 10-MEINS, 11th-KDIND. Index = pos - 1
            modelFormCreate.setData(new Object);
            inFormCreateBUZEI.setValue(" ");
            i2++;
            inFormCreateITEM.setValue(i2);
            inFormCreateSELFLAG.setSelected(" ");
            inFormCreateSUMMA.setValue(0);
            inFormCreateVEROTXT.setValue("");
            inFormCreateWAERS.setValue(" ");
            inFormCreatePKTILI.setValue(cells[0]);
            inFormCreatePRCTR.setValue(cells[1]);
            inFormCreateKUSTPAIKKA.setValue(cells[2]);
            inFormCreateORDER.setValue(cells[3]);
            inFormCreateFUNC_AREA.setValue(cells[4]);
            inFormCreateVERO.setValue(cells[5]);
            inFormCreateWRBTR.setValue(commaToPoint(cells[6]));
            inFormCreateDMBTR.setValue(commaToPoint(cells[7]));
            inFormCreateMENGE.setValue(commaToPoint(cells[8]));
            inFormCreateMEINS.setValue(cells[9]);
            inFormCreateKDIND.setValue(cells[10]);
            inFormCreateRIVITXT.setValue(cells[11]);
            // reviewer name 12
            inFormCreateAPPROVER.setValue(cells[13]);

            var data = modelFormCreate.getData();
            data.SELFLAG = false;
            data.EDITFLAG = true;
            data.EDITFLAG2 = true;
            data.EDITFLAG3 = false;
            arr.push(data);
        }
    }
    modeltblItems.setData(arr);
    intblItemsSUMMAT.fireChange();
    Busy.close();
    Page1.setBusy(false);
    oTxtFromExcel.setValue("");
}
