if (app.documents.length != 0) {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
    app.activeWindow.transformReferencePoint = AnchorPoint.TOP_LEFT_ANCHOR;
    var myDocument = app.activeDocument;
    var myPageCount = myDocument.pages.length;

    var myPrintPageWidth, myPrintPageHeight;
    var myStartPage, myEndPage;
    var myStartNumber;
    var myReversePage;
    var myOffsetBottom;
	var myOffsetTop = 18;
    var myTextBefore, myTextAfter
    var iPage = myDocument.pages.length;
    var myCaptionLength = 200;
	var myOffsetRightBorder = 4;

    myDisplayDialog();
    app.activeWindow.transformReferencePoint = AnchorPoint.CENTER_ANCHOR;
}
else {
    alert("Нет открытых документов!");
    exit();
}

function myDisplayDialog() {
    var myDialog = app.dialogs.add({
        name: "Нумерация спусков Офис-лидер v2.1.3",
        canCancel: true
    });
    with (myDialog) {
        var myPrintPageWidthField, myPrintPageHeightField;
        var myStartPageField, myEndPageField;
        var myStartNumberField, myCheckBoxRev, myOffsetRightBorderField, myOffsetBottomField;
        var myTextBeforeField, myTextAfterField;

        with (dialogColumns.add()) {
            //поле печатного листа
            with (borderPanels.add()) {
                with (dialogColumns.add()) {
                    staticTexts.add({ staticLabel: "Размер печатного листа: W:", minWidth: 170 });
                }
                with (dialogColumns.add()) {
                    myPrintPageWidthField = integerEditboxes.add({ editValue: 497, minWidth: 60 });
                }
                with (dialogColumns.add()) {
                    staticTexts.add({ staticLabel: "H:" });
                }
                with (dialogColumns.add()) {
                    myPrintPageHeightField = integerEditboxes.add({ editValue: 347, minWidth: 60 });
                }
            }
            //поле нумерации
            with (borderPanels.add()) {
                with (dialogColumns.add()) {
                    staticTexts.add({ staticLabel: "Начать со страницы:", minWidth: 170 });
                }
                with (dialogColumns.add()) {
                    myStartPageField = integerEditboxes.add({ editValue: 1, minWidth: 60, minimumValue: 1, maximumValue: myPageCount });
                }
                with (dialogColumns.add()) {
                    staticTexts.add({ staticLabel: "по:" });
                }
                with (dialogColumns.add()) {
                    myEndPageField = integerEditboxes.add({ editValue: myPageCount, minWidth: 60, minimumValue: 1, maximumValue: myPageCount });
                }
            }
            //поле первой страницы и оборота
            with (borderPanels.add()) {
                with (dialogColumns.add()) {
                    staticTexts.add({ staticLabel: "Начать с номера:", minWidth: 170 });
                }
                with (dialogColumns.add()) {
                    myStartNumberField = integerEditboxes.add({ editValue: 1, minWidth: 60 });
                }
                with (dialogColumns.add()) {
                    staticTexts.add({ staticLabel: "Оборот?", minWidth: 56 });
                }
                with (dialogColumns.add()) {
                    myCheckBoxRev = checkboxControls.add({ checkedState: true });
                }
            }
            //поле клапана и отступа
            with (borderPanels.add()) {
                with (dialogColumns.add()) {
                    staticTexts.add({ staticLabel: "Отступ справа:", minWidth: 170 });
                }
                with (dialogColumns.add()) {
                    myOffsetRightBorderField = integerEditboxes.add({ editValue: 4, minWidth: 60 });
                }
                with (dialogColumns.add()) {
                    staticTexts.add({ staticLabel: "Отступ снизу:" });
                }
                with (dialogColumns.add()) {
                    myOffsetBottomField = integerEditboxes.add({ editValue: 20, minWidth: 60 });
                }
            }
            //поле текста
            with (borderPanels.add()) {
                with (dialogColumns.add()) {
                    with (dialogRows.add()) {
                        staticTexts.add({ staticLabel: "Текст перед номером:" });
                    }
                    with (dialogRows.add()) {
                        myTextBeforeField = textEditboxes.add({ editContents: "#0000, 4+4, 347*497, Астарта глянц 120", minWidth: 400 });
                    }
                    with (dialogRows.add()) {
                        staticTexts.add({ staticLabel: "Текст после номера:" });
                    }
                    with (dialogRows.add()) {
                        myTextAfterField = dialogRows.add().textEditboxes.add({ editContents: "", minWidth: 400 });
                    }
                }
            }
        }
    }
    if (myDialog.show() == true) {
        myPrintPageWidth = myPrintPageWidthField.editValue;
        myPrintPageHeight = myPrintPageHeightField.editValue;
        myStartPage = myStartPageField.editValue;
        myEndPage = myEndPageField.editValue;
        myStartNumber = myStartNumberField.editValue;
        myReversePage = myCheckBoxRev.checkedState;
        myOffsetRightBorder = myOffsetRightBorderField.editValue;
        myOffsetBottom = myOffsetBottomField.editValue;
        myTextBefore = myTextBeforeField.editContents;
        myTextAfter = myTextAfterField.editContents;
        myDialog.destroy()

        if (myStartPage == myEndPage)
            myCaptionOnOnePage();
        else if (!myReversePage)
            myCaptionOnPagesNoReverse();
        else if ((myEndPage + 1) - myStartPage == 2)
            myCaptionOnTwoPages();
        else
            myCaptionOnManyPages();
    }
    else
        myDialog.destroy();

}

function myCaptionOnOnePage() {
    var myCurrentPage = myDocument.pages.item(myStartPage - 1);
    var myCurrentPageWidth = myCurrentPage.bounds[3];
    var myCurrentPageHeight = myCurrentPage.bounds[2];
    var myPlaceText;
    var myTextObject;
    var offsetY = myOffsetTop + myPrintPageHeight - myOffsetBottom;
    var offsetX = (myCurrentPageWidth / 2) + (myPrintPageWidth / 2) - myOffsetRightBorder;

    var myTextFrame = myCurrentPage.textFrames.add();
    myTextFrame.geometricBounds = [offsetY, offsetX - 4, offsetY + 4, offsetX + myCaptionLength];
    myTextFrame.rotationAngle = 90;

    myPlaceText = (myTextBefore + " " + myTextAfter);
    myTextFrame.contents = myPlaceText;

    myTextObject = myTextFrame.paragraphs.item(0);
    myTextObject.fillColor = myDocument.colors.item("Registration");
    myTextObject.justification = Justification.leftAlign;
    myTextObject.leading = -1
    try {
        myTextObject.appliedFont = app.fonts.item("Arial");
        myTextObject.fontStyle = "Bold";
    }
    catch (e) { }
    myTextObject.fillColor = "Registration";
    myTextObject.pointSize = 10;
}

function myCaptionOnTwoPages() {
    var myCurrentPage;
    var myCurrentPageWidth;
    var myCurrentPageHeight;
    var myPlaceText;
    var myTextObject;
    var offsetY, offsetX;
    var iPage;
    var iSpusk = myStartNumber;
    var iEven = 1;

    for (iPage = myStartPage; iPage <= myEndPage; iPage++) {
        //caption
        if (iEven % 2 == 1)
            myPlaceText = (myTextBefore + ", оборот чужой, лицо " + myTextAfter);
        else {
            var myPlaceText = (myTextBefore + ", оборот чужой, оборот " + myTextAfter);
            iSpusk++;
        }
        iEven++

        //size page
        myCurrentPage = myDocument.pages.item(iPage - 1);
        myCurrentPageWidth = myCurrentPage.bounds[3];
        myCurrentPageHeight = myCurrentPage.bounds[2];

        //offset X Y
        offsetY = myOffsetTop + myPrintPageHeight - myOffsetBottom;
        offsetX = (myCurrentPageWidth / 2) + (myPrintPageWidth / 2) - myOffsetRightBorder;

        //text frame
        myTextFrame = myCurrentPage.textFrames.add();
        myTextFrame.geometricBounds = [offsetY, offsetX - 4, offsetY + 4, offsetX + myCaptionLength];
        myTextFrame.rotationAngle = 90;
        myTextFrame.contents = myPlaceText;

        //text object options
        myTextObject = myTextFrame.paragraphs.item(0);
        myTextObject.fillColor = myDocument.colors.item("Registration");
        myTextObject.justification = Justification.leftAlign;
        myTextObject.leading = -1
        try {
            myTextObject.appliedFont = app.fonts.item("Arial");
            myTextObject.fontStyle = "Bold";
        }
        catch (e) { }
        myTextObject.fillColor = "Registration";
        myTextObject.pointSize = 10;
    }
}

function myCaptionOnManyPages() {
    var myCurrentPage;
    var myCurrentPageWidth;
    var myCurrentPageHeight;
    var myPlaceText;
    var myTextObject;
    var offsetY, offsetX;
    var iPage;
    var iSpusk = myStartNumber;
    var iEven = 1;

    for (iPage = myStartPage; iPage <= myEndPage; iPage++) {
        //caption
        if (iEven % 2 == 1)
            myPlaceText = (myTextBefore + ", оборот чужой, спуск " + iSpusk + " лицо " + myTextAfter);
        else {
            var myPlaceText = (myTextBefore + ", оборот чужой, спуск " + iSpusk + " оборот " + myTextAfter);
            iSpusk++;
        }
        iEven++

        //size page
        myCurrentPage = myDocument.pages.item(iPage - 1);
        myCurrentPageWidth = myCurrentPage.bounds[3];
        myCurrentPageHeight = myCurrentPage.bounds[2];

        //offset X Y
        offsetY = myOffsetTop + myPrintPageHeight - myOffsetBottom;
        offsetX = (myCurrentPageWidth / 2) + (myPrintPageWidth / 2) - myOffsetRightBorder;

        //text frame
        myTextFrame = myCurrentPage.textFrames.add();
        myTextFrame.geometricBounds = [offsetY, offsetX - 4, offsetY + 4, offsetX + myCaptionLength];
        myTextFrame.rotationAngle = 90;
        myTextFrame.contents = myPlaceText;

        //text object options
        myTextObject = myTextFrame.paragraphs.item(0);
        myTextObject.fillColor = myDocument.colors.item("Registration");
        myTextObject.justification = Justification.leftAlign;
        myTextObject.leading = -1
        try {
            myTextObject.appliedFont = app.fonts.item("Arial");
            myTextObject.fontStyle = "Bold";
        }
        catch (e) { }
        myTextObject.fillColor = "Registration";
        myTextObject.pointSize = 10;
    }
}

function myCaptionOnPagesNoReverse() {
    var myCurrentPage;
    var myCurrentPageWidth;
    var myCurrentPageHeight;
    var myPlaceText;
    var myTextObject;
    var offsetY, offsetX;
    var iPage;
    var iSpusk = myStartNumber;

    for (iPage = myStartPage; iPage <= myEndPage; iPage++) {
        //caption
        myPlaceText = (myTextBefore+", спуск " + iSpusk + myTextAfter);
        iSpusk++;

        //size page
        myCurrentPage = myDocument.pages.item(iPage - 1);
        myCurrentPageWidth = myCurrentPage.bounds[3];
        myCurrentPageHeight = myCurrentPage.bounds[2];

        //offset X Y
        offsetY = myOffsetTop + myPrintPageHeight - myOffsetBottom;
        offsetX = (myCurrentPageWidth / 2) + (myPrintPageWidth / 2) - myOffsetRightBorder;

        //text frame
        myTextFrame = myCurrentPage.textFrames.add();
        myTextFrame.geometricBounds = [offsetY, offsetX - 4, offsetY + 4, offsetX + myCaptionLength];
        myTextFrame.rotationAngle = 90;
        myTextFrame.contents = myPlaceText;

        //text object options
        myTextObject = myTextFrame.paragraphs.item(0);
        myTextObject.fillColor = myDocument.colors.item("Registration");
        myTextObject.justification = Justification.leftAlign;
        myTextObject.leading = -1
        try {
            myTextObject.appliedFont = app.fonts.item("Arial");
            myTextObject.fontStyle = "Bold";
        }
        catch (e) { }
        myTextObject.fillColor = "Registration";
        myTextObject.pointSize = 10;
    }
}

