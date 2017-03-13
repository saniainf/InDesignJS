/**
* @@@BUILDINFO@@@ CropMarks.js !Version1.0! Mon Mar 13 2017 12:09:03 GMT+0700
*/

main();
function main(){
	//Make certain that user interaction (display of dialogs, etc.) is turned on.
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	
	if (app.documents.length != 0){
		if (app.selection.length > 0){
			switch(app.selection[0].constructor.name){
				case "Rectangle":
				case "Oval":
				case "Polygon":
				case "GraphicLine":
				case "Group":
				case "TextFrame":
				case "Button":
					myDisplayDialog();
					break;
				default:			
					alert("Please select a page item and try again.");
					break;
			}
		}
		else{
			alert("Please select an object and try again.");
		}
	}
	else{
		alert("Please open a document, select an object, and try again.");
	}
}
function myDisplayDialog(){
	var myDialog = app.dialogs.add({name:"CropMarks Ofis-Lider v1.0"});
	with(myDialog){
		with(dialogColumns.add()){
			with(borderPanels.add()){
				staticTexts.add({staticLabel:"Options:"});
				with (dialogColumns.add()){
					staticTexts.add({staticLabel:"Length:"});
					staticTexts.add({staticLabel:"Offset:"});
				}
				with (dialogColumns.add()){
					var myCropMarkLengthField = integerEditboxes.add({editValue: 4, minWidth: 60 });
					var myCropMarkOffsetField = integerEditboxes.add({editValue: 1, minWidth: 60 });
				}
			}
		}
	}
	var myReturn = myDialog.show();
	if (myReturn == true){
		//Get the values from the dialog box.
		var myCropMarkLength = myCropMarkLengthField.editValue;
		var myCropMarkOffset = myCropMarkOffsetField.editValue;
		var myCropMarkWidth = .088;
		myDrawPrintersMarks(myCropMarkLength, myCropMarkOffset, myCropMarkWidth);
	}
	else{
		myDialog.destroy();
	}
}

function myDrawPrintersMarks(myCropMarkLength, myCropMarkOffset, myCropMarkWidth){
	var myPageItems, myObject;
	var myDocument = app.activeDocument;
	var myOldRulerOrigin = myDocument.viewPreferences.rulerOrigin;
	myDocument.viewPreferences.rulerOrigin = RulerOrigin.spreadOrigin;
	//Save the current measurement units.
	var myOldXUnits = myDocument.viewPreferences.horizontalMeasurementUnits;
	var myOldYUnits = myDocument.viewPreferences.verticalMeasurementUnits;
	//Set the measurement units to points.
	myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
	myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
	//Create a layer to hold the printers marks (if it does not already exist).
	var myLayer = myDocument.layers.item("myCropMarks");
	try{
		myLayerName = myLayer.name;
	}
	catch (myError){
		var myLayer = myDocument.layers.add({name:"myCropMarks"});
	}
	//Get references to the Registration color and the None swatch.
	var myRegistrationColor = myDocument.colors.item("Registration");
	var myNoneSwatch = myDocument.swatches.item("None");
	myPageItems = myDocument.selection;
	for(var myCounter = 0; myCounter < myPageItems.length; myCounter ++){
		myObject = myPageItems[myCounter];
		myDrawCropMarks (myObject, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myPageItems);
	}
	myDocument.viewPreferences.rulerOrigin = myOldRulerOrigin;
	//Set the measurement units back to their original state.
	myDocument.viewPreferences.horizontalMeasurementUnits = myOldXUnits;
	myDocument.viewPreferences.verticalMeasurementUnits = myOldYUnits;
}

function myDrawCropMarks (myPageItem, myCropMarkLength, myCropMarkOffset, myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer, myPageItems){
	var myX1 = myPageItem.geometricBounds[1];
	var myY1 = myPageItem.geometricBounds[0];
	var myX2 = myPageItem.geometricBounds[3];
	var myY2 = myPageItem.geometricBounds[2];
    
	//Upper left crop mark pair.
    if (myCheckPoint(myX1 - myCropMarkLength, myY1 + myCropMarkOffset, myPageItems) == false){
	myDrawLine([myY1 + myCropMarkOffset, myX1, myY1 + myCropMarkOffset, myX1 - myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
    }
    if (myCheckPoint(myX1 + myCropMarkOffset, myY1 - myCropMarkLength, myPageItems) == false){
	myDrawLine([myY1, myX1 + myCropMarkOffset, myY1 - myCropMarkLength, myX1 + myCropMarkOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
    }

	//Lower left crop mark pair.
    if (myCheckPoint(myX1 - myCropMarkLength, myY2 - myCropMarkOffset, myPageItems) == false){
	myDrawLine([myY2 - myCropMarkOffset, myX1, myY2 - myCropMarkOffset, myX1 - myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
    }
    if (myCheckPoint(myX1 + myCropMarkOffset, myY2 + myCropMarkLength, myPageItems) == false){    
	myDrawLine([myY2, myX1 + myCropMarkOffset, myY2 + myCropMarkLength, myX1 + myCropMarkOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
    }

	//Upper right crop mark pair.
    if (myCheckPoint(myX2 + myCropMarkLength, myY1 + myCropMarkOffset, myPageItems) == false){
	myDrawLine([myY1 + myCropMarkOffset, myX2, myY1 + myCropMarkOffset, myX2 + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
    }
    if (myCheckPoint(myX2 - myCropMarkOffset, myY1 - myCropMarkLength, myPageItems) == false){
	myDrawLine([myY1, myX2 - myCropMarkOffset, myY1 - myCropMarkLength, myX2 - myCropMarkOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
    }

	//Lower left crop mark pair.
    if (myCheckPoint(myX2 + myCropMarkLength, myY2 - myCropMarkOffset, myPageItems) == false){
	myDrawLine([myY2 - myCropMarkOffset, myX2, myY2 - myCropMarkOffset, myX2 + myCropMarkLength], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
    }
    if (myCheckPoint(myX2 - myCropMarkOffset, myY2 + myCropMarkLength, myPageItems) == false){
	myDrawLine([myY2, myX2 - myCropMarkOffset, myY2 + myCropMarkLength, myX2 - myCropMarkOffset], myCropMarkWidth, myRegistrationColor, myNoneSwatch, myLayer);
    }
}

function myDrawLine(myBounds, myStrokeWeight, myRegistrationColor, myNoneSwatch, myLayer){
	app.activeWindow.activeSpread.graphicLines.add(myLayer, undefined, undefined,{strokeWeight:myStrokeWeight, fillColor:myNoneSwatch, strokeColor:myRegistrationColor, geometricBounds:myBounds})
    app.activeWindow.activeSpread.graphicLines.add
}

function myCheckPoint(myX, myY, myPageItems){
    for(var myCounter = 0; myCounter < myPageItems.length; myCounter ++){
		myObject = myPageItems[myCounter];
		var result = myIsPointInside (myX, myY, myObject)
        if (result == true){
            return true;
            }
	}
    return false;
}

function myIsPointInside(myX, myY, myPageItem){
    var myX1 = myPageItem.geometricBounds[1];
    var myY1 = myPageItem.geometricBounds[0];
    var myX2 = myPageItem.geometricBounds[3];
    var myY2 = myPageItem.geometricBounds[2];
    
    if ((myX > myX1) && (myX < myX2)){
        if ((myY > myY1) && (myY < myY2)){
            return true;
            }
        }
    return false;
}