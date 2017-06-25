# Swift-NumbersScripting-Framework
Framework to use Swift with the Scripting Bridge with iWork Numbers

This is a Numbers scripting framework developed via the instructions  on Github from Tony Ingraldi  (tingraldi), many thanks:

[GitHub - tingraldi/SwiftScripting: Utilities and samples to aid in using Swift with the Scripting Bridge](https://github.com/tingraldi/SwiftScripting)

Please refer to his writeup on modifying/changing for your own use.
# Disclaimer
Hey, I’m not the world’s best programmer by any measure.  I am writing this project because I want to see this functionality, I know you can do better, please do.  If you happen to see value in this or other Projects herein and would like to make them happen, please contact me.

# Known Issues
* NumbersCell.value. This property is supposed to return an Any?, enclosing String, Double, Date, missing value (nil) etc.  Instead, it returns an SBObject referring to the Cell itself.    I’d love to hear how to fix this. Workarounds:
	* ```let value: String = cell.formattedValue ??  “”```
	* ```let valueofcell: [Any?] = (row as! NumbersRange).cells!().array(byApplying:  \#selector(getter: NumbersCell.value))```
* NumbersRange.format does not return the proper value: workaround:
	* ```let formatofcell: [NSAppleEventDescriptor?] = row.cells!().array(byApplying:  \#selector(getter: NumbersRange.format)) as! [NSAppleEventDescriptor]```
	
	Use this for compare
	* ```if formatofcell[i]!.typeCodeValue == NumbersNMCT.xxxx.rawValue``` where xxxx is the desired format
	 
# Usage Hints:
* Set Application Object:
```
let nmbrs: NumbersApplication = SBApplication(bundleIdentifier: “com.apple.iWork.numbers”)! as NumbersApplication
```
* Set Document Object:
```
let document: NumbersDocument = nmbrs.documents!().object(atLocation: 0) as! NumbersDocument
```
* Get Cell values:
```
let formattedValue: String = (cell as! NumbersCell).formattedValue ?? ""
```
* Read entire table to array:
```
let valarray = (tbl.cells!() as NSArray).map {($0 as! NumbersCell).formattedValue!}
```
* Read datarange (skip row headers and footers) into array of arrays:
```
var rowdata: [[String]] = []
let firstdatarow: Int = tbl.headerRowCount! + 1
let lastdatarow: Int = tbl.rows!().count - tbl.footerRowCount!
for  r in firstdatarow...lastdatarow {
	let row: NumbersRow = tbl.rows!().object(atLocation: r) as! NumbersRow        
        let coldata = (row.cells!() as NSArray).map {($0 as! NumbersCell).formattedValue!}
        rowdata.append(coldata)
}    	
```
* Set Cell Values. Note: Setting values one at a time can be slower than you might like as the spreadsheet refreshes during the setting process.  If you are doing bulk data setting, use the copy and paste commands, instead.  Note that the table format is tab-delimited.  Copy and Paste ignore hidden and filtered rows/columns, so you can copy only the visible rows of a filtered table. likewise, pasteing occurs from the beginning of the selected cell range in the table and ignores hidden rows/columns.
	* String:	```(tbl.cellRange?.cells!().object(atLocation: 2) as! NumbersCell).setValue!(“test”)```
	* Number:	```(tbl.cellRange?.cells!().object(atLocation: 2) as! NumbersCell).setValue!(4)```
	* Formula:	```(tbl.cellRange?.cells!().object(atLocation: 2) as! NumbersCell).setValue!(“=14*2”)```
* Adding Objects:
	* Sheet:
```
sht  = objectWithApplication(nmbrs, scriptingClass: NumbersScripting.sheet)
document.sheets!().add(sht)
```
	* Table:
```
tbl = objectWithApplication(nmbrs, scriptingClass: NumbersScripting.table)
sht.tables!().add(tbl)
```
* copy and paste commands (keystroke simulation): 
 ```
 func copy() {
     
    let event1 = CGEvent(keyboardEventSource: nil, virtualKey: 0x08, keyDown: true); // cmd-c down
    event1?.flags = CGEventFlags.maskCommand;
    event1?.post(tap: CGEventTapLocation.cghidEventTap);
    
    let event2 = CGEvent(keyboardEventSource: nil, virtualKey: 0x08, keyDown: false); // cmd-c up
    event2?.post(tap: CGEventTapLocation.cghidEventTap);
  }
```
```func paste () {
    
    let event1 = CGEvent(keyboardEventSource: nil, virtualKey: 0x09, keyDown: true); // cmd-v down
    event1?.flags = CGEventFlags.maskCommand;
    event1?.post(tap: CGEventTapLocation.cghidEventTap);
    
    let event2 = CGEvent(keyboardEventSource: nil, virtualKey: 0x09, keyDown: false) // cmd-v up
    event2?.post(tap: CGEventTapLocation.cghidEventTap)
}
```

