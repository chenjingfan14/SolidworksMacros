''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SheetParts
'Joe Weaver ( jeweave4@ncsu.edu)
'
'Traverses an assembly and writes a csv file of matching parts with their expected dimensions
'
'Post-processing of the list is intended for other scripts.
'Examples:
'   -rounding and unit conversion
'   -mapping length, width, height to x,y,z based on SheetThickness,
'   -Ideally, feeding into an implementation of good heuristic for 2d cutting and packing
'
'Input: An active assembly containing one or more parts to be cut from sheet stock
'
'Output: A CSV file of all sheet stock parts with dimensions
'
'Requirements: Sheet stockfinds any parts containing the custom properties 'FromSheetStock'
'with the boolean value 'Yes' and a listed 'SheetThickness'
'
'Written for SolidWorks 2015 - will likely work on earlier versions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WARNING:  This macro was developed for my own personal use. It has not been extensively tested
'in the wild, so there's likely plenty of bugs to be discovered when it runs into assumptions
'that are not valid beyond my assemblies
'
'For example, this assumes that the sheet is aligned with one of the built in reference planes
'(Front, Top, Right). Otherwise, the bounding box will give incorrect values. All my parts
'are aligned this way, so I'm loathe to write a 'smarter' bounding box algorithm.
'
'Apart from overwriting the output, this script should not delete any of your parts/assemblies.
'It WILL induce a rebuild when reading parts, so be warned.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'todo: default config
'- take an assembly filename rather than use active doc
'- log warnings, errors
'- output location
'- custom headers in csv?
'- specify units expected/to use?
'- disallow rebuilds?


'edit this to the filename you want to write output to
Const outFile = "K:\Output.txt"



Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2


Private Sub Main()
    Dim depends                 As Variant
    Dim idx                     As Integer
    Dim depDoc                  As SldWorks.ModelDoc2
    Dim swPart                  As SldWorks.PartDoc
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Dim valOut As String
    Dim rValOut As String
    Dim customVal As String
    Dim wasResolved As Boolean
    Dim customGetResult As swCustomInfoGetResult_e
    Dim acterrs As Long
    Dim vBoundBox As Variant
    Dim xdim As Double
    Dim ydim As Double
    Dim zdim As Double
    Dim swMatDB As String
    Dim sMatName As String
    
    'Set the solidworks app and model pointers to the current instance and active document
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    'If no model currently loaded, then exit
    If swModel Is Nothing Then
        Call MsgBox("A SolidWorks document needs to be loaded!", vbExclamation, "Custom Properties")
        Exit Sub
    End If
       
    'Don't traverse model if it is not an assembly
    If (swModel.GetType <> swDocAssembly) Then
        Call MsgBox("A SolidWorks assembly needs to be loaded!", vbExclamation, "Custom Properties")
        Exit Sub
    End If
    
    'get a list of dependencies
    depends = swApp.GetDocumentDependencies2(swModel.GetPathName, True, False, False)
    
    If IsEmpty(depends) Then
        Call MsgBox("No dependencies found in assembly.")
        Exit Sub
    End If
    
    'TODO user specify output file
    Open outFile For Output Access Write Lock Write As #1
    
    'write the header for the csv file
    Print #1, "x,y,x,dimUnits,thickness,material,name,fileName"
    
    idx = 1
    While idx <= UBound(depends)
    
        'Only look at parts files, based on extension and then based on what the opened file model type returns
        'What I really want here is a CONTINUE statement, but not supported in VBA.
        'Options are a) nesting conditionals, which I don't like for guard clauses
        'b) Alternative would be hacking up some sort of map/filter function, but I don't want to do that for a one-off VBA macro
        'Otherwise, c) a GoTo to NextIdx, which lives at the end end of the loop - about the only time I would consider GoTo
        If (Right(depends(idx), 7) <> ".SLDPRT") Then GoTo NextIdx
    
        'Go ahead and open/activated the doc
        'We're assuming automatic rebuilds are ok. Don't bug the user.
        Set depDoc = swApp.ActivateDoc3(depends(idx), True, swRebuildActiveDoc, acterrs)
        
        'If we had trouble opening this doc, go to next one
        'TODO log these errors?
        If (acterrs <> 0) Then GoTo NextIdx
        'Had to add this as opening a vendor supplied model got to this point without a depdoc
        If (depDoc Is Nothing) Then GoTo NextIdx
        
        'all is well, update our models to the current doc
        Set swModel = swApp.ActiveDoc
    
        'Only bother if the swModle really is a part
        If (swModel.GetType <> swDocPART) Then GoTo NextIdx
      
        'update our part reference
        Set swPart = swApp.ActiveDoc
        
        'lets us read custom properties
        Set swCustPropMgr = swModel.Extension.CustomPropertyManager(Empty)
        
        '(Refactor note) A lot of reading and check custom values here, ripe for subroutine)
        'Check the FromSheetStock property
        customGetResult = swCustPropMgr.Get5("FromSheetStock", False, valOut, rValOut, wasResolved)
        
        'No sheet stock property listed
        If (customGetResult = swCustomInfoGetResult_NotPresent) Then GoTo NextIdx
        If wasResolved Then customVal = rValOut Else customVal = valOut
        
        'Property is listed as not being from sheet stock
        If (customVal <> "Yes") Then GoTo NextIdx
        
        'Get the sheet thickness
        'Todo - output blank if it doesn't exist, rather than skip?
        customGetResult = swCustPropMgr.Get5("SheetThickness", False, valOut, rValOut, wasResolved)
        If (customGetResult = swCustomInfoGetResult_NotPresent) Then GoTo NextIdx
        If wasResolved Then customVal = rValOut Else customVal = valOut
                
        'strip the double quote if the thickness is given as inches (eg 3/8")
        If (Right$(customVal, 1) = Chr(34)) Then
                    customVal = Left$(customVal, Len(customVal) - 1)
        End If
        
        'use this to get a really simple bounding box
        'accurate enough for cutting decisions
        'BUT assumes that the part axes are parallel to the main axes - this is usually the case
        'ran into one issue where I extruded from an angle to a plane and the bounding box height was the rise of the part rather than its thickness
        'should be ablet to check and correct for that case
        vBoundBox = swPart.GetPartBox(False)
    
        'x, y, and z lengths based on bounding box vertices
        xdim = Abs(vBoundBox(0) - vBoundBox(3))
        ydim = Abs(vBoundBox(1) - vBoundBox(4))
        zdim = Abs(vBoundBox(2) - vBoundBox(5))
    
        'get the material of the part
        sMatName = swPart.GetMaterialPropertyName2("Default", swMatDB)
            
            
        'Write out the comma separated dimensions, units,thickness, material, filename, and part name
        Print #1, xdim & "," & ydim & "," & zdim & "," & swModel.GetUserUnit(swLengthUnit).GetUnitsString(True) & "," & customVal & "," & sMatName & "," & depends(idx) & "," & depends(idx - 1)
            
        
NextIdx:
        'keep things clean, close the current doc
        Set depDoc = Nothing
        swApp.CloseDoc (depends(idx))
        
        'technically, we can iterate by 2, since the dependency traversal returns two strings per dep found, but it doesn't noticibly impact my performance to do this
        idx = idx + 1
        
    Wend
    
    'close the output file
    Close #1

End Sub


