export enum VBAElementKind {
    lineContinuationElement = ' _',
    commentElement = "'",
    remElement = 'Rem',

    trueElement = 'True',
    falseElement = 'False',
    
    meElement = 'Me',
    
    byvalElement = 'ByVal',
    byrefElement = 'ByRef',
    optionalElement = 'Optional',
    paramArrayElement = 'ParamArray',

    constElement = 'Const',
    dimElement = 'Dim',
    enumElement = 'Enum',
    eventElement = 'Event',
    functionElement = 'Function',
    getElement = 'Get',
    letElement = 'Let',
    lsetElement = 'LSet',
    preserveElement = 'Preserve',
    propertyElemet = 'Property',
    redimElement = 'ReDim',
    rsetElement = 'RSet',
    setElement = 'Set',
    subElement = 'Sub',
    typeElement = 'Type',
    withEventElement = 'WithEvent',
    
    ifElement = 'If',
    elseElement = 'Else',
    elseifElement = 'ElseIf',
    thenElement = 'Then',
    
    andElement = 'And',
    eqvElement = 'Eqv',
    impElement = 'Imp',
    isElement = 'Is',
    likeElement = 'Like',
    notElement = 'Not',
    orElement = 'Or',
    xorElement = 'Xor',
    
    errorElement = 'Error',
    gosubElement = 'GoSub',
    gotoElement = 'GoTo',
    onElement = 'On',
    resumeElement = 'Resume',
    returnElement = 'Return',
    
    selectElement = 'Select',
    caseElement = 'Case',
    
    doElement = 'Do',
    whileElement = 'While',
    untilElement = 'Until',
    loopElement = 'Loop',
    wendElement = 'Wend',
    
    eachElement = 'Each',
    forElement = 'For',
    inElement = 'In',
    nextElement = 'Next',
    stepElement = 'Step',
    toElement = 'To',
    
    baseElement = 'Base',
    binaryElement = 'Binary',
    compareElement = 'Compare',
    explicitElement = 'Explicit',
    implementsElement = 'Implements',
    optionElement = 'Option',
    textElement = 'Text',
    
    emptyElement = 'Empty',
    nothingElement = 'Nothing',
    nullElement = 'Null',
    
    defBoolElement = 'DefBool',
    defByteElement = 'DefByte',
    defCurElement = 'DefCur',
    defDateElement = 'DefDate',
    defDblElement = 'DefDbl',
    defIntElement = 'DefInt',
    defLngElement = 'DefLng',
    defLngLngElement = 'DefLngLng',
    defLngPtrElement = 'DefLngPtr',
    defObjElement = 'DefObj',
    defSngElement = 'DefSng',
    defStrElement = 'DefStr',
    defVarElement = 'DefVar',

    closeElement = 'Close',
    inputElement = 'Input',
    lineElement = 'Line',
    lockElement = 'Lock',
    openElement = 'Open',
    outputElement = 'Output',
    printElement = 'Print',
    putElement = 'Put',
    readElement = 'Read',
    unlockElement = 'Unlock',
    writeElement = 'Write',

    addressOfElement = 'AddressOf',
    asElement = 'As',
    callSignatureElement = 'Call',
    endElement = 'End',
    eraseElement = 'Erase',
    exitElement = 'Exit',
    modElement = 'Mod',
    newElement = 'New',
    raiseEventElement = 'RaiseEvent',
    stopElement = 'Stop',
    typeOfElement = 'TypeOf',
    withElement = 'With',

    classInitializeElement = '_Initialize',
    classTerminateElement = '_Terminate'
}

export enum VBAElementKindModifier {
    friendModifier = 'Friend',
    privateModifier = 'Private',
    propertyModifier = 'Property',
    publicModifier = 'Public',
    staticModifier = 'Static',
    sharedModifier = 'Shared',
    clsModifier = '.cls',
    basModifier = '.bas',
    frmModifier = '.frm'
}