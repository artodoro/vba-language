'@Interface

''' <summary>
''' При использовании без квалификатора объектов это свойство возвращает объект Application , представляющий приложение Microsoft Excel.
''' Если используется с квалификатором объекта, это свойство возвращает объект Application, который представляет создателя указанного объекта (вы можете задать это свойство с объектом OLE-автоматизации для возврата приложения этого объекта). Только для чтения.
''' </summary>
''' <remarks>
''' Объект AddIn является членом коллекции AddIns . Коллекция AddIns содержит список всех надстроек, доступных для Microsoft Excel, независимо от того, установлены ли они. Этот список соответствует списку надстроек, отображаемым в диалоговом окне Надстройки 
''' <example>
''' В этом примере отображается сообщение о приложении, которое создало myObject.
''' <code>
''' Set myObject = ActiveWorkbook 
''' If myObject.Application.Value = "Microsoft Excel" Then 
'''     MsgBox "This is an Excel Application object." 
''' Else 
'''     MsgBox "This is not an Excel Application object." 
''' End If
''' </code>
''' </example>
Property Get Application() As Application
End Property

''' <summary>
''' Возвращает уникальный идентификатор только для чтения или CLSID, определяющий объект в виде строки.
''' </summary>
Property Get CLSID() As String
End Property

''' <summary>
''' Возвращает 32-битное целое число, указывающее на приложение, в котором объект был создан. Только для чтения, Long.
''' </summary>
''' <returns>XlCreator</returns>
''' <remarks>
''' Если объект был создан в Microsoft Excel, это свойство возвращает строку XCEL, которая эквивалентна шестнадцатеричной цифре 5843454C. Свойство Creator предназначено для использования в Microsoft Excel для Macintosh, где каждое приложение имеет код создателя из четырех символов. Например, Microsoft Excel содержит код создателя XCEL.
''' </remarks>
Property Get Creator() As XlCreator
End Property

''' <summary>
''' Возвращает имя объекта, включая путь к нему на диске, в виде строки. Только для чтения, String.
''' </summary>
''' <example>
''' В этом примере отображается путь и имя файла для каждой доступной надстройки.
''' <code>
''' For Each a In AddIns 
'''     MsgBox a.FullName 
''' Next a
''' </code>
''' </example>
Property Get FullName() As String
End Property

''' <summary>
''' Значение <c>True</c>, если надстройка установлена или для установки надстройки; Значение <c>False</c>, если надстройка удалена или для удаления надстройки. Для чтения и записи, <c>Boolean</c>.
''' </summary>
''' <remarks>
''' Если установить для этого свойства значение <c>True</c>, надстройка устанавливается и вызывается ее <c>Auto_Add</c> функции. При установке для этого свойства значения <c>False</c> надстройка удаляется и вызывается ее <c>Auto_Remove</c> функции.
''' </remarks>
''' <example>
''' В этом примере окно сообщения используется для отображения состояния установки надстройки Решателя.
''' <code>
''' Set a = AddIns("Solver Add-In") 
''' If a.Installed = True Then 
'''     MsgBox "The Solver add-in is installed" 
''' Else 
'''     MsgBox "The Solver add-in is not installed" 
''' End If
''' </code>
''' </example>
Property Get Installed() As Boolean
End Property
Property Let Installed() As Boolean
End Property

''' <summary>
''' Возвращает значение <c>True</c>, если надстройка в настоящее время открыта. Только для чтения, <c>Boolean</c>.
''' </summary>
Property Get IsOpen() As Boolean
End Property

''' <summary>
''' Возвращает значение <c>String</c>, представляющее имя объекта.
''' </summary>
Property Get Name() As String
End Property

''' <summary>
''' Возвращает родительский объект для указанного объекта. Только для чтения.
''' </summary>
Property Get Parent() As Object
End Property

''' <summary>
''' Возвращает значение <c>String</c>, представляющее полный путь к надстройке, за исключением окончательного разделителя и имени надстройки.
''' </summary>
Property Get Path() As String
End Property

''' <summary>
''' Возвращает программные идентификаторы для объекта. Только для чтения, <c>String</c>.
''' </summary>
''' <example>
''' В этом примере создается список программных идентификаторов для объектов OLE на одном листе.
''' <code>
''' rw = 0 
''' For Each o in Worksheets(1).OLEObjects 
'''     With Worksheets(2) 
'''         rw = rw + 1 
'''         .cells(rw, 1).Value = o.ProgId 
'''     End With 
''' Next
''' </code>
''' </example>
Property Get progID() As String
End Property