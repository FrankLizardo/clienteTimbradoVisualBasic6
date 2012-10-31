VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Ejemplo de conexion al Web Service de Facturación Moderna"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTimbrar 
      Caption         =   "&Timbrar"
      Height          =   615
      Left            =   5160
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtFolioFiscal 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "FB59830A-6906-4F00-939D-2981614B98D2"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtLayout 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Text            =   "C:\factura_en_texto_ejemplo.txt"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Timbrar comprobante"
      Height          =   1455
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   6735
      Begin VB.Label Label1 
         Caption         =   "Ruta del layout"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cancelar Comprobante"
      Height          =   1455
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   6735
      Begin VB.Label Label2 
         Caption         =   "Folio Fiscal"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Autor: L.I. Samuel Muñoz Chavez desarrollador de Facturación Moderna.
'Última revisión 03/07/2012
'Referencias necesarias: Microsoft ActiveX Data Objects 2.5 Library y Microsoft XML, v3.0

Private Sub cmdCancelar_Click()
    Dim result As VbMsgBoxResult
    result = MsgBox("Realmente desea cancelar el comprobante con folio fiscal:" + txtFolioFiscal.Text, vbYesNo, "Cancelación")
    If result = vbYes Then
        Dim XMLHTTPTimbrarCFDI As XMLHTTP30 'Objetos encargados de realizar las peticiones http al web service de Facturación Moderna
        Dim CFDIBase64, PDFBase64, responseServer, passwordUser, rfc, userId, UUID, layoutBase64, CDFIXML As String
        On Error GoTo Form_LoadError 'Manejador de excepciones
                
        'A continuación se definen las credenciales de acceso al Web Service, en cuanto se active su servicio deberá cambiar esta información por sus claves de acceso en modo productivo
        userId = "UsuarioPruebasWS"
        rfc = "ESI920427886"
        passwordUser = "b9ec2afa3361a59af4b4d102d3f704eabdf097d4"
        UUID = txtFolioFiscal.Text 'Folio fiscal a cancelar (uuid)
        
        Set XMLHTTPTimbrarCFDI = New XMLHTTP30
        XMLHTTPTimbrarCFDI.Open "POST", "https://t2demo.facturacionmoderna.com/timbrado/soap", False
        XMLHTTPTimbrarCFDI.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        XMLHTTPTimbrarCFDI.setRequestHeader "SOAPAction", "https://t2demo.facturacionmoderna.com/timbrado/soap" ' Dirección del web service
        XMLHTTPTimbrarCFDI.send "<?xml version=""1.0"" encoding=""UTF-8""?><env:Envelope xmlns:env=""http://www.w3.org/2003/05/soap-envelope"" xmlns:ns1=""https://t2demo.facturacionmoderna.com/timbrado/soap"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:enc=""http://www.w3.org/2003/05/soap-encoding""><env:Body><ns1:requestCancelarCFDI env:encodingStyle=""http://www.w3.org/2003/05/soap-encoding""><param0 xsi:type=""enc:Struct""><CancelarCFDI xsi:type=""enc:Struct""><UUID xsi:type=""xsd:string"">" + UUID + "</UUID></CancelarCFDI><UserPass xsi:type=""xsd:string"">" + passwordUser + "</UserPass><UserID xsi:type=""xsd:string"">" + userId + "</UserID><emisorRFC xsi:type=""xsd:string"">" + rfc + "</emisorRFC></param0></ns1:requestCancelarCFDI></env:Body></env:Envelope>" ' Soap request
        While XMLHTTPTimbrarCFDI.readyState <> 4
        Wend
        
        responseServer = XMLHTTPTimbrarCFDI.responseText 'Respuesta del Web Service
        Dim xmldoc As DOMDocument30
        Set xmldoc = New DOMDocument30
        If xmldoc.loadXML(responseServer) Then ' Creamos un objeto capas de recorrer el xml de respuesta del servidor para mayor facilidad.
            If xmldoc.getElementsByTagName("env:Fault").length >= 1 Then 'Si nos retorna un error el WS (Soap Fault) lo visualizamos
                MsgBox xmldoc.getElementsByTagName("env:Text").Item(0).Text, vbOKOnly, "Soap Fault"
                
            Else 'Mensaje retornado por el WS
                mensajeCancelacion = xmldoc.getElementsByTagName("Message").Item(0).Text
                codigoCancelacion = xmldoc.getElementsByTagName("Code").Item(0).Text
                MsgBox "[" + codigoCancelacion + "]" + mensajeCancelacion
        End If
        'xmldoc.save "c:/soapResponse.xml"
        Else
            MsgBox "Ha ocurrido un error."
        End If
        Resume Next
Form_LoadError:         'Manejador de errores en tiempo de ejecución
        If Err.Number <> 20 Then
            MsgBox "Form_Load " & Err.Number & ":" & Err.Description, vbCritical, "Error de sistema"
        End If
    End If
End Sub

Private Sub cmdTimbrar_Click()
    Dim XMLHTTPTimbrarCFDI As XMLHTTP30 'Objetos encargados de realizar las peticiones http al web service de Facturación Moderna
    Dim CFDIBase64, PDFBase64, responseServer, passwordUser, rfc, userId, UUID, layoutBase64, CDFIXML, rutaLayout As String
    On Error GoTo Form_LoadError 'Manejador de excepciones
    rutaLayout = txtLayout.Text
    Open rutaLayout For Input As #1 'Lectura del layout contenedor del comprobante
    Dim Linea As String, layoutTXT As String
        Do Until EOF(1)
            Line Input #1, Linea
            layoutTXT = layoutTXT + Linea + vbCrLf
        Loop
    Close
    
    Dim strData As String
    layoutBase64 = EncodeBase64(StrConv(layoutTXT, vbFromUnicode)) 'Transformación del Layout a formato base64
    passwordUser = "b9ec2afa3361a59af4b4d102d3f704eabdf097d4"
    rfc = "ESI920427886"
    userId = "UsuarioPruebasWS"
    
    
    Set XMLHTTPTimbrarCFDI = New XMLHTTP30
    XMLHTTPTimbrarCFDI.Open "POST", "https://t2demo.facturacionmoderna.com/timbrado/soap", False
    XMLHTTPTimbrarCFDI.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    XMLHTTPTimbrarCFDI.setRequestHeader "SOAPAction", "https://t2demo.facturacionmoderna.com/timbrado/soap" ' Dirección del web service
    XMLHTTPTimbrarCFDI.send "<?xml version=""1.0"" encoding=""UTF-8""?><env:Envelope xmlns:env=""http://www.w3.org/2003/05/soap-envelope"" xmlns:ns1=""https://t2demo.facturacionmoderna.com/timbrado/soap"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:enc=""http://www.w3.org/2003/05/soap-encoding""><env:Body><ns1:requestTimbrarCFDI env:encodingStyle=""http://www.w3.org/2003/05/soap-encoding""><param0 xsi:type=""enc:Struct""><UserPass xsi:type=""xsd:string"">" + passwordUser + "</UserPass><UserID xsi:type=""xsd:string"">" + userId + "</UserID><emisorRFC xsi:type=""xsd:string"">" + rfc + "</emisorRFC><text2CFDI xsi:type=""xsd:string"">" + layoutBase64 + "</text2CFDI></param0></ns1:requestTimbrarCFDI></env:Body></env:Envelope>" 'Soap request
    While XMLHTTPTimbrarCFDI.readyState <> 4
    Wend
    
    responseServer = XMLHTTPTimbrarCFDI.responseText ' Respuesta del web service
    Dim xmldoc As DOMDocument30
    Set xmldoc = New DOMDocument30
    If xmldoc.loadXML(responseServer) Then ' Creamos un objeto capaz de acceder a los nodos de la respuesta en formato XML para mayor facilidad
        If xmldoc.getElementsByTagName("env:Fault").length >= 1 Then ' Buscamos si contiene un mensaje de error(Soap Fault)
            MsgBox xmldoc.getElementsByTagName("env:Text").Item(0).Text
            End
        Else
            CFDIBase64 = xmldoc.getElementsByTagName("xml").Item(0).Text ' En caso de éxito obtenemos el nodo xml contenedor del CFDI
            PDFBase64 = xmldoc.getElementsByTagName("pdf").Item(0).Text  ' Obtenemos la representación impresa del CFDI en formato PDF
            CFDIXML = StrConv(DecodeBase64(CFDIBase64), vbUnicode)
            Dim cfdiXmlDoc As DOMDocument30
            Set cfdiXmlDoc = New DOMDocument30
            If cfdiXmlDoc.loadXML(CFDIXML) Then
                Dim xmlNode As IXMLDOMNode
                Set xmlNode = cfdiXmlDoc.documentElement.getElementsByTagName("tfd:TimbreFiscalDigital").Item(0)
                UUID = xmlNode.Attributes.getNamedItem("UUID").Text ' A manera de ejemplo se almacena el XML y PDF con el folio fiscal(UUID) contenido en el xml
                Dim nameCFDI As String
                nameCFDI = UUID
                Open "c:\" + nameCFDI + ".xml" For Binary Access Write As 2 'Almacenamiento del CFDI en formato xml en C:\
                Put #2, , DecodeBase64(CFDIBase64)
                
                Open "c:\" + nameCFDI + ".pdf" For Binary Access Write As 3 'Almacenamiento de la representación impresa del CFDI en formato pdf en C:\
                Put #3, , DecodeBase64(PDFBase64)
                Close
            End If
        End If
    Else
        MsgBox "Ha ocurrido un error al cargar la respuesta del WS."
        'xmldoc.save "c:/soapResponse.xml"
    End If
    Resume Next
Form_LoadError: 'Manejador de errores en tiempo de ejecución
    If Err.Number <> 20 Then
        MsgBox "Form_Load " & Err.Number & ":" & Err.Description
    Else
        MsgBox "CFDI en formato xml y pdf generados correctamente en C:\" + nameCFDI + ".xml|.pdf"
    End If
End Sub

Private Function EncodeBase64(ByRef arrData() As Byte) As String 'Método utilitario para la codificación del layout a base64
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Private Function DecodeBase64(ByVal strData As String) As Byte() 'Método utilitario para la decodificación de la respuesta en base64
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue
    Set objNode = Nothing
    Set objXML = Nothing
End Function

