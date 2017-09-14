' ==============================================================================================
' Adaptation of JSONToXML() function for enhancements and bugfixes.
' Author: Praveen Nandagiri (pravynandas@gmail.com)
' Enhancement#1: Arrays are now rendered as Text Nodes
'
' Credits:
' Visit: https://stackoverflow.com/a/12171836/1751166
' Author: https://stackoverflow.com/users/881441/stephen-quan
' ==============================================================================================
Const stateRoot = 0
Const stateNameQuoted = 1
Const stateNameFinished = 2
Const stateValue = 3
Const stateValueQuoted = 4
Const stateValueQuotedEscaped = 5
Const stateValueUnquoted = 6
Const stateValueUnquotedEscaped = 7

Function JSONToXML(json)
  Dim dom, xmlElem, i, ch, state, name, value
  Set dom = CreateObject("Microsoft.XMLDOM")
  state = stateRoot
  For i = 1 to Len(json)
    ch = Mid(json, i, 1)
    Select Case state
    Case stateRoot
      Select Case ch
      Case "["
        If dom.documentElement is Nothing Then
          Set xmlElem = dom.CreateElement("ARRAY")
          Set dom.documentElement = xmlElem
        Else
          Set xmlElem = XMLCreateChild(xmlElem, "ARRAY")
        End If
      Case "{"
        If dom.documentElement is Nothing Then
          Set xmlElem = dom.CreateElement("ROOT")
          Set dom.documentElement = xmlElem
        Else
          Set xmlElem = XMLCreateChild(xmlElem, "OBJECT")
        End If
      Case """"
        state = stateNameQuoted 
        name = ""
      Case "}"
        Set xmlElem = xmlElem.parentNode
      Case "]"
        Set xmlElem = xmlElem.parentNode
      End Select
    Case stateNameQuoted 
      Select Case ch
      Case """"
        state = stateNameFinished
      Case Else
        name = name + ch
      End Select
    Case stateNameFinished
      Select Case ch
      Case ":"
        value = ""
        State = stateValue
      Case Else						'@@Enhancement#1: Handling Array values
        Set xmlitem = dom.createTextNode(name)
	xmlElem.appendChild(xmlitem)
        State = stateRoot					
      End Select
    Case stateValue
      Select Case ch
      Case """"
        State = stateValueQuoted
      Case "{"
        Set xmlElem = XMLCreateChild(xmlElem, name)
        State = stateRoot
      Case "["
        Set xmlElem = XMLCreateChild(xmlElem, name)
        State = stateRoot
      Case " "
      Case Chr(9)
      Case vbCr
      Case vbLF
      Case Else
        value = ch
        State = stateValueUnquoted
      End Select
    Case stateValueQuoted
      Select Case ch
      Case """"
        xmlElem.setAttribute name, value
        state = stateRoot
      Case "\"
        state = stateValueQuotedEscaped
      Case Else
        value = value + ch
      End Select
    Case stateValueQuotedEscaped ' @@TODO: Handle escape sequences
      value = value + ch
      state = stateValueQuoted
    Case stateValueUnquoted
      Select Case ch
      Case "}"
        xmlElem.setAttribute name, value
        Set xmlElem = xmlElem.parentNode
        state = stateRoot
      Case "]"
        xmlElem.setAttribute name, value
        Set xmlElem = xmlElem.parentNode
        state = stateRoot
      Case ","
        xmlElem.setAttribute name, value
        state = stateRoot
      Case "\"
         state = stateValueUnquotedEscaped
      Case Else
        value = value + ch
      End Select
    Case stateValueUnquotedEscaped ' @@TODO: Handle escape sequences
      value = value + ch
      state = stateValueUnquoted
    End Select
  Next
  Set JSONToXML = dom
End Function

Function XMLCreateChild(xmlParent, tagName)
  Dim xmlChild
  If xmlParent is Nothing Then
    Set XMLCreateChild = Nothing
    Exit Function
  End If
  If xmlParent.ownerDocument is Nothing Then
    Set XMLCreateChild = Nothing
    Exit Function
  End If
  Set xmlChild = xmlParent.ownerDocument.createElement(tagName)
  xmlParent.appendChild xmlChild
  Set XMLCreateChild = xmlChild
End Function
