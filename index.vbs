' ==============================================================================================
' Adaptation of JSONToXML() function for enhancements and bugfixes.
' Author: Praveen Nandagiri (pravynandas@gmail.com)
' Enhancement#1: Arrays are now rendered as Text Nodes
' Enhancement#2: Handled Escape characters (incl. Hex). Refer: http://www.json.org/
'
' Credits:
' Visit: https://stackoverflow.com/a/12171836/1751166
' Author: https://stackoverflow.com/users/881441/stephen-quan
' ==============================================================================================

Class JSONToXML

  Private stateRoot
  Private stateNameQuoted
  Private stateNameFinished
  Private stateValue
  Private stateValueQuoted
  Private stateValueQuotedEscaped
  Private stateValueQuotedEscapedHex
  Private stateValueUnquoted
  Private stateValueUnquotedEscaped

  Private Sub Class_Initialize
    stateRoot = 0
    stateNameQuoted = 1
    stateNameFinished = 2
    stateValue = 3
    stateValueQuoted = 4
    stateValueQuotedEscaped = 5
    stateValueQuotedEscapedHex = 6
    stateValueUnquoted = 7
    stateValueUnquotedEscaped = 8
  End Sub

  Public Function toXml(json)
    Dim dom, xmlElem, i, ch, state, name, value, sHex
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
      Case stateValueQuotedEscaped ' @@Enhancement#2: Handle escape sequences
      If ch = "u" Then	'Four digit hex. Ex: o = 00f8
        sHex = ""
        state = stateValueQuotedEscapedHex
      Else
        Select Case ch
        Case """"
          value = value + """"
        Case "\"
          value = value + "\"
        Case "/"
          value = value + "/"
        Case "b"	'Backspace
          value = value + chr(08)
        Case "f"	'Form-Feed
          value = value + chr(12)
        Case "n"	'New-line (LineFeed(10))
          value = value + vbLF
        Case "r"	'New-line (CarriageReturn/CRLF(13))
          value = value + vbCR
        Case "t"	'Horizontal-Tab (09)
          value = value + vbTab
        Case Else
          'do not accept any other escape sequence
        End Select
        state = stateValueQuoted
      End If
    Case stateValueQuotedEscapedHex
      sHex = sHex + ch
      If len(sHex) = 4 Then
        on error resume next
        value = value + Chr("&H" & sHex)	'Hex to String conversion
        on error goto 0
        state = stateValueQuoted
      End If
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
    Set toXml = dom
  End Function

  Private Function XMLCreateChild(xmlParent, tagName)
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
End Class