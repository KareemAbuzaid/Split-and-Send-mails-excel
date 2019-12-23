Sub send()
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'VBA function to read file names from the first column of a sheet,            '
  'and the recepients from the 2nd column, then create e-mails attaching        '
  'the files and encrypting those e-mails, after that the e-mails are either    '
  'sent or just displayed.The file names should not include .xlsx as this       '
  'is added by in the method. user emails should be                             '
  'separated by a `;`.                                                          '
  'Anil Coelho <anil.coelho@siemens.com>                                        '
  'Kareem Abuzaid <kareem.abuzaid.ext@siemens.com>                              '
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'variable to be used for looping over rows
  Dim i As Long

  'variable to store the name of the file from the 1st column in the excel sheet
  Dim fileName As String

  'variable to store the e-mail address to which the e-mails will be sent
  Dim emails As String

  'variables to store the e-mail application object
  Dim emailApplication As Object
  Set emailApplication = CreateObject("Outlook.Application")

  'variable to store the path of this file
  Dim filePath As String
  filePath = Application.ActiveWorkbook.Path

  'variable to be used when reading text files
  Dim iFile As Integer
    
  'variable to store the path of the file that has the e-mail subject.
  'The file is stored in the same path and is called `subject.txt`
  Dim subjectFilePath As String
  subjectFilePath = Application.ActiveWorkbook.Path & "/" & "subject.txt"
  
  'variable to store the e-mail subject line
  Dim emailSubject As String
  
  'open body file and read it's contents into the variable called emailSubject
  iFile = FreeFile
  Open subjectFilePath For Input As #iFile
  emailSubject = Input(LOF(iFile), iFile)
  Close #iFile
  
  'variable to store the path of the file that has the e-mail body.
  'The file is stored in the same path and is called `body.txt`
  Dim bodyFilePath As String
  bodyFilePath = Application.ActiveWorkbook.Path & "/" & "body.txt"
  
  'variable to store the e-mail body
  Dim emailBody As String
  
  'open body file and read it's contents into the variable called emailBody
  iFile = FreeFile
  Open bodyFilePath For Input As #iFile
  emailBody = Input(LOF(iFile), iFile)
  Close #iFile
  
  'MAPI property PR_SECURITY_FLAGS used for e-mail encryption
  Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"


  'loop over all rows and extract file names and the emails
  For i = 1 To Range("A1").End(xlDown).Row
    
    'variable to store the e-mail object
    Dim emailItem As Object
    Set emailItem = emailApplication.CreateItem(0)
    
    fileName = Cells(i, 1).Value
    emails = Cells(i, 2).Value
    
    'create variable to store file path and name
    Dim filePathName As String
    filePathName = filePath & "/" & fileName & ".xlsx"
    
    '''''''''''''''''
    'buld the e-mail'
    '''''''''''''''''
    emailItem.To = emails
    emailItem.Subject = emailSubject
    emailItem.Body = emailBody
    'change the sender of the e-mail
    emailItem.SentOnBehalfOfName = "<add the e-mail address here>"
    
    'attach the file
    emailItem.Attachments.Add (filePathName)
    
    ''''''''''''''''''''''''''
    'do the e-mail encryption'
    ''''''''''''''''''''''''''
    Dim prop As Long
    prop = CLng(emailItem.PropertyAccessor.GetProperty(PR_SECURITY_FLAGS))
    
    'this does encrypting
    ulFlags = ulFlags Or &H1 ' SECFLAG_ENCRYPTED
    
    'this does signing and can be removed if not needed
    ulFlags = ulFlags Or &H2 ' SECFLAG_SIGNED
    emailItem.PropertyAccessor.SetProperty PR_SECURITY_FLAGS, ulFlags
    
    ''''''''''''''''''''''''''''
    'send or dispaly the e-mail'
    ''''''''''''''''''''''''''''
    
    'send the Email
    'use this OR .Display, but not both together.
    'emailItem.send

    'display the Email so the user can change it as desired before sending it.
    'use this OR .Send, but not both together.
    emailItem.Display
  Next i
End Sub

