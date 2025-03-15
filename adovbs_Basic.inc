<%
'--------------------------------------------------------------------
' Microsoft ADO
'
' Copyright (c) 1996-1998 Microsoft Corporation.
'
'
'
' ADO constants include file for VBScript
'
'--------------------------------------------------------------------

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- ParameterAttributesEnum Values ----
Const adParamSigned = &H0010
Const adParamNullable = &H0040
Const adParamLong = &H0080

'---- ParameterDirectionEnum Values ----
Const adParamUnknown = &H0000
Const adParamInput = &H0001
Const adParamOutput = &H0002
Const adParamInputOutput = &H0003
Const adParamReturnValue = &H0004

'---- CommandTypeEnum Values ----
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004
Const adCmdFile = &H0100
Const adCmdTableDirect = &H0200

'---- DataTypeEnum Values ----
Const adEmpty = 0
Const adTinyInt = 16
Const adSmallInt = 2
Const adInteger = 3
Const adBigInt = 20
Const adUnsignedTinyInt = 17
Const adUnsignedSmallInt = 18
Const adUnsignedInt = 19
Const adUnsignedBigInt = 21
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14
Const adNumeric = 131
Const adBoolean = 11
Const adError = 10
Const adUserDefined = 132
Const adVariant = 12
Const adIDispatch = 9
Const adIUnknown = 13
Const adGUID = 72
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205
Const adChapter = 136
Const adFileTime = 64
Const adPropVariant = 138
Const adVarNumeric = 139
Const adArray = &H2000

'---- RecordStatusEnum Values ----
Const adRecOK = &H0000000
Const adRecNew = &H0000001
Const adRecModified = &H0000002
Const adRecDeleted = &H0000004
Const adRecUnmodified = &H0000008
Const adRecInvalid = &H0000010
Const adRecMultipleChanges = &H0000040
Const adRecPendingChanges = &H0000080
Const adRecCanceled = &H0000100
Const adRecCantRelease = &H0000400
Const adRecConcurrencyViolation = &H0000800
Const adRecIntegrityViolation = &H0001000
Const adRecMaxChangesExceeded = &H0002000
Const adRecObjectOpen = &H0004000
Const adRecOutOfMemory = &H0008000
Const adRecPermissionDenied = &H0010000
Const adRecSchemaViolation = &H0020000
Const adRecDBDeleted = &H0040000

'---- PersistFormatEnum Values ----
Const adPersistADTG = 0
Const adPersistXML = 1

'---- ConnectModeEnum Values ----
Const adModeUnknown = 0
Const adModeRead = 1
Const adModeWrite = 2
Const adModeReadWrite = 3
Const adModeShareDenyRead = 4
Const adModeShareDenyWrite = 8
Const adModeShareExclusive = &Hc
Const adModeShareDenyNone = &H10
Const adModeRecursive = &H400000

'---- ErrorValueEnum Values ----
Const adErrProviderFailed = &Hbb8
Const adErrInvalidArgument = &Hbb9
Const adErrOpeningFile = &Hbba
Const adErrReadFile = &Hbbb
Const adErrWriteFile = &Hbbc
Const adErrNoCurrentRecord = &Hbcd
Const adErrIllegalOperation = &Hc93
Const adErrCantChangeProvider = &Hc94
Const adErrInTransaction = &Hcae
Const adErrFeatureNotAvailable = &Hcb3
Const adErrItemNotFound = &Hcc1
Const adErrObjectInCollection = &Hd27
Const adErrObjectNotSet = &Hd5c
Const adErrDataConversion = &Hd5d
Const adErrObjectClosed = &He78
Const adErrObjectOpen = &He79
Const adErrProviderNotFound = &He7a
Const adErrBoundToCommand = &He7b
Const adErrInvalidParamInfo = &He7c
Const adErrInvalidConnection = &He7d
Const adErrNotReentrant = &He7e
Const adErrStillExecuting = &He7f
Const adErrOperationCancelled = &He80
Const adErrStillConnecting = &He81
Const adErrInvalidTransaction = &He82
Const adErrUnsafeOperation = &He84
Const adwrnSecurityDialog = &He85
Const adwrnSecurityDialogHeader = &He86
Const adErrIntegrityViolation = &He87
Const adErrPermissionDenied = &He88
Const adErrDataOverflow = &He89
Const adErrSchemaViolation = &He8a
Const adErrSignMismatch = &He8b
Const adErrCantConvertvalue = &He8c
Const adErrCantCreate = &He8d
Const adErrColumnNotOnThisRow = &He8e
Const adErrURLIntegrViolSetColumns = &He8f
Const adErrURLDoesNotExist = &He8f
Const adErrTreePermissionDenied = &He90
Const adErrInvalidURL = &He91
Const adErrResourceLocked = &He92
Const adErrResourceExists = &He93
Const adErrCannotComplete = &He94
Const adErrVolumeNotFound = &He95
Const adErrOutOfSpace = &He96
Const adErrResourceOutOfScope = &He97
Const adErrUnavailable = &He98
Const adErrURLNamedRowDoesNotExist = &He99
Const adErrDelResOutOfScope = &He9a
Const adErrPropInvalidColumn = &He9b
Const adErrPropInvalidOption = &He9c
Const adErrPropInvalidValue = &He9d
Const adErrPropConflicting = &He9e
Const adErrPropNotAllSettable = &He9f
Const adErrPropNotSet = &Hea0
Const adErrPropNotSettable = &Hea1
Const adErrPropNotSupported = &Hea2
Const adErrCatalogNotSet = &Hea3
Const adErrCantChangeConnection = &Hea4
Const adErrFieldsUpdateFailed = &Hea5
Const adErrDenyNotSupported = &Hea6
Const adErrDenyTypeNotSupported = &Hea7
Const adErrProviderNotSpecified = &Hea9
%>
