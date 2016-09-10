**********
**********
ExitResponse structure is used for function InsertGroupMember return value
    Public Structure ExitResponse
        Public Exit_status As Integer   
        Public description As String
    End Structure

Exit_status = 0|1 ; 0 for success and 1 for faillure
description = error message in case there is any error
*************
*************




