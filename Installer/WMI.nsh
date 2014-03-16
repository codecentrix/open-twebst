!ifndef __WMI_Included__
!define __WMI_Included__
!define WMIKillProcess "!Insertmacro WMIKillProcess"
 
!macro WMIKillProcess _ProcName
Push $0
Push $1
nsexec::exectostack "wmic process where name='${_ProcName}' delete"
pop $0
pop $0
StrCpy $0 $0 2
${If} $0 == "No"
pop $0
pop $0
push 0
${Else}
pop $0
pop $0
push 1
${EndIf}
!macroend
!endif
