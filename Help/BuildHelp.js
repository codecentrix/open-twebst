// Build HTML files.
var shellSystem     = WScript.CreateObject("WScript.Shell");
var chmCompiler     = "\"c:\\Program Files (x86)\\HTML Help Workshop\\hhc.exe\"";


// Generate HTML files.
var outExec = shellSystem.Exec("wscript.exe .\\BuildHTMLs.js nomessage");
while (outExec.Status == 0)
{
	WScript.Sleep(100);
}

// Compile the help file.
var fileSystem        = WScript.CreateObject("Scripting.FileSystemObject");
var createHelpCmdLine = chmCompiler + " \"" + fileSystem.GetFolder(".") + "\\CatHelp.hhp\"";

outExec = shellSystem.Exec(createHelpCmdLine);
while (outExec.Status == 0)
{
	WScript.Sleep(100);
}

if (WScript.Arguments.length == 0)
{
	WScript.Echo("HTML help file properly generated !");
}
