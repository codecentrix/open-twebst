var fileSystem  = WScript.CreateObject("Scripting.FileSystemObject");
var xmlDir      = fileSystem.GetFolder(".\\XML\\");
var xmlFilesCol = xmlDir.Files;
var fileEnum    = new Enumerator(xmlFilesCol);


// Create the parser for XML file.
var xmlDoc             = new ActiveXObject("MSXML2.DOMDocument.3.0");
xmlDoc.validateOnParse = false;
xmlDoc.async           = false;

// Create the parser for topic XSL file.
var xslTopicDoc   = new ActiveXObject("MSXML2.DOMDOCUMENT.3.0");
xslTopicDoc.async = false;

// Create the parser for overview XSL file.
var xslOverviewDoc   = new ActiveXObject("MSXML2.DOMDOCUMENT.3.0");
xslOverviewDoc.async = false;


// Load the topic xsl document.
var xslTopicFile = "";
xslTopicFile += fileSystem.GetFolder(".\\XSL\\") + "\\topicCHM.xsl";
if (!xslTopicDoc.load(xslTopicFile))
{
	WScript.Echo("XSL Loading Error: " + xslTopicDoc.parseError.reason);
	WScript.Quit(1);
}

if (xslTopicDoc.parseError.errorCode != 0)
{
	WScript.Echo ("XSL Parse Error : " + xslTopicDoc.parseError.reason);
	WScript.Quit(1);
}


// Load the overview xsl document.
var xslOverviewFile = "";
xslOverviewFile += fileSystem.GetFolder(".\\XSL\\") + "\\overviewCHM.xsl";
if (!xslOverviewDoc.load(xslOverviewFile))
{
	WScript.Echo("XSL Loading Error: " + xslOverviewDoc.parseError.reason);
	WScript.Quit(1);
}

if (xslOverviewDoc.parseError.errorCode != 0)
{
	WScript.Echo ("XSL Parse Error : " + xslOverviewDoc.parseError.reason);
	WScript.Quit(1);
}


var topicXslDate    = fileSystem.GetFile(xslTopicFile).DateLastModified;
var overviewXslDate = fileSystem.GetFile(xslOverviewFile).DateLastModified;
var maxXslDate = topicXslDate;
if (maxXslDate < overviewXslDate)
{
	maxXslDate = overviewXslDate;
}

// Browse the xml files.
while (!fileEnum.atEnd())
{
	var xmlFile = ""
	xmlFile += fileEnum.item();

	// Get the name of the html file to be created.
	var xmlFileName = "";
	xmlFileName += fileEnum.item().Name;
	
	var baseFileName      = xmlFileName.substr(0, xmlFileName.length - 3);
	var htmlFileName      = baseFileName + "htm";
	var htmlFileNameFull  = htmlFileName;
	var destinationFolder = fileSystem.GetFolder(".\\HTML\\");
	
	htmlFileNameFull = destinationFolder + "\\" + htmlFileNameFull;
	
	if (AlreadyCreated(xmlFile, htmlFileNameFull, maxXslDate))
	{
		// The HTML file is already created. Skip to the next xml file.
		fileEnum.moveNext();
		continue;
	}

	if (!xmlDoc.load(xmlFile))
	{
		WScript.Echo("XSL Loading Error: " + xmlDoc.parseError.reason);
		WScript.Quit(1);
	}

	if (xmlDoc.parseError.errorCode != 0)
	{
		WScript.Echo ("XSL Parse Error : " + xmlDoc.parseError.reason);
		WScript.Quit(1);
	}

	var xslDoc = null;
	if (xmlDoc.documentElement.tagName == "topic")
	{
		xslDoc = xslTopicDoc;
	}
	else
	{
		xslDoc = xslOverviewDoc;
	}

	try
	{
		// Add color-sytanx to XML document.
		AddSyntaxColorAndReferences(xmlDoc);
		var newText = xmlDoc.documentElement.xml;
	
		// Generate the html text from xml.
		var htmlText = xmlDoc.transformNode(xslDoc.documentElement);
		
		try
		{
			fileSystem.DeleteFile(htmlFileNameFull, true);
		}
		catch (err)
		{
		}

		var htmlFile = fileSystem.CreateTextFile(htmlFileNameFull, false);
		htmlFile.Write(htmlText);
		htmlFile.Close();
	}
	catch(err)
	{
		WScript.Echo("Transformation Error in " + xmlFile + " : " + err.number + "*" + err.description);
		WScript.Quit(1);
	}

	fileEnum.moveNext();
}

if (WScript.Arguments.length == 0)
{
    WScript.Echo("HTML files properly generatd !");
}


function AddSyntaxColorAndReferences(xmlDocument)
{
	var xmlNewNodeDoc             = new ActiveXObject("MSXML2.DOMDocument.3.0");
	xmlNewNodeDoc.validateOnParse = false;
	xmlNewNodeDoc.async           = false;

	// Add color syntax to <jscode> elements.
	var wordRegExp    = new RegExp("\\w+", "g");
	var stringRegExp  = new RegExp("\"[^\"]*\"", "g");
	var commentRegExp = new RegExp("// [^\n]*", "g");

	var jsCodeCollection = xmlDocument.selectNodes("//jscode");
	for (i = 0; i < jsCodeCollection.length; ++i)
	{
		var currentCode  = jsCodeCollection(i);
		var textToModify = currentCode.xml;
		var modifiedText = textToModify.replace(wordRegExp, ColorJsKeyword);
		modifiedText     = modifiedText.replace(commentRegExp, "<jscomment>$&</jscomment>");
		modifiedText     = modifiedText.replace(stringRegExp, "<string>$&</string>");

		xmlNewNodeDoc.loadXML(modifiedText);
		currentCode.parentNode.replaceChild(xmlNewNodeDoc.documentElement, currentCode);
	}


	// Add <reference> tags.
	var refTags       = "//description|//apply|//arguments|//remarks|//jscode|//seealso|//method|//property|//constant";
	var libKeywordCol = xmlDocument.selectNodes(refTags);

	for (j = 0; j < libKeywordCol.length; ++j)
	{
		var currentNode  = libKeywordCol(j);
		var textToModify = currentNode.xml;
		var modifiedText = textToModify.replace(wordRegExp, ColorLibraryKeyword);

		// Hack for lastError.xml where we have href="core.htm" and we need "Core.htm"
		// If we write Core.htm then Core will be replaced by <reference> etc construction.
		if (currentNode.tagName == "remarks")
		{
			modifiedText = modifiedText.replace("core.htm", "Core.htm");
		}

		xmlNewNodeDoc.loadXML(modifiedText);
		currentNode.parentNode.replaceChild(xmlNewNodeDoc.documentElement, currentNode);
	}
}


function ModifyWord(word, keywordArray, tagToAdd)
{
	for (i = 0; i < keywordArray.length; ++i)
	{
		if (word == keywordArray[i])
		{
			return "<" + tagToAdd + ">" + word + "</" + tagToAdd + ">";
		}
	}

	return word;
}


function ColorLibraryKeyword(word)
{
	// The list doesn't contain: item, length, core, document, title, text and url keywords because they are too general.
	var libraryKeywords  =
		[
			"AddSelection",           "AddSelectionRange",    "asyncHtmlEvents",   "AttachToHWND",          "AttachToNativeBrowser", "AttachToNativeElement",
			"AttachToNativeFrame",    "Browser",              "Check",             "ClearSelection",        "Click",
			"Close",                  "closeBrowserPopups",   "ClosePopup",        "ClosePrompt",           "Core",
			"Element",			      "ElementList",          "FindAllElements",   "FindBrowser",           "FindChildElement",
			"FindChildFrame",		  "FindChildrenElements", "FindElement",       "FindFrame",             "FindModalHtmlDialog",
			"FindModelessHtmlDialog", "FindParentElement",    "foregroundBrowser", "Frame",                 "GetAllSelectedOptions",
			"GetAttribute",			  "GetClipboardText",     "Highlight",         "IEVersion",             "InputText",
			"isChecked",			  "isLoading",			  "lastError",		   "loadTimeout",           "loadTimeoutIsError",
			"Navigate",			      "nativeBrowser",        "nativeElement",     "nativeFrame",           "navigationError",
			"nextSiblingElement",	  "parentBrowser",        "parentElement",     "parentFrame",           "previousSiblingElement",
			"productName",			  "productVersion",       "RemoveAttribute",   "Reset",                 "RightClick",
			"searchTimeout",          "Select",               "SelectRange",       "selectedOption",        "SetAttribute", 
			"SetClipboardText",       "StartBrowser",         "tagName",           "topFrame",              "Uncheck",
			"useHardwareInputEvents", "WaitToLoad"
	    ];

	var result = ModifyWord(word, libraryKeywords, "reference");
	if (result != word)
	{
		return result;
	}
	
	var coreConstants = ["OK_ERROR", "FAIL_ERROR", "INVALID_ARG_ERROR", "LOAD_TIMEOUT_ERROR", "INDEX_OUT_OF_BOUND_ERROR",
						 "BROWSER_CONNECTION_LOST_ERROR", "INVALID_OPERATION_ERROR", "NOT_FOUND_ERROR"]
	for (i = 0; i < coreConstants.length; ++i)
	{
		if (word == coreConstants[i])
		{
			return "<reference href=\"Core.htm#constants\">" + word + "</reference>";
		}
	}

	return word;
}


function ColorJsKeyword(word)
{
	var jsKeywords = ["break", "delete", "function", "return", "typeof", "case", "do", "if", "switch", "var",
	                  "catch", "else", "in", "this", "void", "continue", "false", "instanceof", "throw", "while",
	                  "debugger", "finally", "new", "true", "with", "default", "for", "null", "try"];
	return ModifyWord(word, jsKeywords, "keyword");
}


function AlreadyCreated(sXmlFileName, sHtmlFileName, xslDate)
{
	var xmlFileDate  = fileSystem.GetFile(sXmlFileName).DateLastModified;
	try
	{
		var htmlFileDate = fileSystem.GetFile(sHtmlFileName).DateLastModified;
		return ((htmlFileDate > xmlFileDate) && (htmlFileDate > xslDate));
	}
	catch (err)
	{
		return false;
	}
}
