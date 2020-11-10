component accessors="true" output="false" {

	public any function init(string pythonExePath='C:\_docs\python3') {
		// actual path 2 python
		variables.pythonPath = arguments.pythonExePath;

		// place where to put all files 4 python
		variables.pythonAppPath = expandPath('../python');
		variables.pythonFilePath = variables.pythonAppPath&'\pyfiles';
		variables.pythonPatchPath = variables.pythonAppPath&'\patchfiles';
		variables.jsonFilePath = variables.pythonAppPath&'\jsonfiles';
		variables.wordTemplatePath = variables.pythonAppPath&'\wordtemplates';
		variables.wordSavedFilePath = variables.pythonAppPath&'\wordsaved';

		if (!directoryExists(variables.pythonAppPath)) directoryCreate(variables.pythonAppPath);
		if (!directoryExists(variables.pythonFilePath)) directoryCreate(variables.pythonFilePath);
		if (!directoryExists(variables.pythonPatchPath)) directoryCreate(variables.pythonPatchPath);
		if (!directoryExists(variables.jsonFilePath)) directoryCreate(variables.jsonFilePath);
		if (!directoryExists(variables.wordTemplatePath)) directoryCreate(variables.wordTemplatePath);
		if (!directoryExists(variables.wordSavedFilePath)) directoryCreate(variables.wordSavedFilePath);

		// create python file 4 generating a word-document
		variables.pythonFile = createPythonFile();

		return this;
	}

	public boolean function genWordDoc(string jsonFile="test", boolean downloadfile=false, any docStruct, string templatefile) hint="generating Word-documents with Python 3" {
		param name="arguments.docStruct" default={ "text": "Hello World" };
		param name="arguments.templatefile" default="demo.docx";

		var pythonFile = variables.pythonFilePath&"/"&variables.pythonFile;
		var fileName = 'example.docx';

		// create a json file; essential for generating word-document
		var jsonFile = variables.jsonFilePath&"/"&arguments.jsonFile&".json";
		var docStruct = !isStruct(arguments.docStruct)? deserializeJSON(arguments.docStruct) : arguments.docStruct;
		docStruct = isStruct(docStruct)? docStruct : {};
		var result = {
			"uploadpath" : variables.wordSavedFilePath
			, "filename" : fileName
			, "docStruct" : docStruct
			, "templatefile" : variables.wordTemplatePath&"\"&arguments.templatefile
		};
		fileWrite(jsonFile, serializeJSON(result));

		// create a bat file to execute.
		var createdBatchFile = "genDocument.bat";
		var execFile = variables.pythonPatchPath&"\"&createdBatchFile;
		var pythonString = variables.pythonPath&'/python '&pythonFile&' '&jsonFile;
		fileWrite(execFile, pythonString);
		cfexecute(name=execFile variable="endResult" timeout="10000");

		// download the generated word-document
		if (arguments.downloadfile) {
			var downloadPath = variables.wordSavedFilePath&"/"&fileName;
			cfheader(name="Content-Disposition" value="attachment;filename=""#fileName#""");
			cfcontent(file="#downloadPath#" reset="true" type="rc.filetType");
		}

		// delete both .json and .bat file
		fileDelete(jsonFile);
		fileDelete(execFile);

		return true;
	}


	private string function createPythonFile(string fileName='generateWordDoc') {
		var fileName = arguments.fileName&'.py';
		var pythonFile = variables.pythonFilePath&'\'&fileName;

		if(!fileExists(pythonFile)) {
			var strFile = "import json, sys"&Chr(13);
				strFile &= "from docxtpl import DocxTemplate"&Chr(13);
				strFile &= "strFile = sys.argv[1]"&Chr(13);
				strFile &= "with open(strFile, 'r', encoding='utf-8') as f:"&Chr(13);
				strFile &= "	data = f.read()"&Chr(13);
				strFile &= "	y = json.loads(data)"&Chr(13);
				strFile &= "	uploadpath = y['uploadpath']+'\\'"&Chr(13);
				strFile &= "	filename =  y['filename']"&Chr(13);
				strFile &= "	docInput = y['docStruct']"&Chr(13);
				strFile &= "	templateFile = y['templatefile']"&Chr(13);
				strFile &= "	context = docInput"&Chr(13);
				strFile &= "	template = DocxTemplate(templateFile)"&Chr(13);
				strFile &= "	template.render(context)"&Chr(13);
				strFile &= "	template.save(uploadpath+filename)";
			fileWrite(pythonFile, strFile);
		}
		return fileName;
	}
}