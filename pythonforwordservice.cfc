component accessors="true" output="false" {

	/**
	 * @Hint: Initial setup for creating the necessary directories and python-script
	 * @param pythonExePath: path to the Python directory
	*/
	public any function init(string pythonExePath='') {
		// actual path to python
		variables.pythonPath = arguments.pythonExePath;

		// place where to put all files for python
		if (variables.pythonPath != '') {
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

			// create python file for generating a static word-document
			variables.pythonFile = createPythonFile();
		}

		return this;
	}

	/**
	 *	@Hint: Generates Word-document
	 *	@param jsonFile: name of the json file wich will temporarily be saved.
	 *	@param downloadfile: yes or no downloading document; default false
	 *	@param docStruct: structure holding all variables for filleing the tmplatefile
	 *	@param templatefile: word-document templatefile holding all variables to fill
	 *
	*/
	public boolean function genWordDoc(string jsonFile="test", boolean downloadfile=false, any docStruct, string templatefile) hint="generating Word-documents with Python 3" {
		param name="arguments.docStruct" default={ "text": "Hello World" };
		param name="arguments.templatefile" default="demo.docx";

		// is path to the Python directory present
		if (variables.pythonPath != '') {
			var pythonFile = variables.pythonFilePath&"/"&variables.pythonFile;
			var fileName = 'example.docx';

			// create a json file; essential for generating word-document
			var jsonFileLocal = variables.jsonFilePath&"/"&arguments.jsonFile&".json";
			var docStructLocal = !isStruct(arguments.docStruct)? deserializeJSON(arguments.docStruct) : arguments.docStruct;
			var templateFileLocal = variables.wordTemplatePath&"\"&arguments.templatefile;
			docStructLocal = isStruct(docStructLocal)? docStructLocal : {};
			var result = {
				"uploadpath" : variables.wordSavedFilePath
				, "filename" : fileName
				, "docStruct" : docStructLocal
				, "templatefile" : templateFileLocal
			};
			fileWrite(jsonFileLocal, serializeJSON(result));

			// create a .bat file to execute
			var createdBatchFile = "genDocument.bat";
			var execFile = variables.pythonPatchPath&"\"&createdBatchFile;
			var pythonString = variables.pythonPath&'/python '&pythonFile&' '&jsonFileLocal;
			fileWrite(execFile, pythonString);
			
			// when python.exe and word-document templatefile exists, execute .bat file
			var pythonFile = variables.pythonPath&'/python.exe';
			if (fileExists(pythonFile) && fileExists(templateFileLocal)) {
				try {
					cfexecute(name=execFile, variable="endResult", timeout="10000");
				} catch(any a) {
					var strError = '<strong>NOTE: '&lcase(a.Message)&'</strong><br /><em>'&a.Detail&'</em><br />';
					writeOutput(strError);
				}
			}

			// delete both .json and .bat file
			fileDelete(jsonFileLocal);
			fileDelete(execFile);

			// download the generated word-document
			if (arguments.downloadfile) {
				var downloadPath = variables.wordSavedFilePath&"/"&fileName;
				cfheader(name='Content-Disposition', value='attachment; filename=#fileName#');
				cfcontent(file='#downloadPath#', reset='true', type='application/msword');
			}
			return true;
		}
		else {
			return false;
		}
	}

	/**
	 * @Hint: Python-script wich will be running to generate the static Word-document
	 * @param fileName: name for the python-file
	*/
	private string function createPythonFile(string fileName='generateWordDoc') {
		var fileNameLocal = arguments.fileName&'.py';
		var pythonFile = variables.pythonFilePath&'\'&fileNameLocal;
		var strFile = "";

		if(!fileExists(pythonFile)) {
			strFile &= "import json, sys"&Chr(13);
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
		return fileNameLocal;
	}

}
