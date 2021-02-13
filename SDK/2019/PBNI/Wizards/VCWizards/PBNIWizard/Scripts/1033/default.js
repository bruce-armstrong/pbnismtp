function OnFinish(selProj, selObj)
{
   try
   {
      var strProjectPath = wizard.FindSymbol("PROJECT_PATH");
      var strProjectName = wizard.FindSymbol("PROJECT_NAME");
      var strClassName = wizard.FindSymbol("PBNI_CLASS_NAME");
      var strSafeProjectName = CreateSafeName(strProjectName);
      var strSafeClassName = CreateSafeName(strClassName);
      
      wizard.AddSymbol("SAFE_PROJECT_NAME", strSafeProjectName);
      wizard.AddSymbol("LOWER_CASE_PROJECT_NAME", strSafeProjectName.toLowerCase());
      wizard.AddSymbol("UPPER_CASE_PROJECT_NAME", strSafeProjectName.toUpperCase());
      wizard.AddSymbol("SAFE_CLASS_NAME", strSafeClassName);
      wizard.AddSymbol("UPPER_CLASS_NAME", strSafeClassName.toUpperCase());
      wizard.AddSymbol("LOWER_CLASS_NAME", strSafeClassName.toLowerCase());

      selProj = CreateCustomProject(strSafeProjectName, strProjectPath);
      AddConfig(selProj, strSafeProjectName);
      AddFilters(selProj);

      var InfFile = CreateCustomInfFile();
      AddFilesToCustomProj(selProj, strSafeProjectName, strProjectPath, InfFile);
      PchSettings(selProj);
      InfFile.Delete();

      selProj.Object.Save();
   }
   catch(e)
   {
      if (e.description.length != 0)
         SetErrorInfo(e);
      return e.number
   }
}

function CreateCustomProject(strProjectName, strProjectPath)
{
   try
   {
      var strProjTemplatePath = wizard.FindSymbol('PROJECT_TEMPLATE_PATH');
      var strProjTemplate = '';
      strProjTemplate = strProjTemplatePath + '\\default.vcproj';

      var Solution = dte.Solution;
      var strSolutionName = "";
      if (wizard.FindSymbol("CLOSE_SOLUTION"))
      {
         Solution.Close();
         strSolutionName = wizard.FindSymbol("VS_SOLUTION_NAME");
         if (strSolutionName.length)
         {
            var strSolutionPath = strProjectPath.substr(0, strProjectPath.length - strProjectName.length);
            Solution.Create(strSolutionPath, strSolutionName);
         }
      }

      var strProjectNameWithExt = '';
      strProjectNameWithExt = strProjectName + '.vcproj';

      var oTarget = wizard.FindSymbol("TARGET");
      var prj;
      if (wizard.FindSymbol("WIZARD_TYPE") == vsWizardAddSubProject)  // vsWizardAddSubProject
      {
         var prjItem = oTarget.AddFromTemplate(strProjTemplate, strProjectNameWithExt);
         prj = prjItem.SubProject;
      }
      else
      {
         prj = oTarget.AddFromTemplate(strProjTemplate, strProjectPath, strProjectNameWithExt);
      }
      return prj;
   }
   catch(e)
   {
      throw e;
   }
}

function AddFilters(proj)
{
   try
   {
      // Add the folders to your project
      var strSrcFilter = wizard.FindSymbol('SOURCE_FILTER');
      var group1 = proj.Object.AddFilter('Source Files');
      group1.Filter = strSrcFilter;

      var strHeaderFilter = wizard.FindSymbol('HEADER_FILTER');
      var group2 = proj.Object.AddFilter('Header Files');
      group2.Filter = strHeaderFilter;
   }
   catch(e)
   {
      throw e;
   }
}

function AddConfig(proj, strProjectName)
{
   try
   {
      var config = proj.Object.Configurations('Debug');

      var nUnicodeSupport = wizard.FindSymbol('PB10UNICODE');
      if (nUnicodeSupport != 0)
      {
         config.CharacterSet = charSetUnicode;
      }
      else
      {
         config.CharacterSet = charSetMBCS;
      }
      config.ConfigurationType = typeDynamicLibrary;
      config.IntermediateDirectory = 'Debug';
      config.OutputDirectory = 'Debug';

      var CLTool = config.Tools('VCCLCompilerTool');
      CLTool.RuntimeLibrary = rtMultiThreadedDebug;
      CLTool.UsePrecompiledHeader = pchGenerateAuto;
      var strDefines = GetPlatformDefine(config);
      strDefines += "_DEBUG";
      strDefines += ";_WINDOWS;_USRDLL;";
      if (nUnicodeSupport != 0)
      {
         strDefines += "UNICODE;";
      }      
      var strExports = wizard.FindSymbol("UPPER_CASE_PROJECT_NAME") + "_EXPORTS";
      strDefines += strExports;
      //if (nUnicodeSupport != 0)
      //{      
         CLTool.AdditionalIncludeDirectories = "\"C:\\sybase\\PowerBuilder 10.0\\SDK\\PBNI\\include\"";
     //}
      //else
      //{
        // CLTool.AdditionalIncludeDirectories = "\"C:\\sybase\\PowerBuilder 9.0\\SDK\\PBNI\\include\"";
      //}
      CLTool.PreprocessorDefinitions = strDefines;
      CLTool.DebugInformationFormat = debugEditAndContinue;
      CLTool.Optimization = optimizeDisabled;

      var LinkTool = config.Tools('VCLinkerTool');
      LinkTool.SubSystem = subSystemWindows;
      LinkTool.TargetMachine = machineX86;
      LinkTool.ProgramDatabaseFile = "$(OutDir)/" + strProjectName + ".pdb";
      LinkTool.GenerateDebugInformation = true;
      LinkTool.LinkIncremental = linkIncrementalYes;
      LinkTool.ImportLibrary = "$(OutDir)/" + strProjectName + ".lib";
      LinkTool.OutputFile = "$(OutDir)/" + strProjectName + ".pbx";
      /*if (nUnicodeSupport != 0)
      {      
         // PB10 no longer uses a static library
         // LinkTool.AdditionalLibraryDirectories = "\"C:\\sybase\\PowerBuilder 10.0\\SDK\\PBNI\\lib\"";
      }
      else
      {
         //LinkTool.AdditionalLibraryDirectories = "\"C:\\sybase\\PowerBuilder 9.0\\SDK\\PBNI\\lib\"";
         //LinkTool.AdditionalDependencies = "pbni.lib";
      }*/
      var oldNames = LinkTool.IgnoreDefaultLibraryNames;
      LinkTool.IgnoreDefaultLibraryNames = "libc.lib " + oldNames

      var DebugTool = config.DebugSettings
      //if (nUnicodeSupport != 0)
      //{      
         DebugTool.Command = "pb100.exe"
      //}
      //else
      //{
        // DebugTool.Command = "pb90.exe"
     // }
            
      config = proj.Object.Configurations('Release');
      if (nUnicodeSupport != 0)
      {
         config.CharacterSet = charSetUnicode;
      }
      else
      {
         config.CharacterSet = charSetMBCS;
      }
      config.ConfigurationType = typeDynamicLibrary;
      config.IntermediateDirectory = 'Release';
      config.OutputDirectory = 'Release';

      var CLTool = config.Tools('VCCLCompilerTool');
      CLTool.RuntimeLibrary = rtMultiThreaded;
      CLTool.UsePrecompiledHeader = pchGenerateAuto;
      var strDefines = GetPlatformDefine(config);
      strDefines += "NDEBUG";
      strDefines += ";_WINDOWS;_USRDLL;";
      if (nUnicodeSupport != 0)
      {
         strDefines += "UNICODE;";
      }      
      
      var strExports = wizard.FindSymbol("UPPER_CASE_PROJECT_NAME") + "_EXPORTS";
      strDefines += strExports;
      CLTool.PreprocessorDefinitions = strDefines;
      //if (nUnicodeSupport != 0)
      //{      
         CLTool.AdditionalIncludeDirectories = "\"C:\\sybase\\PowerBuilder 10.0\\SDK\\PBNI\\include\"";
      //}
      //else
      //{
         //CLTool.AdditionalIncludeDirectories = "\"C:\\sybase\\PowerBuilder 9.0\\SDK\\PBNI\\include\"";
      //}
      CLTool.DebugInformationFormat = debugEnabled;
      CLTool.InlineFunctionExpansion = expandOnlyInline;

      var LinkTool = config.Tools('VCLinkerTool');
      LinkTool.SubSystem = subSystemWindows;
      LinkTool.TargetMachine = machineX86;
      LinkTool.GenerateDebugInformation = true;
      LinkTool.LinkIncremental = linkIncrementalNo;
      LinkTool.ImportLibrary = "$(OutDir)/" + strProjectName + ".lib";
      LinkTool.OutputFile = "$(OutDir)/" + strProjectName + ".pbx";
      if (nUnicodeSupport != 0)
      {      
         // PB10 no longer uses a static library
         // LinkTool.AdditionalLibraryDirectories = "\"C:\\sybase\\PowerBuilder 10.0\\SDK\\PBNI\\lib\"";
      }
      else
      {
         //LinkTool.AdditionalLibraryDirectories = "\"C:\\sybase\\PowerBuilder 9.0\\SDK\\PBNI\\lib\"";
         //LinkTool.AdditionalDependencies = "pbni.lib";
      }
      var oldNames = LinkTool.IgnoreDefaultLibraryNames;
      LinkTool.IgnoreDefaultLibraryNames = "libc.lib " + oldNames
   }
   catch(e)
   {
      throw e;
   }
}

function PchSettings(proj)
{
   // TODO: specify pch settings
}

function DelFile(fso, strWizTempFile)
{
   try
   {
      if (fso.FileExists(strWizTempFile))
      {
         var tmpFile = fso.GetFile(strWizTempFile);
         tmpFile.Delete();
      }
   }
   catch(e)
   {
      throw e;
   }
}

function CreateCustomInfFile()
{
   try
   {
      var fso, TemplatesFolder, TemplateFiles, strTemplate;
      fso = new ActiveXObject('Scripting.FileSystemObject');

      var TemporaryFolder = 2;
      var tfolder = fso.GetSpecialFolder(TemporaryFolder);
      var strTempFolder = tfolder.Drive + '\\' + tfolder.Name;

      var strWizTempFile = strTempFolder + "\\" + fso.GetTempName();

      var strTemplatePath = wizard.FindSymbol('TEMPLATES_PATH');
      var strInfFile = strTemplatePath + '\\Templates.inf';
      wizard.RenderTemplate(strInfFile, strWizTempFile);

      var WizTempFile = fso.GetFile(strWizTempFile);
      return WizTempFile;
   }
   catch(e)
   {
      throw e;
   }
}

function GetTargetName(strName, strProjectName, strClassName)
{
   try
   {
      // TODO: set the name of the rendered file based on the template filename
      var strTarget = strName;

      if (strName == 'readme.txt')
         strTarget = 'ReadMe.txt';

      if (strName == 'main.cpp')
         strTarget = strProjectName + '.cpp';

      if (strName == 'main.h')
         strTarget = strProjectName + '.h';

      if (strName == 'classname.h')
         strTarget = strClassName + '.h';

      if (strName == 'classname.cpp')
         strTarget = strClassName + '.cpp';

      if (strName == 'i_classname.h')
         strTarget = 'I' + strClassName + '.h';

      return strTarget; 
   }
   catch(e)
   {
      throw e;
   }
}

function AddFilesToCustomProj(proj, strProjectName, strProjectPath, InfFile)
{
   try
   {
      var projItems = proj.ProjectItems
      var strTemplatePath = wizard.FindSymbol("TEMPLATES_PATH");
      var strClassName = wizard.FindSymbol("SAFE_CLASS_NAME");
      var strTpl = '';
      var strName = '';

      var strTextStream = InfFile.OpenAsTextStream(1, -2);
      while (!strTextStream.AtEndOfStream)
      {
         strTpl = strTextStream.ReadLine();
         if (strTpl != '')
         {
            strName = strTpl;
            var strTarget = GetTargetName(strName, strProjectName, strClassName);
            var strTemplate = strTemplatePath + '\\' + strTpl;
            var strFile = strProjectPath + '\\' + strTarget;

            var bCopyOnly = false;  //"true" will only copy the file from strTemplate to strTarget without rendering/adding to the project
            var strExt = strName.substr(strName.lastIndexOf("."));
            if (strExt==".bmp" ||
                strExt==".ico" ||
                strExt==".gif" ||
                strExt==".rtf" ||
                strExt==".css" /*||
                strName == "i_classname.h"*/)
               bCopyOnly = true;
            wizard.RenderTemplate(strTemplate, strFile, bCopyOnly);
            proj.Object.AddFile(strFile);
         }
      }
      strTextStream.Close();
   }
   catch(e)
   {
      throw e;
   }
}
