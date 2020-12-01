/******************************************************************************************
The zip and unzip macro in windows environment. 
Created by: Shenzhen YAO
Date: August 24, 2014. 
Version: 1.2.0.
LICENSE: MIT standard.
Github repository: https://github.com/ShenzhenYAO/sasmacros.
******************************************************************************************/

/*Instructions and examples*/
%macro example;
	%let targetPath =C:\Users\Z70\Desktop\Canada DPD\database;
	%let srcPath = C:\Users\Z70\Desktop\Canada DPD\database;

	/*zip a SAS data set from a specific path to the target path*/
	%vbsZip(
		tgtDir=&targetPath, 
		zipFile=dpd.zip, 
		srcFile=dpd.sas7bdat,
		srcDir=&srcPath /*from a source path that is NOT the working directory*/
	);

	/*unzip a SAS data set to the work directory */
	%vbsUnzip(
		zipFile=%str(&srcPath.\dpd.zip), /*with full path, zip name and zip extension name*/
		srcFile=%str(dpd.sas7bdat) /*with the extension name .sas7bdat*/
	);
	proc print data=work.dpd (obs=10);run;

	/*zip a SAS data set from the SAS work directory to the target path*/
	%vbsZip(
		tgtDir=&targetPath, 
		zipFile=dpd1.zip, 
		srcFile=dpd.sas7bdat,
		srcDir= /*if srcDir is not specified, the macro looks for the data set from the working directory*/
	);
%mend example;


/****************************************************************************************/
/*
vbsZip a single file into a new zip (the existing zip file will be overwritten)
By default, the source directory (srcDir) = the SAS work directory
example: 
%vbsZip(tgtDir=c:\Users\z70\Desktop, zipFile=new.zip, 
srcFile=0914.wma, srcDir=c:\Users\z70\Desktop);
*/
%macro vbsZip(tgtDir=, zipFile=, srcFile=, srcDir=);
	%local tgtDir zipFile srcFile srcDir overWrite tPath tZip sPath sFile;
	%if %length(&zipFile) = 0 or %length(&srcFile)=0 %then %goto exit;
	%if %length(&srcFile) ne 0 and %index(&srcFile, .) =0 %then %let srcFile=&srcFile..sas7bdat;
	%put >>>  &srcFile;

	/*by default, set the target folder to the physical location of the SAS temp 'work' folder*/
	%if %length(&srcDir) = 0 %then %do;
		%let srcDir=%sysfunc(pathname(work));
		%let srcDir=%substr(&srcDir, 1, %length(&srcDir));
	%end;

	%if %length(&tgtDir) =0 %then %do;
		%let tgtDir=&srcDir;
	%end;

	%put >>>  so far so good;
	/*remove the tailing backslash \ from the tgtDir and srcDir*/
	%if %substr(&srcDir, %length(&srcDir), 1)=\ %then %do;
		%let srcDir=%substr(&srcDir, 1, %eval(%length(&srcDir)-1));
	%end;
	%if %substr(&tgtDir, %length(&tgtDir),1)=\ %then %do;
		%let tgtDir=%substr(&tgtDir, 1, %eval(%length(&tgtDir)-1));
	%end;

	/*if the zipfile does not have an extension name, add .zip*/
	%if %eval(%sysfunc(lowcase(%substr(&zipFile, %length(&zipFile)-3, 4
																)
												)
									) 
					ne .zip	)=1 %then %do;
		%let zipFile=&zipFile..zip;
	%end;

	%put >>>  &srcDir;
	%put >>>  &tgtDir;
	%put >>>  &srcfile;
	%put >>>  &zipFile;

	/*delete the existing tmp vbs file*/
	%if %sysfunc(fileExist(&srcDir.\vbsZiptmp.vbs)) %then %do;
		x "del ""&srcDir.\vbsZiptmp.vbs";
	%end;
	/*assign filename vbspt to the temporary vbs file*/
	filename vbspt "&srcDir.\vbsZiptmp.vbs";

	/*%bquote works for unmated quotes, also, it resolves variables*/
	%let tPath=%bquote('TargetPath="&tgtDir"');
	%let tZip=%bquote('TargetZip="&zipFile"');
	%let sPath=%bquote('SourcePath="&srcDir"');
	%let sFile=%bquote('SourceFile="&srcFile"');
	
	/*write lines into the tmp vbs file*/
	data _null_;
		file vbspt;
		put 'Set FSO=createObject("Scripting.FileSystemObject")';
		put 'Set objShell=CreateObject("Shell.Application")';
		/*the bquote contents must be unquoted*/
		put %unquote(&tPath);
		put %unquote(&tZip);
		put %unquote(&sPath);
		put %unquote(&sFile);
		put 'theTarget=TargetPath & "\" & TargetZip';
		put ' Set file = fso.CreateTextFile(theTarget, True)';
		put ' file.write("PK" & chr(5) & chr(6) & string(18, chr(0)))';
		put ' file.close';
		put 'set objTarget=objShell.Namespace(theTarget)';
		put 'set objSource=objShell.Namespace(SourcePath)';
		put 'objTarget.CopyHere objSource.Items.Item(SourceFile)';
		put 'do until objTarget.items.count=1';
		put ' wscript.sleep 1000';
		put 'loop';
		put 'Set FSO=Nothing';
		put 'Set objShell=Nothing';
		put 'Set objSource=Nothing';
		put 'Set objTarget=Nothing'; 
	run;

	/*use x commands to run vbs script*/
	x "%str(cscript.exe %"&srcDir.\vbsZiptmp.vbs %")";

	/*delete the existing tmp vbs file*/
	%if %sysfunc(fileExist(&srcDir.\vbsZiptmp.vbs)) %then %do;
		x "del ""&srcDir.\vbsZiptmp.vbs";
	%end;
	filename vbspt clear;
%exit:
%mend vbsZip;
/****************************************************************************************/

/*
This macro is better then 62 'unzip', as it does not require to clumsily open excel files. 
Example:

%vbsUnzip(
	tgtDir=S:\inbox,
	zipFile=%str(S:\Master Datasets\zips\md.zip),
	srcFile=%str(msbdoc.sas7bdat)
);
*/

%macro vbsUnzip(tgtDir=,zipFile=,srcFile=, overWrite=);
	%local tgtDir zipFile srcFile objT objS overWrite ;

	/*by default, set the target folder to the physical location of the SAS temp 'work' folder*/
	%if %length(&tgtDir )=0 %then %do;
		%let tgtDir=%sysfunc(pathname(work));
		%let tgtDir=%substr(&tgtDir, 1, %length(&tgtDir));
	%end;

	/*delete the existing tmp vbs file*/
	%if %sysfunc(fileExist(&tgtDir.\vbsUnziptmp.vbs)) %then %do;
		x "del ""&tgtDir.\vbsUnziptmp.vbs"" ";
	%end;

	/*assign filename vbspt to the temporary vbs file*/
	filename vbspt "&tgtDir.\vbsUnziptmp.vbs";

	/*create a string to set objTarget terms*/
	%let objT = %str(%'set objTarget=objShell.Namespace%(%")&tgtDir%str(%"%)%');
	%let objT = %substr(&objT, 1, %length(&objT));
	%put >>> objT=&objT;

	/*create a string to set objSource terms*/
	%let objS = %str(%'set objSource = objShell.Namespace%(%")&zipFile%str(%"%).items.item%(%")&srcFile%str(%"%)%');
	%let objS = %substr(&objS, 1, %length(&objS));
	%put >>> objS=&objS;

	%if %length(&overWrite)=0 %then %let overWrite =1;

	%put >>> overWrite=&overWrite;

	/*write lines into the tmp vbs file*/
	data _null_;
		file vbspt;
		put 'Set FSO=createObject("Scripting.FileSystemObject")';
		put 'Set objShell = CreateObject("Shell.Application")';
		put &objT;
		put &objS;
		%if &overWrite=1 %then %do;
			put 'objTarget.CopyHere objSource, 16'; /*256 = ask before overwriting, 16=always overwrite*/
		%end;
		%else %do;
			put 'objTarget.CopyHere objSource, 256'; /*256 = ask before overwriting, 16=always overwrite*/
		%end;
		put 'Set FSO=Nothing';
		put 'Set objShell=Nothing';
		Put 'Set objSource=Nothing';
		Put 'Set objTarget=Nothing';
	run;

	/*use x command to run vbs script*/
	x "%str(cscript.exe %"&tgtDir.\vbsUnziptmp.vbs %")";

	/*delete the tmp vbs script file*/
	%if %sysfunc(fileExist(&tgtDir.\vbsUnziptmp.vbs)) %then %do;
		x "del ""&tgtDir.\vbsUnziptmp.vbs"" ";
	%end;

%mend vbsUnzip;
/****************************************************************************************/
