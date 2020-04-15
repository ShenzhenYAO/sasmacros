/******************************************************************************************
The zip and unzip macro in unix/linux environment. 
Created by: Shenzhen YAO
Date: August 24, 2014. 
Version: 1.2.0.
LICENSE: MIT standard.
Github repository: https://github.com/ShenzhenYAO/sasmacros.
******************************************************************************************/


/****************************************************************************************/
/*to zip files in Unix (the remote server)
!!! IN UNIX, EVERYTHING IS CASE SENSATIVE
%UnixZip(
	sPath=, 
	source=, 
	data=PHRS_INFO_OccuER, 
	tPath=&remoteTargetPath, 
	target=PHRS_INFO_OccuER.zip, 
	debug=);

*/
%macro UnixZip(sPath=, source=, data=, tPath=, target=, debug=);
	%local sPath tPath data source target debug;

	/*If the source file/data or the target file is not specified, quit */
	%if (%length(&source)=0 and %length(&data) =0) or (%length(&target)=0) %then %do;
		%put >>>  The name of the source/data, and the target file must be specified; 
		%goto exit;
	%end;

	/*by default, use the work directory as the source path*/
	%if %length(&sPath)=0 %then %do;
		%put >>>  using path of the work directory as the default source Path;
		%let sPath=%sysfunc(pathname(work));
	%end;
	%else %do;
		/*if there are quotes around the string, strip them*/
		%let sPath=%sysfunc(dequote(&sPath));
	%end;

	/*If there is no slash at the end of sPath, add it
			note that the char '/' must be masked by %bquote()
			Otherwise SAS will treat it as an operend and report error*/
	%if %bquote(%substr(&sPath, %length(&sPath), 1)) ne %bquote(/) %then %do;
		%put >>>  add a slash to the end of the source path;
		%let sPath=&sPath./;
	%end; 

	/*by default, use the work directory as the target path*/
	%if %length(&tPath)=0 %then %do;
		%put >>>  using path of the work directory as the default target Path;
		%let tPath=%sysfunc(pathname(work));
	%end;
	%else %do;
		/*if there are quotes around the string, strip them*/
		%let tPath=%sysfunc(dequote(&tPath));
	%end;

	%if %bquote(%substr(&tPath, %length(&tPath), 1)) ne %bquote(/) %then %do;
		%put >>>  add a slash to the end of the target path;
		%let tPath=&tPath./;
	%end; 

	/*user can enter source file in var &data, or in var &source
	by default, use the value specified in var &data*/
	%if %length(&data) ne 0 %then %do;
		%let source=&data;
	%end;

	/*by default, add the source file extension name as sas7bdat*/
	%if %index(&source, .)=0 %then %do;
		%let source=&source..sas7bdat;
		%put >>>  source file default extension name (.sas7bdat) added;
	%end;

	/*add the extension name .zip to target file*/
	%if %index(&target, .zip)=0 %then %do;
		%let target=&target..zip;
		%put >>>  target file default extension name (.zip) added;
	%end;

	%let source=%sysfunc(lowcase(&source));
	%let target=%sysfunc(lowcase(&target));

	%put >>>  sourcePath= &sPath;
	%put >>>  source file = &source;
	%put >>>  targetPath= &tPath;
	%put >>>  target file =&target;

	/*No matter the user add/not add quotes, the dequote funciton ensures
		that in the command line the quotes are stripped*/
	/*In the following commands, the option -X is to specify NOT to save extra file attributes
		(note, it is in uppercase -X, not in lower case -x. The latter is for excluding files when zipping)
		
		If a file is larger than 2 or 4 GB, in the current linux system which uses zip 3.0, the file will be automatically zipped
	as 64 bit format. By default (if not have -X), extra information about the 64 bit format will be saved. 
		However, the extra information cannot be read correctly by Windows system. The Windows, although can handle large 
	files, fails to recognize the correct size of the original files in 64bit zips by linux. As the result, the Windows often asks for 
	PB (1PB= 1 million GB) size of space! This is because Windows has a bug that confuses 32-bit header and 64-bit headers of 
	the files. The bug has been there since XP. It witnessed the downhill-going of Microsoft after Steve Ballmer took over,
	as well as the futher sliding down under the leadership of Satya Nadella. 

		As Windows is not accountable for solving such problems, the alternative is to use the option -X (i.e., NOT to save 
	the extra information of the 64 bit format). Without the extra information, it is way easier for Windows to identify 
	the correct file size! 

		ref. 
		https://ardamis.com/2011/07/28/native-windows-compressed-folders-utility-5-99-pb/
		https://www.linux.org/docs/man1/zip.html*/
	data _null_;
		command2 =trim(
		"zip -j -X " ||trim(symget('tPath')) || trim(symget('target')) || " " || trim(symget('sPath')) || trim(symget('source'))
			);
		command = 'zip -j -X /SAS_data/dw0/HQC_share/Shenzhen/Projects/EC_NeuroAdh/Source/pt_newdrugs.zip /SAS_data/dw0/HQC_share/Shenzhen/Projects/EC_NeuroAdh/Source/pt_newdrugs.sas7bdat';
		call system(command2);
	run;
%exit:
%mend UnixZip;
/****************************************************************************************/
/*the following macro helps exract files from a zip file on SASApp (unix)
!!! IN UNIX, EVERYTHING IS CASE SENSATIVE
by default, the sPath/tPath is the location of the work directory
by default, the zip file's extention is zip
by default, the zipped extention name is sas7bdat
the sPath/tPath can be specified with or without quate symbols
if the zipped file is not specified, all files in the zip will be extracted, and the extracted files 
	cannot be renamed as a bat
if the zipped file is not specified, the extracted file, even is a single file, 
	cannot be renamed (even if the target file is specified using a different name)


example:
%UnixUnZip(
	sPath=/SAS_data/dw0/HQC_share/Shenzhen/Projects/EC_NeuroAdh/Source/,
	zip=try,
	zipped=try,
	tPath=/SAS_data/dw0/HQC_share/Shenzhen/Projects/EC_NeuroAdh/Source,
	target=try1
);
*/
%macro UnixUnZip(sPath=, zip=, zipped=, tPath=, target=, debug=);
	%local sPath zip zipped tPath target debug;

	/*If the zip name is not specified, quit */
	%if %length(&zip)=0 %then %do;
		%put >>>  The name of the zip must be specified; 
		%goto exit;
	%end;

	/*add the extension name .zip to the zip*/
	%if %index(&zip, .zip)=0 %then %do;
		%let zip=&zip..zip;
		%put >>>  the zip file default extension name (.zip) added;
	%end;

	/*by default, use the work directory as the source path*/
	%if %length(&sPath)=0 %then %do;
		%put >>>  using path of the work directory as the default source Path;
		%let sPath=%sysfunc(pathname(work));
	%end;
	%else %do;
		/*if there are quotes around the string, strip them*/
		%let sPath=%sysfunc(dequote(&sPath));
	%end;

	/*If there is no slash at the end of sPath, add it
			note that the char '/' must be masked by %bquote()
			Otherwise SAS will treat it as an operend and report error*/
	%if %bquote(%substr(&sPath, %length(&sPath), 1)) ne %bquote(/) %then %do;
		%put >>>  add a slash to the end of the source path;
		%let sPath=&sPath./;
	%end; 

	/*by default, use the work directory as the target path*/
	%if %length(&tPath)=0 %then %do;
		%put >>>  using path of the work directory as the default target Path;
		%let tPath=%sysfunc(pathname(work));
	%end;
	%else %do;
		/*if there are quotes around the string, strip them*/
		%let tPath=%sysfunc(dequote(&tPath));
	%end;

	%if %bquote(%substr(&tPath, %length(&tPath), 1)) ne %bquote(/) %then %do;
		%put >>>  add a slash to the end of the target path;
		%let tPath=&tPath./;
	%end; 

	/*by default, add the zipped file extension name as sas7bdat*/
	%if %index(&zipped, .)=0 and %length(&zipped) ne 0 %then %do;
		%let zipped=&zipped..sas7bdat;
		%put >>> the zipped file default extension name (.sas7bdat) added;
	%end;

	/*if the target file is not specified*/
	%if %length(&target) =0 %then %do;
		/*1. if the zipped file is specified, let target=zipped*/
		%if %length(&zipped) ne 0 %then %do;
			%put >>>  be default, set target = the zipped file, &zipped;
			%let target=&zipped;
		%end;
		%else %do;
			/*2. if the zipped file is not specified, there are more than 1 file 
				to be extracted. There is no need to specify target file names*/
		%end;
	%end;

	/*if the zipped file is not specified (implying to extract all files from the zip),
			let target = nothing, so as to be aligned with the setting in the zipped*/
	%if %length(&zipped)=0 %then %do;
		%let target=;
	%end; 

	/*by default, add the target file extension name as sas7bdat*/
	%if %index(&target, .)=0 and %length(&target) ne 0 %then %do;
		%let target=&target..sas7bdat;
		%put >>> the target file default extension name (.sas7bdat) added;
	%end;
	
	%if %length(&zipped) ne 0 %then %do;
		%let zipped=%sysfunc(lowcase(&zipped));
	%end;
	%if %length(&target) ne 0 %then %do;
		%let target=%sysfunc(lowcase(&target));
	%end;

	%put >>>  sourcePath= &sPath;
	%put >>>  the zip file = &zip;
	%put >>>  the zipped file = &zipped;
	%put >>>  targetPath= &tPath;
	%put >>>  target file =&target;

	/*No matter the user add/not add quotes, the dequote funciton ensures
		that in the command line the quotes are stripped
		wrap the folder name and the file name with apostrophes ''
		if the names contain '-' or '()', the apostrophes ensure that the unix command read the folder names
			correctly and won't treat '-' as sort of options
	*/
	data UnixUnZip_tmp1;
		step1 =	"unzip '" ||trim(symget('sPath')) || trim(symget('zip')) || "' " 
    	|| trim(symget('zipped')) || " -d '" ||trim(symget('tPath')) || "'";

		step1a = 'unzip /SAS_data/dw0/HQC_share/Shenzhen/Projects/EC_NeuroAdh/Source/try.zip try.sas7bdat -d /SAS_data/dw0/HQC_share/Shenzhen/Projects/EC_NeuroAdh/Source/';
		step2a = 'mv /SAS_data/dw0/HQC_share/Shenzhen/Projects/EC_NeuroAdh/Source/try.sas7bdat /SAS_data/dw0/HQC_share/Shenzhen/Projects/EC_NeuroAdh/Source/try1.sas7bdat';
		call system(step1);
		/*Only excute the step2 (rename) if the zipped and the target are different*/
		%if &zipped ne &target and %length(&zipped) ne 0 %then %do;
			step2="mv '" ||trim(symget('tPath')) || trim(symget('zipped')) || "' '"
			||trim(symget('tPath')) ||trim(symget('target')) ||"'";
			call system(step2);
		%end;
	run;
%exit:

	%if &debug ne 1 %then %do;
		proc datasets nolist;
			delete UnixUnZip_tmp:;
		run;
	%end;
%mend UnixUnZip;