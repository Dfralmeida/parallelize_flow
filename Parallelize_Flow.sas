/******************************************************************************
Parallelize_Flow.sas
Author:		David Kratz
Email:		David.Kratz@d-wise.com
Date:		3/2/2017

Description:
This SAS Script, when run on a Windows operating system services generated .vbs
flow script will back it up, and create a new flow script, as well as a subscript
for each job invoked in the flow.  When run, this new script will run the flow's jobs
in parallel, where allowed by the dependencies defined for the jobs.

Instructions: 
1. Modify the line:
targetScript =  "C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/test3/test3.vbs";
below to contain the path
to a Windows operating system services flow .  Note that the slashes used in this path
are the reverse of Windows standard path names.
2. Submit the code.
3. Run the newly generated .vbs script to run your flow in parallel.
******************************************************************************/


proc groovy;
	submit; 

		/*
		targetScript holds the location of the target flow script.  To use the script, replace the quoted path with
		the location of the script you'd like to convert.
		Note the slashes are reversed in this path.  This removes the need to escape the Windows path.
		*/
		targetScript = "C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/test3/test3.vbs";

		/* The following have global scope, and can be accessed by any of this script's functions. */
		/* [] in this context are arraylists. */
		scriptLines = [];
		headerLines = [];
		loopBeginLines = [];
		jobHeaders = [];
		jobBeginnings = [];
		jobOldInvocations = [];
		jobNewInvocations = [];
		jobEndings = [];
		jobNames = [];
		jobStatusFileNames = [];
		monitorLines = [];
		footerLines = [];
		loopEndFound = false ;
		scriptHasDependencies = false;
		invalidScriptReasons = "";

		/* Create a file object for the target script*/
		originalScript = new File(targetScript);

		/*
		Read the text of the target script into a single string, split it by the line seperator characters (\r, \n)
		and convert the returned array into an array list.
		*/
		scriptLines = originalScript.text.split("\\r?\\n").toList();

		/* Perform simple validation on the contents of the script. */
		boolean validScript = validateScript();

		/* If the script has been deemed valid, perform the following: */
		if(validScript){
		
			/* Extract the directory of the targetScript from its path. */
			targetScriptDir = targetScript.substring(0,targetScript.lastIndexOf('/') + 1);

			/* Parse the text of the target script, and separate it into its constituent pieces. */
			breakScriptIntoChunks();
			
			/* Modify the chunks of the target script as needed to support parallel runs. */
			modifyChunksForParallelProcessing();
			
			/* Output a back up copy of the original script, then output the  modified script and the new subscripts. */
			outputParallelScriptFiles();

			/* Output a notification to the SAS log that script modification has completed. */
			println "Modification of your script has completed successfully.";
			println "A backup of your script (.bak) has been placed in the same folder as your original script";
		
		}
		/* If the script has been deemed invalid, perform the following: */
		else{
		
			/* Output a notification to the SAS log that script modification has failed. */
			println "Modification of your script has failed.";
			println invalidScriptReasons;
		}

		/* Perform simple validation on the script and return true if the script is deemed valid, false otherwise. */
		boolean validateScript(){

			boolean validScript = false;

			/* Perform a check to see if the script has already been modified for parallel processing. */
			boolean scriptModified = checkScriptForModification();

			/* Perform a check to see if the script has any dependencies. */
			scriptHasDependencies = checkScriptForDependencies();

			/* Count how many job invocations the script contains. */
			int numberOfJobsInvoked = countScriptJobs();

			/* If the necessary conditions have been met, deem the script valid. */
			if (scriptModified == false && (numberOfJobsInvoked > 1)){
				validScript = true;
			}

			/* If the script has already been modified, report the issue. */
			if (scriptModified){
				invalidScriptReasons = invalidScriptReasons + "The script has already been modified.  ";
			}

			/* If the script contains only one job, report the issue. */
			if (numberOfJobsInvoked <= 1){
				invalidScriptReasons = invalidScriptReasons + "The script contains at most one job, and does not need to modified."
			}

			return validScript;
		}

		/* Return an integer value containing the number of jobs invoked by the script flow. */
		int countScriptJobs(){

			int numberOfJobsFound = 0;
			
			/* Count the number of times the script contains the comment "Begin Job Event" as this happens for every job invoked. */
			for (int i = 0; i < scriptLines.size; i ++){

				if (scriptLines[i].trim().startsWith("' *** Begin Job Event ***")){
					numberOfJobsFound = numberOfJobsFound + 1;
				}
			}

			return numberOfJobsFound;
		}

		/* Return a boolean value of true if the script contains a string that indicates it has been modified, false otherwise. */
		boolean checkScriptForModification(){

			boolean modificationFound = false;

			/*
			Scan the text of the script for a string that is added to every parallelized flow.  If it is present, the script is
			assumed to have been modified.
			*/
			for (int i = 0; i < scriptLines.size; i ++){

				if (scriptLines[i].trim().startsWith("' This script has been modified to support running jobs in parallel.")){
					modificationFound = true;
				}
			}

			return modificationFound;
		}

		/* Return a boolean value of true if the script contains a string that indicates it has no dependencies, false otherwise. */
		boolean checkScriptForDependencies(){

			boolean dependencyFound = true;

			/*
			Scan the text of the script for a string that is added to every flow without dependencies.  
			If it is present, the script is assumed to have no dependencies.
			*/
			for (int i = 0; i < scriptLines.size; i ++){

				if (scriptLines[i].trim().startsWith("' *** No Dependencies ***")){
					dependencyFound = false;
				}
			}

			return dependencyFound;
		}
		
		/* Break the script into its component pieces for processing. */
		void breakScriptIntoChunks(){
			getHeader();
			
			/* 
			In the event that the script has no dependencies, it will have no internal job invocation loop.
			So it is necessary to inject one for the parallelization to work.
			*/
			if(scriptHasDependencies){
				getLoopBeginning();
			}
			else {
				createLoopBeginning();
			}
			getJobParts();
			getFooter();
		}

		/* Populate headerLines with every line up to and including the line "statusFile.WriteLine("Flow STARTING...")" */
		void getHeader(){
			headerLines = consumeLinesUntilText('statusFile.WriteLine("Flow STARTING...")', true, false );
		}

		/* Populate loopBeginLines with every line up to the first instance of "If Not" */
		void getLoopBeginning(){
			loopBeginLines = consumeLinesUntilText("If Not", false, false );
		}

		/* Populate loopBeginLines with a constructed loop beginning. */
		void createLoopBeginning(){
			
			loopBeginLines.add("dLoop = True");
			loopBeginLines.add("");
			loopBeginLines.add("Do While dLoop = True");
			loopBeginLines.add("");
			loopBeginLines.add("  dLoop = False");
			loopBeginLines.add("");

			consumeLinesUntilText("' *** No Dependencies ***", false, false );
		}

		/* Break apart the job invocations into their component parts */
		void getJobParts(){

			loopEndFound = false;
			
			/* Continue consuming text from the target script until the end of the job invocation loop is encountered. */
			while (! loopEndFound){
				getJobBeginning();
				getJobInvocation();
				getJobEnding();
			}
		
		}
		
		/* Consume text up to and including the line "' *** Begin Job Event ***" */
		void getJobBeginning(){
			jobBeginnings.add(consumeLinesUntilText("' *** Begin Job Event ***", true, false ));
		}
		
		/* Consume text up to and including the line "' *** End Job Event ***" */
		void getJobInvocation(){
			jobOldInvocations.add(consumeLinesUntilText("' *** End Job Event ***", false, false ));
		}
		
		/* Consume text until the beginning of the next job. Stop if the end of the loop is encountered.*/
		void getJobEnding(){
		
			/*Consume text lines until the next occurence of 'If Not' which is the start of the next job invocation */
			if(scriptHasDependencies){
				jobEndings.add(consumeLinesUntilText("If Not", false, true ));
			}
			/* 
			In the event that the script has no dependencies, there will be no dependency checks, so consume
			until the line "' *** Begin Job Event ***" is encountered.
			*/
			else{
				jobEndings.add(consumeLinesUntilText("' *** Begin Job Event ***", false, true));
			}
		}

		/* Consume the remaining lines of the script. */
		void getFooter(){
			
			/* 
			In the event that the script has no dependencies, it will lack a internal loop.
			Thus a loop ending must be created, as a loop beginning was created previously.
			*/
			if(!scriptHasDependencies){
				footerLines.add("Loop");
				footerLines.add("");
			}
			
			footerLines.addAll(scriptLines[0 .. scriptLines.size - 1].toList());
		}
		
		/*
		Consume lines of text from the scriptLines array until a line starts with target text.
		Depending on the passed in value of inclusive, either stop before that line, or consume it as well.
		Depending on the passed in value of stopAtLoopEnd, certain values will stop consumption,
		as they signify the end of the current loop.
		Return the consumed lines as an arrayList.
		*/
		ArrayList consumeLinesUntilText(String targetText, boolean inclusive, boolean stopAtLoopEnd){
		
			def i = 0;
			def targetTextFound = false;
			
			def returnLines;
			
			while (i < scriptLines.size && ! targetTextFound) {	
				if (scriptLines[i].trim().startsWith(targetText)){
					targetTextFound = true;
				}
				else {
					if(stopAtLoopEnd){
						if(scriptLines[i].trim() == "Loop" || scriptLines[i].trim() == "' Update date and time variables"){
							targetTextFound = true;
							loopEndFound = true;
						}
						else{
							i++;
						}
					}
					else{
						i++;	
					}	
				}
			}
			if(inclusive){
				returnLines = scriptLines[0..i].toList();
				scriptLines.removeRange(0, i + 1);
			}
			else{
				returnLines = scriptLines[0.. i - 1].toList();
				scriptLines.removeRange(0, i);
			}
			
			return returnLines;
		
		}
		
		/* Modify the chunks of the script to support Parallel Processing */
		void modifyChunksForParallelProcessing(){

			/* Pull the name of the jobs from the job invocations. */
			deriveJobNamesFromInvocations();
			
			/* Create headers specific to the subscripts. */
			createJobSpecificHeaders();
			
			/* Add the logic to the main script header which deletes the signal files if they are already present. */
			modifyJobHeader();
			
			/* 
			Add necessary additional variables to the dependency logic and deal with special case of the Job Starts Event
			*/
			modifyJobBeginnings();
			
			/* 
			Modify the old invocations to include the signal file creation.
			Create new invocations which call the subscript.
			*/
			modifyjobInvocations();

			/* 
			If the target script had no dependencies, close the if block of the new job invocation.
			The original invocation will not have one.
			*/
			if (!scriptHasDependencies){
				modifyjobEndings();
			}
			
			/* Create the logic which will be injected into the main script to detect the presence of signal files. */
			createMonitorLines();
		
		}
		
		/* Populate jobNames with the name pulled from the invocations */
		void deriveJobNamesFromInvocations(){

			for (invocation in jobOldInvocations){

				def statusLine = 0;

				/* Locate a line of the invocation in which the job name is present. */
				for (int i = 0; i < invocation.size(); i++)
				{
					if(invocation[i].trim().endsWith("= errorLevel")){
						statusLine = i;
					}
				}

				/* Parse out the job name from the discovered line. */
				def nameStart = invocation[statusLine].indexOf('status_') + "status_".length();
				
				def nameEnd = invocation[statusLine].indexOf(' =');
				
				jobNames.add(invocation[statusLine].substring(nameStart,nameEnd));
			}

		}
		
		/* 
		Populate job headers with a copy of the main script header for each job invoked. 
		Prune the copy as appropriate, and modify the referenced names to the header's respective job.
		*/
		void createJobSpecificHeaders(){
		
			/* Make a copy of the header for each job */
			for (i = 0; i < jobNames.size; i++){
			
				def tempHeader = [];
				
				for(line in headerLines){
					tempHeader.add(line);
				}
				
				jobHeaders.add(tempHeader);
			}
			
			/* Modify the job headers to write the log out to a file reflecting the name of the job. */
			for (i = 0; i < jobHeaders.size; i++){
				def jobHeaderEnd = 0;
				for (j = 0; j < jobHeaders[i].size; j++){
					if (jobHeaders[i][j].trim().startsWith("statusFilename")){
						jobHeaders[i][j] = jobHeaders[i][j].replace("timeStamp", '"' + jobNames[i] + '"');

						/*Store the jobStatusFileName declaration so that it can be used in the monitoring code.*/
						jobStatusFileNames.add(jobHeaders[i][j].replace("statusFilename", "jobStatusFilename"))
					}
					
					if (jobHeaders[i][j].trim().startsWith("' *** Start of flow ***")){
						jobHeaderEnd = j;
					}
				}
				
				/* Do not need several of the lines after the Start of the flow statement. */
				jobHeaders[i].removeRange(jobHeaderEnd, jobHeaders[i].size);
			}
			
		}
		
		/* Add code to the main script header which will delete signal files if they are present when the script is run. */
		void modifyJobHeader(){
		
			for (i = 0; i < jobNames.size; i ++){
			
				headerLines.add('If fileSys.FileExists("' + targetScriptDir + jobNames[i] + '_simple_status.log") Then');
				headerLines.add('');
				headerLines.add('fileSys.DeleteFile("' + targetScriptDir + jobNames[i] + '_simple_status.log")');
				headerLines.add('');
				headerLines.add("End If");
			}
		
		}
		
		/* 
		Modify the job beginnings to include a new variable Running_<job name>
		This injected logic keeps multiple copies of the job from being launched.
		A job invocation happens only if the job has not executed, and is not already running.
		*/
		void modifyJobBeginnings(){
		
			if(scriptHasDependencies){
				/* Inject the check for the running_<job name> variable */
				for (i = 0; i < jobBeginnings.size; i++){
					jobBeginnings[i][0] = jobBeginnings[i][0].replace("Then", "and Not running_" + jobNames[i] + " Then");

					dependencyCheckLine = -1;
					dependencyJobs = [];

					for ( j = 0; j < jobBeginnings[i].size; j++){
						/* Locate the portion of the old invocation which contains the dependency check. */
						if(jobBeginnings[i][j].trim().startsWith("If (")){
							dependencyCheckLine = j;
							break;
						}
					}

					/* 
					If we find a dependency check line, replace any instances of (exec_<jobname>) with
					(running_<jobname>).  To explain, when a job is dependent on another job starting, the logical
					expression follows this pattern.
					*/
					if( dependencyCheckLine != -1){
						
						for( j = 0; j < jobNames.size; j++){
							jobBeginnings[i][dependencyCheckLine] = jobBeginnings[i][dependencyCheckLine].replace("(exec_" + jobNames[j] + ")", "(running_" + jobNames[j] + ")");
						}
					}
				}
			}
			else{
					/* If the job had no dependencies, there is not a dependency check to inject into, so a new check must be created. */
					for (i = 0; i < jobBeginnings.size; i++){
						jobBeginnings[i][0] = "If Not exec_" + jobNames[i] + " and Not running_" + jobNames[i] + " Then";
					}
			}
		}
		
		/* 
		Populate newInvocations with subscript calls.
		Add the creation of signal files to the old invocations.
		*/
		void modifyjobInvocations(){
		
			for (i = 0; i < jobOldInvocations.size; i++){
			
				jobBeganLogEntry = [];
				jobBeganLogEntryLine = 0;

				for (j = 0; j < jobOldInvocations[i].size; j++){
					/* Locate the portion of the old invocation which writes that the job has started to the log. */
					if(jobOldInvocations[i][j].trim().startsWith("' Update date and time variables")){
						jobBeganLogEntryLine = j;
						break;
					}
				}
				/* 
				Extract the log writing portion of the code.  The use of + 7 in this case is a magic number
				but it was derived by observation of the files which have not deviated from this pattern.
				*/
				for (j = jobBeganLogEntryLine; j < jobBeganLogEntryLine + 7 && j < jobOldInvocations[i].size; j++)
				{
					jobBeganLogEntry.add(jobOldInvocations[i][j]);
				}

				/* Remove the log writing portion of the code from what will be written out to the subscript. */
				jobOldInvocations[i].removeRange(jobBeganLogEntryLine, jobBeganLogEntryLine + 7);

				def newInvocation = [];
				
				newInvocation.addAll(jobBeganLogEntry);
				newInvocation.add("");
				newInvocation.add("running_" + jobNames[i] + " = True");
				newInvocation.add("");
				newInvocation.add('shell.Run("' + targetScriptDir + jobNames[i] + '.vbs")');
			
				jobNewInvocations.add(newInvocation);
			
				jobOldInvocations[i].add('simpleStatusFilename = "' + targetScriptDir + jobNames[i] + '_simple_status.log"');
				jobOldInvocations[i].add("");
				jobOldInvocations[i].add("' Open status file");
				jobOldInvocations[i].add("Set simpleStatusFile = fileSys.OpenTextFile(simpleStatusFilename, ForWriting, True)");
				jobOldInvocations[i].add("' Log completion of job and exit code to status file");
				jobOldInvocations[i].add("simpleStatusFile.WriteLine(status_" + jobNames[i] + ")")
				jobOldInvocations[i].add("' Close status file");
				jobOldInvocations[i].add("simpleStatusFile.Close");
				
			}
		
		}

		/* This is only called when a flow has no dependencies.  Close the dependency if block created previously. */
		void modifyjobEndings(){
			
			for (ending in jobEndings){

				ending.add("End If");

			}

		}
		
		/* Create the logic which monitors for the signal files. */
		void createMonitorLines(){

			monitorLines.add("");
			
			/* 
			Add a chunk of code for each job which test for the presence of a signal file, and reads it in if it is present.
			Only check for jobs which haven't finished executing, and are currently running.
			If a file is found, pull the job's return code from it, and set status variables appropriately.
			*/
			for(i = 0; i < jobNames.size; i++){
			
				monitorLines.add("If ( ( Not exec_" + jobNames[i] + " ) and running_" + jobNames[i] + " ) Then");
				monitorLines.add("");
				monitorLines.add("  dLoop = True");
				monitorLines.add("");
				monitorLines.add('  If fileSys.FileExists("' + targetScriptDir + jobNames[i] + '_simple_status.log") Then');
				monitorLines.add("");
				monitorLines.add('    exec_' + jobNames[i] + " = True");
				monitorLines.add('    Set sf = fileSys.OpenTextFile("' + targetScriptDir + jobNames[i] + '_simple_status.log", ForReading, True)');
				monitorLines.add('    strStatus = sf.ReadLine');
				monitorLines.add("    status_" + jobNames[i] + " = CInt(strStatus)");
				monitorLines.add('    ' + jobStatusFileNames[i]);
				monitorLines.add('    Set sf = fileSys.OpenTextFile( jobStatusFilename , ForReading, True)');
				monitorLines.add('    strLog = sf.ReadLine');
				monitorLines.add('    statusFile.WriteLine(strLog)');
				monitorLines.add("    If flowStatus = 0 Then");
				monitorLines.add("      flowStatus = status_" + jobNames[i]);
				monitorLines.add("    End If");
				monitorLines.add("  End If");
				monitorLines.add("End If");	
				monitorLines.add("");
				
			}
			
			monitorLines.add("");
			monitorLines.add("Wscript.Sleep(15000)");
		
		}
		
		/* Create a back up of the original script, and output the new main script and subscripts. */
		void outputParallelScriptFiles()
		{
			backupOriginalScriptFile();
			outputModifiedMainScript();
			outputModifiedJobScripts();
		}
		
		/* Create a backup of the original script. */
		void backupOriginalScriptFile(){
		
			File scriptBackup = new File(targetScript + ".bak");
			
			if (! scriptBackup.exists()){
				originalScript.renameTo(scriptBackup)
			}
		
		}
		
		/* Output the new main script by constructing it from the modified pieces of the old script. */
		void outputModifiedMainScript(){
		
			/* 
			Add the modified flag string to the text of the script.  This allows us to avoid trying to parallelize
			a flow we've already parallelized.
			*/
			def newScriptText = "' This script has been modified to support running jobs in parallel." + System.getProperty("line.separator");
			
			/* Append the header. */
			for (line in headerLines){
				newScriptText = newScriptText + line + System.getProperty("line.separator");
			}
			
			/* Append the beginning of the loop. */
			for (line in loopBeginLines){
				newScriptText = newScriptText + line + System.getProperty("line.separator");
			}
			
			/* Append the job invocations. */
			for (i = 0; i < jobNames.size; i ++){
			
				for(line in jobBeginnings[i]){
					newScriptText = newScriptText + line + System.getProperty("line.separator");
				}
				for(line in jobNewInvocations[i]){
					newScriptText = newScriptText + line + System.getProperty("line.separator");
				}
				for(line in jobEndings[i]){
					newScriptText = newScriptText + line + System.getProperty("line.separator");
				}
			}
			
			/* Append the monitoring logic. */
			for (line in monitorLines){
				newScriptText = newScriptText + line + System.getProperty("line.separator");
			}
			/* Append the footer lines. */
			for (line in footerLines){
				newScriptText = newScriptText + line + System.getProperty("line.separator");
			}
			
			/* Output the constructed string as a new main script. */
			originalScript.write(newScriptText);
		
		}
		
		/* Output the job subscripts by constructing them from modified pieces of the original script. */
		void outputModifiedJobScripts(){
		
			for (i = 0; i < jobNames.size; i++)
			{
			
				def newJobScriptText = "";
				
				/* Append the job specific headers. */
				for (line in jobHeaders[i]){
					newJobScriptText = newJobScriptText + line + System.getProperty("line.separator");
				}
				
				/* Append the modified job invocations. */
				for (line in jobOldInvocations[i]){
					newJobScriptText = newJobScriptText + line + System.getProperty("line.separator");
				}
				
				File jobScript = new File(targetScriptDir + jobNames[i] + ".vbs");
				
				/* Output the constructed string as a subscript. */
				jobScript.write(newJobScriptText);
			
			}
		
		}

	endsubmit;
quit;
