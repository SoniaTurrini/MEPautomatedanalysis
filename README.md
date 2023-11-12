# MEPautomatedanalysis
These matlab scripts allow you to analyze EMG traces to extract the basic features of MEPs automatically (MEP size, AUC, latency + background EMG activity). 

One script allows you to run the analysis quickly on multiple traces (MEPanalysis_multiplefiles), another to do the analysis one file at the time (MEPanalysis_singlefile). The single file script has the option to be ran without a digital marker.
The output is exported as an excel file. 

MEPanalysis_singlefile
The first lines of code are the ones you need to set to your needs: select the EMG channel you are analysing for MEP size, the digital channel (if you have it), and specify the sampling rate of your recording.
Also choose an output name for the excel file the script saves at the end of the analysis!
If you DO NOT have a digital marker in your EMG trace and need to upload the MEP positions from a separate file (e.g., an excel file), uncomment the corresponding section and specify the name of the file you are extracting MEP positions from.

MEPanalysis_multiplefiles
this script will run on all files that are in the same directory as the script. I have not coded the option to do this if there are no digital marker on the EMG traces, but it is pretty simple to do if necessary.
Once again, the first lines of code are the only one you need to modify, to define the EMG channel to analyse, the digital channel and the sampling rate. 
No need to define the output file name, it is set automatically based on the name of each file.
Of note: all files should have the EMG trace on the same channel and the same sampling rate for this loop to work.

