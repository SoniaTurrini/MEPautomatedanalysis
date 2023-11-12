%% Sonia Turrini
%turrinisonia@gmail.com

% this script allows you to automatically process an EMG trace and extract
% the MEP size (peak to peak and area under the curve), the MEP onset
% latency using two different methods, and the EMG activity preceding the
% TMS pulse, which can be useful to exclude MEPs preceded by excessive
% muscular contraction.

clearvars
close all
clc
%% These are the things you'll need to do manually
% 0: Save your EMG as a .mat file
% 1: double click on the .mat EMG file to load it in your workspace

% 2: Set the following variables
muscle1 = 1; % # of the muscle channel
digital = 3; % # of your digital channel
srate = 20000; %your recording sampling rate in Hz
winbyms = srate/1000; %don't touch this!
MEPstart = 15; % MEP analysis starts 15ms after TMS pulse
MEPstop = 60; % MEP analysis ends 60ms after TMS pulse
EMGprestart = 5; %background EMG analysis ends 5 ms before the pulse
EMGprestop = 105; %background EMG analysis starts 105 ms before the pulse
output_filename = 'outputMAT.xls'; %change name of file to whatever you would like the output name to be

%% uncomment this section if you DO NOT have a digital marker on your trace but rather an excel file that informs you of MEP latencies (in ms)
%put the Excel file in the same folder as this script and change the next
%line to match the name.
%MEPpositions_filename = 'MEPpos.xlsx';  

% Skip Option 1 and move to %Option 2 below.

    %% Option 1: find marker based on a digital channel
    %comment this entire section if you do not have a digital channel
    levelstart = data(1,digital);
    seqmrk5 = [];
    seqmrk0 = [];
    idpnt = 1;
    k = 1;
    while idpnt <= length(data(:,digital)) %for every data point of the EMG trace
        if levelstart ~= 0 && idpnt == 1
            idpnt = find(data(:,digital) == 0,1,'first');
        end
       
        tmp = find(data(idpnt:end,digital) > 0,1,'first'); % finds the start of the next marker
        
        if isempty(tmp) % when the last marker has been found and there aren't any more, the while loop breaks
            break
        end
        
        seqmrk5(1,k) = idpnt + tmp - 1; % the start of this marker is noted in the matrix seqmrk5
        idpnt = idpnt + tmp - 1;
        
        %finds the end of this marker
        tmp = find(data(idpnt:end,digital) == 0,1,'first'); %finds the first datapoint after the start of the marker where the digital channel gets back to 0
        
        if isempty(tmp) % when the last marker has been found and there aren't any more, the while loop breaks
            break
        end
        
        seqmrk0(1,k) = idpnt + tmp - 1; % the end of this marker is noted in the matrix seqmrk0
        idpnt = idpnt + tmp - 1;
        k = k+1;
    end
    
  % clean up errors (e.g., digital channel was >0 at the end of
  % recording)
    if seqmrk5(1,end) > seqmrk0(1,end)
        mrkOK = seqmrk5(1,1:(end-1));
    else
        mrkOK = seqmrk5;
    end
   
    %% Option 2: the latency of each MEP is extracted from a separate excel file
    % in this instance, an excel file with the latency in ms of each MEP
    %from the start of the recording is necessary
    winbyms = srate/1000; % number of recorded data points per ms

    mrklatency = xlsread(MEPpositions_filename); %extracts MEP positions (in ms) from the excel file named above
    mrkOK = (mrklatency.*srate)'; %adjusts the latencies to the sampling rate of the EMG file

    %% MEP analysis 
     outmat = zeros(size(mrkOK,2),6); % create an empty matrix
    for thismarker = 1:size(mrkOK,2) % for each of the markers
        disp(num2str(thismarker))
        
        % MEP size
        peakmin1 = min(data((mrkOK(thismarker)+MEPstart*winbyms):(mrkOK(thismarker)+MEPstop*winbyms),muscle1));
        peakmax1 = max(data((mrkOK(thismarker)+MEPstart*winbyms):(mrkOK(thismarker)+MEPstop*winbyms),muscle1));
        outmat(thismarker, 1) = peakmax1-peakmin1; %p-p amplitude of MEP
        thisMEP= abs(data(mrkOK(thismarker)+MEPstart*winbyms:mrkOK(thismarker)+MEPstop*winbyms,muscle1));
        outmat(thismarker, 2) = trapz (1:size(thisMEP,1), thisMEP); %area under the curve

        % BACKGROUND EMG ACTIVATION
        emgpeakmin1 = min(data((mrkOK(thismarker)-EMGprestop*winbyms):(mrkOK(thismarker)-EMGprestart*winbyms),muscle1));
        emgpeakmax1 = max(data((mrkOK(thismarker)-EMGprestop*winbyms):(mrkOK(thismarker)-EMGprestart*winbyms),muscle1));
        outmat(thismarker, 3) = emgpeakmax1-emgpeakmin1; % p-p amplitude of activity in the 100ms window before the TMS pulse
        outmat(thismarker, 4) = mean(abs(data((mrkOK(thismarker)-EMGprestop*winbyms):(mrkOK(thismarker)-EMGprestart*winbyms),muscle1))); % rectified average of activity in the 100ms window before the TMS pulse
        
        %MEP onset
       % METHOD 1: automatic method that defines latency as the time at which the post-pulse EMG activity exceeds 2 SD above the EMG activity in the 50 ms preceding the pulse (Huang and Mouraux 2015; Hordacre et al. 2017; İşcan et al. 2018; Torrecillos et al. 2020). 
        thisMEP= data(mrkOK(thismarker)+MEPstart*winbyms:mrkOK(thismarker)+MEPstop*winbyms,muscle1);
        thisPRE = data(mrkOK(thismarker)-51*winbyms:mrkOK(thismarker)-1*winbyms,muscle1);
        threshold=mean(thisPRE)+(2*std(thisPRE));
       if find(thisMEP>threshold,1) > 0
             outmat(thismarker, 5) = find(thisMEP>threshold,1)/srate*1000+MEPstart;
        else
             outmat(thismarker, 5) = 0;
        end 

        % METHOD 2: this is calculated as the start of the ascending leg of the first
        %positive peak of the MEP, i.e. the start of the positive segment
        %of the first derivative.
        firstder = gradient(thisMEP(:));
        peakpos=find(thisMEP==peakmax1,1);
        tmp=flip(firstder(1:peakpos));
        l = find(tmp < 0, 1);
        outmat(thismarker, 6) = (peakpos-l)/srate*1000+MEPstart; %onset time in ms
    end
%% writes the results matrix into an excel file
    xlswrite(output_filename,outmat)

disp('End')

%% HOW TO READ THE OUTPUT
%column1 = p-p MEP amplitude
%column 2 = area under the curve of the MEP
%column 3 = p-p of the background EMG activity (100 ms before pulse)
%column 4= rectified mean of the background EMG activity (100 ms before pulse)
%column 5 = MEP onset latency, method 1
%column 6 = MEP onset latency, method 2
