%% setup
%clear all variables and graphs
clear all; 
close all;
%dir must be a string to a directory containing a folder named 'XLS_files'
%and a folder named 'Results'
%dataDir = uigetdir; 
dataDir = /Volumes/rapearce/Kai/Ai32_Series/CD/20260218; %TODO: delete
genotype = "AI32xL5cre";
date = "20260218";
LEDdur = 10; %in ms
conditions = ["saline", "3uM KSEB", "10uM KSEB", "30uM KSEB"];
trials = [%first and last trial for each of the conditions
    2 3; %saline
    7 8; %3uM KSEB
    9 11; %10uM KSEB
    15 18]; %30uM KSEB
isiValues = [10 20 40 80 160 320 640 1280];
%%
%combine xls files
xlsDir = append(dataDir, "/XLS_files");
results = combineXlsFiles(xlsDir);

condition = strings(numel(results(:,1)),1); %create a new column to label condidion
for(i=1:numel(trials(:,1)))
    s = trials(i,1);
    e = trials(i,2);
    range = s:e;
    condition(range) = conditions(i);
end

