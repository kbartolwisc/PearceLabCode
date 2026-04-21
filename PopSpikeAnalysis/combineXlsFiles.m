function cTab = combineXlsFiles(dataDir)
%UNTITLED4 Summary of this function goes here
%   Detailed explanation goes here
%TODO: remove

%Extract the names of the XLS files in the folder
xlsDirListing = dir(dataDir);
xlsDirTbl = struct2table(xlsDirListing);
namedFiles = xlsDirTbl(xlsDirTbl.isdir==false,:);

i = 1;
iPath = dataDir + "/" + namedFiles(i,:).name;
cTab = readtable(iPath, 'VariableNamingRule','preserve');

for i=2:height(namedFiles)
    iPath = dataDir + "/" + namedFiles(i,:).name;
    cTab = [cTab; readtable(iPath, 'VariableNamingRule','preserve')];
end
end