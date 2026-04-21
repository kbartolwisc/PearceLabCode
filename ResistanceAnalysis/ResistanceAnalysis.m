close all
clear all
%file = uigetfile('*.xlsx');
file = "001-042.xlsx";
tab = readtable(file);

%%
close all
xlimits = [50 120]; %zoom around voltage step in ms
vStep = -10; %voltage step in mv

xlimitsIndex = (tab.Time >= xlimits(1)) & (tab.Time <= xlimits(2));
xlimitsTrace = tab(xlimitsIndex, :);
xampleTrace = xlimitsTrace(:, [1 2]);
plot(xampleTrace.Time, xampleTrace.Trace1)
hold on

fitlims = [88.2 92];
fitlimsIndex = (xampleTrace.Time >= fitlims(1)) & (xampleTrace.Time <= fitlims(2));
fitlimsTrace = xampleTrace(fitlimsIndex, :);

f = fit(fitlimsTrace.Time, fitlimsTrace.Trace1, 'exp2');

plot(f, fitlimsTrace.Time, fitlimsTrace.Trace1)
xline(fitlims(1));
xline(fitlims(2));

%%




baseline = [60 85]; %baseline region to compute in mv
step = [92 97];
hold on
xline(baseline(1));
xline(baseline(2));
xline(step(1));
xline(step(2));

numCols = numel(tab(1, :));
%prealocate table to hold our results
sz = [numCols-1, 3];
varTypes = ["string", "double", "double"];
varNames = ["Trace", "Ra", "Rm"];
resistances = table('Size', sz, 'VariableTypes', varTypes, 'VariableNames',varNames);

for i=2:numCols % first column is Time_ms_
    trace = tab.Properties.VariableNames{i};
    baselineAvg = mean(tab.(trace)(tab.Time >=baseline(1) ...
                         & tab.Time <=baseline(2)));

    stepAvg =  mean(tab.(trace)(tab.Time >=step(1) ...
                         & tab.Time <=step(2)));

    peak = min(tab.(trace)(tab.Time >= baseline(2) & ...
                      tab.Time <= step(1)));

    stepDiff = baselineAvg - stepAvg; % in pA
    peakDiff = baselineAvg - peak;

    accessR = abs(vStep) / abs(peakDiff) * 1000;
    membraneR = abs(vStep) / abs(stepDiff) * 1000;

    resistances(i-1,:) = {trace, accessR, membraneR};
end
