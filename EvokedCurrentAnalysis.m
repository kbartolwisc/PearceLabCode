clear all
close all
[filename, path] = uigetfile('*.xlsx');
%file = "001-042.xlsx";
tab = readtable(strcat(path, filename));

%%
close all
x_values = tab.Time;
% Extract all other variables into a matrix
y_values_matrix = tab{:, tab.Properties.VariableNames ~= "Time"};
plot(x_values, y_values_matrix);
xlimits = [2000 3000]; %zoom around voltage step in ms
xlim(xlimits);
hold on 

window = [2500 2600];
xline(window(1));
xline(window(2));

numCols = numel(tab(1, :));
for i=2:numCols % first column is Time
    trace = tab.Properties.VariableNames{i};
    windowedTrace = [tab.Time(tab.Time >= window(1) & tab.Time <= window(2)), tab.(trace)(tab.Time >= window(1) & tab.Time <= window(2))];
    [minX, minIndex] = min(windowedTrace(:, 2));
    minY = windowedTrace(minIndex, 1);
    plot(1100, -810, 'o-');

end


