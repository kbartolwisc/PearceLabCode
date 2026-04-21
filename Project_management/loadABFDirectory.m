function experiment = loadABFDirectory()
% Prompts for a folder, loads all .abf files, groups them by recording date
% read from the ABF file header, and saves experiment.mat plus a blank
% Excel metadata template in that folder.
%
% If experiment.mat already exists in the selected folder, only new .abf
% files are loaded. Existing metadata in the Excel template is preserved
% and new rows are appended for the new files only.
%
% Usage:
%   experiment = loadABFDirectory();
%
% After filling in the Excel template, merge metadata with:
%   experiment = applyMetadata(experiment);
%   save('experiment.mat', 'experiment');
%
% Data structure:
%   experiment(i).date          'YYYY-MM-DD'
%   experiment(i).notes         day-level notes (from Excel Sheet 2)
%   experiment(i).recordings(j) struct array, one per .abf file
%       .filename               original filename
%       .expNumber              last 3 digits of filename stem
%       .data                   raw array from abfload
%                               (2D if gap-free, 3D if episodic,
%                                cell array if variable-length episodic)
%       .si                     sampling interval (microseconds)
%       .fs                     sampling rate (Hz)
%       .header                 full header struct from abfload
%       .mode                   'gap-free' | 'episodic' | 'variable-length episodic'
%       .channelNames           cell array of channel name strings
%       .metadata               struct populated by applyMetadata()

abfDir = uigetdir('', 'Select folder containing .abf files');
if isequal(abfDir, 0)
    error('No folder selected.');
end

% Add abfload to path relative to this file
abfloadDir = fullfile(fileparts(mfilename('filepath')), '..', 'abfload');
if exist(fullfile(abfloadDir, 'abfload.m'), 'file')
    addpath(abfloadDir);
end

matPath  = fullfile(abfDir, 'experiment.mat');
xlsxPath = fullfile(abfDir, 'experiment_metadata.xlsx');

% Pre-populate dateMap from existing experiment if one exists
dateMap       = containers.Map('KeyType', 'char', 'ValueType', 'any');
existingFiles = {};

if exist(matPath, 'file')
    loaded = load(matPath, 'experiment');
    for i = 1:numel(loaded.experiment)
        e = loaded.experiment(i);
        day.date       = e.date;
        day.notes      = e.notes;
        day.recordings = e.recordings;
        dateMap(e.date) = day;
        for j = 1:numel(e.recordings)
            existingFiles{end+1} = e.recordings(j).filename; %#ok<AGROW>
        end
    end
    fprintf('Existing experiment loaded: %d recording(s) already present.\n\n', ...
        numel(existingFiles));
end

% Find .abf files and filter to only new ones
allFiles = dir(fullfile(abfDir, '*.abf'));
if isempty(allFiles)
    error('No .abf files found in: %s', abfDir);
end

isNew    = ~ismember({allFiles.name}, existingFiles);
newFiles = allFiles(isNew);

if isempty(newFiles)
    fprintf('No new .abf files found. Experiment is already up to date.\n');
    experiment = loaded.experiment;
    return;
end

fprintf('Found %d new .abf file(s) to load...\n\n', numel(newFiles));

% Load new files into dateMap
newFilenames = {};
datesBefore  = keys(dateMap);

for k = 1:numel(newFiles)
    fname = newFiles(k).name;
    fprintf('  [%d/%d] %s ... ', k, numel(newFiles), fname);

    try
        [d, si, h] = abfload(fullfile(abfDir, fname), 'doDispInfo', false);
    catch ME
        fprintf('FAILED (%s)\n', ME.message);
        continue;
    end
    fprintf('ok  (%s)\n', detectMode(d));

    dateStr = dateFromHeader(h);
    expNum  = expNumFromFilename(fname);

    rec.filename     = fname;
    rec.expNumber    = expNum;
    rec.data         = d;
    rec.si           = si;
    rec.fs           = 1e6 / si;
    rec.header       = h;
    rec.mode         = detectMode(d);
    rec.channelNames = channelNamesFromHeader(h);
    rec.metadata     = emptyMetadata(dateStr);

    if isKey(dateMap, dateStr)
        day = dateMap(dateStr);
        day.recordings(end+1) = rec;
        dateMap(dateStr) = day;
    else
        day.date       = dateStr;
        day.notes      = '';
        day.recordings = rec;
        dateMap(dateStr) = day;
    end

    newFilenames{end+1} = fname; %#ok<AGROW>
end

if isempty(newFilenames)
    error('All new files failed to load.');
end

% Build sorted struct array
dateKeys   = sort(keys(dateMap));
experiment = struct([]);
for i = 1:numel(dateKeys)
    e = dateMap(dateKeys{i});
    experiment(i).date       = e.date;
    experiment(i).notes      = e.notes;
    experiment(i).recordings = e.recordings;
end

% Identify dates that are brand new (not previously in the experiment)
newDates = setdiff(keys(dateMap), datesBefore);

save(matPath, 'experiment');
updateTemplate(experiment, xlsxPath, newFilenames, newDates);

nNew  = numel(newFilenames);
nDays = numel(experiment);
nRecs = sum(arrayfun(@(e) numel(e.recordings), experiment));
fprintf('\nAdded %d new recording(s). Experiment now has %d recording(s) across %d day(s).\n', ...
    nNew, nRecs, nDays);
fprintf('Saved:    %s\n', matPath);
fprintf('Template: %s\n', xlsxPath);
fprintf('\nNext steps:\n');
fprintf('  1. Fill in new rows in experiment_metadata.xlsx\n');
fprintf('  2. experiment = applyMetadata(experiment);\n');
fprintf('  3. save(''%s'', ''experiment'');\n', matPath);
end


% -------------------------------------------------------------------------
function dateStr = dateFromHeader(h)
if isfield(h, 'uFileStartDate') && h.uFileStartDate > 0
    d      = h.uFileStartDate;
    year   = floor(d / 10000);
    month  = floor(mod(d, 10000) / 100);
    day    = mod(d, 100);
    dateStr = sprintf('%04d-%02d-%02d', year, month, day);
else
    dateStr = 'unknown';
    warning('loadABFDirectory:noDate', ...
        'Could not read date from ABF header (ABF1 format?). Grouping under "unknown".');
end
end

function expNum = expNumFromFilename(fname)
[~, name] = fileparts(fname);
tok = regexp(name, '(\d{3})$', 'tokens');
if ~isempty(tok)
    expNum = str2double(tok{1}{1});
else
    expNum = 0;
    warning('loadABFDirectory:expNum', ...
        'Could not parse 3-digit experiment number from filename: %s', fname);
end
end

function mode = detectMode(d)
if iscell(d)
    mode = 'variable-length episodic';
elseif ndims(d) == 3
    mode = 'episodic';
else
    mode = 'gap-free';
end
end

function names = channelNamesFromHeader(h)
if isfield(h, 'recChNames')
    names = h.recChNames;
else
    names = {};
end
end

function m = emptyMetadata(dateStr)
m.animalID  = '';
m.cellID    = '';
m.condition = '';
m.date      = dateStr;
m.genotype  = '';
m.sex       = '';
m.age       = '';
m.notes     = '';
end

function updateTemplate(experiment, xlsxPath, newFilenames, newDates)
% Build new rows for recordings that were just added
recCols = {'date','filename','expNumber', ...
           'animalID','cellID','condition','genotype','sex','age','recordingNotes'};
newRows = {};
for i = 1:numel(experiment)
    for j = 1:numel(experiment(i).recordings)
        r = experiment(i).recordings(j);
        if ismember(r.filename, newFilenames)
            newRows(end+1, :) = {experiment(i).date, r.filename, r.expNumber, ...
                                 '', '', '', '', '', '', ''}; %#ok<AGROW>
        end
    end
end
T1new = cell2table(newRows, 'VariableNames', recCols);

% Build new rows for dates that did not previously exist
newDateRows = {};
for i = 1:numel(experiment)
    if ismember(experiment(i).date, newDates)
        newDateRows(end+1, :) = {experiment(i).date, ''}; %#ok<AGROW>
    end
end
T2new = cell2table(newDateRows, 'VariableNames', {'date', 'dayNotes'});

if exist(xlsxPath, 'file')
    % Append to existing Excel, preserving filled-in metadata
    T1existing = readtable(xlsxPath, 'Sheet', 'Recordings', 'TextType', 'string');
    T2existing = readtable(xlsxPath, 'Sheet', 'DayNotes',   'TextType', 'string');

    % Match column sets: add any user-added columns as empty in new rows
    T1new = reconcileColumns(T1new, T1existing);

    writetable([T1existing; T1new], xlsxPath, 'Sheet', 'Recordings');
    writetable([T2existing; T2new], xlsxPath, 'Sheet', 'DayNotes');
else
    writetable(T1new, xlsxPath, 'Sheet', 'Recordings');
    writetable(T2new, xlsxPath, 'Sheet', 'DayNotes');
end
end

function Tnew = reconcileColumns(Tnew, Texisting)
% Ensures Tnew has the same columns as Texisting, filling any extras with ''.
existingCols = Texisting.Properties.VariableNames;
newCols      = Tnew.Properties.VariableNames;
missing      = setdiff(existingCols, newCols);
for c = 1:numel(missing)
    Tnew.(missing{c}) = repmat({''}, height(Tnew), 1);
end
% Reorder to match existing column order
Tnew = Tnew(:, existingCols);
end
