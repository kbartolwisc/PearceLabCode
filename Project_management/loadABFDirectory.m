function loadABFDirectory()
% Scans a project root folder for YYYYMMDD date subfolders, loads .abf
% files from each, and saves a per-date experiment.mat inside each
% subfolder. A single master experiment_metadata.xlsx is maintained at
% the project root.
%
% Run this whenever new recordings are added. Only new .abf files are
% processed; existing experiment.mat files are updated incrementally.
%
% After filling in experiment_metadata.xlsx, apply metadata with:
%   applyMetadata()
%
% Data structure (per-date experiment.mat):
%   experiment.date          'YYYY-MM-DD'
%   experiment.notes         day-level notes (from Excel DayNotes sheet)
%   experiment.recordings(j) struct array, one per .abf file
%       .filename            .abf filename (basename only)
%       .expNumber           last 3 digits of filename stem
%       .data                raw array from abfload
%                            (2D if gap-free, 3D if episodic,
%                             cell array if variable-length episodic)
%       .si                  sampling interval (microseconds)
%       .fs                  sampling rate (Hz)
%       .header              full header struct from abfload
%       .mode                'gap-free' | 'episodic' | 'variable-length episodic'
%       .channelNames        cell array of channel name strings
%       .metadata            struct populated by applyMetadata()

projectDir = uigetdir('', 'Select project root folder');
if isequal(projectDir, 0)
    error('No folder selected.');
end

abfloadDir = fullfile(fileparts(mfilename('filepath')), '..', 'abfload');
if exist(fullfile(abfloadDir, 'abfload.m'), 'file')
    addpath(abfloadDir);
end

analysisDir = fullfile(projectDir, 'analysis');
if ~exist(analysisDir, 'dir')
    mkdir(analysisDir);
end

xlsxPath = fullfile(analysisDir, 'experiment_metadata.xlsx');

listing  = dir(projectDir);
dateDirs = listing([listing.isdir] & ...
    ~cellfun(@isempty, regexp({listing.name}, '^\d{8}$', 'match')));

if isempty(dateDirs)
    error('No YYYYMMDD subfolders found in: %s', projectDir);
end

fprintf('Found %d date folder(s).\n\n', numel(dateDirs));

totalNew   = 0;
newEntries = struct('date', {}, 'filename', {});
newDates   = {};

for iDay = 1:numel(dateDirs)
    folderName = dateDirs(iDay).name;
    dateStr    = folderNameToDate(folderName);
    dateDir    = fullfile(projectDir, folderName);
    matPath    = fullfile(analysisDir, [folderName '_experiment.mat']);

    if exist(matPath, 'file')
        loaded     = load(matPath, 'experiment');
        experiment = loaded.experiment;
        if isempty(experiment.recordings)
            existingFiles = {};
        else
            existingFiles = {experiment.recordings.filename};
        end
        fprintf('[%s] %d recording(s) already present.\n', folderName, numel(existingFiles));
    else
        experiment    = struct('date', dateStr, 'notes', '', 'recordings', []);
        existingFiles = {};
        newDates{end+1} = dateStr; %#ok<AGROW>
    end

    abfFiles = dir(fullfile(dateDir, '*.abf'));
    isNew    = ~ismember({abfFiles.name}, existingFiles);
    newFiles = abfFiles(isNew);

    if isempty(newFiles)
        fprintf('[%s] No new files.\n\n', folderName);
        continue;
    end

    fprintf('[%s] Loading %d new file(s)...\n', folderName, numel(newFiles));

    nLoaded = 0;
    for k = 1:numel(newFiles)
        fname = newFiles(k).name;
        fprintf('  [%d/%d] %s ... ', k, numel(newFiles), fname);

        try
            [data, si, h] = abfload(fullfile(dateDir, fname), 'doDispInfo', false);
        catch ME
            fprintf('FAILED (%s)\n', ME.message);
            continue;
        end
        fprintf('ok  (%s)\n', detectMode(data));

        rec.filename     = fname;
        rec.expNumber    = expNumFromFilename(fname);
        rec.data         = data;
        rec.si           = si;
        rec.fs           = 1e6 / si;
        rec.header       = h;
        rec.mode         = detectMode(data);
        rec.channelNames = channelNamesFromHeader(h);
        rec.metadata     = struct();

        if isempty(experiment.recordings)
            experiment.recordings = rec;
        else
            experiment.recordings(end+1) = rec;
        end

        newEntries(end+1) = struct('date', dateStr, 'filename', fname); %#ok<AGROW>
        nLoaded = nLoaded + 1;
    end

    if nLoaded > 0
        save(matPath, 'experiment');
        fprintf('  -> Saved %s\n', matPath);
        totalNew = totalNew + nLoaded;
    end
    fprintf('\n');
end

if totalNew == 0
    fprintf('No new recordings found. Everything is up to date.\n');
    return;
end

% Reload all experiments to build master Excel
allExperiments = struct('date', {}, 'notes', {}, 'recordings', {});
for iDay = 1:numel(dateDirs)
    matPath = fullfile(analysisDir, [dateDirs(iDay).name '_experiment.mat']);
    if exist(matPath, 'file')
        loaded = load(matPath, 'experiment');
        allExperiments(end+1) = loaded.experiment; %#ok<AGROW>
    end
end

updateTemplate(allExperiments, xlsxPath, newEntries, newDates);

fprintf('Added %d new recording(s) across %d date folder(s).\n', totalNew, numel(dateDirs));
fprintf('Template: %s\n', xlsxPath);
fprintf('\nNext steps:\n');
fprintf('  1. Fill in new rows in experiment_metadata.xlsx\n');
fprintf('  2. applyMetadata()\n');
end


% -------------------------------------------------------------------------
function dateStr = folderNameToDate(folderName)
dateStr = sprintf('%s-%s-%s', folderName(1:4), folderName(5:6), folderName(7:8));
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

function updateTemplate(allExperiments, xlsxPath, newEntries, newDates)
recCols = {'date','filename','expNumber', ...
           'animalID','cellID','condition','genotype','sex','age','recordingNotes'};

newKeys = arrayfun(@(e) [e.date '/' e.filename], newEntries, 'UniformOutput', false);

newRows = {};
for i = 1:numel(allExperiments)
    for j = 1:numel(allExperiments(i).recordings)
        r   = allExperiments(i).recordings(j);
        key = [allExperiments(i).date '/' r.filename];
        if ismember(key, newKeys)
            newRows(end+1, :) = {allExperiments(i).date, r.filename, r.expNumber, ...
                                 '', '', '', '', '', '', ''}; %#ok<AGROW>
        end
    end
end
T1new = cell2table(newRows, 'VariableNames', recCols);

newDateRows = {};
for i = 1:numel(allExperiments)
    if ismember(allExperiments(i).date, newDates)
        newDateRows(end+1, :) = {allExperiments(i).date, ''}; %#ok<AGROW>
    end
end
T2new = cell2table(newDateRows, 'VariableNames', {'date', 'dayNotes'});

if exist(xlsxPath, 'file')
    T1existing = readtable(xlsxPath, 'Sheet', 'Recordings', 'TextType', 'string');
    T2existing = readtable(xlsxPath, 'Sheet', 'DayNotes',   'TextType', 'string');
    T1new = reconcileColumns(T1new, T1existing);
    writetable([T1existing; T1new], xlsxPath, 'Sheet', 'Recordings');
    writetable([T2existing; T2new], xlsxPath, 'Sheet', 'DayNotes');
else
    writetable(T1new, xlsxPath, 'Sheet', 'Recordings');
    writetable(T2new, xlsxPath, 'Sheet', 'DayNotes');
end
end

function Tnew = reconcileColumns(Tnew, Texisting)
existingCols = Texisting.Properties.VariableNames;
newCols      = Tnew.Properties.VariableNames;
missing      = setdiff(existingCols, newCols);
for c = 1:numel(missing)
    Tnew.(missing{c}) = repmat({''}, height(Tnew), 1);
end
Tnew = Tnew(:, existingCols);
end
