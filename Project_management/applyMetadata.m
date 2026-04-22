function applyMetadata(projectDir)
% Reads the master experiment_metadata.xlsx at the project root and
% applies it to each per-date experiment.mat in YYYYMMDD subfolders.
%
% Usage:
%   applyMetadata()
%   applyMetadata('/path/to/project')
%
% After running, each date folder's experiment.mat will have
% experiment.recordings(j).metadata populated and experiment.notes set.

if nargin < 1
    projectDir = uigetdir('', 'Select project root folder');
    if isequal(projectDir, 0)
        error('No folder selected.');
    end
end

xlsxPath = fullfile(projectDir, 'analysis', 'experiment_metadata.xlsx');
if ~exist(xlsxPath, 'file')
    error('No experiment_metadata.xlsx found in: %s', fullfile(projectDir, 'analysis'));
end

T1 = readtable(xlsxPath, 'Sheet', 'Recordings', 'TextType', 'string');
T2 = readtable(xlsxPath, 'Sheet', 'DayNotes',   'TextType', 'string');

fixedCols = {'date', 'filename', 'expNumber'};
metaCols  = T1.Properties.VariableNames(~ismember(T1.Properties.VariableNames, fixedCols));

listing  = dir(projectDir);
dateDirs = listing([listing.isdir] & ...
    ~cellfun(@isempty, regexp({listing.name}, '^\d{8}$', 'match')));

nApplied = 0;
nMissing = 0;

for iDay = 1:numel(dateDirs)
    matPath = fullfile(projectDir, 'analysis', [dateDirs(iDay).name '_experiment.mat']);
    if ~exist(matPath, 'file')
        continue;
    end

    loaded     = load(matPath, 'experiment');
    experiment = loaded.experiment;

    mask = strcmp(T2.date, experiment.date);
    if any(mask)
        experiment.notes = char(T2.dayNotes(find(mask, 1)));
    end

    for j = 1:numel(experiment.recordings)
        fname = experiment.recordings(j).filename;
        row   = T1(strcmp(T1.date, experiment.date) & strcmp(T1.filename, fname), :);

        if isempty(row)
            warning('applyMetadata:missingRow', ...
                'No metadata row for: %s / %s', experiment.date, fname);
            nMissing = nMissing + 1;
            continue;
        end

        for c = 1:numel(metaCols)
            col = metaCols{c};
            val = row.(col)(1);
            if isstring(val)
                val = char(val);
            end
            experiment.recordings(j).metadata.(matlab.lang.makeValidName(col)) = val;
        end
        nApplied = nApplied + 1;
    end

    save(matPath, 'experiment');
end

fprintf('Metadata applied to %d recording(s)', nApplied);
if nMissing > 0
    fprintf(' (%d missing rows in Excel)', nMissing);
end
fprintf('.\n');
end
