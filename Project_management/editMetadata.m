function editMetadata(projectDir)
% GUI editor for experiment_metadata.xlsx.
%
% Usage:
%   editMetadata()
%   editMetadata('/path/to/project')
%
% Controls:
%   Date dropdown + Prev/Next buttons  — navigate recording days
%   Day notes field                    — free-text notes for the whole day
%   Table                              — edit per-recording metadata inline
%   Fill Down button                   — copies the selected cell value to
%                                        all rows below it in that column
%   Save / Ctrl+S                      — write changes to Excel

if nargin < 1
    projectDir = uigetdir('', 'Select project root folder');
    if isequal(projectDir, 0), return; end
end

xlsxPath = fullfile(projectDir, 'analysis', 'experiment_metadata.xlsx');
if ~exist(xlsxPath, 'file')
    error('No experiment_metadata.xlsx found in: %s', fullfile(projectDir, 'analysis'));
end

T1 = readtable(xlsxPath, 'Sheet', 'Recordings', 'TextType', 'string');
T2 = readtable(xlsxPath, 'Sheet', 'DayNotes',   'TextType', 'string');

% readtable infers numeric type for columns that contain only numbers
% (e.g. animal IDs like 1001). Force all columns except expNumber to string.
for i_ = 1:width(T1)
    col_ = T1.Properties.VariableNames{i_};
    if ~strcmp(col_, 'expNumber') && isnumeric(T1.(col_))
        v_             = string(T1.(col_));
        v_(isnan(T1.(col_))) = "";
        T1.(col_)      = v_;
    end
end
clear i_ col_ v_

fixedCols   = {'date','filename','expNumber'};
allCols     = T1.Properties.VariableNames;
metaCols    = allCols(~ismember(allCols, fixedCols));
displayCols = [{'filename','expNumber'}, metaCols];
colEditable = [false, false, true(1, numel(metaCols))];

dates = cellstr(unique(T1.date));
if isempty(dates)
    error('No recordings found in metadata.');
end

currentDate    = dates{1};
unsavedChanges = false;
lastSelection  = [];

% -------------------------------------------------------------------------
% Build UI
% -------------------------------------------------------------------------
fig = uifigure('Name', 'Experiment Metadata Editor', ...
    'Position',          [100 100 920 530], ...
    'KeyPressFcn',       @onKeyPress, ...
    'CloseRequestFcn',   @onClose);

% Project path
uilabel(fig, 'Position', [10 496 900 22], ...
    'Text',      ['Project: ' projectDir], ...
    'FontSize',  10, 'FontColor', [0.45 0.45 0.45], 'Interpreter', 'none');

% Date row
uilabel(fig, 'Position', [10 462 42 26], 'Text', 'Date:', 'FontWeight', 'bold');
dateDropdown = uidropdown(fig, 'Position', [55 460 165 28], ...
    'Items',              dates, ...
    'Value',              dates{1}, ...
    'ValueChangedFcn',    @onDateChanged);
uibutton(fig, 'Position', [228 460 70 28], 'Text', '< Prev', ...
    'ButtonPushedFcn', @onPrev);
uibutton(fig, 'Position', [304 460 70 28], 'Text', 'Next >', ...
    'ButtonPushedFcn', @onNext);
statusLabel = uilabel(fig, 'Position', [420 462 490 26], ...
    'Text',              'All changes saved.', ...
    'HorizontalAlignment', 'right', ...
    'FontColor',         [0.18 0.52 0.18]);

% Day notes row
uilabel(fig, 'Position', [10 425 78 26], 'Text', 'Day notes:', 'FontWeight', 'bold');
dayNotesField = uieditfield(fig, 'text', 'Position', [91 423 819 28], ...
    'Placeholder',        'Notes for this recording day...', ...
    'ValueChangedFcn',    @(~,~) markUnsaved());

% Table
tbl = uitable(fig, ...
    'Position',              [10 60 900 355], ...
    'ColumnEditable',        colEditable, ...
    'RowName',               [], ...
    'FontSize',              11, ...
    'CellEditCallback',      @onCellEdited, ...
    'SelectionChangedFcn',   @onSelectionChanged);

% Bottom row
uibutton(fig, 'Position', [10 15 120 32], 'Text', 'Fill Down', ...
    'Tooltip',         'Copy selected cell value to all rows below in that column', ...
    'ButtonPushedFcn', @onFillDown);
uilabel(fig, 'Position', [140 15 530 32], ...
    'Text',      'Select a cell then click Fill Down to apply its value to all rows below.', ...
    'FontColor', [0.5 0.5 0.5], 'FontSize', 10);
uibutton(fig, 'Position', [780 15 130 32], 'Text', 'Save  (Ctrl+S)', ...
    'BackgroundColor', [0.16 0.50 0.20], 'FontColor', 'white', ...
    'FontWeight',      'bold', ...
    'ButtonPushedFcn', @onSave);

refreshTable();

% -------------------------------------------------------------------------
% Callbacks
% -------------------------------------------------------------------------

    function refreshTable()
        mask  = strcmp(cellstr(T1.date), currentDate);
        rows  = T1(mask, :);
        nRows = height(rows);
        nCols = numel(displayCols);
        data  = cell(nRows, nCols);

        for c = 1:nCols
            v = rows.(displayCols{c});
            for r = 1:nRows
                if isnumeric(v)
                    data{r,c} = v(r);
                else
                    data{r,c} = char(v(r));
                end
            end
        end

        tbl.Data        = data;
        tbl.ColumnName  = displayCols;
        tbl.ColumnWidth = colWidths(displayCols);

        mask2 = strcmp(cellstr(T2.date), currentDate);
        if any(mask2)
            dayNotesField.Value = char(T2.dayNotes(find(mask2, 1)));
        else
            dayNotesField.Value = '';
        end
    end

    function switchToDate(newDate)
        if strcmp(newDate, currentDate), return; end
        if unsavedChanges
            choice = uiconfirm(fig, ...
                sprintf('Save changes to %s before switching?', currentDate), ...
                'Unsaved Changes', ...
                'Options',       {'Save','Discard','Cancel'}, ...
                'DefaultOption', 'Save');
            if strcmp(choice, 'Cancel')
                dateDropdown.Value = currentDate;
                return;
            end
            if strcmp(choice, 'Save'), saveToExcel(); end
            unsavedChanges = false;
        end
        currentDate   = newDate;
        lastSelection = [];
        dateDropdown.Value = currentDate;
        refreshTable();
    end

    function onDateChanged(~, event)
        switchToDate(event.Value);
    end

    function onPrev(~,~)
        idx = find(strcmp(dates, currentDate));
        if idx > 1, switchToDate(dates{idx-1}); end
    end

    function onNext(~,~)
        idx = find(strcmp(dates, currentDate));
        if idx < numel(dates), switchToDate(dates{idx+1}); end
    end

    function onSelectionChanged(src, ~)
        lastSelection = src.Selection;
    end

    function onCellEdited(~, event)
        colName = displayCols{event.Indices(2)};
        if ismember(colName, fixedCols), return; end

        mask    = strcmp(cellstr(T1.date), currentDate);
        rowsIdx = find(mask);
        t1Row   = rowsIdx(event.Indices(1));

        val = event.NewData;
        if isnumeric(T1.(colName))
            if ~isnumeric(val), val = str2double(val); end
            T1.(colName)(t1Row) = val;
        else
            T1.(colName)(t1Row) = string(val);
        end

        markUnsaved();
    end

    function onFillDown(~,~)
        if isempty(lastSelection)
            uialert(fig, 'Click a cell in the table first, then press Fill Down.', ...
                'No Cell Selected');
            return;
        end

        startRow = lastSelection(1,1);
        colIdx   = lastSelection(1,2);
        colName  = displayCols{colIdx};

        if ismember(colName, fixedCols)
            uialert(fig, sprintf('"%s" is read-only.', colName), 'Fill Down');
            return;
        end

        val     = tbl.Data{startRow, colIdx};
        mask    = strcmp(cellstr(T1.date), currentDate);
        rowsIdx = find(mask);

        for r = startRow:numel(rowsIdx)
            tbl.Data{r, colIdx} = val;
            t1Row = rowsIdx(r);
            if isnumeric(T1.(colName))
                T1.(colName)(t1Row) = str2double(char(val));
            else
                T1.(colName)(t1Row) = string(val);
            end
        end

        markUnsaved();
    end

    function onSave(~,~)
        saveToExcel();
    end

    function onKeyPress(~, event)
        if strcmp(event.Key, 's') && ismember('control', event.Modifier)
            saveToExcel();
        end
    end

    function onClose(~,~)
        if unsavedChanges
            choice = uiconfirm(fig, 'Save changes before closing?', 'Unsaved Changes', ...
                'Options',       {'Save','Discard','Cancel'}, ...
                'DefaultOption', 'Save');
            if strcmp(choice, 'Cancel'), return; end
            if strcmp(choice, 'Save'), saveToExcel(); end
        end
        delete(fig);
    end

    function saveToExcel()
        % Sync day notes for the current date into T2
        mask2 = strcmp(cellstr(T2.date), currentDate);
        if any(mask2)
            T2.dayNotes(find(mask2, 1)) = string(dayNotesField.Value);
        else
            T2 = [T2; table(string(currentDate), string(dayNotesField.Value), ...
                'VariableNames', {'date','dayNotes'})];
        end

        writetable(T1, xlsxPath, 'Sheet', 'Recordings');
        writetable(T2, xlsxPath, 'Sheet', 'DayNotes');

        unsavedChanges = false;
        c = clock;
        statusLabel.Text = sprintf('Saved at %02d:%02d:%02d', c(4), c(5), floor(c(6)));
        statusLabel.FontColor = [0.18 0.52 0.18];
    end

    function markUnsaved()
        unsavedChanges = true;
        statusLabel.Text      = 'Unsaved changes';
        statusLabel.FontColor = [0.72 0.33 0.08];
    end

end


% -------------------------------------------------------------------------
function widths = colWidths(cols)
known = struct('filename', 140, 'expNumber', 70, 'animalID', 90, ...
               'cellID', 70, 'condition', 110, 'genotype', 90, ...
               'sex', 50, 'age', 50, 'recordingNotes', 180);
widths = cell(1, numel(cols));
for i = 1:numel(cols)
    field = matlab.lang.makeValidName(cols{i});
    if isfield(known, field)
        widths{i} = known.(field);
    else
        widths{i} = 100;
    end
end
end
