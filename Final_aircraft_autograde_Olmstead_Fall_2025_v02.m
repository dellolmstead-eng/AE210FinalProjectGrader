
% Final Aircraft Autograder: Grades AE210 Jet11 Excel submissions, logs feedback, and optionally exports Blackboard-compatible scores.


%--------------------------------------------------------------------------
% AE210 Final Aircraft Autograder Script â€“ Fall 2025
% Version: v3.1.1
%
% Description:
% This script automates grading for the AE210 Final Aircraft assignment by processing Jet11 Excel files (*.xlsm). It evaluates
% multiple design criteria, generates detailed feedback, and outputs both a
% summary log and an optional Blackboard-compatible grade import file.
% 
% Sometimes the parallel pool requires 2 runs to start correctly. Just stop
% the run and restart if it freezes for more than 30 seconds.
%
% Key Features:
% - Supports both single-file and batch-folder grading via GUI
% - Parallel-safe execution using MATLAB's parpool
% - Robust Excel reading with fallback for missing data
% - Detailed feedback log per cadet with scoring breakdown
% - Optional export to Blackboard offline grade format (SMART_TEXT)
% - Histogram visualization of score distribution
%
% Inputs:
% - User-selected Excel file or folder of files
%
% Outputs:
% - Text log file: textout_<timestamp>.txt
% - Histogram of scores
% - Optional Blackboard CSV: FinalAircraft_Blackboard_Offline_<timestamp>.csv
%
% Embedded Functions:
% - gradeCadet: Grades a single cadet's file and returns score and feedback
% - loadAllJet11Sheets: Loads all required sheets from a Jet11 Excel file
% - safeReadMatrix: Robustly reads numeric data from Excel, with fallback to readcell
% - cell2sub: Converts Excel cell references (e.g., 'G4') to row/col indices
% - sub2excel: Converts row/col indices back to Excel cell references
% - logf: Appends formatted text to a log string
% - selectRunMode: GUI for selecting single file or folder mode
% - promptAndGenerateBlackboardCSV: Dialog + export to Blackboard SMART_TEXT format
%
% Author: Lt Col Dell Olmstead, based on work by Capt Carol Bryant and Capt Anna Mason
% Heavy ChatGPT CODEX help in Nov 2025
% Last Updated: 10 Dec 2025 15:55
%--------------------------------------------------------------------------
clear; close all; clc;


%% Choose directory and get Excel files
% fprintf('Executing %s\n',mfilename);
% I recommend updating the below line to point to your Final Aircraft files. It works
% as is, but will default to the right place if this is updated.

% folderAnalyzed = uigetdir('C:\Users\dell.olmstead\OneDrive - afacademy.af.edu\Documents 1\01 Classes\AE210 FA24\Design Project\Final Aircraft files');
% fprintf('%s\n\n', folderAnalyzed);
% files = dir(fullfile(folderAnalyzed, '*.xlsm'));



%% Select run mode: single file or folder, start parallel pool if folder
[mode, selectedPath] = selectRunMode();
tic
if strcmp(mode, 'cancelled')
    disp('Operation cancelled by user.');
    return;
elseif strcmp(mode, 'single')
    folderAnalyzed = fileparts(selectedPath);
    files = dir(selectedPath);  % single file
elseif strcmp(mode, 'folder')
    % Ensure a process-based parallel pool is active
    poolobj = gcp('nocreate'); % Get the current pool, if any
    if isempty(poolobj)
        % Create a new local pool, ensuring process-based if possible
        try
            p = parpool('local'); % Try the simplest form first
        catch ME
            if contains(ME.message, 'ExecutionMode') % Check for specific error message
                p = parpool('local', 'ExecutionMode', 'Processes'); % Use ExecutionMode if supported
            else
                rethrow(ME); % If it's a different error, re-throw it
            end
        end

        if ~isempty(p)
            if isa(p, 'parallel.ThreadPool')
                warning('Created a thread-based pool despite requesting "local". Attempting to delete and recreate as process-based.');
                delete(p);
                parpool('local', 'ExecutionMode', 'Processes'); % Explicitly use ExecutionMode
            elseif isa(p, 'parallel.Pool')
                fprintf('Successfully created a process-based local parallel pool.\n');
            end
        end
    elseif isa(poolobj, 'parallel.ThreadPool')
        % If an existing pool is thread-based, delete it and create a process-based one
        warning('Existing parallel pool is thread-based. Deleting and creating a process-based local pool.');
        delete(poolobj);
        parpool('local', 'ExecutionMode', 'Processes'); % Explicitly use ExecutionMode
    elseif isa(poolobj, 'parallel.Pool')
        fprintf('A process-based local parallel pool is already running.\n');
    end
    folderAnalyzed = selectedPath;
    files = [dir(fullfile(folderAnalyzed, '*.xlsm')); dir(fullfile(folderAnalyzed, '*.xlsx')); dir(fullfile(folderAnalyzed, '*.xls'))];
else
    error('Unknown mode selected.');
end


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%% Iterate through cadets %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
textout = strings(numel(files), 1);
points = 10*ones(numel(files),1);  % Initialize points for each file
feedback = cell(1,numel(files));
numFiles = numel(files);

fprintf('Reading %d files\n', numel(files));

if strcmp(mode, 'folder')        
    % Combined parallel read + grade
    
parfor cadetIdx = 1:numel(files)
    filename = fullfile(folderAnalyzed, files(cadetIdx).name);
    try
        [pt, fb] = gradeCadet(filename);
        points(cadetIdx) = pt;
        feedback{cadetIdx} = fb; % keep raw feedback for log file
        dispText = sprintf('#%d/%d %s\n%s', cadetIdx, numFiles, files(cadetIdx).name, char(fb));
        fprintf('Feedback (%d/%d) for %s:\n%s\n', cadetIdx, numFiles, files(cadetIdx).name, dispText);
    catch ME
        points(cadetIdx) = NaN;
        location = '';
        if ~isempty(ME.stack)
            topFrame = ME.stack(1);
            location = sprintf(' (%s line %d)', topFrame.name, topFrame.line);
        end
        feedback{cadetIdx} = sprintf('Error reading or grading file: %s\n%s%s', files(cadetIdx).name, ME.message, location);
        fprintf('Error grading %s: %s\n', files(cadetIdx).name, feedback{cadetIdx});
    end
end

else %       %%% Use the below code to run a single cadet

    filename = fullfile(folderAnalyzed, files(1).name);
    [points, feedback{1}] = gradeCadet(filename);
    dispText = sprintf('#%d/%d %s\n%s', 1, numFiles, files(1).name, char(feedback{1}));
    fprintf('Feedback (1/%d) for %s:\n%s\n', numFiles, files(1).name, dispText);

end


%% Set up log file
timestamp = char(datetime('now', 'Format', 'yyyy-MM-dd_HH-mm-ss'));
logFilePath = fullfile(folderAnalyzed, ['textout_', timestamp, '.txt']);
finalout = fopen(logFilePath,'w');

% Log file header
fprintf(finalout, 'FinalAircraft Autograder Log\n');
fprintf(finalout, 'Script Name: %s.m\n', mfilename);
fprintf(finalout, 'Run Date: %s\n', string(datetime('now', 'Format', 'yyyy-MM-dd HH:mm:ss')));
fprintf(finalout, 'Analyzed Folder: %s\n', folderAnalyzed);
fprintf(finalout, 'Files to Analyze (%d):\n', numel(files));
for i = 1:numel(files)
    fprintf(finalout, '  - %s\n', files(i).name);
end
fprintf(finalout, '\n');

%% Concatenate all outputs into one text file and write it.
fbCells = cellfun(@(x) strtrim(string(x)), feedback, 'UniformOutput', false);
allLogText = strjoin([fbCells{:}], '\n');
fprintf(finalout, '%s', allLogText); % Write accumulated log text
fclose(finalout);


%% Prompt user to export Blackboard CSV

promptAndGenerateBlackboardCSV(folderAnalyzed, files, points, feedback, timestamp);



%%  Create a histogram with 10 bins
figure;  % Open a new figure window
histogram(points, 10);
% Add labels and title
xlabel('Scores');
ylabel('Count');
title('Distribution of Scores');

duration=toc;
fprintf('Average time was %0.1f seconds per cadet\n',duration/numel(files))
%% Give link to the log file
fprintf('Open the output file here:\n <a href="matlab:system(''notepad %s'')">%s</a>\n', logFilePath, logFilePath);




%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%Embedded functions%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% This is the main code that does all the evaluations. It is here so
% it can be called using a for loop for one file, and a parfor loop for
% many files.
function [pt, fb] = gradeCadet(filename) % Read the sheet

sheets = loadAllJet11Sheets(filename);

Aero = sheets.Aero;
Miss = sheets.Miss;
Main = sheets.Main;
Consts = sheets.Consts;
Gear = sheets.Gear;
Geom = sheets.Geom;

pt = 45;
logText = "";
bonusPoints = 0;

[~, name, ext] = fileparts(filename);
logText = logf(logText, '%s\n', [name, ext]);

tol = 1e-3;
altTol = 1;
machTol = 1e-2;
timeTol = 1e-2;
distTol = 1e-3;
fuelTol = 5e-2;
bonusFullEps = 1e-6;
bonusMinDisplay = 1e-2;
NON_VIABLE_CAP = 22.5;

% Pre-initialize key inputs, then populate from the workbook
fuel_available = NaN;
fuel_capacity = NaN;
fuel_required = NaN;
volume_remaining = NaN;
cost = NaN;
numaircraft = NaN;
radius = NaN;
aim120 = NaN;
aim9 = NaN;
takeoff_dist = NaN;
landing_dist = NaN;

fuel_available = Main(18, 15);
fuel_capacity = Main(15, 15);
fuel_required = Main(40, 24);
volume_remaining = Main(23, 17);
cost = Main(31, 17);
numaircraft = Main(31, 14);
radius = Main(37, 25);
aim120 = Main(3, 28);
aim9 = Main(4, 28);
takeoff_dist = Main(12, 24);
landing_dist = Main(13, 24);

if ~isnan(fuel_available) && ~isnan(fuel_capacity) && fuel_capacity ~= 0
    betaDefault = 1 - fuel_available/(2*fuel_capacity);
else
    betaDefault = 0.87620980519917;
end
betaExpected = betaDefault;

ConstraintsMach = Main(4, 21);

% Sheet validation to catch missing geometry inputs early
geomBlock = Main(18:27, 2:8);
geomNaN = isnan(geomBlock);
skipCells = [...
    7 1;  % B24 intentionally blank
    7 2;  % C24 intentionally blank
    5 4;  % E22 intentionally blank
    5 6;  % G22 intentionally blank
    10 3; % D27 intentionally blank
    10 4; % E27 intentionally blank
    10 5; % F27 intentionally blank
    10 6; % G27 intentionally blank
    9 7]; % H26 intentionally blank
for k = 1:size(skipCells,1)
    geomNaN(skipCells(k,1), skipCells(k,2)) = false;
end
if any(geomNaN, 'all')
    [rIdx, cIdx] = find(geomNaN);
    badCells = arrayfun(@(r,c) sub2excel(r+17, c+1), rIdx, cIdx, 'UniformOutput', false);
    logText = logf(logText, 'Sheet validation: Geometry inputs B18:H27 must be numeric (missing at %s).\n', strjoin(badCells, ', '));
    fb = char(logText);
    pt = 0;
    return;
end

geomBlock2 = Main(34:53, 3:6);
geomNaN2 = isnan(geomBlock2);
if any(geomNaN2, 'all')
    [rIdx, cIdx] = find(geomNaN2);
    badCells = arrayfun(@(r,c) sub2excel(r+33, c+2), rIdx, cIdx, 'UniformOutput', false);
    logText = logf(logText, 'Sheet validation: Geometry inputs C34:F53 must be numeric (missing at %s).\n', strjoin(badCells, ', '));
    fb = char(logText);
    pt = 0;
    return;
end

% Aero tab programming (3 pts)
aeroIssues = 0;
if isequal(Aero(3,7), Aero(4,7)), aeroIssues = aeroIssues + 1; end
if isequal(Aero(10,7), Aero(11,7)), aeroIssues = aeroIssues + 1; end
if isequal(Aero(15,1), Aero(16,1)), aeroIssues = aeroIssues + 1; end

aeroDeduction = min(3, aeroIssues);
if aeroDeduction > 0
    pt = pt - aeroDeduction;
    logText = logf(logText, '-%d pts Aero tab formulas not active in cells A15, G3, or G10\n', aeroDeduction);
end
% Mission table (2 pts max deduction)
MissionArray = Main(33:44, 11:25);
colIdx = [1, 2, 3, 4, 6, 8, 9, 12, 13];
alt = MissionArray(1, colIdx);
mach = MissionArray(3, colIdx);
ab = MissionArray(4, colIdx);
dist = MissionArray(6, colIdx);
timeLeg = MissionArray(7, colIdx);

missionErrors = 0;

if abs(alt(1)) > altTol || abs(ab(1) - 100) > tol
    logText = logf(logText, 'Leg 1: Altitude must be 0 ft and AB = 100%% (found alt=%.1f, AB=%.1f)\n', alt(1), ab(1));
    missionErrors = missionErrors + 1;
end

if alt(2) < min(alt(1), alt(3)) - altTol || alt(2) > max(alt(1), alt(3)) + altTol
    logText = logf(logText, 'Leg 2: Altitude must remain between legs 1 and 3 (found alt2=%.1f, alt1=%.1f, alt3=%.1f)\n', alt(2), alt(1), alt(3));
    missionErrors = missionErrors + 1;
end
if mach(2) < min(mach(1), mach(3)) - machTol || mach(2) > max(mach(1), mach(3)) + machTol
    logText = logf(logText, 'Leg 2: Mach must remain between legs 1 and 3 (found mach2=%.2f, mach1=%.2f, mach3=%.2f)\n', mach(2), mach(1), mach(3));
    missionErrors = missionErrors + 1;
end
if abs(ab(2)) > tol
    logText = logf(logText, 'Leg 2: AB must be 0%% (found AB=%.1f)\n', ab(2));
    missionErrors = missionErrors + 1;
end

if alt(3) < 35000 - altTol || abs(mach(3) - 0.9) > machTol || abs(ab(3)) > tol
    logText = logf(logText, 'Leg 3: Must be >= 35,000 ft, Mach = 0.9, AB = 0%% (found alt=%.1f, mach=%.2f, AB=%.1f)\n', alt(3), mach(3), ab(3));
    missionErrors = missionErrors + 1;
end

if alt(4) < 35000 - altTol || abs(mach(4) - 0.9) > machTol || abs(ab(4)) > tol
    logText = logf(logText, 'Leg 4: Must be >= 35,000 ft, Mach = 0.9, AB = 0%% (found alt=%.1f, mach=%.2f, AB=%.1f)\n', alt(4), mach(4), ab(4));
    missionErrors = missionErrors + 1;
end

if alt(5) < 35000 - altTol || abs(mach(5) - ConstraintsMach) > machTol || abs(ab(5)) > tol || dist(5) < 150 - distTol
    logText = logf(logText, 'Leg 5: Must be >= 35,000 ft, Mach = constraint Supercruise Mach (Main!U4), AB = 0%%, Distance >= 150 nm (found alt=%.1f, mach=%.2f, AB=%.1f, dist=%.1f)\n', alt(5), mach(5), ab(5), dist(5));
    missionErrors = missionErrors + 1;
end

if alt(6) < 30000 - altTol || mach(6) < 1.2 - machTol || abs(ab(6) - 100) > tol
    logText = logf(logText, 'Leg 6: Must be >= 30,000 ft, Mach >= 1.2, AB = 100%% (found alt=%.1f, mach=%.2f, AB=%.1f)\n', alt(6), mach(6), ab(6));
    missionErrors = missionErrors + 1;
end
if Main(39, 28) < 720
    logText = logf(logText, 'Two full 360 turns are required. Increase total turn angle (cell AB39) to 720 degrees or greater to meet the combat turn requirement\n');
    missionErrors = missionErrors + 1;
end

if alt(7) < 35000 - altTol || abs(mach(7) - ConstraintsMach) > machTol || abs(ab(7)) > tol || dist(7) < 150 - distTol
    logText = logf(logText, 'Leg 7: Must be >= 35,000 ft, Mach = constraint Supercruise Mach (Main!U4), AB = 0%%, Distance >= 150 nm (found alt=%.1f, mach=%.2f, AB=%.1f, dist=%.1f)\n', alt(7), mach(7), ab(7), dist(7));
    missionErrors = missionErrors + 1;
end

if alt(8) < 35000 - altTol || abs(mach(8) - 0.9) > machTol || abs(ab(8)) > tol
    logText = logf(logText, 'Leg 8: Must be >= 35,000 ft, Mach = 0.9, AB = 0%% (found alt=%.1f, mach=%.2f, AB=%.1f)\n', alt(8), mach(8), ab(8));
    missionErrors = missionErrors + 1;
end

if abs(alt(9) - 10000) > altTol || abs(mach(9) - 0.4) > machTol || abs(ab(9)) > tol || abs(timeLeg(9) - 20) > timeTol
    logText = logf(logText, 'Leg 9: Must be 10,000 ft, Mach = 0.4, AB = 0%%, Time = 20 min (found alt=%.1f, mach=%.2f, AB=%.1f, time=%.2f)\n', alt(9), mach(9), ab(9), timeLeg(9));
    missionErrors = missionErrors + 1;
end

if ~isnan(radius)
    if radius < 375 - distTol
        logText = logf(logText, 'Mission radius below threshold (375 nm): %.1f\n', radius);
        missionErrors = missionErrors + 1;
    elseif radius >= 410 - distTol
        logText = logf(logText, 'Mission radius meets objective (410 nm): %.1f\n', radius);
    end
end

missionDeduction = min(2, missionErrors);
if missionDeduction > 0
    pt = pt - missionDeduction;
    logText = logf(logText, '-%d pts Mission profile inputs incorrect\n', missionDeduction);
end

% Tavailable > D check (3 pts)
thrust_drag = Miss(48:49, 3:14);
thrustShort = thrust_drag(2, :) <= thrust_drag(1, :);
thrustFailures = sum(thrustShort);
if thrustFailures > 0
    deduction = min(3, thrustFailures);
    pt = pt - deduction;
    logText = logf(logText, '-%d pts Not enough thrust: Tavailable <= D for %d mission segment(s)\n', deduction, thrustFailures);
end
% Control surface attachment (2 pts)
controlFailures = 0;
VALUE_TOL = 1e-3;
AR_TOL = 0.1;
VT_WING_FRACTION = 0.8;
EDGE_ALIGN_TOL = 0.2;

fuselage_length = Main(32, 2);
fuselage_end = fuselage_length;
PCS_area = Main(18, 3);
VT_area = Main(18, 8);
strake_area = Main(18, 4);
PCS_x = Main(23, 3);
PCS_root_chord = Geom(8, 3);
if PCS_area >= 1
if any(isnan([fuselage_end, PCS_x, PCS_root_chord]))
    logText = logf(logText, 'Unable to verify PCS placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif PCS_x > (fuselage_end - 0.25 * PCS_root_chord)
    logText = logf(logText, 'PCS X-location too far aft. Must overlap at least 25%% of root chord.\n');
    controlFailures = controlFailures + 1;
end
end

VT_x = Main(23, 8);
VT_root_chord = Geom(10, 3);
if VT_area >= 1
if any(isnan([fuselage_end, VT_x, VT_root_chord]))
    logText = logf(logText, 'Unable to verify vertical tail placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif VT_x > (fuselage_end - 0.25 * VT_root_chord)
    logText = logf(logText, 'VT X-location too far aft. Must overlap at least 25%% of root chord.\n');
    controlFailures = controlFailures + 1;
end
end

PCS_z = Main(25, 3);
fuse_z_center = Main(52, 4);
fuse_z_height = Main(52, 6);
if PCS_area >= 1
if any(isnan([PCS_z, fuse_z_center, fuse_z_height]))
    logText = logf(logText, 'Unable to verify PCS vertical placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif PCS_z < (fuse_z_center - fuse_z_height/2) || PCS_z > (fuse_z_center + fuse_z_height/2)
    logText = logf(logText, 'PCS Z-location outside fuselage vertical bounds.\n');
    controlFailures = controlFailures + 1;
end
end

VT_y = Main(24, 8);
fuse_width = Main(52, 5);
vtMountedOffFuselage = false;
if VT_area >= 1
if any(isnan([VT_y, fuse_width]))
    logText = logf(logText, 'Unable to verify vertical tail lateral placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif abs(VT_y) > fuse_width/2 + VALUE_TOL
    vtMountedOffFuselage = true;
    logText = logf(logText, 'Vertical tail mounted off the fuselage; ensure structural support at the wing.\n');
end
end

if strake_area >= 1
    sweep = Geom(15, 11);
    y = Geom(152, 13);
    strake = Geom(155, 12);
    apex = Geom(38, 12);
    if any(isnan([sweep, y, strake, apex]))
        logText = logf(logText, 'Unable to verify strake attachment due to missing geometry data\n');
        controlFailures = controlFailures + 1;
    else
        wing = (y / tand(90 - sweep) + apex);
        if wing >= (strake + 0.5)
            logText = logf(logText, 'Strake disconnected.\n');
            controlFailures = controlFailures + 1;
        end
    end
end

wingLeadingRoot = geomPlanformPoint(Geom, 38);
wingLeadingTip = geomPlanformPoint(Geom, 39);
wingTrailingRoot = geomPlanformPoint(Geom, 41);
wingTrailingTip = geomPlanformPoint(Geom, 40);

if Main(18, 5) >= 1
    elevonLEInboard = geomPlanformPoint(Geom, 174);
    elevonLEOutboard = geomPlanformPoint(Geom, 175);
    elevonTEOutboard = geomPlanformPoint(Geom, 176);
    elevonTEInboard = geomPlanformPoint(Geom, 177);
    [logText, controlFailures] = checkWingDevicePlacement(logText, controlFailures, ...
        'Elevon', 'trailing edge', elevonTEInboard, elevonTEOutboard, elevonLEInboard, elevonLEOutboard, ...
        wingLeadingRoot, wingLeadingTip, wingTrailingRoot, wingTrailingTip, EDGE_ALIGN_TOL);
end

if Main(18, 6) >= 1
    lefLEInboard = geomPlanformPoint(Geom, 186);
    lefLEOutboard = geomPlanformPoint(Geom, 187);
    lefTEOutboard = geomPlanformPoint(Geom, 188);
    lefTEInboard = geomPlanformPoint(Geom, 189);
    [logText, controlFailures] = checkWingDevicePlacement(logText, controlFailures, ...
        'LE Flap', 'leading edge', lefLEInboard, lefLEOutboard, lefTEInboard, lefTEOutboard, ...
        wingLeadingRoot, wingLeadingTip, wingTrailingRoot, wingTrailingTip, EDGE_ALIGN_TOL);
end

if Main(18, 7) >= 1
    tefLEInboard = geomPlanformPoint(Geom, 198);
    tefLEOutboard = geomPlanformPoint(Geom, 199);
    tefTEOutboard = geomPlanformPoint(Geom, 200);
    tefTEInboard = geomPlanformPoint(Geom, 201);
    [logText, controlFailures] = checkWingDevicePlacement(logText, controlFailures, ...
        'TE Flap', 'trailing edge', tefTEInboard, tefTEOutboard, tefLEInboard, tefLEOutboard, ...
        wingLeadingRoot, wingLeadingTip, wingTrailingRoot, wingTrailingTip, EDGE_ALIGN_TOL);
end

component_positions = Main(23, 2:8);
component_areas = Main(18, 2:8);
active_components = component_positions(component_areas >= 1);
if ~isempty(active_components)
    if isnan(fuselage_end)
        logText = logf(logText, 'Unable to verify component X-location due to missing fuselage length\n');
        controlFailures = controlFailures + 1;
    elseif any(active_components >= fuselage_end)
        logText = logf(logText, 'One or more components X-location extend beyond the fuselage end (B32 = %.2f)\n', fuselage_end);
        controlFailures = controlFailures + 1;
    end
end

if vtMountedOffFuselage
    vtApex = geomPlanformPoint(Geom, 163);
    vtRootTE = geomPlanformPoint(Geom, 166);
    wingTE = geomPlanformPoint(Geom, 41);
    if any(isnan([vtApex(1), vtRootTE(1), wingTE(1)]))
        logText = logf(logText, 'Unable to verify vertical tail overlap with wing due to missing geometry data\n');
        controlFailures = controlFailures + 1;
    else
        chord = vtRootTE(1) - vtApex(1);
        overlap = max(0, min(wingTE(1), vtRootTE(1)) - vtApex(1));
        if ~(chord > 0) || overlap + VALUE_TOL < VT_WING_FRACTION * chord
            logText = logf(logText, 'Vertical tail mounted on the wing must overlap at least 80%% of its root chord with the wing trailing edge.\n');
            controlFailures = controlFailures + 1;
        end
    end
end

wingAR = Main(19, 2);
pcsAR = Main(19, 3);
vtAR = Main(19, 8);
if ~isnan(wingAR) && ~isnan(pcsAR) && pcsAR > wingAR + AR_TOL
    logText = logf(logText, 'Pitch control surface aspect ratio (%.2f) must be lower than wing aspect ratio (%.2f).\n', pcsAR, wingAR);
    controlFailures = controlFailures + 1;
end
if ~isnan(wingAR) && ~isnan(vtAR) && vtAR >= wingAR - AR_TOL
    logText = logf(logText, 'Vertical tail aspect ratio (%.2f) must be lower than wing aspect ratio (%.2f).\n', vtAR, wingAR);
    controlFailures = controlFailures + 1;
end

engine_diameter = Main(29, 8);
inlet_x = Main(31, 6);
compressor_x = Main(32, 6);
engine_start = inlet_x + compressor_x;
widthValues = [];
if ~isnan(engine_start)
    for row = 34:53
        station_x = Main(row, 2);
        width = Main(row, 5);
        if ~isnan(station_x) && ~isnan(width) && station_x >= engine_start
            widthValues(end+1) = width; %#ok<AGROW>
        end
    end
end
if isempty(widthValues) || isnan(engine_diameter)
    logText = logf(logText, 'Unable to verify fuselage width clearance for engines\n');
    controlFailures = controlFailures + 1;
else
    minWidth = min(widthValues);
    maxWidth = max(widthValues);
    requiredWidth = engine_diameter + 0.5;
    if minWidth + VALUE_TOL <= requiredWidth
        logText = logf(logText, 'Fuselage minimum width (%.2f ft) must exceed engine diameter + 0.5 ft (%.2f ft).\n', minWidth, requiredWidth);
        controlFailures = controlFailures + 1;
    end
    allowedOverhang = 2 * maxWidth;
    if ~isnan(fuselage_end)
        pcsTipX = max(Geom(117, 12), Geom(118, 12));
        vtTipX = max(Geom(165, 12), Geom(166, 12));
        if ~isnan(pcsTipX)
            overhang = pcsTipX - fuselage_end;
            if overhang > allowedOverhang + VALUE_TOL
                logText = logf(logText, '%s extends %.2f ft beyond the fuselage end (limit %.2f ft).\n', 'Pitch control surface', overhang, allowedOverhang);
                controlFailures = controlFailures + 1;
            end
        end
        if ~isnan(vtTipX)
            overhang = vtTipX - fuselage_end;
            if overhang > allowedOverhang + VALUE_TOL
                logText = logf(logText, '%s extends %.2f ft beyond the fuselage end (limit %.2f ft).\n', 'Vertical tail', overhang, allowedOverhang);
                controlFailures = controlFailures + 1;
            end
        end
    end
end

engine_length = Main(29, 9);
if any(isnan([engine_diameter, fuselage_end, inlet_x, compressor_x, engine_length]))
    logText = logf(logText, 'Unable to verify engine protrusion due to missing geometry data\n');
    controlFailures = controlFailures + 1;
else
    protrusion = inlet_x + compressor_x + engine_length - fuselage_end;
    if protrusion > engine_diameter + VALUE_TOL
        logText = logf(logText, 'Engine nacelles protrude %.2f ft past the fuselage end (limit %.2f ft).\n', protrusion, engine_diameter);
        controlFailures = controlFailures + 1;
    end
end

if controlFailures > 0
    deduction = min(2, controlFailures);
    pt = pt - deduction;
    logText = logf(logText, '-%d pts Control surface placement issues\n', deduction);
end

% Stealth shaping 
STEALTH_TOL = 5;
CORNER_REFLECTOR_TARGET = 45;
CORNER_REFLECTOR_TOL = 5;
CORNER_REFLECTOR_EDGE_TOL = 0.1;
stealthFailures = 0;
stealthHeaderLogged = false;

wingLeadingAngle = computeEdgeAngleDeg(Geom, 38, 39);
wingTrailingAngle = computeWingTrailingPlanformAngleDeg(Geom);
pcsLeadingAngle = computeEdgeAngleDeg(Geom, 115, 116);
pcsTrailingAngle = computePcsTrailingPlanformAngleDeg(Geom);
strakeLeadingAngle = computeEdgeAngleDeg(Geom, 152, 153);
strakeTrailingAngle = computeEdgeAngleDeg(Geom, 154, 155);
vtLeadingAngle = computeEdgeAngleDeg(Geom, 163, 164);
vtTrailingAngle = computeEdgeAngleDeg(Geom, 165, 166);
wingDihedral = Main(26, 2);
pcsDihedral = Main(26, 3);
vtTilt = Main(27, 8);
wingArea = Main(18, 2);
pcsArea = Main(18, 3);
strakeArea = Main(18, 4);
vtArea = Main(18, 8);
wingActive = isnan(wingArea) || wingArea >= 1;
pcsActive = isnan(pcsArea) || pcsArea >= 1;
strakeActive = isnan(strakeArea) || strakeArea >= 1;
vtActive = isnan(vtArea) || vtArea >= 1;

if pcsActive && wingActive && ~anglesParallel(pcsLeadingAngle, wingLeadingAngle, STEALTH_TOL)
    [logText, stealthHeaderLogged] = logStealth(logText, stealthHeaderLogged, 'Pitch control surface leading edge sweep %.1f° must match the wing leading edge sweep %.1f° (+/- %.1f°).\n', pcsLeadingAngle, wingLeadingAngle, STEALTH_TOL);
    stealthFailures = stealthFailures + 1;
end

if wingActive && ~isnan(wingDihedral) && abs(wingDihedral - CORNER_REFLECTOR_TARGET) < (CORNER_REFLECTOR_TOL - CORNER_REFLECTOR_EDGE_TOL)
    [logText, stealthHeaderLogged] = logStealth(logText, stealthHeaderLogged, 'Main!B26 wing dihedral angle %.1f° creates a corner reflector because it is within %.1f° of 45°. Increase or decrease at least %.1f° off 45° to avoid significant stealth signature increase.\n', wingDihedral, CORNER_REFLECTOR_TOL, CORNER_REFLECTOR_TOL);
    stealthFailures = stealthFailures + 1;
end
if pcsActive && ~isnan(pcsDihedral) && abs(pcsDihedral - CORNER_REFLECTOR_TARGET) < (CORNER_REFLECTOR_TOL - CORNER_REFLECTOR_EDGE_TOL)
    [logText, stealthHeaderLogged] = logStealth(logText, stealthHeaderLogged, 'Main!C26 pitch control surface dihedral angle %.1f° creates a corner reflector because it is within %.1f° of 45°. Increase or decrease at least %.1f° off 45° to avoid significant stealth signature increase.\n', pcsDihedral, CORNER_REFLECTOR_TOL, CORNER_REFLECTOR_TOL);
    stealthFailures = stealthFailures + 1;
end
if vtActive && ~isnan(vtTilt) && abs(vtTilt - CORNER_REFLECTOR_TARGET) < (CORNER_REFLECTOR_TOL - CORNER_REFLECTOR_EDGE_TOL)
    [logText, stealthHeaderLogged] = logStealth(logText, stealthHeaderLogged, 'Main!H27 vertical tail tilt angle %.1f° creates a corner reflector because it is within %.1f° of 45°. Increase or decrease at least %.1f° off 45° to avoid significant stealth signature increase.\n', vtTilt, CORNER_REFLECTOR_TOL, CORNER_REFLECTOR_TOL);
    stealthFailures = stealthFailures + 1;
end

wingTipTE = geomPlanformPoint(Geom, 40);
wingCenterTE = geomPlanformPoint(Geom, 41);
pcsTipTE = geomPlanformPoint(Geom, 117);
pcsInnerTE = geomPlanformPoint(Geom, 118);
vtTipTE = geomPlanformPoint(Geom, 165);
vtInnerTE = geomPlanformPoint(Geom, 166);
VT_z = Main(25, 8);
if wingActive
    if ~(anglesParallel(wingTrailingAngle, wingLeadingAngle, STEALTH_TOL) || teNormalHitsCenterline(wingTipTE, wingCenterTE, fuselage_length))
        [logText, stealthHeaderLogged] = logStealth(logText, stealthHeaderLogged, 'Wing trailing edge %.1f° is not parallel to the leading edge and its normal does not reach the fuselage centerline (+/- %.1f°).\n', wingTrailingAngle, STEALTH_TOL);
        stealthFailures = stealthFailures + 1;
    end
elseif pcsActive || strakeActive || vtActive
    [logText, stealthHeaderLogged] = logStealth(logText, stealthHeaderLogged, 'Unable to verify stealth shaping due to missing geometry data\n');
    stealthFailures = stealthFailures + 1;
end

if pcsActive && ~isnan(pcsDihedral) && pcsDihedral > 5
    [logText, stealthFailures, stealthHeaderLogged] = requireParallelAngle(logText, stealthFailures, stealthHeaderLogged, pcsLeadingAngle, wingLeadingAngle, STEALTH_TOL, 'Pitch control surface leading edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
    [logText, stealthFailures, stealthHeaderLogged] = requireParallelAngleOrCenterlineIfWithinFuselageHeight(logText, stealthFailures, stealthHeaderLogged, pcsTrailingAngle, wingLeadingAngle, STEALTH_TOL, 'Pitch control surface trailing edge sweep %.1f° must be parallel to the wing leading edge %.1f° or its normal must reach the fuselage centerline when the surface remains within the fuselage average height (+/- %.1f°).\n', pcsTipTE, pcsInnerTE, isSurfaceWithinFuselageHeight(PCS_z, pcsDihedral, pcsTipTE, pcsInnerTE, fuse_z_center, fuse_z_height), fuselage_length);
end

if strakeActive
    [logText, stealthFailures, stealthHeaderLogged] = requireParallelAngle(logText, stealthFailures, stealthHeaderLogged, strakeLeadingAngle, wingLeadingAngle, STEALTH_TOL, 'Strake leading edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
    [logText, stealthFailures, stealthHeaderLogged] = requireParallelAngle(logText, stealthFailures, stealthHeaderLogged, strakeTrailingAngle, wingLeadingAngle, STEALTH_TOL, 'Strake trailing edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
end

if ~vtActive
    % ignore
elseif isnan(vtTilt)
    [logText, stealthHeaderLogged] = logStealth(logText, stealthHeaderLogged, 'Unable to verify stealth shaping due to missing geometry data\n');
    stealthFailures = stealthFailures + 1;
elseif vtTilt < 85
    [logText, stealthFailures, stealthHeaderLogged] = requireParallelAngle(logText, stealthFailures, stealthHeaderLogged, vtLeadingAngle, wingLeadingAngle, STEALTH_TOL, 'Vertical tail leading edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
    [logText, stealthFailures, stealthHeaderLogged] = requireParallelAngleOrCenterlineIfWithinFuselageHeight(logText, stealthFailures, stealthHeaderLogged, vtTrailingAngle, wingLeadingAngle, STEALTH_TOL, 'Vertical tail trailing edge sweep %.1f° must be parallel to the wing leading edge %.1f° or its normal must reach the fuselage centerline when the tail remains within the fuselage average height (+/- %.1f°).\n', vtTipTE, vtInnerTE, isSurfaceWithinFuselageHeight(VT_z, vtTilt, vtTipTE, vtInnerTE, fuse_z_center, fuse_z_height), fuselage_length);
end

stealthDeduction = min(5, stealthFailures); % up to 5-point hit for stealth issues
if stealthFailures > 0
    pt = pt - stealthDeduction;
    logText = logf(logText, '-%d pts Stealth shaping issues detected.\n', stealthDeduction);
end

% Constraint table values (2 pts max deduction)
constraintErrors = 0;
objectiveSet = struct('MaxMach', false, 'CruiseMach', false, 'CmbtTurn1', false, 'CmbtTurn2', false, 'Ps1', false, 'Ps2', false);
rowErrorsMap = objectiveSet;
curveStatus = struct();
labelNames = struct('MaxMach', 'MaxMach', 'CruiseMach', 'CruiseMach', 'CmbtTurn1', 'Cmbt Turn1', 'CmbtTurn2', 'Cmbt Turn2', 'Ps1', 'Ps1', 'Ps2', 'Ps2');

if Main(3,20) < 35000 - altTol
    logText = logf(logText, 'MaxMach: Altitude = %.0f, must be >= %.0f\n', Main(3,20), 35000);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.MaxMach = true;
end
if Main(3,21) < 2.0 - machTol
    logText = logf(logText, 'MaxMach: Mach = %.2f, must be >= %.2f\n', Main(3,21), 2.0);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.MaxMach = true;
elseif Main(3,21) >= 2.2 - machTol
    objectiveSet.MaxMach = true;
end
if Main(3,22) < 1 - tol
    logText = logf(logText, 'MaxMach: n = %.3f, expected %.3f\n', Main(3,22), 1);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.MaxMach = true;
end
if abs(Main(3,23) - 100) > tol
    logText = logf(logText, 'MaxMach: AB = %.0f%%, expected %.0f%%\n', Main(3,23), 100);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.MaxMach = true;
end
if abs(Main(3,24)) > tol
    logText = logf(logText, 'MaxMach: Ps = %.0f, expected %.0f\n', Main(3,24), 0);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.MaxMach = true;
end
if abs(Main(3,19) - betaExpected) > 1e-3
    logText = logf(logText, 'MaxMach: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(3,19));
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.MaxMach = true;
end

if Main(4,20) < 35000 - altTol
    logText = logf(logText, 'CruiseMach: Altitude = %.0f, must be >= %.0f\n', Main(4,20), 35000);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CruiseMach = true;
end
if Main(4,21) < 1.5 - machTol
    logText = logf(logText, 'CruiseMach: Mach = %.2f, must be >= %.2f\n', Main(4,21), 1.5);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CruiseMach = true;
elseif Main(4,21) >= 1.8 - machTol
    objectiveSet.CruiseMach = true;
end
if Main(4,22) < 1 - tol
    logText = logf(logText, 'CruiseMach: n = %.3f, expected %.3f\n', Main(4,22), 1);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CruiseMach = true;
end
if abs(Main(4,23)) > tol
    logText = logf(logText, 'CruiseMach: AB = %.0f%%, expected %.0f%%\n', Main(4,23), 0);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CruiseMach = true;
end
if abs(Main(4,24)) > tol
    logText = logf(logText, 'CruiseMach: Ps = %.0f, expected %.0f\n', Main(4,24), 0);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CruiseMach = true;
end
if abs(Main(4,19) - betaExpected) > 1e-3
    logText = logf(logText, 'CruiseMach: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(4,19));
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CruiseMach = true;
end

if abs(Main(6,20) - 30000) > altTol
    logText = logf(logText, 'Cmbt Turn1: Altitude = %.0f, expected %.0f\n', Main(6,20), 30000);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn1 = true;
end
if abs(Main(6,21) - 1.2) > machTol
    logText = logf(logText, 'Cmbt Turn1: Mach = %.2f, expected %.2f\n', Main(6,21), 1.2);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn1 = true;
end
if Main(6,22) < 3 - tol
    logText = logf(logText, 'Cmbt Turn1: g-load = %.3f, must be >= %.1f\n', Main(6,22), 3.0);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn1 = true;
elseif Main(6,22) >= 4 - tol
    objectiveSet.CmbtTurn1 = true;
end
if abs(Main(6,23) - 100) > tol
    logText = logf(logText, 'Cmbt Turn1: AB = %.0f%%, expected %.0f%%\n', Main(6,23), 100);
   constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn1 = true;
end
if abs(Main(6,24)) > tol
    logText = logf(logText, 'Cmbt Turn1: Ps = %.0f, expected %.0f\n', Main(6,24), 0);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn1 = true;
end
if abs(Main(6,19) - betaExpected) > 1e-3
    logText = logf(logText, 'Cmbt Turn1: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(6,19));
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn1 = true;
end

if abs(Main(7,20) - 10000) > altTol
    logText = logf(logText, 'Cmbt Turn2: Altitude = %.0f, expected %.0f\n', Main(7,20), 10000);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn2 = true;
end
if abs(Main(7,21) - 0.9) > machTol
    logText = logf(logText, 'Cmbt Turn2: Mach = %.2f, expected %.2f\n', Main(7,21), 0.9);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn2 = true;
end
if Main(7,22) < 4 - tol
    logText = logf(logText, 'Cmbt Turn2: g-load = %.3f, must be >= %.1f\n', Main(7,22), 4.0);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn2 = true;
elseif Main(7,22) >= 4.5 - tol
    objectiveSet.CmbtTurn2 = true;
end
if abs(Main(7,23) - 100) > tol
    logText = logf(logText, 'Cmbt Turn2: AB = %.0f%%, expected %.0f%%\n', Main(7,23), 100);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn2 = true;
end
if abs(Main(7,24)) > tol
    logText = logf(logText, 'Cmbt Turn2: Ps = %.0f, expected %.0f\n', Main(7,24), 0);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn2 = true;
end
if abs(Main(7,19) - betaExpected) > 1e-3
    logText = logf(logText, 'Cmbt Turn2: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(7,19));
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.CmbtTurn2 = true;
end

if abs(Main(8,20) - 30000) > altTol
    logText = logf(logText, 'Ps1: Altitude = %.0f, expected %.0f\n', Main(8,20), 30000);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps1 = true;
end
if abs(Main(8,21) - 1.15) > machTol
    logText = logf(logText, 'Ps1: Mach = %.2f, expected %.2f\n', Main(8,21), 1.15);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps1 = true;
end
if Main(8,22) < 1 - tol
    logText = logf(logText, 'Ps1: n = %.3f, expected %.3f\n', Main(8,22), 1);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps1 = true;
end
if abs(Main(8,23) - 100) > tol
    logText = logf(logText, 'Ps1: AB = %.0f%%, expected %.0f%%\n', Main(8,23), 100);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps1 = true;
end
if Main(8,24) < 400 - distTol
    logText = logf(logText, 'Ps1: Ps = %.0f, must be >= %.0f\n', Main(8,24), 400);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps1 = true;
elseif ~isnan(Main(8,24)) && Main(8,24) >= 500 - distTol
    objectiveSet.Ps1 = true;
end
if abs(Main(8,19) - betaExpected) > 1e-3
    logText = logf(logText, 'Ps1: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(8,19));
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps1 = true;
end

if abs(Main(9,20) - 10000) > altTol
    logText = logf(logText, 'Ps2: Altitude = %.0f, expected %.0f\n', Main(9,20), 10000);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps2 = true;
end
if abs(Main(9,21) - 0.9) > machTol
    logText = logf(logText, 'Ps2: Mach = %.2f, expected %.2f\n', Main(9,21), 0.9);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps2 = true;
end
if Main(9,22) < 1 - tol
    logText = logf(logText, 'Ps2: n = %.3f, expected %.3f\n', Main(9,22), 1);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps2 = true;
end
if abs(Main(9,23)) > tol
    logText = logf(logText, 'Ps2: AB = %.0f%%, expected %.0f%%\n', Main(9,23), 0);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps2 = true;
end
if Main(9,24) < 400 - distTol
    logText = logf(logText, 'Ps2: Ps = %.0f, must be >= %.0f\n', Main(9,24), 400);
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps2 = true;
elseif ~isnan(Main(9,24)) && Main(9,24) >= 500 - distTol
    objectiveSet.Ps2 = true;
end
if abs(Main(9,19) - betaExpected) > 1e-3
    logText = logf(logText, 'Ps2: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(9,19));
    constraintErrors = constraintErrors + 1;
    rowErrorsMap.Ps2 = true;
end

if abs(Main(12,20)) > altTol
    logText = logf(logText, 'Takeoff: Altitude = %.0f, expected %.0f\n', Main(12,20), 0);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,21) - 1.2) > machTol
    logText = logf(logText, 'Takeoff: Mach = %.2f, expected %.2f\n', Main(12,21), 1.2);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,22) - 0.03) > 5e-4
    logText = logf(logText, 'Takeoff: n = %.3f, expected %.3f\n', Main(12,22), 0.03);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,23) - 100) > tol
    logText = logf(logText, 'Takeoff: AB = %.0f%%, expected %.0f%%\n', Main(12,23), 100);
    constraintErrors = constraintErrors + 1;
end
if takeoff_dist > 3000 + distTol
    logText = logf(logText, 'Takeoff distance exceeds threshold (3000 ft): %.0f\n', takeoff_dist);
    constraintErrors = constraintErrors + 1;
elseif ~isnan(takeoff_dist) && takeoff_dist <= 2500 + distTol
    logText = logf(logText, 'Takeoff distance meets objective (<= 2500 ft): %.0f\n', takeoff_dist);
end
if abs(Main(12,19) - 1.0) > tol
    logText = logf(logText, 'Takeoff: W/WTO must be 1.000 (full fuel) within +/- %.3f; found %.3f\n', tol, Main(12,19));
    constraintErrors = constraintErrors + 1;
end

cdxTakeoff = Main(12,25);
if ~(abs(cdxTakeoff) <= tol || abs(cdxTakeoff - 0.035) <= tol)
    logText = logf(logText, 'Takeoff: CDx = %.3f must be one of 0.000, 0.035\n', cdxTakeoff);
    constraintErrors = constraintErrors + 1;
end

if abs(Main(13,20)) > altTol
    logText = logf(logText, 'Landing: Altitude = %.0f, expected %.0f\n', Main(13,20), 0);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,21) - 1.3) > machTol
    logText = logf(logText, 'Landing: Mach = %.2f, expected %.2f\n', Main(13,21), 1.3);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,22) - 0.5) > tol
    logText = logf(logText, 'Landing: n = %.3f, expected %.3f\n', Main(13,22), 0.5);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,23)) > tol
    logText = logf(logText, 'Landing: AB = %.0f%%, expected %.0f%%\n', Main(13,23), 0);
    constraintErrors = constraintErrors + 1;
end
if landing_dist > 5000 + distTol
    logText = logf(logText, 'Landing distance exceeds threshold (5000 ft): %.0f\n', landing_dist);
    constraintErrors = constraintErrors + 1;
elseif ~isnan(landing_dist) && landing_dist <= 3500 + distTol
    logText = logf(logText, 'Landing distance meets objective (<= 3500 ft): %.0f\n', landing_dist);
end
if abs(Main(13,19) - 1.0) > tol
    logText = logf(logText, 'Landing: W/WTO must be 1.000 (loaded worst case) within +/- %.3f; found %.3f\n', tol, Main(13,19));
    constraintErrors = constraintErrors + 1;
end

cdxLanding = Main(13,25);
if ~(abs(cdxLanding) <= tol || abs(cdxLanding - 0.045) <= tol)
    logText = logf(logText, 'Landing: CDx = %.3f must be one of 0.000, 0.045\n', cdxLanding);
    constraintErrors = constraintErrors + 1;
end

constraintDeduction = min(2, constraintErrors);
if constraintDeduction > 0
    pt = pt - constraintDeduction;
    logText = logf(logText, '-%d pts One or more constraint table entries are incorrect\n', constraintDeduction);
end
% Constraint curve compliance (8 pts)
constraintCurveFailures = 0;
failedCurves = {};
curveDiagnostics = strings(0,1);
curveSuffixFew = '';
curveSuffixMany = ' Consider seeking EI; multiple constraints remain unmet.';

try
    WS_axis = Consts(22, 11:31);
    WS_axis = double(WS_axis);

    constraintRows = [23, 24, 26, 27, 28, 29, 32];
    columnLabels = {"MaxMach", "Supercruise", "CombatTurn1", "CombatTurn2", "Ps1", "Ps2", "Takeoff"};

    WS_design = Main(13, 16);
    TW_design = Main(13, 17);

    for idx = 1:numel(constraintRows)
        row = constraintRows(idx);
        TW_curve = Consts(row, 11:31);
        TW_curve = double(TW_curve);
        estimatedTWvalue = interp1(WS_axis, TW_curve, WS_design, 'pchip', 'extrap');
        if ~isnan(estimatedTWvalue)
            fieldName = mapCurveField(columnLabels{idx});
            passes = TW_design >= estimatedTWvalue - tol;
            curveStatus.(fieldName) = passes;
            if ~passes
                constraintCurveFailures = constraintCurveFailures + 1;
                failedCurves{end+1} = columnLabels{idx}; %#ok<AGROW>
                curveDiagnostics(end+1,1) = sprintf('Constraint curve %s: T/W=%.1f below required %.1f at W/S=%.1f', columnLabels{idx}, round(TW_design,1), round(estimatedTWvalue,1), round(WS_design,1)); %#ok<AGROW>
            end
        end
    end

    WS_limit_landing = Consts(33, 12);
    landingPass = ~(WS_design > WS_limit_landing);
    curveStatus.Landing = landingPass;
    if ~landingPass
        constraintCurveFailures = constraintCurveFailures + 1;
        failedCurves{end+1} = 'Landing'; %#ok<AGROW>
        logText = logf(logText, 'Landing constraint violated: W/S = %.2f exceeds limit of %.2f\n', WS_design, WS_limit_landing);
    end
catch ME
    logText = logf(logText, 'Could not perform constraint curve check due to error: %s\n', ME.message);
    constraintCurveFailures = 0;
    failedCurves = {};
end

for idx = 1:numel(curveDiagnostics)
    logText = logf(logText, '%s\n', curveDiagnostics(idx));
end

if constraintCurveFailures == 1
    pt = pt - 4;
    logText = logf(logText, '-4 pts Design did not meet the following constraint: %s. Your design is not above those limits; increase T/W or relax the offending constraint values toward their thresholds.%s\n', char(failedCurves{1}), curveSuffixFew);
elseif constraintCurveFailures >= 2
    pt = pt - 8;
    summary = strjoin(string(failedCurves), ', ');
    suffix = curveSuffixFew;
    if constraintCurveFailures > 6
        suffix = curveSuffixMany;
    end
    logText = logf(logText, '-8 pts Design did not meet the following constraints: %s. Your design is not above those limits; increase T/W or relax the offending constraint values toward their thresholds.%s\n', char(summary), suffix);
end

objectiveFields = fieldnames(objectiveSet);
for idx = 1:numel(objectiveFields)
    key = objectiveFields{idx};
    if ~objectiveSet.(key)
        continue;
    end
    label = labelNames.(key);
    if isfield(curveStatus, key) && curveStatus.(key) && ~rowErrorsMap.(key)
        logText = logf(logText, 'Constraint %s set above threshold and satisfied.\n', label);
    elseif isfield(curveStatus, key) && ~curveStatus.(key)
        logText = logf(logText, 'Constraint %s set at or above objective. Design fails to meet this constraint; consider lowering it toward the threshold value.\n', label);
    end
end

% Payload (4 pts)
if isnan(aim120) || aim120 < 8 - tol
    count = aim120;
    if isnan(count), count = 0; end
    pt = pt - 4;
    logText = logf(logText, '-4 pts Payload must include at least 8 AIM-120Ds (found %.0f)\n', count);
elseif ~isnan(aim9) && aim9 >= 2 - tol
    logText = logf(logText, 'Payload meets objective: %.0f AIM-120s + %.0f AIM-9s\n', aim120, aim9);
end
% Stability (3 pts)
SM = Main(10, 13);
clb = Main(10, 15);
cnb = Main(10, 16);
rat = Main(10, 17);

stabilityErrors = 0;
if ~(SM >= -0.1 && SM <= 0.11)
    logText = logf(logText, 'Static margin out of bounds (M10 = %.3f)\n', SM);
    stabilityErrors = stabilityErrors + 1;
    if SM < 0
        logText = logf(logText, 'Warning: aircraft is statically unstable (SM < 0)\n');
    end
end
if clb >= -0.001
    logText = logf(logText, 'Clb must be < -0.001 (O10 = %.6f)\n', clb);
    stabilityErrors = stabilityErrors + 1;
end
if cnb <= 0.002
    logText = logf(logText, 'Cnb must be > 0.002 (P10 = %.6f)\n', cnb);
    stabilityErrors = stabilityErrors + 1;
end
if ~(abs(rat) >= 0.3 && abs(rat) <= 1)
    logText = logf(logText, 'Cnb/Clb ratio magnitude must be between 0.3 and 1.0 (Q10 = %.3f)\n', rat);
    stabilityErrors = stabilityErrors + 1;
end

stabilityDeduction = min(3, stabilityErrors);
if stabilityDeduction > 0
    pt = pt - stabilityDeduction;
    logText = logf(logText, '-%d pts Stability parameters outside limits\n', stabilityDeduction);
end

% Fuel (2 pts) and volume (2 pts)
fuelFail = false;
volumeFail = false;
if isnan(fuel_available) || isnan(fuel_required)
    pt = pt - 2;
    logText = logf(logText, '-2 pts Fuel check could not be evaluated because O18 or X40 is not numeric\n');
    fuelFail = true;
elseif fuel_available + fuelTol < fuel_required
    pt = pt - 2;
    logText = logf(logText, '-2 pts Fuel available (%.1f) is less than required (%.1f)\n', fuel_available, fuel_required);
    fuelFail = true;
end

if isnan(volume_remaining) || volume_remaining <= 0
    pt = pt - 2;
    logText = logf(logText, '-2 pts Volume remaining must be positive (Q23 = %.2f)\n', volume_remaining);
    volumeFail = true;
end

% Recurring cost (5 pts)
costDeduction = 0;
if abs(numaircraft - 187) < 1e-3
    if isnan(cost)
        costDeduction = 5;
        logText = logf(logText, '-5 pts Recurring cost missing for 187-aircraft estimate\n');
    elseif cost > 115 + tol
        costDeduction = 5;
        logText = logf(logText, '-5 pts Recurring cost exceeds threshold ($115M): $%.1fM\n', cost);
    elseif cost <= 100 + tol
        logText = logf(logText, 'Recurring cost meets objective (<= $100M): $%.1fM\n', cost);
    end
elseif abs(numaircraft - 800) < 1e-3
    if isnan(cost)
        costDeduction = 5;
        logText = logf(logText, '-5 pts Recurring cost missing for 800-aircraft estimate\n');
    elseif cost > 75 + tol
        costDeduction = 5;
        logText = logf(logText, '-5 pts Recurring cost exceeds threshold ($75M): $%.1fM\n', cost);
    elseif cost <= 61 + tol
        logText = logf(logText, 'Recurring cost meets objective (<= $61M): $%.1fM\n', cost);
    end
else
    costDeduction = 5;
    logText = logf(logText, '-5 pts Number of aircraft (N31) must be 187 or 800 (found %.0f)\n', numaircraft);
end

if costDeduction > 0
    pt = pt - costDeduction;
end

% Landing gear geometry (4 pts)
gearFailures = 0;
takeoffSpeedFail = false;

g90 = Gear(20, 10);
if isnan(g90) || g90 < 80 - tol || g90 > 90.5 + tol
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Violates main gear 90/10 rule share at J20: %.1f%% (must be between 80.0%% and 90.5%%)\n', g90);
end

tipbackActual = Gear(20, 12);
tipbackLimit = Gear(21, 12);
if isnan(tipbackActual) || isnan(tipbackLimit) || tipbackActual >= tipbackLimit - 1e-2
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Violates tipback angle requirement: upper %.2f%s must be less than lower %.2f%s\n', tipbackActual, char(176), tipbackLimit, char(176));
end

rolloverActual = Gear(20, 13);
rolloverLimit = Gear(21, 13);
if isnan(rolloverActual) || isnan(rolloverLimit) || rolloverActual >= rolloverLimit - 1e-2
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Violates rollover angle requirement: upper %.2f%s must be less than lower %.2f%s\n', rolloverActual, char(176), rolloverLimit, char(176));
end

rotationSpeed = Gear(20, 14);
takeoffSpeed = Gear(21, 14);
if isnan(rotationSpeed)
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Violates takeoff rotation speed: value missing (must be < 200 kts and less than N21)\n');
    takeoffSpeedFail = true;
else
    if rotationSpeed >= 200 - tol
        gearFailures = gearFailures + 1;
        logText = logf(logText, 'Violates takeoff rotation speed: %.1f kts (must be < 200 kts)\n', rotationSpeed);
        takeoffSpeedFail = true;
    end
    if ~isnan(takeoffSpeed)
        if rotationSpeed >= takeoffSpeed
            gearFailures = gearFailures + 1;
            logText = logf(logText, 'Takeoff speed margin failed: N20 must be less than N21 (N20 = %.2f, N21 = %.2f)\n', rotationSpeed, takeoffSpeed);
            takeoffSpeedFail = true;
        end
        if takeoffSpeed > 200 + tol
            gearFailures = gearFailures + 1;
            logText = logf(logText, 'Takeoff speed too high: N21 = %.1f kts (must be <= 200 kts)\n', takeoffSpeed);
            takeoffSpeedFail = true;
        end
    else
        gearFailures = gearFailures + 1;
        logText = logf(logText, 'Takeoff speed reference missing (N21); cannot verify N20 margin.\n');
        takeoffSpeedFail = true;
    end
end

if gearFailures > 0
    deduction = min(4, gearFailures);
    pt = pt - deduction;
    logText = logf(logText, '-%d pts Landing gear geometry outside limits\n', deduction);
end
% Final score and bonuses
baseScore = max(0, pt);
pt = baseScore;
bonusPoints = 0;
bonusEligible = constraintDeduction == 0 && constraintCurveFailures == 0;

if bonusEligible && ~isnan(radius)
    radiusBonusRaw = linearBonus(radius, 375, 410);
    radiusBonus = roundToTenth(radiusBonusRaw);
    if radiusBonus > 0
        bonusPoints = bonusPoints + radiusBonus;
        logText = logf(logText, 'Mission radius bonus [+%.1f bonus]: %.1f nm\n', radiusBonus, radius);
    end
end
if bonusEligible && ~isnan(aim120) && ~isnan(aim9)
    payloadBonusRaw = double(aim120 >= 8 - tol && aim9 >= 2 - tol);
    payloadBonus = roundToTenth(payloadBonusRaw);
    if payloadBonus > 0
        bonusPoints = bonusPoints + payloadBonus;
        logText = logf(logText, 'Payload bonus [+%.1f bonus]: %.0f AIM-120s + %.0f AIM-9s\n', payloadBonus, aim120, aim9);
    end
end
if bonusEligible && ~isnan(takeoff_dist)
    takeoffBonusRaw = linearBonusInv(takeoff_dist, 3000, 2500);
    takeoffBonus = roundToTenth(takeoffBonusRaw);
    if takeoffBonus > 0
        bonusPoints = bonusPoints + takeoffBonus;
        logText = logf(logText, 'Takeoff distance bonus [+%.1f bonus]: %.0f ft\n', takeoffBonus, takeoff_dist);
    end
end
if bonusEligible && ~isnan(landing_dist)
    landingBonusRaw = linearBonusInv(landing_dist, 5000, 3500);
    landingBonus = roundToTenth(landingBonusRaw);
    if landingBonus > 0
        bonusPoints = bonusPoints + landingBonus;
        logText = logf(logText, 'Landing distance bonus [+%.1f bonus]: %.0f ft\n', landingBonus, landing_dist);
    end
end
maxMachValue = Main(3,21);
if bonusEligible && ~isnan(maxMachValue)
    maxMachBonusRaw = linearBonus(maxMachValue, 2.0, 2.2);
    maxMachBonus = roundToTenth(maxMachBonusRaw);
    if maxMachBonus > 0
        bonusPoints = bonusPoints + maxMachBonus;
        logText = logf(logText, 'Max Mach bonus [+%.1f bonus]: Mach %.2f\n', maxMachBonus, maxMachValue);
    end
end
superValue = Main(4,21);
if bonusEligible && ~isnan(superValue)
    superBonusRaw = linearBonus(superValue, 1.5, 1.8);
    superBonus = roundToTenth(superBonusRaw);
    if superBonus > 0
        bonusPoints = bonusPoints + superBonus;
        logText = logf(logText, 'Supercruise Mach bonus [+%.1f bonus]: Mach %.2f\n', superBonus, superValue);
    end
end
psHighValue = Main(8,24);
if bonusEligible && ~isnan(psHighValue)
    psHighBonusRaw = linearBonus(psHighValue, 400, 500);
    psHighBonus = roundToTenth(psHighBonusRaw);
    if psHighBonus > 0
        bonusPoints = bonusPoints + psHighBonus;
        logText = logf(logText, 'Ps @30k ft bonus [+%.1f bonus]: %.0f ft/s\n', psHighBonus, psHighValue);
    end
end
psLowValue = Main(9,24);
if bonusEligible && ~isnan(psLowValue)
    psLowBonusRaw = linearBonus(psLowValue, 400, 500);
    psLowBonus = roundToTenth(psLowBonusRaw);
    if psLowBonus > 0
        bonusPoints = bonusPoints + psLowBonus;
        logText = logf(logText, 'Ps @10k ft bonus [+%.1f bonus]: %.0f ft/s\n', psLowBonus, psLowValue);
    end
end
gHighValue = Main(6,22);
if bonusEligible && ~isnan(gHighValue)
    gHighBonusRaw = linearBonus(gHighValue, 3.0, 4.0);
    gHighBonus = roundToTenth(gHighBonusRaw);
    if gHighBonus > 0
        bonusPoints = bonusPoints + gHighBonus;
        logText = logf(logText, 'Combat turn (30k ft) bonus [+%.1f bonus]: %.2f g\n', gHighBonus, gHighValue);
    end
end
gLowValue = Main(7,22);
if bonusEligible && ~isnan(gLowValue)
    gLowBonusRaw = linearBonus(gLowValue, 4.0, 4.5);
    gLowBonus = roundToTenth(gLowBonusRaw);
    if gLowBonus > 0
        bonusPoints = bonusPoints + gLowBonus;
        logText = logf(logText, 'Combat turn (10k ft) bonus [+%.1f bonus]: %.2f g\n', gLowBonus, gLowValue);
    end
end
if bonusEligible && ~isnan(cost)
    if abs(numaircraft - 187) < 1e-3
        costBonusRaw = linearBonusInv(cost, 115, 100);
        costBonus = roundToTenth(costBonusRaw);
        if costBonus > 0
            bonusPoints = bonusPoints + costBonus;
            logText = logf(logText, 'Recurring cost bonus [+%.1f bonus]: %.0f aircraft, $%.1fM\n', costBonus, numaircraft, cost);
        end
    elseif abs(numaircraft - 800) < 1e-3
        costBonusRaw = linearBonusInv(cost, 75, 61);
        costBonus = roundToTenth(costBonusRaw);
        if costBonus > 0
            bonusPoints = bonusPoints + costBonus;
            logText = logf(logText, 'Recurring cost bonus [+%.1f bonus]: %.0f aircraft, $%.1fM\n', costBonus, numaircraft, cost);
        end
    end
end

if ~bonusEligible
    logText = logf(logText, 'Bonus points unavailable because one or more constraints miss threshold values or the design is below a constraint curve.\n');
end

bonusPoints = roundToTenth(bonusPoints);
pt = roundToTenth(baseScore + bonusPoints);

nonViableReasons = strings(0,1);
if fuelFail
    nonViableReasons(end+1,1) = "fuel check failed";
end
if volumeFail
    nonViableReasons(end+1,1) = "insufficient volume remaining";
end
if takeoffSpeedFail
    nonViableReasons(end+1,1) = "takeoff speed check failed";
end
if ~isempty(nonViableReasons)
    pt = roundToTenth(min(pt, NON_VIABLE_CAP));
    logText = logf(logText, 'Final score capped at %.1f out of 45 because the aircraft is non-viable (%s).\n', NON_VIABLE_CAP, strjoin(nonViableReasons, ', '));
end

logText = logf(logText, 'Jet11 base score: %d out of 45\n', baseScore);
logText = logf(logText, 'Bonus points: +%.1f (final score %.1f)\n', bonusPoints, pt);

% Clean up formatting: trim lines, drop empties, join with single newlines
logText = strjoin(string(logText), ""); % ensure single string
cleanLines = splitlines(logText);
cleanLines = strtrim(cleanLines);
cleanLines = cleanLines(~strcmp(cleanLines, ""));
fbStr = strjoin(cleanLines, newline);
fb = char(fbStr);
end
%% Function to read all useful sheets, and verify numbers returned for used cells
function sheets = loadAllJet11Sheets(filename)
%LOADALLJET11SHEETS Load all required Jet11 sheets using safeReadMatrix
%   sheets = loadAllJet11Sheets(filename) returns a struct with fields:
%   Aero, Miss, Main, Consts, Gear, Geom

sheets.Aero   = safeReadMatrix(filename, 'Aero',   {'G3','G4','G10','G11','A15','A16'});
sheets.Miss   = safeReadMatrix(filename, 'Miss',   {'C48','C49'});
sheets.Main   = safeReadMatrix(filename, 'Main',   {'S3','T3','U3','V3','W3','X3','Y3','S4','T4','U4','V4','W4','X4','Y4',...
    'S5','S6','S7','S8','S9','T6','U6','V6','W6','X6','Y6','T7','U7','V7',...
    'W7','X7','Y7','T8','U8','V8','W8','X8','Y8','T9','U9','V9','W9','X9',...
    'Y9','S12','S13','AB3','AB4','X12','X13','Y37','M10','O10','P10','Q10',...
    'O18','X40','Q23','Q31','N31','P13','Q13','B32','B19','C19','D19','H19','B21','C21',...
    'D21','H21','B23','C23','D23','H23','D24','H24','C26','D26',...
    'B27','C27','H27','F31','F32','H29','I29','O15','B34','B35','B36','B37',...
    'B38','B39','B40','B41','B42','B43','B44','B45','B46','B47','B48','B49',...
    'B50','B51','B52','B53','E34','E35','E36','E37','E38','E39','E40','E41',...
    'E42','E43','E44','E45','E46','E47','E48','E49','E50','E51','E52','E53',...
    'D18','D23','D52','F52'});
sheets.Consts = safeReadMatrix(filename, 'Consts', {'K22','K23','K24','K26','K27','K28','K29','K32','AO42','AQ41','K33'});
sheets.Gear   = safeReadMatrix(filename, 'Gear',   {'J20','L20','L21','M20','M21','N20','N21'});
sheets.Geom   = safeReadMatrix(filename, 'Geom',   {'C8','C10','M152','K15','L155','L38','L40','L41','N44','L117','L118','N121'});

% Constants is off by three rows. Row 22 of the Consts tab comes in as
% row 19 in matlab Consts variable. Adding three rows of NaN to the top
% so it can be addressed accurately.

% sheets.Consts = [NaN(3, size(sheets.Consts, 2)); sheets.Consts];

end

%% Function to read the data from the excel sheets as quickly and accurately as possible
function data = safeReadMatrix(filename, sheetname, fallbackCells)
% safeReadMatrix - Efficiently reads numeric data from an Excel sheet.
%   Attempts fast readmatrix first. If key cells are NaN, falls back to readcell.
%
% Inputs:
%   filename      - Excel file path
%   sheetname     - Sheet name to read
%   fallbackCells - Cell array of cell references to verify (e.g., {'G4', 'G10'})
%
% Output:
%   data - Numeric matrix with fallback values patched in if needed

% Try fast read without invoking Excel UI
data = readmatrix(filename, 'Sheet', sheetname,'DataRange','A1:AQ250','UseExcel', false);


% Convert cell references to row/col indices
fallbackIndices = cellfun(@(c) cell2sub(c), fallbackCells, 'UniformOutput', false);

% Check for NaNs in fallback cells
needsPatch = false;
for i = 1:numel(fallbackIndices)
    idx = fallbackIndices{i};
    if idx(1) > size(data,1) || idx(2) > size(data,2) || isnan(data(idx(1), idx(2)))
        needsPatch = true;
        %                 fprintf('Patching %s %d\n',sheetname, idx)
        fprintf('Patched %s cell %s with value %.4f\n', sheetname, sub2excel(idx(1), idx(2)), data(idx(1), idx(2)));
        break;
    end
end

% If needed, patch from readcell
if needsPatch
    raw = readcell(filename, 'Sheet', sheetname, 'UseExcel', false);
    for i = 1:numel(fallbackIndices)
        idx = fallbackIndices{i};
        if idx(1) <= size(raw,1) && idx(2) <= size(raw,2)
            val = raw{idx(1), idx(2)};
            if isnumeric(val)
                data(idx(1), idx(2)) = val;
            elseif ischar(val) || isstring(val)
                data(idx(1), idx(2)) = str2double(val);
            end
        end
    end
end
end

function idx = cell2sub(cellref)
% Converts Excel cell reference (e.g., 'G4') to row/col indices
col = regexp(cellref, '[A-Z]+', 'match', 'once');
row = str2double(regexp(cellref, '\d+', 'match', 'once'));
colNum = 0;
for i = 1:length(col)
    colNum = colNum * 26 + (double(col(i)) - double('A') + 1);
end
idx = [row, colNum];
end

function angle = computeEdgeAngleDeg(Geom, rowA, rowB)
p1 = geomPlanformPoint(Geom, rowA);
p2 = geomPlanformPoint(Geom, rowB);
if any(isnan([p1, p2]))
    angle = NaN;
    return;
end
dx = abs(p2(1) - p1(1));
dy = abs(p2(2) - p1(2));
if dx == 0 && dy == 0
    angle = 0;
else
    angle = atan2d(dy, dx);
end
end

function angle = computeWingTrailingPlanformAngleDeg(Geom)
angle = computeEdgeAngleDeg(Geom, 40, 41);
end

function angle = computePcsTrailingPlanformAngleDeg(Geom)
angle = computeEdgeAngleDeg(Geom, 117, 118);
end

function point = geomPlanformPoint(Geom, row)
x = Geom(row, 12);
yCandidates = [Geom(row, 13), Geom(row, 14)];
yCandidates = yCandidates(~isnan(yCandidates));
if isempty(yCandidates)
    y = 0;
else
    y = max(abs(yCandidates));
end
point = [x, y];
end

function [logText, failures] = checkWingDevicePlacement(logText, failures, deviceName, edgeName, relevantA, relevantB, oppositeA, oppositeB, wingLeadingRoot, wingLeadingTip, wingTrailingRoot, wingTrailingTip, tol)
points = [relevantA, relevantB, oppositeA, oppositeB, wingLeadingRoot, wingLeadingTip, wingTrailingRoot, wingTrailingTip];
if any(isnan(points))
    logText = logf(logText, 'Unable to verify %s wing placement due to missing geometry data\n', deviceName);
    failures = failures + 1;
    return;
end

[relevantInboard, relevantOutboard, oppositeInboard, oppositeOutboard] = sortEdgePairsByY(relevantA, relevantB, oppositeA, oppositeB);
spanFail = false;
edgeFail = false;
envelopeFail = false;
wingSpan = max([wingLeadingRoot(2), wingLeadingTip(2), wingTrailingRoot(2), wingTrailingTip(2)]);

for idx = 1:2
    if idx == 1
        relevantPoint = relevantInboard;
        oppositePoint = oppositeInboard;
    else
        relevantPoint = relevantOutboard;
        oppositePoint = oppositeOutboard;
    end

    y = relevantPoint(2);
    if y < -tol || y > wingSpan + tol
        spanFail = true;
        continue;
    end

    [wingLeadingX, leadingInRange] = interpolateEdgeXAtY(wingLeadingRoot, wingLeadingTip, y);
    [wingTrailingX, trailingInRange] = interpolateEdgeXAtY(wingTrailingRoot, wingTrailingTip, y);
    if ~leadingInRange || ~trailingInRange
        spanFail = true;
        continue;
    end

    if strcmp(edgeName, 'leading edge')
        targetX = wingLeadingX;
    else
        targetX = wingTrailingX;
    end

    if abs(relevantPoint(1) - targetX) > tol
        edgeFail = true;
    end

    lower = min(wingLeadingX, wingTrailingX) - tol;
    upper = max(wingLeadingX, wingTrailingX) + tol;
    if oppositePoint(1) < lower || oppositePoint(1) > upper
        envelopeFail = true;
    end
end

if spanFail
    logText = logf(logText, '%s extends outside the wing span in top view.\n', deviceName);
end
if edgeFail
    logText = logf(logText, '%s %s must align with the wing %s within %.2f ft in top view.\n', deviceName, edgeName, edgeName, tol);
end
if envelopeFail
    logText = logf(logText, '%s must remain within the wing planform envelope in top view.\n', deviceName);
end
if spanFail || edgeFail || envelopeFail
    failures = failures + 1;
end
end

function [relevantInboard, relevantOutboard, oppositeInboard, oppositeOutboard] = sortEdgePairsByY(relevantA, relevantB, oppositeA, oppositeB)
if relevantA(2) <= relevantB(2)
    relevantInboard = relevantA;
    relevantOutboard = relevantB;
    oppositeInboard = oppositeA;
    oppositeOutboard = oppositeB;
else
    relevantInboard = relevantB;
    relevantOutboard = relevantA;
    oppositeInboard = oppositeB;
    oppositeOutboard = oppositeA;
end
end

function [x, inRange] = interpolateEdgeXAtY(pointA, pointB, y)
if any(isnan([pointA, pointB])) || isnan(y)
    x = NaN;
    inRange = false;
    return;
end
y1 = pointA(2);
y2 = pointB(2);
lower = min(y1, y2);
upper = max(y1, y2);
if y < lower - 1e-6 || y > upper + 1e-6
    x = NaN;
    inRange = false;
    return;
end
if abs(y2 - y1) < 1e-9
    x = pointA(1);
    inRange = abs(y - y1) <= 1e-6;
    return;
end
t = (y - y1) / (y2 - y1);
x = pointA(1) + t * (pointB(1) - pointA(1));
inRange = true;
end

function username = extractBlackboardUsername(fname)
% Blackboard export filenames place the username in the token immediately
% before "_attempt". Keep this tolerant of dots, dashes, and underscores.
userTok = regexp(fname, '_([^\s]+?)_attempt(?:_|\.|$)', 'tokens', 'once');
if ~isempty(userTok)
    username = userTok{1};
else
    username = 'UNKNOWN';
end
end

function hit = teNormalHitsCenterline(tipPoint, innerPoint, fuselageLength)
if any(isnan([tipPoint, innerPoint]))
    hit = false;
    return;
end
dir = innerPoint - tipPoint;
normals = [dir(2), -dir(1); -dir(2), dir(1)];
hit = false;
for k = 1:2
    normal = normals(k, :);
    if abs(normal(2)) < 1e-6
        continue;
    end
    t = -tipPoint(2) / normal(2);
    if t <= 0
        continue;
    end
    x = tipPoint(1) + normal(1) * t;
    if ~isnan(fuselageLength) && x >= -1e-6 && x <= fuselageLength + 1e-6
        hit = true;
        break;
    end
end
end

function tf = anglesParallel(angle, wingAngle, tol)
if isnan(angle) || isnan(wingAngle)
    tf = false;
    return;
end
a = mod(angle, 180);
b = mod(wingAngle, 180);
diffVal = abs(a - b);
alt = 180 - diffVal;
tf = min(diffVal, alt) <= tol;
end

function [logText, failures, headerLogged] = requireParallelAngle(logText, failures, headerLogged, angle, wingAngle, tol, template)
if isnan(angle) || isnan(wingAngle)
    [logText, headerLogged] = logStealth(logText, headerLogged, 'Unable to verify stealth shaping due to missing geometry data\n');
    failures = failures + 1;
elseif ~anglesParallel(angle, wingAngle, tol)
    [logText, headerLogged] = logStealth(logText, headerLogged, template, angle, wingAngle, tol);
    failures = failures + 1;
end
end

function [logText, failures, headerLogged] = requireParallelAngleOrCenterlineIfWithinFuselageHeight(logText, failures, headerLogged, angle, wingAngle, tol, template, tipPoint, innerPoint, withinFuselageHeight, fuselageLength)
if isnan(angle) || isnan(wingAngle)
    [logText, headerLogged] = logStealth(logText, headerLogged, 'Unable to verify stealth shaping due to missing geometry data\n');
    failures = failures + 1;
elseif ~(anglesParallel(angle, wingAngle, tol) || (withinFuselageHeight && teNormalHitsCenterline(tipPoint, innerPoint, fuselageLength)))
    [logText, headerLogged] = logStealth(logText, headerLogged, template, angle, wingAngle, tol);
    failures = failures + 1;
end
end

function tf = isWithinFuselageHeight(componentZ, fuselageCenterZ, fuselageHeight)
if any(isnan([componentZ, fuselageCenterZ, fuselageHeight]))
    tf = false;
    return;
end
tf = componentZ >= (fuselageCenterZ - fuselageHeight/2) && componentZ <= (fuselageCenterZ + fuselageHeight/2);
end

function tf = isSurfaceWithinFuselageHeight(componentZ, dihedralAngle, tipPoint, innerPoint, fuselageCenterZ, fuselageHeight)
if any(isnan([componentZ, dihedralAngle, tipPoint, innerPoint, fuselageCenterZ, fuselageHeight]))
    tf = false;
    return;
end
spanOffset = abs(tipPoint(2) - innerPoint(2));
tipZ = componentZ + spanOffset * tand(dihedralAngle);
lower = fuselageCenterZ - fuselageHeight/2;
upper = fuselageCenterZ + fuselageHeight/2;
tf = componentZ >= lower && componentZ <= upper && tipZ >= lower && tipZ <= upper;
end

function [logText, headerLogged] = logStealth(logText, headerLogged, fmt, varargin)
if ~headerLogged
    logText = logf(logText, 'Stealth shaping violations:\n');
    headerLogged = true;
end
logText = logf(logText, fmt, varargin{:});
end

function field = mapCurveField(label)
switch label
    case 'MaxMach'
        field = 'MaxMach';
    case 'Supercruise'
        field = 'CruiseMach';
    case 'CombatTurn1'
        field = 'CmbtTurn1';
    case 'CombatTurn2'
        field = 'CmbtTurn2';
    case 'Ps1'
        field = 'Ps1';
    case 'Ps2'
        field = 'Ps2';
    case 'Takeoff'
        field = 'Takeoff';
    otherwise
        field = matlab.lang.makeValidName(label);
end
end

function ref = sub2excel(row, col)
letters = '';
while col > 0
    rem = mod(col - 1, 26);
    letters = [char(65 + rem), letters]; %#ok<AGROW>
    col = floor((col - 1) / 26);
end
ref = sprintf('%s%d', letters, row);
end

%% Function to do an fprintf like function to a local variable for future use
function logText = logf(logText, varargin)
logEntry = sprintf(varargin{:});  % Format input like fprintf
logText = [logText, logEntry];      % Append to string
end


function [mode, selectedPath] = selectRunMode()
% SELECTRUNMODE - Launches a GUI to choose between single file or folder mode



cursorPos = get(0, 'PointerLocation');
dialogWidth = 300;
dialogHeight = 150;

% Position just below the cursor
dialogLeft = cursorPos(1) - dialogWidth / 2;
dialogBottom = cursorPos(2) - dialogHeight - 20;  % 20 pixels below the cursor

d = dialog('Position', [dialogLeft, dialogBottom, dialogWidth, dialogHeight], ...
    'Name', 'Select Run Mode');


txt = uicontrol('Parent',d,...
    'Style','text',...
    'Position',[20 90 260 40],...
    'String','Choose how you want to run the autograder:',...
    'FontSize',10); %#ok<NASGU>

btn1 = uicontrol('Parent',d,...
    'Position',[30 40 100 30],...
    'String','Single File',...
    'Callback',@singleFile); %#ok<NASGU>

btn2 = uicontrol('Parent',d,...
    'Position',[170 40 100 30],...
    'String','Folder of Files',...
    'Callback',@folderRun); %#ok<NASGU>

mode = '';
selectedPath = '';

uiwait(d);  % Wait for user to close dialog

    function singleFile(~,~)
        [file, path] = uigetfile('*.xls*','Select a Jet11 Excel file');
        if isequal(file,0)
            mode = 'cancelled';
        else
            mode = 'single';
            selectedPath = fullfile(path, file);
        end
        delete(d);
    end

    function folderRun(~,~)
        path = uigetdir(pwd, 'Select folder containing Jet11 files');
        if isequal(path,0)
            mode = 'cancelled';
        else
            mode = 'folder';
            selectedPath = path;
        end
        delete(d);
    end
end

%% Prompt user and generate Blackboard CSV (combined function)
function promptAndGenerateBlackboardCSV(folderAnalyzed, files, points, feedback, timestamp)
% Position dialog below cursor
cursorPos = get(0, 'PointerLocation');
dialogWidth = 300;
dialogHeight = 150;
dialogLeft = cursorPos(1) - dialogWidth / 2;
dialogBottom = cursorPos(2) - dialogHeight - 20;

% Create dialog
d = dialog('Position', [dialogLeft, dialogBottom, dialogWidth, dialogHeight], ...
    'Name', 'Blackboard Export');

uicontrol('Parent', d, ...
    'Style', 'text', ...
    'Position', [20 90 260 40], ...
    'String', 'Generate Blackboard CSV for grade import?', ...
    'FontSize', 10);

uicontrol('Parent', d, ...
    'Position', [30 40 100 30], ...
    'String', 'Yes', ...
    'Callback', @(~,~) doExport(true, d));

uicontrol('Parent', d, ...
    'Position', [170 40 100 30], ...
    'String', 'No', ...
    'Callback', @(~,~) doExport(false, d));

    function doExport(shouldExport, dialogHandle)
        delete(dialogHandle);
        if shouldExport
            %% Create Blackboard Offline Grade CSV (SMART_TEXT format)
            csvFilename = fullfile(folderAnalyzed, ['FinalAircraft_Blackboard_Offline_', timestamp, '.csv']);
            fid = fopen(csvFilename, 'w');

            % Assignment title column (update if needed)
            assignmentTitle = 'Final AATF Jet10 Design [Total Pts: 45 Score] |409581';

            % Write header (username-only export)
            % Column order must match Blackboard export: Username | Assignment score | Grading Notes | Notes Format | Feedback to Learner | Feedback Format
            fprintf(fid, '"Username","%s","Grading Notes","Notes Format","Feedback to Learner","Feedback Format"\n', assignmentTitle);

            for i = 1:numel(files)
                fname = files(i).name;

                username = extractBlackboardUsername(fname);

                % Get score and feedback
                score = max(0, min(100, roundToTenth(points(i))));
                fbText = char(strtrim(string(feedback{i})));

                % Sanitize feedback for SMART_TEXT (HTML-safe but readable)
                fbText = strrep(fbText, 'â‰¥', '&ge;');
                fbText = strrep(fbText, 'â‰¤', '&le;');
                fbText = strrep(fbText, 'â‰ ', '&ne;');
                fbText = strrep(fbText, 'âœ”', '&#10004;');
                fbText = strrep(fbText, 'âœ˜', '&#10008;');
                fbText = strrep(fbText, 'âœ…', '&#9989;');
                fbText = strrep(fbText, 'âŒ', '&#10060;');
                fbText = strrep(fbText, '<', '&lt;');
                fbText = strrep(fbText, '>', '&gt;');
                fbText = strrep(fbText, '"', '&quot;');
                fbText = strrep(fbText, newline, '<br>');

                % Write row: Username | Score | Grading Notes | Notes Format | Feedback | Feedback Format
                fprintf(fid, '"%s","%.1f","","","%s","%s"\n', ...
                    username, score, fbText, 'SMART_TEXT');
            end

            fclose(fid);
            fprintf('Blackboard offline grade CSV created: %s\n', csvFilename);

        end
    end
end

function y = clamp01(x)
y = max(0, min(1, x));
end

function bonus = linearBonus(value, threshold, objective)
if isnan(value)
    bonus = 0;
    return;
end
if abs(objective - threshold) < eps
    bonus = double(value >= objective);
    return;
end
bonus = clamp01((value - threshold) / (objective - threshold));
end

function bonus = linearBonusInv(value, threshold, objective)
if isnan(value)
    bonus = 0;
    return;
end
if abs(objective - threshold) < eps
    bonus = double(value <= objective);
    return;
end
bonus = clamp01((threshold - value) / (threshold - objective));
end

function rounded = roundToTenth(value)
if isnan(value)
    rounded = NaN;
else
    rounded = round(value*10)/10;
end
end
