
% Final Exam Autograder: Grades AE210 Jet11 Excel submissions, logs feedback, and optionally exports Blackboard-compatible scores.


%--------------------------------------------------------------------------
% AE210 Final Exam Autograder Script â€“ Fall 2025
%
% Description:
% This script automates grading for the AE210 Final Exam by processing Jet11 Excel files (*.xlsm). It evaluates 
% multiple design criteria, generates detailed feedback, and outputs both a
% summary log and an optional Blackboard-compatible grade import file.
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
% - Optional Blackboard CSV: FinalProject_Blackboard_Offline_<timestamp>.csv
%
% Scoring:
% - 85 pts for meeting threshold requirements (range 500 nm radius, cost <$120M @187, 8x AIM-120D, constraint table, Tavailable > Drag, control surface attachment, stability, positive volume, landing gear)
% - 15 pts for objectives (+5 radius >= 800 nm, +5 cost <$110M @187, +5 payload adds 2x AIM-9)
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
% Last Updated: 09 Dec 2025 19:49
%--------------------------------------------------------------------------
clear; close all; clc;


%% Choose directory and get Excel files
% fprintf('Executing %s\n',mfilename);
% I recommend updating the below line to point to your Final Exam files. It works
% as is, but will default to the right place if this is updated.

% folderAnalyzed = uigetdir('C:\Users\dell.olmstead\OneDrive - afacademy.af.edu\Documents 1\01 Classes\AE210 FA24\Design Project\Final Exam files');
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
points = zeros(numel(files),1);  % Initialize points for each file
feedback = cell(1,numel(files));

fprintf('Reading %d files\n', numel(files));

if strcmp(mode, 'folder')
    % Combined parallel read + grade

    parfor cadetIdx = 1:numel(files)
        filename = fullfile(folderAnalyzed, files(cadetIdx).name);
        fprintf('Grading %s\n', files(cadetIdx).name);
        try
            [pt, fb] = gradeCadet(filename);
            points(cadetIdx) = pt;
            feedback{cadetIdx} = fb;
            fprintf('Finished %s\n', files(cadetIdx).name);
        catch ME
            if strcmp(ME.identifier, 'FinalExam:PasswordProtected')
                points(cadetIdx) = 0;
                feedback{cadetIdx} = sprintf('Skipped password-protected file: %s', files(cadetIdx).name);
                continue;
            end
            points(cadetIdx) = 0;
            feedback{cadetIdx} = sprintf('Error reading or grading file: %s. Details: %s', files(cadetIdx).name, ME.message);
        end
    end

else %       %%% Use the below code to run a single cadet

    filename = fullfile(folderAnalyzed, files(1).name);
    fprintf('Grading %s\n', files(1).name);
    try
        [points, feedback{1}] = gradeCadet(filename);
    catch ME
        points = 0;
        feedback{1} = sprintf('Error reading or grading file: %s. Details: %s', files(1).name, ME.message);
    end

end


%% Set up log file
timestamp = char(datetime('now', 'Format', 'yyyy-MM-dd_HH-mm-ss'));
logFilePath = fullfile(folderAnalyzed, ['textout_', timestamp, '.txt']);
finalout = fopen(logFilePath,'w');

% Log file header
fprintf(finalout, 'Final Exam Autograder Log\n');
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

[~, name, ext] = fileparts(filename);
sheets = loadAllJet11Sheets(filename);

Aero = sheets.Aero;
Miss = sheets.Miss;
Main = sheets.Main;
Consts = sheets.Consts;
Gear = sheets.Gear;
Geom = sheets.Geom;
fprintf('Read Complete %s\n', [name, ext]);


BASE_TOTAL = 85;
OBJECTIVE_TOTAL = 15;
pt = 0;
logText = "";
thresholdScore = 0;
objectiveScore = 0;
% Pre-initialize key inputs to avoid undefined-variable errors if read fails
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

logText = logf(logText, '%s\n', [name, ext]);

tol = 1e-3;
wtoTol = 1e-2; % tolerance specific to W/WTO checks
altTol = 1;
machTol = 1e-2;
timeTol = 1e-2;
distTol = 1e-3;
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
rangePass = false;
rangeObjectivePass = false;
costPass = false;
costObjectivePass = false;
payloadPass = false;
payloadObjectivePass = false;
controlPass = false;
thrustPass = false;
stabilityPass = false;
volumePass = false;
gearPass = false;
stealthPass = false;
missionPass = false;
efficiencyPass = false;
fuelPass = false;
dataValidPass = true;
rangeFailDetail = "";
fuelFailDetail = "";
volumeFailDetail = "";
constraintFailDetails = {};

% Aero tab programming (informational)
aeroIssues = 0;
if isequal(Aero(3,7), Aero(4,7)), aeroIssues = aeroIssues + 1; end
if isequal(Aero(10,7), Aero(11,7)), aeroIssues = aeroIssues + 1; end
if isequal(Aero(15,1), Aero(16,1)), aeroIssues = aeroIssues + 1; end

if aeroIssues > 0
    logText = logf(logText, 'Aero tab formulas inactive in %d key cell(s); check A15, G3, and G10.\n', aeroIssues);
end

% Sheet validation (geometry inputs) before further checks
geomBlock = Main(18:27, 2:8);
geomNaN = isnan(geomBlock);
% Skip known blank cells that are intentionally empty
skipCells = [...
    7 1;  % B24
    7 2;  % C24
    10 3; % D27
    10 4; % E27
    10 5; % F27
    10 6; % G27
    9 7]; % H26
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
% Mission table validation (exact matches)
MissionArray = Main(33:44, 11:25);
colIdx = 1:14; % Takeoff through Landing
alt = MissionArray(1, colIdx);
mach = MissionArray(3, colIdx);
ab = MissionArray(4, colIdx);
dist = MissionArray(6, colIdx);
timeLeg = MissionArray(7, colIdx);

% Expected values from provided mission profile
altExpected = [0, 2000, 35000, 35000, 35000, 35000, 35000, 30000, 35000, 35000, 35000, 35000, 10000, 0];
machExpected = [0.268473504, 0.88, 0.88, 0.88, 0.88, 1.5, 0.8, 0.8, 1.5, 0.8, 0.88, 0.88, 0.4, 0.0];
abExpected = [100, 0, 0, 0, 0, 0, 0, 100, 0, 0, 0, 0, 0, 0];
supercruiseCols = [6, 9];
distExpected = zeros(1, numel(colIdx));
distExpected(supercruiseCols) = [400, 400];
combatCol = 8;
loiterCol = 13;
timeExpected = nan(1, numel(colIdx));
timeExpected(combatCol) = 2;
timeExpected(loiterCol) = 20;

missionErrors = 0;
for i = 1:numel(colIdx)
    if abs(alt(i) - altExpected(i)) > altTol
        logText = logf(logText, 'Leg %d Altitude must be %.0f (found %.0f)\n', i, altExpected(i), alt(i));
        missionErrors = missionErrors + 1;
    end
    % Leg 1 Mach is advisory only; skip rigid check
    if i ~= 1 && abs(mach(i) - machExpected(i)) > machTol
        missionErrors = missionErrors + 1;
    end
    % Leg 14 AB cell has other content; skip AB check there
    if i ~= 14 && abs(ab(i) - abExpected(i)) > tol
        missionErrors = missionErrors + 1;
    end
    if ismember(i, supercruiseCols)
        if abs(dist(i) - distExpected(i)) > distTol
            logText = logf(logText, 'Leg %d Supercruise distance must be %.0f (found %.2f)\n', i, distExpected(i), dist(i));
            missionErrors = missionErrors + 1;
        end
    end
    if i == combatCol || i == loiterCol
        if abs(timeLeg(i) - timeExpected(i)) > timeTol
            logText = logf(logText, 'Leg %d Time must be %.2f min (found %.2f)\n', i, timeExpected(i), timeLeg(i));
            missionErrors = missionErrors + 1;
        end
    end
end

if ~isnan(radius)
    if radius >= 800 - distTol
        rangePass = true;
        rangeObjectivePass = true;
    elseif radius >= 500 - distTol
        rangePass = true;
    else
        logText = logf(logText, 'Range below threshold: mission radius = %.1f nm (needs >= 500 nm)\n', radius);
        missionErrors = missionErrors + 1;
    end
else
    logText = logf(logText, 'Mission radius missing; unable to verify range requirement.\n');
    missionErrors = missionErrors + 1;
end
% Keep mission error logging for awareness; scoring handled later.
if missionErrors == 0
    missionPass = true;
end

% Efficiency guardrails (must not be altered)
o1 = Main(1, 15);
q1 = Main(1, 17);
c30 = Main(30, 3);
d30 = Main(30, 4);
efficiencyErrors = 0;
if isnan(o1) || abs(o1 - 0.0037) > tol
    efficiencyErrors = efficiencyErrors + 1;
    logText = logf(logText, 'O1 must be 0.0037 (found %.4f)\n', o1);
end
if isnan(q1) || abs(q1 - 2.2) > tol
    efficiencyErrors = efficiencyErrors + 1;
    logText = logf(logText, 'Q1 must be 2.2 (found %.4f)\n', q1);
end
if isnan(c30) || abs(c30 - 0.8) > tol
    efficiencyErrors = efficiencyErrors + 1;
    logText = logf(logText, 'C30 must be 0.8 (found %.4f)\n', c30);
end
if isnan(d30) || abs(d30 - 2.0) > tol
    efficiencyErrors = efficiencyErrors + 1;
    logText = logf(logText, 'D30 must be 2.0 (found %.4f)\n', d30);
end
efficiencyPass = efficiencyErrors == 0;

% Thrust available vs drag check
thrust_drag = Miss(48:49, 3:14);
thrustShort = thrust_drag(2, :) <= thrust_drag(1, :);
thrustFailures = sum(thrustShort);
thrustPass = thrustFailures == 0;
if ~thrustPass
    logText = logf(logText, 'Thrust shortfall: Tavailable <= Drag for %d mission segment(s).\n', thrustFailures);
end
% Control surface attachment
controlFailures = 0;
VALUE_TOL = 1e-3;
AR_TOL = 0.1;
VT_WING_FRACTION = 0.8;

fuselage_length = Main(32, 2);
fuselage_end = fuselage_length;
PCS_x = Main(23, 3);
PCS_root_chord = Geom(8, 3);
if any(isnan([fuselage_end, PCS_x, PCS_root_chord]))
    logText = logf(logText, 'Unable to verify PCS placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif PCS_x > (fuselage_end - 0.25 * PCS_root_chord)
    logText = logf(logText, 'PCS X-location too far aft. Must overlap at least 25%% of root chord.\n');
    controlFailures = controlFailures + 1;
end

VT_x = Main(23, 8);
VT_root_chord = Geom(10, 3);
if any(isnan([fuselage_end, VT_x, VT_root_chord]))
    logText = logf(logText, 'Unable to verify vertical tail placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif VT_x > (fuselage_end - 0.25 * VT_root_chord)
    logText = logf(logText, 'VT X-location too far aft. Must overlap at least 25%% of root chord.\n');
    controlFailures = controlFailures + 1;
end

PCS_z = Main(25, 3);
fuse_z_center = Main(52, 4);
fuse_z_height = Main(52, 6);
if any(isnan([PCS_z, fuse_z_center, fuse_z_height]))
    logText = logf(logText, 'Unable to verify PCS vertical placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif PCS_z < (fuse_z_center - fuse_z_height/2) || PCS_z > (fuse_z_center + fuse_z_height/2)
    logText = logf(logText, 'PCS Z-location outside fuselage vertical bounds.\n');
    controlFailures = controlFailures + 1;
end

VT_y = Main(24, 8);
fuse_width = Main(52, 5);
vtMountedOffFuselage = false;
if any(isnan([VT_y, fuse_width]))
    logText = logf(logText, 'Unable to verify vertical tail lateral placement due to missing geometry data\n');
    controlFailures = controlFailures + 1;
elseif abs(VT_y) > fuse_width/2 + VALUE_TOL
    vtMountedOffFuselage = true;
    logText = logf(logText, 'Vertical tail mounted off the fuselage; ensure structural support at the wing.\n');
end

if Main(18, 4) > 1
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

component_positions = Main(23, 2:8);
if any(component_positions >= fuselage_end)
    logText = logf(logText, 'One or more components X-location extend beyond the fuselage end (B32 = %.2f)\n', fuselage_end);
    controlFailures = controlFailures + 1;
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

controlPass = controlFailures == 0;
if controlFailures > 0
    logText = logf(logText, 'Control surface attachment has %d issue(s).\n', controlFailures);
end

% Stealth shaping 
STEALTH_TOL = 5;
stealthFailures = 0;

wingLeadingAngle = computeEdgeAngleDeg(Geom, 38, 39);
wingTrailingAngle = computeEdgeAngleDeg(Geom, 40, 41);
pcsLeadingAngle = computeEdgeAngleDeg(Geom, 115, 116);
pcsTrailingAngle = computeEdgeAngleDeg(Geom, 117, 118);
strakeLeadingAngle = computeEdgeAngleDeg(Geom, 152, 153);
strakeTrailingAngle = computeEdgeAngleDeg(Geom, 154, 155);
vtLeadingAngle = computeEdgeAngleDeg(Geom, 163, 164);
vtTrailingAngle = computeEdgeAngleDeg(Geom, 165, 166);
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
    logText = logf(logText, 'Pitch control surface leading edge sweep %.1f° must match the wing leading edge sweep %.1f° (+/- %.1f°).\n', pcsLeadingAngle, wingLeadingAngle, STEALTH_TOL);
    stealthFailures = stealthFailures + 1;
end

wingTipTE = geomPlanformPoint(Geom, 40);
wingCenterTE = geomPlanformPoint(Geom, 41);
if ~(wingActive && (anglesParallel(wingTrailingAngle, wingLeadingAngle, STEALTH_TOL) || teNormalHitsCenterline(wingTipTE, wingCenterTE)))
    logText = logf(logText, 'Wing trailing edge %.1f° is not parallel to the leading edge and its normal does not reach the fuselage centerline (+/- %.1f°).\n', wingTrailingAngle, STEALTH_TOL);
    stealthFailures = stealthFailures + 1;
end

if pcsActive && ~isnan(pcsDihedral) && pcsDihedral > 5
    [logText, stealthFailures] = requireParallelAngle(logText, stealthFailures, pcsLeadingAngle, wingLeadingAngle, STEALTH_TOL, 'Pitch control surface leading edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
    [logText, stealthFailures] = requireParallelAngle(logText, stealthFailures, pcsTrailingAngle, wingLeadingAngle, STEALTH_TOL, 'Pitch control surface trailing edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
end

if strakeActive
    [logText, stealthFailures] = requireParallelAngle(logText, stealthFailures, strakeLeadingAngle, wingLeadingAngle, STEALTH_TOL, 'Strake leading edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
    [logText, stealthFailures] = requireParallelAngle(logText, stealthFailures, strakeTrailingAngle, wingLeadingAngle, STEALTH_TOL, 'Strake trailing edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
end

if ~vtActive
    % ignore
elseif isnan(vtTilt)
    logText = logf(logText, 'Unable to verify stealth shaping due to missing geometry data\n');
    stealthFailures = stealthFailures + 1;
elseif vtTilt < 85
    [logText, stealthFailures] = requireParallelAngle(logText, stealthFailures, vtLeadingAngle, wingLeadingAngle, STEALTH_TOL, 'Vertical tail leading edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
    [logText, stealthFailures] = requireParallelAngle(logText, stealthFailures, vtTrailingAngle, wingLeadingAngle, STEALTH_TOL, 'Vertical tail trailing edge sweep %.1f° must be parallel to the wing leading edge %.1f° (+/- %.1f°).\n');
end

if stealthFailures > 0
    logText = logf(logText, 'Stealth shaping issues flagged in %d area(s).\n', stealthFailures);
else
    stealthPass = true;
end

% Constraint table values
constraintErrors = 0;
objectiveSet = struct('MaxMach', false, 'CruiseMach', false, 'CmbtTurn1', false, 'CmbtTurn2', false, 'Ps1', false, 'Ps2', false);
rowErrorsMap = objectiveSet;
curveStatus = struct();
labelNames = struct('MaxMach', 'MaxMach', 'CruiseMach', 'CruiseMach', 'CmbtTurn1', 'Cmbt Turn1', 'CmbtTurn2', 'Cmbt Turn2', 'Ps1', 'Ps1', 'Ps2', 'Ps2');

% MaxMach row: 35k ft, Mach 2.0, n=1, AB=100%, Ps=0, W/WTO=betaExpected
if isnan(Main(3,20)) || abs(Main(3,20) - 35000) > altTol
    logText = logf(logText, 'MaxMach: Altitude must be 35000 (found %.0f)\n', Main(3,20));
    constraintErrors = constraintErrors + 1; rowErrorsMap.MaxMach = true;
end
if isnan(Main(3,21)) || abs(Main(3,21) - 2.0) > machTol
    logText = logf(logText, 'MaxMach: Mach must be 2.0 (found %.2f)\n', Main(3,21));
    constraintErrors = constraintErrors + 1; rowErrorsMap.MaxMach = true;
end
if isnan(Main(3,22)) || abs(Main(3,22) - 1) > tol
    logText = logf(logText, 'MaxMach: n must be 1 (found %.3f)\n', Main(3,22));
    constraintErrors = constraintErrors + 1; rowErrorsMap.MaxMach = true;
end
if isnan(Main(3,23)) || abs(Main(3,23) - 100) > tol
    logText = logf(logText, 'MaxMach: AB must be 100%% (found %.0f%%)\n', Main(3,23));
    constraintErrors = constraintErrors + 1; rowErrorsMap.MaxMach = true;
end
if isnan(Main(3,24)) || abs(Main(3,24) - 0) > tol
    logText = logf(logText, 'MaxMach: Ps must be 0 (found %.0f)\n', Main(3,24));
    constraintErrors = constraintErrors + 1; rowErrorsMap.MaxMach = true;
end
if isnan(Main(3,25)) || abs(Main(3,25) - 0) > tol
    logText = logf(logText, 'MaxMach: CDx must be 0.000 (found %.3f)\n', Main(3,25));
    constraintErrors = constraintErrors + 1; rowErrorsMap.MaxMach = true;
end
if abs(Main(3,19) - betaExpected) > wtoTol
    logText = logf(logText, 'MaxMach: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(3,19));
    constraintErrors = constraintErrors + 1; rowErrorsMap.MaxMach = true;
end

% CruiseMach row: 35k ft, Mach 1.5, n=1, AB=0, Ps=0, W/WTO=betaExpected
if isnan(Main(4,20)) || abs(Main(4,20) - 35000) > altTol
    logText = logf(logText, 'CruiseMach: Altitude must be 35000 (found %.0f)\n', Main(4,20));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CruiseMach = true;
end
if isnan(Main(4,21)) || abs(Main(4,21) - 1.5) > machTol
    logText = logf(logText, 'CruiseMach: Mach must be 1.5 (found %.2f)\n', Main(4,21));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CruiseMach = true;
end
if isnan(Main(4,22)) || abs(Main(4,22) - 1) > tol
    logText = logf(logText, 'CruiseMach: n must be 1 (found %.3f)\n', Main(4,22));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CruiseMach = true;
end
if isnan(Main(4,23)) || abs(Main(4,23) - 0) > tol
    logText = logf(logText, 'CruiseMach: AB must be 0%% (found %.0f%%)\n', Main(4,23));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CruiseMach = true;
end
if isnan(Main(4,24)) || abs(Main(4,24) - 0) > tol
    logText = logf(logText, 'CruiseMach: Ps must be 0 (found %.0f)\n', Main(4,24));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CruiseMach = true;
end
if isnan(Main(4,25)) || abs(Main(4,25) - 0) > tol
    logText = logf(logText, 'CruiseMach: CDx must be 0.000 (found %.3f)\n', Main(4,25));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CruiseMach = true;
end
if abs(Main(4,19) - betaExpected) > wtoTol
    logText = logf(logText, 'CruiseMach: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(4,19));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CruiseMach = true;
end

% Supercruise row (50k ft, Mach 1.5, n=1, AB=100%, Ps=0, W/WTO=betaExpected)
if isnan(Main(5,20)) || abs(Main(5,20) - 50000) > altTol
    logText = logf(logText, 'Supercruise: Altitude must be 50000 (found %.0f)\n', Main(5,20));
    constraintErrors = constraintErrors + 1;
end
if isnan(Main(5,21)) || abs(Main(5,21) - 1.5) > machTol
    logText = logf(logText, 'Supercruise: Mach must be 1.5 (found %.2f)\n', Main(5,21));
    constraintErrors = constraintErrors + 1;
end
if isnan(Main(5,22)) || abs(Main(5,22) - 1) > tol
    logText = logf(logText, 'Supercruise: n must be 1 (found %.3f)\n', Main(5,22));
    constraintErrors = constraintErrors + 1;
end
if isnan(Main(5,23)) || abs(Main(5,23) - 100) > tol
    logText = logf(logText, 'Supercruise: AB must be 100%% (found %.0f%%)\n', Main(5,23));
    constraintErrors = constraintErrors + 1;
end
if isnan(Main(5,24)) || abs(Main(5,24) - 0) > tol
    logText = logf(logText, 'Supercruise: Ps must be 0 (found %.0f)\n', Main(5,24));
    constraintErrors = constraintErrors + 1;
end
if isnan(Main(5,25)) || abs(Main(5,25) - 0) > tol
    logText = logf(logText, 'Supercruise: CDx must be 0.000 (found %.3f)\n', Main(5,25));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(5,19) - betaExpected) > wtoTol
    logText = logf(logText, 'Supercruise: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(5,19));
    constraintErrors = constraintErrors + 1;
end

if abs(Main(6,20) - 30000) > altTol
    logText = logf(logText, 'Cmbt Turn1: Altitude must be 30000 (found %.0f)\n', Main(6,20));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn1 = true;
end
if abs(Main(6,21) - 1.2) > machTol
    logText = logf(logText, 'Cmbt Turn1: Mach must be 1.2 (found %.2f)\n', Main(6,21));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn1 = true;
end
if isnan(Main(6,22)) || abs(Main(6,22) - 3.0) > tol
    logText = logf(logText, 'Cmbt Turn1: g-load must be 3.0 (found %.3f)\n', Main(6,22));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn1 = true;
end
if abs(Main(6,23) - 100) > tol
    logText = logf(logText, 'Cmbt Turn1: AB must be 100%% (found %.0f%%)\n', Main(6,23));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn1 = true;
end
if abs(Main(6,24)) > tol
    logText = logf(logText, 'Cmbt Turn1: Ps must be 0 (found %.0f)\n', Main(6,24));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn1 = true;
end
if isnan(Main(6,25)) || abs(Main(6,25) - 0) > tol
    logText = logf(logText, 'Cmbt Turn1: CDx must be 0.000 (found %.3f)\n', Main(6,25));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn1 = true;
end
if abs(Main(6,19) - betaExpected) > wtoTol
    logText = logf(logText, 'Cmbt Turn1: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(6,19));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn1 = true;
end

if abs(Main(7,20) - 10000) > altTol
    logText = logf(logText, 'Cmbt Turn2: Altitude must be 10000 (found %.0f)\n', Main(7,20));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn2 = true;
end
if abs(Main(7,21) - 0.9) > machTol
    logText = logf(logText, 'Cmbt Turn2: Mach must be 0.9 (found %.2f)\n', Main(7,21));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn2 = true;
end
if isnan(Main(7,22)) || abs(Main(7,22) - 4.0) > tol
    logText = logf(logText, 'Cmbt Turn2: g-load must be 4.0 (found %.3f)\n', Main(7,22));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn2 = true;
end
if abs(Main(7,23) - 100) > tol
    logText = logf(logText, 'Cmbt Turn2: AB must be 100%% (found %.0f%%)\n', Main(7,23));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn2 = true;
end
if abs(Main(7,24)) > tol
    logText = logf(logText, 'Cmbt Turn2: Ps must be 0 (found %.0f)\n', Main(7,24));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn2 = true;
end
if isnan(Main(7,25)) || abs(Main(7,25) - 0) > tol
    logText = logf(logText, 'Cmbt Turn2: CDx must be 0.000 (found %.3f)\n', Main(7,25));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn2 = true;
end
if abs(Main(7,19) - betaExpected) > wtoTol
    logText = logf(logText, 'Cmbt Turn2: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(7,19));
    constraintErrors = constraintErrors + 1; rowErrorsMap.CmbtTurn2 = true;
end

if abs(Main(8,20) - 30000) > altTol
    logText = logf(logText, 'Ps1: Altitude must be 30000 (found %.0f)\n', Main(8,20));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps1 = true;
end
if abs(Main(8,21) - 1.15) > machTol
    logText = logf(logText, 'Ps1: Mach must be 1.15 (found %.2f)\n', Main(8,21));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps1 = true;
end
if isnan(Main(8,22)) || abs(Main(8,22) - 1) > tol
    logText = logf(logText, 'Ps1: n must be 1 (found %.3f)\n', Main(8,22));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps1 = true;
end
if abs(Main(8,23) - 100) > tol
    logText = logf(logText, 'Ps1: AB must be 100%% (found %.0f%%)\n', Main(8,23));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps1 = true;
end
if isnan(Main(8,24)) || abs(Main(8,24) - 400) > distTol
    logText = logf(logText, 'Ps1: Ps must be 400 (found %.0f)\n', Main(8,24));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps1 = true;
end
if isnan(Main(8,25)) || abs(Main(8,25) - 0) > tol
    logText = logf(logText, 'Ps1: CDx must be 0.000 (found %.3f)\n', Main(8,25));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps1 = true;
end
if abs(Main(8,19) - betaExpected) > wtoTol
    logText = logf(logText, 'Ps1: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(8,19));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps1 = true;
end

if abs(Main(9,20) - 10000) > altTol
    logText = logf(logText, 'Ps2: Altitude must be 10000 (found %.0f)\n', Main(9,20));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps2 = true;
end
if abs(Main(9,21) - 0.9) > machTol
    logText = logf(logText, 'Ps2: Mach must be 0.9 (found %.2f)\n', Main(9,21));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps2 = true;
end
if isnan(Main(9,22)) || abs(Main(9,22) - 1) > tol
    logText = logf(logText, 'Ps2: n must be 1 (found %.3f)\n', Main(9,22));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps2 = true;
end
if abs(Main(9,23)) > tol
    logText = logf(logText, 'Ps2: AB must be 0%% (found %.0f%%)\n', Main(9,23));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps2 = true;
end
if isnan(Main(9,24)) || abs(Main(9,24) - 400) > distTol
    logText = logf(logText, 'Ps2: Ps must be 400 (found %.0f)\n', Main(9,24));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps2 = true;
end
if isnan(Main(9,25)) || abs(Main(9,25) - 0) > tol
    logText = logf(logText, 'Ps2: CDx must be 0.000 (found %.3f)\n', Main(9,25));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps2 = true;
end
if abs(Main(9,19) - betaExpected) > wtoTol
    logText = logf(logText, 'Ps2: W/WTO must be set for 50%% fuel load (%.3f); found %.3f\n', betaExpected, Main(9,19));
    constraintErrors = constraintErrors + 1; rowErrorsMap.Ps2 = true;
end

if abs(Main(12,20)) > altTol
    logText = logf(logText, 'Takeoff: Altitude must be 0 (found %.0f)\n', Main(12,20));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,21) - 1.2) > machTol
    logText = logf(logText, 'Takeoff: V/Vstall must be 1.2 (found %.2f)\n', Main(12,21));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,22) - 0.03) > 5e-4
    logText = logf(logText, 'Takeoff: mu must be 0.03 (found %.3f)\n', Main(12,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,23) - 100) > tol
    logText = logf(logText, 'Takeoff: AB must be 100%% (found %.0f%%)\n', Main(12,23));
    constraintErrors = constraintErrors + 1;
end
if isnan(takeoff_dist) || abs(takeoff_dist - 3000) > distTol
    logText = logf(logText, 'Takeoff distance must be 3000 ft (found %.0f)\n', takeoff_dist);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(12,19) - 1.0) > wtoTol
    logText = logf(logText, 'Takeoff: W/WTO must be 1.000 within ±%.3f (found %.3f)\n', wtoTol, Main(12,19));
    constraintErrors = constraintErrors + 1;
end

cdxTakeoff = Main(12,25);
if isnan(cdxTakeoff) || abs(cdxTakeoff - 0.035) > tol
    logText = logf(logText, 'Takeoff: CDx must be 0.035 (found %.3f)\n', cdxTakeoff);
    constraintErrors = constraintErrors + 1;
end

if abs(Main(13,20)) > altTol
    logText = logf(logText, 'Landing: Altitude must be 0 (found %.0f)\n', Main(13,20));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,21) - 1.3) > machTol
    logText = logf(logText, 'Landing: V/Vstall must be 1.3 (found %.2f)\n', Main(13,21));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,22) - 0.5) > tol
    logText = logf(logText, 'Landing: mu must be 0.5 (found %.3f)\n', Main(13,22));
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,23)) > tol
    logText = logf(logText, 'Landing: AB must be 0%% (found %.0f%%)\n', Main(13,23));
    constraintErrors = constraintErrors + 1;
end
if isnan(landing_dist) || abs(landing_dist - 5000) > distTol
    logText = logf(logText, 'Landing distance must be 5000 ft (found %.0f)\n', landing_dist);
    constraintErrors = constraintErrors + 1;
end
if abs(Main(13,19) - 1.0) > wtoTol
    logText = logf(logText, 'Landing: W/WTO must be 1.000 within ±%.3f (found %.3f)\n', wtoTol, Main(13,19));
    constraintErrors = constraintErrors + 1;
end

cdxLanding = Main(13,25);
if isnan(cdxLanding) || abs(cdxLanding - 0.045) > tol
    logText = logf(logText, 'Landing: CDx must be 0.045 (found %.3f)\n', cdxLanding);
    constraintErrors = constraintErrors + 1;
end

if constraintErrors > 0
    logText = logf(logText, 'Constraint table has %d entry issue(s).\n', constraintErrors);
end
% Constraint curve compliance
constraintCurveFailures = 0;
failedCurves = {};
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
                logText = logf(logText, 'Constraint curve %s: T/W=%.3f below required %.3f at W/S=%.2f\n', columnLabels{idx}, TW_design, estimatedTWvalue, WS_design);
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

if constraintCurveFailures == 1
    logText = logf(logText, 'Design did not meet the following constraint curve: %s.%s\n', char(failedCurves{1}), curveSuffixFew);
elseif constraintCurveFailures >= 2
    summary = strjoin(string(failedCurves), ', ');
    suffix = curveSuffixFew;
    if constraintCurveFailures > 6
        suffix = curveSuffixMany;
    end
    logText = logf(logText, 'Design did not meet the following constraint curves: %s.%s\n', char(summary), suffix);
end

constraintsPass = constraintErrors == 0 && constraintCurveFailures == 0;
if ~constraintsPass
    logText = logf(logText, 'Constraint compliance not met; adjust design to satisfy all threshold constraints.\n');
end

% Payload
if isnan(aim120) || aim120 < 8 - tol
    count = aim120;
    if isnan(count), count = 0; end
    logText = logf(logText, 'Payload missing: need at least 8 AIM-120Ds (found %.0f)\n', count);
else
    payloadPass = true;
    if ~isnan(aim9) && aim9 >= 2 - tol
        payloadObjectivePass = true;
    end
end
% Stability
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
if ~(rat >= -1 && rat <= -0.3)
    logText = logf(logText, 'Cnb/Clb ratio must be between -1 and -0.3 (Q10 = %.3f)\n', rat);
    stabilityErrors = stabilityErrors + 1;
end

stabilityPass = stabilityErrors == 0;
if stabilityErrors > 0
    logText = logf(logText, 'Stability criteria failed in %d area(s).\n', stabilityErrors);
end

% Fuel (informational) and volume (scored later)
if isnan(fuel_available) || isnan(fuel_required) || fuel_available + tol < fuel_required
    logText = logf(logText, 'Fuel available (%.1f) is less than required (%.1f); check reserves.\n', fuel_available, fuel_required);
    fuelPass = false;
else
    fuelPass = true;
end

volumePass = ~(isnan(volume_remaining) || volume_remaining <= 0);
if ~volumePass
    logText = logf(logText, 'Volume remaining must be positive (Q23 = %.2f).\n', volume_remaining);
end

% Recurring cost threshold (scored later)
if abs(numaircraft - 187) > 1e-3
    logText = logf(logText, 'Number of aircraft (N31) must be 187 to evaluate cost thresholds (found %.0f).\n', numaircraft);
elseif isnan(cost)
    logText = logf(logText, 'Recurring cost missing for 187-aircraft estimate.\n');
else
    if cost < 120 + tol
        costPass = true;
    else
        logText = logf(logText, 'Cost above threshold: $%.1fM for 187 aircraft (needs <$120M).\n', cost);
    end
    if cost < 110 + tol
        costObjectivePass = true;
    end
end

% Landing gear geometry
gearFailures = 0;

g90 = Gear(20, 10);
if isnan(g90) || g90 < 80 - tol || g90 > 95 + tol
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Violates nose gear 90/10 rule: %.1f%% (must be between 80%% and 95%%)\n', g90);
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

rotationSpeed = Gear(20, 14); % N20
rotationRef = Gear(21, 14);   % N21
if isnan(rotationSpeed)
    gearFailures = gearFailures + 1;
    logText = logf(logText, 'Takeoff rotation speed (N20) missing; must be <200 kts and below N21.\n');
else
    if rotationSpeed >= 200 - tol
        gearFailures = gearFailures + 1;
        logText = logf(logText, 'Violates takeoff rotation speed: N20 = %.1f kts (must be < 200 kts)\n', rotationSpeed);
    end
    if isnan(rotationRef)
        gearFailures = gearFailures + 1;
        logText = logf(logText, 'Takeoff speed margin failed: N21 missing; N20 must be below N21.\n');
    else
        if rotationSpeed >= rotationRef - tol
            gearFailures = gearFailures + 1;
            logText = logf(logText, 'Takeoff speed margin failed: N20 must be less than N21 (N20 = %.1f, N21 = %.1f)\n', rotationSpeed, rotationRef);
        end
        if rotationRef > 200 + tol
            logText = logf(logText, 'Takeoff speed reference warning: N21 = %.1f kts (should be <= 200 kts)\n', rotationRef);
        end
    end
end

gearPass = gearFailures == 0;
if gearFailures > 0
    logText = logf(logText, 'Landing gear geometry outside limits in %d area(s).\n', gearFailures);
end

% Final score based on new rubric: start at 85, -5 per failed major section
pt = BASE_TOTAL;

constraintsBucketPass = constraintsPass && payloadPass && efficiencyPass && thrustPass && dataValidPass;
geometryBucketPass = controlPass && stabilityPass;

if ~constraintsBucketPass, pt = pt - 5; end
if ~rangePass,             pt = pt - 5; end
if ~geometryBucketPass,    pt = pt - 5; end
if ~gearPass,              pt = pt - 5; end
if ~fuelPass,              pt = pt - 5; end
if ~volumePass,            pt = pt - 5; end

pt = max(0, pt); % floor at 0

objectiveScore = 0;
if rangeObjectivePass,   objectiveScore = objectiveScore + 5; end
if costObjectivePass,    objectiveScore = objectiveScore + 5; end
if payloadObjectivePass, objectiveScore = objectiveScore + 5; end

thresholdScore = roundToTenth(pt);
objectiveScore = roundToTenth(objectiveScore);
pt = roundToTenth(thresholdScore + objectiveScore);

missing = {};
if ~constraintsBucketPass, missing{end+1} = 'constraints/payload/efficiency/Tavail/sheet validity'; end %#ok<AGROW>
if ~rangePass,             missing{end+1} = 'range'; end %#ok<AGROW>
if ~geometryBucketPass,    missing{end+1} = 'geometry (controls/stability)'; end %#ok<AGROW>
if ~gearPass,              missing{end+1} = 'landing gear'; end %#ok<AGROW>
if ~fuelPass,              missing{end+1} = 'fuel'; end %#ok<AGROW>
if ~volumePass,            missing{end+1} = 'volume remaining'; end %#ok<AGROW>
% Stealth and N21 checks are informational only
if ~stealthPass,           missing{end+1} = 'stealth shaping (no deduction)'; end %#ok<AGROW>
if ~missionPass,           missing{end+1} = 'mission table (no deduction)'; end %#ok<AGROW>

logText = logf(logText, 'Threshold score after deductions: %.1f / %d\n', thresholdScore, BASE_TOTAL);
if ~isempty(missing)
    logText = logf(logText, 'Checks not met: %s\n', strjoin(missing, ', '));
end
% Text summary: bucket status + reasons (assembled later for readability)
constraintReasons = {};
if ~constraintsPass, constraintReasons{end+1} = 'constraint table/curves'; end %#ok<AGROW>
if ~payloadPass,     constraintReasons{end+1} = 'payload'; end %#ok<AGROW>
if ~efficiencyPass,  constraintReasons{end+1} = 'efficiency guards'; end %#ok<AGROW>
if ~thrustPass,      constraintReasons{end+1} = 'Tavail>Drag'; end %#ok<AGROW>
if ~dataValidPass,   constraintReasons{end+1} = 'sheet validation'; end %#ok<AGROW>

geometryReasons = {};
if ~controlPass,   geometryReasons{end+1} = 'controls'; end %#ok<AGROW>
if ~stabilityPass, geometryReasons{end+1} = 'stability'; end %#ok<AGROW>

gearReasons = {};
if ~gearPass
    gearReasons{end+1} = sprintf('%d issue(s)', gearFailures);
end

bucketSummary = sprintf(['Bucket summary:\n', ...
    '  Constraints: %s%s\n', ...
    '  Range: %s\n', ...
    '  Geometry: %s%s\n', ...
    '  Gear: %s%s\n', ...
    '  Fuel: %s\n', ...
    '  Volume: %s\n', ...
    'Objectives: Range %s, Cost %s, Payload %s => +%.1f / %d\n'], ...
    ternary(constraintsBucketPass, 'PASS', 'FAIL (-5)'), ternary(isempty(constraintReasons), '', [' [' strjoin(constraintReasons, '; ') ']']), ...
    ternary(rangePass, 'PASS', 'FAIL (-5)'), ...
    ternary(geometryBucketPass, 'PASS', 'FAIL (-5)'), ternary(isempty(geometryReasons), '', [' [' strjoin(geometryReasons, '; ') ']']), ...
    ternary(gearPass, 'PASS', 'FAIL (-5)'), ternary(isempty(gearReasons), '', [' [' strjoin(gearReasons, '; ') ']']), ...
    ternary(fuelPass, 'PASS', 'FAIL (-5)'), ...
    ternary(volumePass, 'PASS', 'FAIL (-5)'), ...
    ternary(rangeObjectivePass, 'PASS', 'FAIL'), ...
    ternary(costObjectivePass, 'PASS', 'FAIL'), ...
    ternary(payloadObjectivePass, 'PASS', 'FAIL'), ...
    objectiveScore, OBJECTIVE_TOTAL);

scoreSummary = sprintf(['Threshold score after deductions: %.1f / %d\n', ...
    'Final score: %.1f / 100\n'], thresholdScore, BASE_TOTAL, pt);

% Reorder log: filename, bucket summary, all prior messages, then scores
logLines = splitlines(string(logText));
if isempty(logLines)
    logLines = strings(1,1);
end
firstLine = logLines(1);
remainingLines = logLines(2:end);
logText = strjoin([firstLine; splitlines(string(bucketSummary)); remainingLines; splitlines(string(scoreSummary))], newline);

% Command window summary (GE3-style quick view)
bucketLabel = @(pass, reasons) ternary(pass, 'PASS', ...
    ['FAIL (-5)', ternary(isempty(reasons), '', [' [' strjoin(reasons, '; ') ']'])]);

constraintReasons = {};
if ~constraintsPass, constraintReasons{end+1} = 'constraint table/curves'; end %#ok<AGROW>
if ~payloadPass,     constraintReasons{end+1} = 'payload'; end %#ok<AGROW>
if ~efficiencyPass,  constraintReasons{end+1} = 'efficiency guards'; end %#ok<AGROW>
if ~thrustPass,      constraintReasons{end+1} = 'Tavail>Drag'; end %#ok<AGROW>
if ~dataValidPass,   constraintReasons{end+1} = 'sheet validation'; end %#ok<AGROW>

geometryReasons = {};
if ~controlPass,   geometryReasons{end+1} = 'controls'; end %#ok<AGROW>
if ~stabilityPass, geometryReasons{end+1} = 'stability'; end %#ok<AGROW>

gearReasons = {};
if ~gearPass
    gearReasons{end+1} = sprintf('%d issue(s)', gearFailures);
end

fprintf('%s completed\n', [name, ext]);
fprintf('Base after deductions: %.1f / %d\n', thresholdScore, BASE_TOTAL);
fprintf('Buckets: Constraints %s | Range %s | Geometry %s | Gear %s | Fuel %s | Volume %s\n', ...
    bucketLabel(constraintsBucketPass, constraintReasons), ...
    bucketLabel(rangePass, {}), ...
    bucketLabel(geometryBucketPass, geometryReasons), ...
    bucketLabel(gearPass, gearReasons), ...
    bucketLabel(fuelPass, {}), ...
    bucketLabel(volumePass, {}));
if ~stealthPass
    fprintf('Advisory: Stealth shaping issues (no deduction).\n');
end
if ~missionPass
    fprintf('Advisory: Mission table mismatches (no deduction).\n');
end
fprintf('Objectives: Range %s, Cost %s, Payload %s => +%.1f / %d\n', ...
    ternary(rangeObjectivePass, 'PASS', 'FAIL'), ...
    ternary(costObjectivePass, 'PASS', 'FAIL'), ...
    ternary(payloadObjectivePass, 'PASS', 'FAIL'), ...
    objectiveScore, OBJECTIVE_TOTAL);
fprintf('Final score: %.1f / 100\n\n', pt);

fb = char(logText);
% Collapse excessive blank lines for readability
fb = regexprep(string(fb), '\n{3,}', '\n\n');
fb = char(strtrim(fb));
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
sheets.Gear   = safeReadMatrix(filename, 'Gear',   {'J20','L20','L21','M20','M21','N20'});
sheets.Geom   = safeReadMatrix(filename, 'Geom',   {'C8','C10','M152','K15','L155','L38'});

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

% Try fast read without invoking Excel UI (prevents password dialogs)
try
    data = readmatrix(filename, 'Sheet', sheetname, 'DataRange', 'A1:AQ250', 'UseExcel', false);
catch ME
    if contains(ME.message, 'password', 'IgnoreCase', true)
        error('FinalExam:PasswordProtected', 'Password-protected file: %s', filename);
    end
    rethrow(ME);
end


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

function hit = teNormalHitsCenterline(tipPoint, innerPoint)
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
    hit = true;
    break;
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

function [logText, failures] = requireParallelAngle(logText, failures, angle, wingAngle, tol, template)
if isnan(angle) || isnan(wingAngle)
    logText = logf(logText, 'Unable to verify stealth shaping due to missing geometry data\n');
    failures = failures + 1;
elseif ~anglesParallel(angle, wingAngle, tol)
    logText = logf(logText, template, angle, wingAngle, tol);
    failures = failures + 1;
end
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
logText = logText + string(logEntry); % Append while keeping a single string scalar
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
            csvFilename = fullfile(folderAnalyzed, ['FinalProject_Blackboard_Offline_', timestamp, '.csv']);
            fid = fopen(csvFilename, 'w');

            % Assignment title column (update if needed)
            assignmentTitle = 'Final Exam Jet11 Problem [Total Pts: 100 Score] |409586';

            % Write header (username only; no first/last name columns)
            fprintf(fid, '"Username","%s","Feedback to Learner","Feedback Format","Grading Notes","Notes Format"\n', assignmentTitle);

            for i = 1:numel(files)
                fname = files(i).name;

                % Extract username from the canonical segment "_<username>_attempt"
                userTok = regexp(fname, '_([A-Za-z0-9\\.\\-]+)_attempt', 'tokens', 'once');
                if ~isempty(userTok)
                    username = userTok{1};
                else
                    username = 'UNKNOWN';
                end

                % Get score and feedback
                score = max(0, min(100, roundToTenth(points(i))));
                % Match Blackboard feedback to the per-cadet log entry (same text as .txt)
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

                % Write row
                fprintf(fid, '"%s","%.1f","%s","%s","",""\n', ...
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

function out = ternary(cond, a, b)
if cond
    out = a;
else
    out = b;
end
end
