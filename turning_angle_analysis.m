%% ============================================================
%  FINAL HIGH-SPEED TURNING-ANGLE ANALYSIS
%
%  FINAL STABLE VERSION
%
%  ============================================================
%
%  TRACK ID:
%
%       Sheet + mRNA_ID
%
%  TURNING ANGLE:
%
%       v1 = P(i+LAG)   - P(i)
%       v2 = P(i+2LAG) - P(i+LAG)
%
%       theta = atan2(cross(v1,v2),dot(v1,v2))
%
%  Angle is assigned to:
%
%       original row i + 2*LAG
%
%
%  ============================================================
%  MAIN FEATURES
%  ============================================================
%
%  1. High-speed trajectory processing
%  2. Sheet + mRNA_ID defines unique track
%  3. LAG = 8
%  4. Minimum track length = 17 frames
%  5. PARFOR acceleration
%  6. No polarhistogram
%  7. No patch on polaraxes
%  8. Rose plot = ordinary axes + Cartesian patch
%  9. Rose plot uses only 72 bins
% 10. Rose radial scale has 1.20x outer margin
% 11. EPS vector output
% 12. PNG output
% 13. Folded-angle distribution
% 14. Circular statistics
% 15. Forward cone statistics
% 16. Excel output split automatically
%
%  ============================================================

clc;
clear;
close all force;

fprintf('\n');
fprintf('============================================================\n');
fprintf(' FINAL HIGH-SPEED TURNING-ANGLE ANALYSIS\n');
fprintf(' FINAL STABLE ROSE VERSION\n');
fprintf(' ROSE OUTER MARGIN = 1.20 x MAX DATA\n');
fprintf('============================================================\n');


%% ============================================================
% USER SETTINGS
% ============================================================

filename = 'Method2_Selected_20Frames.xlsx';


%% ============================================================
% LOGICAL GROUPS
% ============================================================

logicalSheetNames = { ...
    'AllQualified_Directed','Normaldiffusive'
    };


%% ============================================================
% TURNING ANGLE LAG
% ============================================================

LAG = 8;


%% ============================================================
% ROSE PLOT
% ============================================================

N_ROSE_BINS = 72;


%% ============================================================
% ROSE OUTER MARGIN
%
% IMPORTANT:
%
%   rDataMax = maximum actual rose probability
%
%   rMax = rDataMax * 1.20
%
% Therefore the tallest bar occupies only ~83.3%
% of the outer radius.
% ============================================================

ROSE_MARGIN_FACTOR = 1.05;


%% ============================================================
% FORWARD CONES
% ============================================================

CONES = [30 45 60 90];


%% ============================================================
% OUTPUT DIRECTORY
% ============================================================

outputDir = sprintf( ...
    'Method2TurningAngle_Lag%d', ...
    LAG);


%% ============================================================
% OUTPUT FILES
% ============================================================

angleFile = fullfile( ...
    outputDir, ...
    sprintf('Angles_Lag%d.xlsx',LAG));


statsFile = fullfile( ...
    outputDir, ...
    sprintf('Stats_Lag%d.xlsx',LAG));


%% ============================================================
% PARFOR
% ============================================================

USE_PARFOR = true;


%% ============================================================
% FULL TABLE
% ============================================================

WRITE_FULL_TABLE = false;


%% ============================================================
% EXCEL LIMIT
% ============================================================

EXCEL_MAX_ROWS = 1048576;

MAX_DATA_ROWS_PER_SHEET = ...
    EXCEL_MAX_ROWS - 1;


%% ============================================================
% FOLDED ANGLE
% ============================================================

FOLDED_EDGES = ...
    0:15:180;


FOLDED_CENTERS = ...
    FOLDED_EDGES(1:end-1) + ...
    diff(FOLDED_EDGES)/2;


%% ============================================================
% PLOT COLORS
% ============================================================

plotColors = [ ...
    0.30 0.45 0.69;
    0.77 0.31 0.32;
    0.33 0.66 0.41;
    0.51 0.45 0.70;
    0.80 0.73 0.27];


%% ============================================================
% CREATE OUTPUT DIRECTORY
% ============================================================

if ~exist(outputDir,'dir')

    mkdir(outputDir);

end


%% ============================================================
% DELETE OLD OUTPUT
% ============================================================

if exist(angleFile,'file')

    try

        delete(angleFile);

        fprintf( ...
            'Old angle workbook deleted.\n');

    catch ME

        warning( ...
            'Could not delete old angle workbook: %s', ...
            ME.message);

    end

end


if exist(statsFile,'file')

    try

        delete(statsFile);

        fprintf( ...
            'Old statistics workbook deleted.\n');

    catch ME

        warning( ...
            'Could not delete old statistics workbook: %s', ...
            ME.message);

    end

end


%% ============================================================
% INPUT CHECK
% ============================================================

if ~exist(filename,'file')

    error( ...
        'Input Excel file does not exist:\n%s', ...
        filename);

end


%% ============================================================
% SETTINGS REPORT
% ============================================================

fprintf('\n');
fprintf('Input file          : %s\n',filename);
fprintf('LAG                 : %d\n',LAG);
fprintf('Minimum frames      : %d\n',2*LAG+1);
fprintf('Rose bins           : %d\n',N_ROSE_BINS);
fprintf('Rose margin factor  : %.2fx\n',ROSE_MARGIN_FACTOR);
fprintf('PARFOR              : %d\n',USE_PARFOR);
fprintf('Full table output   : %d\n',WRITE_FULL_TABLE);
fprintf('Output directory    : %s\n',outputDir);

fprintf('============================================================\n');


%% ============================================================
% START PARALLEL POOL
% ============================================================

if USE_PARFOR

    try

        poolobj = gcp('nocreate');


        if isempty(poolobj)

            fprintf('\nStarting parallel pool...\n');

            poolobj = parpool;

        end


        fprintf( ...
            'Parallel workers: %d\n', ...
            poolobj.NumWorkers);


    catch ME

        warning( ...
            'PARFOR unavailable: %s', ...
            ME.message);

        USE_PARFOR = false;

    end

end


%% ============================================================
% GET SHEET NAMES
% ============================================================

fprintf('\n');
fprintf('Reading worksheet names...\n');


actualSheetNames = ...
    string(sheetnames(filename));


fprintf('\nAvailable worksheets:\n');


for k = 1:numel(actualSheetNames)

    fprintf( ...
        '  %s\n', ...
        actualSheetNames(k));

end


%% ============================================================
% BUILD LOGICAL GROUPS
% ============================================================

groupNames = {};
groupSourceSheets = {};


for g = 1:numel(logicalSheetNames)


    logicalName = ...
        logicalSheetNames{g};


    %% ========================================================
    % ALLQUALIFIED DIRECTED
    %% ========================================================

    if strcmp( ...
            logicalName, ...
            'AllQualified_Directed')


        matchIdx = ...
            startsWith( ...
            actualSheetNames, ...
            'AllQualified_Directed');


        matchedSheets = ...
            actualSheetNames(matchIdx);


        if isempty(matchedSheets)

            warning( ...
                'No AllQualified_Directed worksheets found.');

            continue;

        end


        matchedSheets = ...
            sort(matchedSheets);


        groupNames{end+1,1} = ...
            'AllQualified_Directed_Part1';


        groupSourceSheets{end+1,1} = ...
            cellstr(matchedSheets);


    else


        %% ====================================================
        % NORMAL GROUP
        %% ====================================================

        matchIdx = ...
            strcmp( ...
            actualSheetNames, ...
            logicalName);


        if ~any(matchIdx)

            warning( ...
                'Worksheet "%s" not found.', ...
                logicalName);

            continue;

        end


        groupNames{end+1,1} = ...
            logicalName;


        groupSourceSheets{end+1,1} = ...
            {logicalName};

    end

end


%% ============================================================
% SHOW GROUPS
% ============================================================

fprintf('\n');
fprintf('============================================================\n');
fprintf('LOGICAL GROUPS\n');
fprintf('============================================================\n');


for g = 1:numel(groupNames)

    fprintf('\n%s\n',groupNames{g});


    sourceSheets = ...
        groupSourceSheets{g};


    for j = 1:numel(sourceSheets)

        fprintf( ...
            '    <- %s\n', ...
            sourceSheets{j});

    end

end


%% ============================================================
% STATISTICS STORAGE
% ============================================================

statRows = cell(0,13);

coneRows = cell(0,7);


%% ============================================================
% FOLDED HISTOGRAM STORAGE
% ============================================================

foldedHistAll = struct();

foldedNAll = struct();


%% ============================================================
% PROCESS GROUPS
% ============================================================

for s = 1:numel(groupNames)


    excelSheet = ...
        groupNames{s};


    sourceSheets = ...
        groupSourceSheets{s};


    fprintf('\n');
    fprintf('============================================================\n');
    fprintf('PROCESSING GROUP: %s\n',excelSheet);
    fprintf('============================================================\n');


    groupTimer = tic;


    %% ========================================================
    % READ SOURCE SHEETS
    %% ========================================================

    dataParts = ...
        cell(numel(sourceSheets),1);


    validPart = ...
        false(numel(sourceSheets),1);


    totalRows = 0;


    for j = 1:numel(sourceSheets)


        sourceSheet = ...
            sourceSheets{j};


        fprintf( ...
            '\n  Reading sheet: %s\n', ...
            sourceSheet);


        readTimer = tic;


        %% ====================================================
        % READ TABLE
        %% ====================================================

        try

            tempData = ...
                readtable( ...
                filename, ...
                'Sheet',sourceSheet);

        catch ME

            warning( ...
                'Could not read sheet "%s": %s', ...
                sourceSheet, ...
                ME.message);

            continue;

        end


        %% ====================================================
        % REQUIRED COLUMNS
        %% ====================================================

        requiredColumns = { ...
            'Sheet', ...
            'mRNA_ID', ...
            'X', ...
            'Y', ...
            'Time'};


        varNames = ...
            tempData.Properties.VariableNames;


        if ~all( ...
                ismember( ...
                requiredColumns, ...
                varNames))


            warning( ...
                ['Sheet "%s" does not contain all ', ...
                 'required columns.'], ...
                sourceSheet);


            clear tempData;

            continue;

        end


        %% ====================================================
        % EMPTY
        %% ====================================================

        if height(tempData) == 0

            fprintf( ...
                '  Empty sheet.\n');

            clear tempData;

            continue;

        end


        %% ====================================================
        % KEEP REQUIRED COLUMNS
        %% ====================================================

        if ~WRITE_FULL_TABLE

            tempData = ...
                tempData(:, ...
                { ...
                'Sheet', ...
                'mRNA_ID', ...
                'X', ...
                'Y', ...
                'Time'});

        end


        dataParts{j} = ...
            tempData;


        validPart(j) = true;


        nThis = ...
            height(tempData);


        totalRows = ...
            totalRows + nThis;


        fprintf( ...
            '  Rows: %d\n', ...
            nThis);


        fprintf( ...
            '  Read time: %.2f s\n', ...
            toc(readTimer));

    end


    %% ========================================================
    % CHECK
    %% ========================================================

    if ~any(validPart)

        warning( ...
            'No valid data found for %s.', ...
            excelSheet);

        continue;

    end


    %% ========================================================
    % MERGE
    %% ========================================================

    fprintf('\n');
    fprintf('  Merging source sheets...\n');


    mergeTimer = tic;


    data = ...
        vertcat(dataParts{validPart});


    clear dataParts;


    nRows = ...
        height(data);


    fprintf( ...
        '  Total rows: %d\n', ...
        nRows);


    fprintf( ...
        '  Merge time: %.2f s\n', ...
        toc(mergeTimer));


    %% ========================================================
    % EXTRACT NUMERIC ARRAYS
    %% ========================================================

    fprintf( ...
        '\n  Extracting trajectory arrays...\n');


    X = ...
        double(data.X);


    Y = ...
        double(data.Y);


    Time = ...
        double(data.Time);


    %% ========================================================
    % TRACK IDENTIFIERS
    %% ========================================================

    fprintf( ...
        '  Creating numeric track IDs... ');


    trackTimer = tic;


    sheetID = ...
        data.Sheet;


    idData = ...
        data.mRNA_ID;


    if ~isstring(sheetID)

        sheetID = ...
            string(sheetID);

    end


    if ~isstring(idData)

        idData = ...
            string(idData);

    end


    %% ========================================================
    % TRACK = SHEET + mRNA_ID
    %% ========================================================

    [groupID,~] = ...
        findgroups( ...
        sheetID, ...
        idData);


    groupID = ...
        double(groupID);


    fprintf( ...
        'done %.2f s\n', ...
        toc(trackTimer));


    %% ========================================================
    % ORIGINAL ROW
    %% ========================================================

    originalRow = ...
        (1:nRows)';


    %% ========================================================
    % SORT
    %% ========================================================

    fprintf( ...
        '  Sorting trajectory rows... ');


    sortTimer = tic;


    [~,sortIdx] = ...
        sortrows( ...
        [groupID Time], ...
        [1 2]);


    %% ========================================================
    % DUPLICATE CHECK
    %% ========================================================

    sortedGroupTmp = ...
        groupID(sortIdx);


    sortedTimeTmp = ...
        Time(sortIdx);


    duplicatePair = ...
        false(nRows,1);


    if nRows > 1

        duplicatePair(2:end) = ...
            sortedGroupTmp(2:end) == ...
            sortedGroupTmp(1:end-1) & ...
            sortedTimeTmp(2:end) == ...
            sortedTimeTmp(1:end-1);

    end


    if any(duplicatePair)


        fprintf( ...
            '\n  Duplicate TrackID + Time detected.\n');


        stableMatrix = [ ...
            sortedGroupTmp, ...
            sortedTimeTmp, ...
            originalRow(sortIdx)];


        [~,stableIdx] = ...
            sortrows( ...
            stableMatrix, ...
            [1 2 3]);


        sortIdx = ...
            sortIdx(stableIdx);


        clear stableMatrix;
        clear stableIdx;

    end


    clear sortedGroupTmp;
    clear sortedTimeTmp;
    clear duplicatePair;


    fprintf( ...
        'done %.2f s\n', ...
        toc(sortTimer));


    %% ========================================================
    % SORTED ARRAYS
    %% ========================================================

    fprintf( ...
        '  Creating sorted arrays... ');


    arrayTimer = tic;


    groupSorted = ...
        groupID(sortIdx);


    Xsorted = ...
        X(sortIdx);


    Ysorted = ...
        Y(sortIdx);


    originalRowSorted = ...
        originalRow(sortIdx);


    fprintf( ...
        'done %.2f s\n', ...
        toc(arrayTimer));


    %% ========================================================
    % RELEASE ORIGINAL ARRAYS
    %% ========================================================

    clear X;
    clear Y;
    clear Time;

    clear groupID;
    clear originalRow;

    clear sheetID;
    clear idData;


    %% ========================================================
    % TRACK BOUNDARIES
    %% ========================================================

    fprintf( ...
        '  Finding track boundaries... ');


    boundaryTimer = tic;


    newTrack = ...
        [ ...
        true; ...
        groupSorted(2:end) ~= ...
        groupSorted(1:end-1)];


    groupStart = ...
        find(newTrack);


    groupEnd = ...
        [ ...
        groupStart(2:end)-1; ...
        nRows];


    clear newTrack;
    clear groupSorted;


    %% ========================================================
    % TRACK LENGTH
    %% ========================================================

    trackLength = ...
        groupEnd - groupStart + 1;


    nTracksTotal = ...
        numel(groupStart);


    minTrackLength = ...
        2*LAG + 1;


    validLength = ...
        trackLength >= minTrackLength;


    validTrackIndices = ...
        find(validLength);


    nValidLengthTracks = ...
        numel(validTrackIndices);


    fprintf( ...
        'done %.2f s\n', ...
        toc(boundaryTimer));


    fprintf( ...
        '  Total tracks        : %d\n', ...
        nTracksTotal);


    fprintf( ...
        '  Tracks >= %d frames : %d\n', ...
        minTrackLength, ...
        nValidLengthTracks);


    %% ========================================================
    % ORIGINAL ROW ANGLE ARRAY
    %% ========================================================

    angleOriginal = ...
        nan(nRows,1);


    %% ========================================================
    % PER-TRACK CIRCULAR STATISTICS
    %% ========================================================

    trackCos = ...
        nan(nValidLengthTracks,1);


    trackSin = ...
        nan(nValidLengthTracks,1);


    nAngPerTrk = ...
        zeros(nValidLengthTracks,1);


    %% ========================================================
    % TURNING ANGLE
    %% ========================================================

    fprintf('\n');
    fprintf('  Calculating turning angles...\n');


    angleTimer = tic;


    if USE_PARFOR


        resultCells = ...
            cell(nValidLengthTracks,1);


        parfor q = 1:nValidLengthTracks


            u = ...
                validTrackIndices(q);


            iStart = ...
                groupStart(u);


            iEnd = ...
                groupEnd(u);


            x = ...
                Xsorted(iStart:iEnd);


            y = ...
                Ysorted(iStart:iEnd);


            orig = ...
                originalRowSorted(iStart:iEnd);


            nLocal = ...
                numel(x);


            if nLocal < 2*LAG+1

                continue;

            end


            %% -----------------------------------------------
            % V1
            %% -----------------------------------------------

            v1x = ...
                x(1+LAG:end-LAG) ...
                - x(1:end-2*LAG);


            v1y = ...
                y(1+LAG:end-LAG) ...
                - y(1:end-2*LAG);


            %% -----------------------------------------------
            % V2
            %% -----------------------------------------------

            v2x = ...
                x(1+2*LAG:end) ...
                - x(1+LAG:end-LAG);


            v2y = ...
                y(1+2*LAG:end) ...
                - y(1+LAG:end-LAG);


            %% -----------------------------------------------
            % VECTOR VALIDITY
            %% -----------------------------------------------

            v1sq = ...
                v1x.^2 + v1y.^2;


            v2sq = ...
                v2x.^2 + v2y.^2;


            validVector = ...
                isfinite(v1x) & ...
                isfinite(v1y) & ...
                isfinite(v2x) & ...
                isfinite(v2y) & ...
                v1sq > 0 & ...
                v2sq > 0;


            %% -----------------------------------------------
            % CROSS
            %% -----------------------------------------------

            crossZ = ...
                v1x .* v2y ...
                - v1y .* v2x;


            %% -----------------------------------------------
            % DOT
            %% -----------------------------------------------

            dotP = ...
                v1x .* v2x ...
                + v1y .* v2y;


            %% -----------------------------------------------
            % ANGLE
            %% -----------------------------------------------

            theta = ...
                atan2(crossZ,dotP);


            %% -----------------------------------------------
            % VALID ANGLES
            %% -----------------------------------------------

            good = ...
                validVector & ...
                isfinite(theta);


            if any(good)


                idx = ...
                    find(good);


                %% -------------------------------------------
                % ANGLE ASSIGNED TO i + 2*LAG
                %% -------------------------------------------

                idx = ...
                    idx + 2*LAG;


                localRows = ...
                    orig(idx);


                thetaGood = ...
                    theta(good);


                resultCells{q} = { ...
                    localRows, ...
                    thetaGood, ...
                    mean(cos(thetaGood)), ...
                    mean(sin(thetaGood)), ...
                    numel(thetaGood)};

            end

        end


        %% ====================================================
        % RESTORE ORIGINAL ROW POSITIONS
        %% ====================================================

        fprintf( ...
            '  Restoring original row positions... ');


        restoreTimer = tic;


        for q = 1:nValidLengthTracks


            if isempty(resultCells{q})

                continue;

            end


            result = ...
                resultCells{q};


            localRows = ...
                result{1};


            thetaGood = ...
                result{2};


            angleOriginal(localRows) = ...
                rad2deg(thetaGood);


            trackCos(q) = ...
                result{3};


            trackSin(q) = ...
                result{4};


            nAngPerTrk(q) = ...
                result{5};

        end


        fprintf( ...
            'done %.2f s\n', ...
            toc(restoreTimer));


        clear resultCells;


    else


        %% ====================================================
        % NORMAL FOR LOOP
        %% ====================================================

        for q = 1:nValidLengthTracks


            u = ...
                validTrackIndices(q);


            iStart = ...
                groupStart(u);


            iEnd = ...
                groupEnd(u);


            x = ...
                Xsorted(iStart:iEnd);


            y = ...
                Ysorted(iStart:iEnd);


            orig = ...
                originalRowSorted(iStart:iEnd);


            %% -----------------------------------------------
            % V1
            %% -----------------------------------------------

            v1x = ...
                x(1+LAG:end-LAG) ...
                - x(1:end-2*LAG);


            v1y = ...
                y(1+LAG:end-LAG) ...
                - y(1:end-2*LAG);


            %% -----------------------------------------------
            % V2
            %% -----------------------------------------------

            v2x = ...
                x(1+2*LAG:end) ...
                - x(1+LAG:end-LAG);


            v2y = ...
                y(1+2*LAG:end) ...
                - y(1+LAG:end-LAG);


            %% -----------------------------------------------
            % VECTOR VALIDITY
            %% -----------------------------------------------

            v1sq = ...
                v1x.^2 + v1y.^2;


            v2sq = ...
                v2x.^2 + v2y.^2;


            validVector = ...
                isfinite(v1x) & ...
                isfinite(v1y) & ...
                isfinite(v2x) & ...
                isfinite(v2y) & ...
                v1sq > 0 & ...
                v2sq > 0;


            %% -----------------------------------------------
            % CROSS
            %% -----------------------------------------------

            crossZ = ...
                v1x .* v2y ...
                - v1y .* v2x;


            %% -----------------------------------------------
            % DOT
            %% -----------------------------------------------

            dotP = ...
                v1x .* v2x ...
                + v1y .* v2y;


            %% -----------------------------------------------
            % ANGLE
            %% -----------------------------------------------

            theta = ...
                atan2(crossZ,dotP);


            %% -----------------------------------------------
            % VALID
            %% -----------------------------------------------

            good = ...
                validVector & ...
                isfinite(theta);


            if any(good)


                idx = ...
                    find(good);


                idx = ...
                    idx + 2*LAG;


                localRows = ...
                    orig(idx);


                thetaGood = ...
                    theta(good);


                angleOriginal(localRows) = ...
                    rad2deg(thetaGood);


                trackCos(q) = ...
                    mean(cos(thetaGood));


                trackSin(q) = ...
                    mean(sin(thetaGood));


                nAngPerTrk(q) = ...
                    numel(thetaGood);

            end

        end

    end


    fprintf( ...
        '  Turning-angle calculation time: %.2f s\n', ...
        toc(angleTimer));


    %% ========================================================
    % VALID RESULTS
    %% ========================================================

    validResult = ...
        nAngPerTrk > 0 & ...
        isfinite(trackCos) & ...
        isfinite(trackSin);


    trackCos = ...
        trackCos(validResult);


    trackSin = ...
        trackSin(validResult);


    nAngPerTrk = ...
        nAngPerTrk(validResult);


    nTrk = ...
        numel(trackCos);


    nAng = ...
        sum(nAngPerTrk);


    %% ========================================================
    % CHECK
    %% ========================================================

    if nAng == 0

        warning( ...
            'No valid turning angles in group %s.', ...
            excelSheet);

        continue;

    end


    %% ========================================================
    % EXTRACT ANGLES
    %% ========================================================

    fprintf( ...
        '\n  Preparing angle statistics...\n');


    allAngles = ...
        angleOriginal( ...
        isfinite(angleOriginal));


    %% ========================================================
    % POOLED CIRCULAR STATISTICS
    %% ========================================================

    cosAll = ...
        cosd(allAngles);


    sinAll = ...
        sind(allAngles);


    meanCos = ...
        mean(cosAll);


    meanSin = ...
        mean(sinAll);


    Rpooled = ...
        hypot( ...
        meanCos, ...
        meanSin);


    meanDir = ...
        mod( ...
        atan2d( ...
        meanSin, ...
        meanCos), ...
        360);


    %% ========================================================
    % PER-TRACK CIRCULAR STATISTICS
    %% ========================================================

    meanTrackCos = ...
        mean(trackCos);


    meanTrackSin = ...
        mean(trackSin);


    RperTrack = ...
        hypot( ...
        meanTrackCos, ...
        meanTrackSin);


    meanDirTrack = ...
        mod( ...
        atan2d( ...
        meanTrackSin, ...
        meanTrackCos), ...
        360);


    %% ========================================================
    % RAYLEIGH
    %% ========================================================

    RayleighZ = ...
        nTrk * RperTrack^2;


    if nTrk > 0


        pRayleigh = ...
            exp(-RayleighZ) .* ...
            ( ...
            1 ...
            + (2*RayleighZ-RayleighZ^2) ...
            /(4*nTrk) ...
            - (24*RayleighZ ...
            -132*RayleighZ^2 ...
            +76*RayleighZ^3 ...
            -9*RayleighZ^4) ...
            /(288*nTrk^2) ...
            );


        pRayleigh = ...
            min( ...
            max(pRayleigh,0), ...
            1);

    else

        pRayleigh = NaN;

    end


    %% ========================================================
    % V TEST
    %% ========================================================

    if nTrk > 0


        Vstat = ...
            nTrk * meanTrackCos;


        uStat = ...
            Vstat * sqrt(2/nTrk);


        pV = ...
            0.5 * erfc( ...
            uStat/sqrt(2));

    else

        uStat = NaN;
        pV = NaN;

    end


    %% ========================================================
    % FOLDED ANGLE
    %% ========================================================

    folded = ...
        abs(allAngles);


    %% ========================================================
    % FORWARD CONES
    %% ========================================================

    coneVals = ...
        zeros(1,numel(CONES));


    for c = 1:numel(CONES)

        coneVals(c) = ...
            100 * mean( ...
            folded <= CONES(c));

    end


    %% ========================================================
    % FOLDED HISTOGRAM
    %% ========================================================

    foldedHist = ...
        histcounts( ...
        folded, ...
        FOLDED_EDGES);


    foldedTotal = ...
        numel(folded);


    fieldName = ...
        matlab.lang.makeValidName( ...
        excelSheet);


    foldedHistAll.(fieldName) = ...
        foldedHist;


    foldedNAll.(fieldName) = ...
        foldedTotal;


    %% ========================================================
    % ROSE HISTOGRAM
    %
    % Only 72 bins are retained for plotting.
    %% ========================================================

    fprintf( ...
        '  Calculating %d-bin rose histogram...\n', ...
        N_ROSE_BINS);


    roseEdgesDeg = ...
        linspace( ...
        -180, ...
        180, ...
        N_ROSE_BINS+1);


    roseCentersDeg = ...
        ( ...
        roseEdgesDeg(1:end-1) + ...
        roseEdgesDeg(2:end)) / 2;


    roseCounts = ...
        histcounts( ...
        allAngles, ...
        roseEdgesDeg);


    roseTotal = ...
        sum(roseCounts);


    if roseTotal > 0

        roseProbability = ...
            roseCounts ./ roseTotal;

    else

        roseProbability = ...
            zeros(size(roseCounts));

    end


    %% ========================================================
    % RELEASE LARGE ANGLE ARRAYS
    %% ========================================================

    clear allAngles;
    clear cosAll;
    clear sinAll;
    clear folded;


    %% ========================================================
    % REPORT
    %% ========================================================

    fprintf('\n');
    fprintf('------------------------------------------------------------\n');
    fprintf('RESULT: %s\n',excelSheet);
    fprintf('------------------------------------------------------------\n');


    fprintf( ...
        'Rows               : %d\n', ...
        nRows);


    fprintf( ...
        'Total tracks       : %d\n', ...
        nTracksTotal);


    fprintf( ...
        'Valid tracks       : %d\n', ...
        nTrk);


    fprintf( ...
        'Turning angles     : %d\n', ...
        nAng);


    fprintf( ...
        'Angles / track     : %.3f\n', ...
        mean(nAngPerTrk));


    fprintf( ...
        'Pooled mean cos    : %+0.5f\n', ...
        meanCos);


    fprintf( ...
        'Pooled R           : %.5f\n', ...
        Rpooled);


    fprintf( ...
        'Pooled mean dir    : %.3f deg\n', ...
        meanDir);


    fprintf( ...
        'Per-track R        : %.5f\n', ...
        RperTrack);


    fprintf( ...
        'Rayleigh Z         : %.5f\n', ...
        RayleighZ);


    fprintf( ...
        'Rayleigh p         : %.6g\n', ...
        pRayleigh);


    fprintf( ...
        'V u                : %.5f\n', ...
        uStat);


    fprintf( ...
        'V p                : %.6g\n', ...
        pV);


    for c = 1:numel(CONES)

        fprintf( ...
            'Within +/- %d deg  : %.3f %%\n', ...
            CONES(c), ...
            coneVals(c));

    end


    %% ========================================================
    % SAVE STATISTICS
    %% ========================================================

    statRows(end+1,:) = { ...
        excelSheet, ...
        nTrk, ...
        nAng, ...
        mean(nAngPerTrk), ...
        meanCos, ...
        Rpooled, ...
        meanDir, ...
        RperTrack, ...
        meanDirTrack, ...
        RayleighZ, ...
        pRayleigh, ...
        uStat, ...
        pV};


    %% ========================================================
    % SAVE CONE RESULTS
    %% ========================================================

    coneRows(end+1,:) = { ...
        excelSheet, ...
        nTrk, ...
        nAng, ...
        coneVals(1), ...
        coneVals(2), ...
        coneVals(3), ...
        coneVals(4)};


    %% ========================================================
    % FAST ROSE PLOT
    %
    % IMPORTANT:
    %
    % rDataMax = actual maximum probability
    %
    % rMax = rDataMax * 1.20
    %
    % This creates 20% empty radial space around the
    % tallest wedge.
    %% ========================================================

    fprintf('\n');
    fprintf('  Creating FAST rose plot...\n');


    roseTimer = tic;


    %% ========================================================
    % FIGURE
    %% ========================================================

    figRose = ...
        figure( ...
        'Visible','off', ...
        'Color','w', ...
        'Position',[100 100 620 620]);


    axRose = ...
        axes( ...
        'Parent',figRose);


    hold(axRose,'on');


    axis(axRose,'equal');

    axis(axRose,'off');


    %% ========================================================
    % COLOR
    %% ========================================================

    barColor = ...
        plotColors( ...
        min(s,size(plotColors,1)),:);


    %% ========================================================
    % ROSE EDGE ANGLES
    %% ========================================================

    roseEdgesRad = ...
        roseEdgesDeg * pi / 180;


    %% ========================================================
    % ACTUAL DATA MAXIMUM
    %% ========================================================

    rDataMax = ...
        max(roseProbability);


    if isempty(rDataMax) || ...
            ~isfinite(rDataMax) || ...
            rDataMax <= 0

        rDataMax = 1;

    end


    %% ========================================================
    % OUTER RADIUS
    %
    % 1.20x maximum data
    %% ========================================================

    rMax = ...
        ROSE_MARGIN_FACTOR * rDataMax;


    %% ========================================================
    % RADIAL GRID LEVELS
    %% ========================================================

    gridLevels = ...
        [0.25 0.50 0.75 1.00] * rMax;


    %% ========================================================
    % CIRCLE ANGLES
    %% ========================================================

    thetaGrid = ...
        linspace( ...
        0, ...
        2*pi, ...
        361);


    %% ========================================================
    % DRAW CONCENTRIC CIRCLES
    %% ========================================================

    for gg = 1:numel(gridLevels)


        rg = ...
            gridLevels(gg);


        xg = ...
            rg*cos(thetaGrid);


        yg = ...
            rg*sin(thetaGrid);


        plot( ...
            axRose, ...
            xg, ...
            yg, ...
            '-', ...
            'Color',[0.75 0.75 0.75], ...
            'LineWidth',0.6);


    end


    %% ========================================================
    % MAIN ANGULAR LINES
    %% ========================================================

    mainAngles = ...
        0:45:315;


    for aa = 1:numel(mainAngles)


        theta = ...
            mainAngles(aa) * pi / 180;


        xLine = ...
            [-rMax rMax] * cos(theta);


        yLine = ...
            [-rMax rMax] * sin(theta);


        plot( ...
            axRose, ...
            xLine, ...
            yLine, ...
            '-', ...
            'Color',[0.82 0.82 0.82], ...
            'LineWidth',0.6);


    end


    %% ========================================================
    % DRAW 72 ROSE WEDGES
    %% ========================================================

    for b = 1:N_ROSE_BINS


        r = ...
            roseProbability(b);


        if ~isfinite(r) || r <= 0

            continue;

        end


        theta1 = ...
            roseEdgesRad(b);


        theta2 = ...
            roseEdgesRad(b+1);


        thetaPatch = ...
            linspace( ...
            theta1, ...
            theta2, ...
            12);


        rPatch = ...
            r * ones(size(thetaPatch));


        %% ----------------------------------------------------
        % POLAR TO CARTESIAN
        %% ----------------------------------------------------

        xPatch = ...
            [ ...
            0, ...
            rPatch .* cos(thetaPatch), ...
            0];


        yPatch = ...
            [ ...
            0, ...
            rPatch .* sin(thetaPatch), ...
            0];


        %% ----------------------------------------------------
        % DRAW WEDGE
        %
        % NO FaceAlpha:
        % better EPS/painters compatibility.
        %% ----------------------------------------------------

        patch( ...
            axRose, ...
            xPatch, ...
            yPatch, ...
            barColor, ...
            'EdgeColor','none');


    end


    %% ========================================================
    % OUTER CIRCLE
    %% ========================================================

    xOuter = ...
        rMax*cos(thetaGrid);


    yOuter = ...
        rMax*sin(thetaGrid);


    plot( ...
        axRose, ...
        xOuter, ...
        yOuter, ...
        'k-', ...
        'LineWidth',0.8);


    %% ========================================================
    % ANGLE LABELS
    %
    % Extra room outside the outer circle.
    %% ========================================================

    labelRadius = ...
        1.12*rMax;


    labelText = { ...
        '0°  Forward', ...
        '45°', ...
        '90°', ...
        '135°', ...
        '180°  Reverse', ...
        '225°', ...
        '270°', ...
        '315°'};


    for aa = 1:numel(mainAngles)


        theta = ...
            mainAngles(aa) * pi / 180;


        xText = ...
            labelRadius*cos(theta);


        yText = ...
            labelRadius*sin(theta);


        text( ...
            axRose, ...
            xText, ...
            yText, ...
            labelText{aa}, ...
            'HorizontalAlignment','center', ...
            'VerticalAlignment','middle', ...
            'FontSize',10, ...
            'FontName','Arial');


    end


    %% ========================================================
    % RADIAL SCALE LABELS
    %% ========================================================

    for gg = 1:numel(gridLevels)


        rg = ...
            gridLevels(gg);


        text( ...
            axRose, ...
            rg, ...
            0, ...
            sprintf(' %.3f',rg), ...
            'FontSize',8, ...
            'FontName','Arial', ...
            'HorizontalAlignment','left', ...
            'VerticalAlignment','bottom');


    end


    %% ========================================================
    % CENTER POINT
    %% ========================================================

    plot( ...
        axRose, ...
        0, ...
        0, ...
        'k.', ...
        'MarkerSize',8);


    %% ========================================================
    % TITLE
    %% ========================================================

    titleText = sprintf( ...
        '%s   lag %d\nMean cos = %+0.3f    Within \\pm30^\\circ = %.1f%%', ...
        strrep(excelSheet,'_','\_'), ...
        LAG, ...
        meanCos, ...
        coneVals(1));


    title( ...
        axRose, ...
        titleText, ...
        'FontSize',11, ...
        'FontName','Arial');


    %% ========================================================
    % AXIS LIMIT
    %
    % Additional space outside radial labels.
    %% ========================================================

    limValue = ...
        1.25*rMax;


    xlim( ...
        axRose, ...
        [-limValue limValue]);


    ylim( ...
        axRose, ...
        [-limValue limValue]);


    %% ========================================================
    % EPS OUTPUT
    %% ========================================================

    roseEPS = ...
        fullfile( ...
        outputDir, ...
        sprintf( ...
        '%s_rose_lag%d.eps', ...
        excelSheet, ...
        LAG));


    fprintf( ...
        '    Writing rose EPS...\n');


    print( ...
        figRose, ...
        '-depsc2', ...
        '-painters', ...
        roseEPS);


    %% ========================================================
    % PNG OUTPUT
    %% ========================================================

    rosePNG = ...
        fullfile( ...
        outputDir, ...
        sprintf( ...
        '%s_rose_lag%d.png', ...
        excelSheet, ...
        LAG));


    fprintf( ...
        '    Writing rose PNG...\n');


    try

        exportgraphics( ...
            figRose, ...
            rosePNG, ...
            'Resolution',300);

    catch

        print( ...
            figRose, ...
            '-dpng', ...
            '-r300', ...
            rosePNG);

    end


    %% ========================================================
    % CLOSE ROSE FIGURE
    %% ========================================================

    close(figRose);


    fprintf( ...
        '  FAST rose plot finished: %.2f s\n', ...
        toc(roseTimer));


    %% ========================================================
    % WRITE ANGLE EXCEL
    %% ========================================================

    fprintf('\n');
    fprintf('  Writing angle Excel...\n');


    writeTimer = tic;


    nOutputSheets = ...
        ceil( ...
        nRows / ...
        MAX_DATA_ROWS_PER_SHEET);


    if nOutputSheets < 1

        nOutputSheets = 1;

    end


    fprintf( ...
        '  Number of output sheets: %d\n', ...
        nOutputSheets);


    %% ========================================================
    % EXCEL CHUNKS
    %% ========================================================

    for part = 1:nOutputSheets


        firstRow = ...
            (part-1) * ...
            MAX_DATA_ROWS_PER_SHEET + 1;


        lastRow = ...
            min( ...
            part * ...
            MAX_DATA_ROWS_PER_SHEET, ...
            nRows);


        %% ====================================================
        % OUTPUT SHEET NAME
        %% ====================================================

        if strcmp( ...
                excelSheet, ...
                'AllQualified_Directed_Part1')


            outputSheet = ...
                sprintf( ...
                'AllQualified_Directed_Part%d', ...
                part);


        elseif nOutputSheets == 1


            outputSheet = ...
                excelSheet;


        else


            outputSheet = ...
                sprintf( ...
                '%s_Part%d', ...
                excelSheet, ...
                part);


        end


        %% ====================================================
        % CLEAN ILLEGAL CHARACTERS
        %% ====================================================

        outputSheet = ...
            strrep(outputSheet,':','_');


        outputSheet = ...
            strrep(outputSheet,'\','_');


        outputSheet = ...
            strrep(outputSheet,'/','_');


        outputSheet = ...
            strrep(outputSheet,'?','_');


        outputSheet = ...
            strrep(outputSheet,'*','_');


        outputSheet = ...
            strrep(outputSheet,'[','_');


        outputSheet = ...
            strrep(outputSheet,']','_');


        %% ====================================================
        % MAX 31 CHARACTERS
        %% ====================================================

        if length(outputSheet) > 31

            outputSheet = ...
                outputSheet(1:31);

        end


        fprintf( ...
            '    Writing %s: rows %d-%d\n', ...
            outputSheet, ...
            firstRow, ...
            lastRow);


        %% ====================================================
        % OUTPUT TABLE
        %% ====================================================

        if WRITE_FULL_TABLE


            data.TurningAngle_deg = ...
                angleOriginal;


            outputChunk = ...
                data(firstRow:lastRow,:);


        else


            outputChunk = ...
                table( ...
                data.Sheet(firstRow:lastRow), ...
                data.mRNA_ID(firstRow:lastRow), ...
                data.X(firstRow:lastRow), ...
                data.Y(firstRow:lastRow), ...
                data.Time(firstRow:lastRow), ...
                angleOriginal(firstRow:lastRow), ...
                'VariableNames', ...
                { ...
                'Sheet', ...
                'mRNA_ID', ...
                'X', ...
                'Y', ...
                'Time', ...
                'TurningAngle_deg'});

        end


        %% ====================================================
        % WRITE
        %% ====================================================

        try

            writetable( ...
                outputChunk, ...
                angleFile, ...
                'Sheet', ...
                outputSheet);


        catch ME

            error( ...
                ['Excel writing failed.\n', ...
                 'Sheet: %s\n', ...
                 'Rows: %d-%d\n', ...
                 'Reason: %s'], ...
                outputSheet, ...
                firstRow, ...
                lastRow, ...
                ME.message);

        end


        clear outputChunk;


    end


    fprintf( ...
        '  Excel writing finished: %.2f s\n', ...
        toc(writeTimer));


    %% ========================================================
    % RELEASE MEMORY
    %% ========================================================

    clear data;

    clear Xsorted;
    clear Ysorted;
    clear originalRowSorted;

    clear groupStart;
    clear groupEnd;

    clear trackLength;
    clear validLength;
    clear validTrackIndices;

    clear angleOriginal;

    clear trackCos;
    clear trackSin;
    clear nAngPerTrk;

    clear roseEdgesDeg;
    clear roseCentersDeg;
    clear roseEdgesRad;
    clear roseCounts;
    clear roseProbability;

    clear foldedHist;


    %% ========================================================
    % GROUP FINISHED
    %% ========================================================

    fprintf('\n');
    fprintf( ...
        'GROUP %s FINISHED: %.2f s\n', ...
        excelSheet, ...
        toc(groupTimer));


end


%% ============================================================
% FINAL FOLDED ANGLE PLOT
% ============================================================

fprintf('\n');
fprintf('============================================================\n');
fprintf('Creating final folded-angle plot...\n');
fprintf('============================================================\n');


fn = ...
    fieldnames(foldedHistAll);


if ~isempty(fn)


    figFolded = ...
        figure( ...
        'Visible','off', ...
        'Color','w', ...
        'Position',[100 100 660 430]);


    axFolded = ...
        axes( ...
        'Parent',figFolded);


    hold(axFolded,'on');


    for k = 1:numel(fn)


        histCount = ...
            foldedHistAll.(fn{k});


        totalN = ...
            foldedNAll.(fn{k});


        if totalN > 0

            histPct = ...
                100 * histCount / totalN;

        else

            histPct = ...
                zeros(size(histCount));

        end


        thisColor = ...
            plotColors( ...
            min(k,size(plotColors,1)),:);


        plot( ...
            axFolded, ...
            FOLDED_CENTERS, ...
            histPct, ...
            '-o', ...
            'LineWidth',1.8, ...
            'MarkerSize',5, ...
            'Color',thisColor, ...
            'MarkerFaceColor',thisColor, ...
            'DisplayName', ...
            strrep( ...
            fn{k}, ...
            '_', ...
            '\_'));

    end


    %% ========================================================
    % UNIFORM REFERENCE
    %% ========================================================

    yline( ...
        axFolded, ...
        100/numel(FOLDED_CENTERS), ...
        'k--', ...
        'Uniform', ...
        'LineWidth',1);


    %% ========================================================
    % AXES
    %% ========================================================

    xlim( ...
        axFolded, ...
        [0 180]);


    xticks( ...
        axFolded, ...
        0:45:180);


    xlabel( ...
        axFolded, ...
        '|Turning angle| (deg)', ...
        'FontName','Arial');


    ylabel( ...
        axFolded, ...
        '% of angles', ...
        'FontName','Arial');


    title( ...
        axFolded, ...
        sprintf( ...
        'Folded turning-angle distribution, lag %d', ...
        LAG), ...
        'FontName','Arial');


    legend( ...
        axFolded, ...
        'Location','northeast', ...
        'Box','off');


    set( ...
        axFolded, ...
        'FontName','Arial', ...
        'FontSize',11, ...
        'Box','off');


    %% ========================================================
    % EPS
    %% ========================================================

    foldedEPS = ...
        fullfile( ...
        outputDir, ...
        sprintf( ...
        'Folded_lag%d.eps', ...
        LAG));


    print( ...
        figFolded, ...
        '-depsc2', ...
        '-painters', ...
        foldedEPS);


    %% ========================================================
    % PNG
    %% ========================================================

    foldedPNG = ...
        fullfile( ...
        outputDir, ...
        sprintf( ...
        'Folded_lag%d.png', ...
        LAG));


    try

        exportgraphics( ...
            figFolded, ...
            foldedPNG, ...
            'Resolution',300);

    catch

        print( ...
            figFolded, ...
            '-dpng', ...
            '-r300', ...
            foldedPNG);

    end


    close(figFolded);


end


%% ============================================================
% WRITE STATISTICS WORKBOOK
% ============================================================

fprintf('\n');
fprintf('============================================================\n');
fprintf('Writing statistics workbook...\n');
fprintf('============================================================\n');


if ~isempty(statRows)


    %% ========================================================
    % STATISTICS TABLE
    %% ========================================================

    statTable = ...
        cell2table( ...
        statRows, ...
        'VariableNames', ...
        { ...
        'ExcelWorksheet', ...
        'N_tracks', ...
        'N_angles', ...
        'Angles_per_track', ...
        'MeanCos', ...
        'R_pooled', ...
        'MeanDir_pooled_deg', ...
        'R_perTrack', ...
        'MeanDir_perTrack_deg', ...
        'Rayleigh_Z', ...
        'Rayleigh_p', ...
        'V_u', ...
        'V_p'});


    %% ========================================================
    % CONE TABLE
    %% ========================================================

    coneTable = ...
        cell2table( ...
        coneRows, ...
        'VariableNames', ...
        { ...
        'ExcelWorksheet', ...
        'N_tracks', ...
        'N_angles', ...
        'Within_30deg_pct', ...
        'Within_45deg_pct', ...
        'Within_60deg_pct', ...
        'Within_90deg_pct'});


    %% ========================================================
    % UNIFORM EXPECTATION
    %% ========================================================

    uniformTable = ...
        table( ...
        CONES(:), ...
        100*CONES(:)/180, ...
        'VariableNames', ...
        { ...
        'Cone_half_angle_deg', ...
        'Uniform_expectation_pct'});


    %% ========================================================
    % WRITE
    %% ========================================================

    writetable( ...
        statTable, ...
        statsFile, ...
        'Sheet','Statistics');


    writetable( ...
        coneTable, ...
        statsFile, ...
        'Sheet','ForwardCones');


    writetable( ...
        uniformTable, ...
        statsFile, ...
        'Sheet','UniformExpectation');


end


%% ============================================================
% FINAL CLEANUP
% ============================================================

clear foldedHistAll;
clear foldedNAll;
clear statRows;
clear coneRows;


%% ============================================================
% FINAL MESSAGE
% ============================================================

fprintf('\n');
fprintf('============================================================\n');
fprintf(' ANALYSIS FINISHED SUCCESSFULLY\n');
fprintf('============================================================\n');


fprintf('\n');
fprintf('Angle workbook:\n');
fprintf('  %s\n',angleFile);


fprintf('\n');
fprintf('Statistics workbook:\n');
fprintf('  %s\n',statsFile);


fprintf('\n');
fprintf('Output directory:\n');
fprintf('  %s\n',outputDir);


fprintf('\n');
fprintf('Generated files:\n');
fprintf('  Angles_Lag%d.xlsx\n',LAG);
fprintf('  Stats_Lag%d.xlsx\n',LAG);
fprintf('  *_rose_lag%d.eps\n',LAG);
fprintf('  *_rose_lag%d.png\n',LAG);
fprintf('  Folded_lag%d.eps\n',LAG);
fprintf('  Folded_lag%d.png\n',LAG);


fprintf('\n');
fprintf('============================================================\n');
fprintf(' ROSE PLOT FINAL SETTINGS\n');
fprintf('============================================================\n');
fprintf(' polarhistogram       : REMOVED\n');
fprintf(' polaraxes patch      : REMOVED\n');
fprintf(' Rose bins            : %d\n',N_ROSE_BINS);
fprintf(' Maximum data radius  : actual maximum probability\n');
fprintf(' Outer radius         : %.2fx maximum data\n',ROSE_MARGIN_FACTOR);
fprintf(' Outer empty margin   : 20%%\n');
fprintf(' FaceAlpha             : OFF\n');
fprintf(' EPS renderer          : painters\n');
fprintf('============================================================\n');