%% ============================================================
%  mRNA MOTION CLASSIFICATION - MULTI-SHEET
%  FULL-LENGTH VERSION (no track-length filter)
%
%  Window metrics (per 20-frame sliding window)
%     Method 1 : r2(tau) from window start, lags 1 .. w-1
%     Method 2 : time-averaged MSD(tau),    lags 1 .. w-1
%     Method 3 : time-averaged MSD(tau),    lags 1 .. floor(w/2)
%     Method 4 : time-averaged MSD(tau),    lags 1 .. floor(w/3)
%
%  Classification
%     Static          : D_track <= 0.01  OR  netdisp2 <= 0.15
%     Non-static      : every track with at least one full
%                       20-frame window is classified; there is
%                       no track-length filter in this version.
%     Superdiffusion  : at least one window with
%                       alpha_win > 1.25, R2_win >= 0.75,
%                       MSD/tau > 0.36
%     Subdiffusion    : alpha_track < 0.75
%     Diffusion       : otherwise
%
%  Localization error
%     D_track    : fitted from the OBSERVED MSD with a free
%                  intercept, MSD = 4*D*tau + c. The intercept
%                  absorbs 2*(sigma_x^2 + sigma_y^2); the error
%                  is NOT subtracted beforehand.
%     alpha      : log-log fit of the CORRECTED curve,
%                  MSD - 2*(sigma_x^2 + sigma_y^2), because a
%                  log-log fit has no intercept term that could
%                  absorb a constant offset.
%
%  Also exports, for methods 1 and 2, the coordinates of one
%  representative 20-frame window per classified track.
%
%  NOTE ON TRACK LENGTH
%     A track is called Superdiffusive if ANY window qualifies,
%     so the number of windows, and therefore the chance of a
%     chance hit, grows with track length. Without a length
%     filter the classification of long tracks is dominated by
%     the number of windows tested. Compare results against a
%     null generated with the same track-length distribution.
%
%  MSD/tau definition in this file: median of MSD(tau)/tau across all lags
% ============================================================

clear; clc;

%% ============================================================
%  CONFIGURATION
% ============================================================

DISP_OVER_T_MODE = 'median';   % 'median' in this file

OUTPUT_FOLDER_NAME = '1a1.25fulllength_classification_median';

FRAME_INTERVAL     = 0.016241;   % s
WINDOW             = 20;         % frames
TRACK_LAG_FRACTION = 0.5;        % full-track MSD uses lags 1 .. n*this

ALPHA_TH           = 1.25;       % window exponent for a directed call
R2_TH              = 0.75;       % window log-log fit quality
DISP_OVER_T_TH     = 0.36;       % um^2/s, equals 4D at D = 0.09

D_STATIC_TH        = 0.01;       % um^2/s
NETDISP2_TH        = 0.15;       % um^2
ALPHA_TRACK_TH     = 0.75;       % below this a motile track is subdiffusive

SIGMA_X            = 0.015;      % um
SIGMA_Y            = 0.0132;     % um

N_METHODS          = 4;
EXPORT_METHODS     = [1 2 3 4];      % export selected windows for these methods

LOC_OFFSET = 2*(SIGMA_X^2 + SIGMA_Y^2);

fprintf('MSD/tau mode        : %s\n', DISP_OVER_T_MODE);
fprintf('Frame filter        : none, all track lengths analysed\n');
fprintf('Localization offset : %.3e um^2\n', LOC_OFFSET);

%% ============================================================
%  SELECT INPUT FILE
% ============================================================

[fileName, filePath] = uigetfile( ...
    {'*.xlsx;*.xls','Excel files'}, 'Select trajectory file');
if isequal(fileName,0)
    error('No file selected.');
end
inputFile = fullfile(filePath, fileName);
fprintf('Input file          : %s\n', inputFile);

sheetNames = sheetnames(inputFile);
fprintf('Sheets found        : %d\n\n', numel(sheetNames));

outputRoot = fullfile(filePath, OUTPUT_FOLDER_NAME);
if ~exist(outputRoot,'dir')
    mkdir(outputRoot);
end

%% ============================================================
%  PROCESS EVERY SHEET
% ============================================================

allTracks   = table();
allExcluded = zeros(numel(sheetNames),1);

% selected 20-frame windows, one table per exported method and class
selWin = struct();
for mm = EXPORT_METHODS
    selWin.(sprintf('m%d',mm)).Diffusion      = table();
    selWin.(sprintf('m%d',mm)).Subdiffusion   = table();
    selWin.(sprintf('m%d',mm)).Superdiffusion = table();
end

for s = 1:numel(sheetNames)

    sheet = sheetNames{s};
    fprintf('[%d/%d] %s ... ', s, numel(sheetNames), sheet);

    try
        T = readtable(inputFile,'Sheet',sheet);
    catch
        fprintf('unreadable, skipped\n');
        continue;
    end
    if isempty(T)
        fprintf('empty, skipped\n');
        continue;
    end

    [ok, colTime, colID, colX, colY] = findColumns(T);
    if ~ok
        fprintf('missing columns, skipped\n');
        continue;
    end

    [tracks, nExcluded, selWin] = classifySheet( ...
        T, colTime, colID, colX, colY, sheet, ...
        FRAME_INTERVAL, WINDOW, TRACK_LAG_FRACTION, ...
        ALPHA_TH, R2_TH, DISP_OVER_T_TH, ...
        D_STATIC_TH, NETDISP2_TH, ALPHA_TRACK_TH, ...
        LOC_OFFSET, N_METHODS, DISP_OVER_T_MODE, ...
        EXPORT_METHODS, selWin);

    allExcluded(s) = nExcluded;

    if ~isempty(tracks)
        allTracks = [allTracks; tracks];   %#ok<AGROW>
    end

    fprintf('%d tracks kept, %d shorter than one window\n', ...
            height(tracks)/N_METHODS, nExcluded);
end

if isempty(allTracks)
    error('No tracks were classified.');
end

%% ============================================================
%  SUMMARY PER SHEET AND METHOD
% ============================================================

sheetsUsed = unique(allTracks.Sheet,'stable');
rows = {};

for m = 1:N_METHODS
    for k = 1:numel(sheetsUsed)

        sel = strcmp(allTracks.Sheet, sheetsUsed{k}) & allTracks.Method == m;
        cls = allTracks.motion_class(sel);

        nSuper  = sum(strcmp(cls,'Superdiffusion'));
        nDiff   = sum(strcmp(cls,'Diffusion'));
        nSub    = sum(strcmp(cls,'Subdiffusion'));
        nStatic = sum(strcmp(cls,'Static'));
        nTotal  = numel(cls);
        nMoving = nSuper + nDiff + nSub;

        rows(end+1,:) = { ...
            sheetsUsed{k}, m, nTotal, nMoving, ...
            nSuper, nDiff, nSub, nStatic, ...
            pct(nSuper,nMoving), pct(nDiff,nMoving), pct(nSub,nMoving), ...
            pct(nSuper,nTotal),  pct(nDiff,nTotal),  pct(nSub,nTotal), ...
            pct(nStatic,nTotal) };   %#ok<AGROW>
    end
end

summaryTable = cell2table(rows, 'VariableNames', { ...
    'Sheet','Method','Total_mRNA','Total_Moving', ...
    'Superdiffusion_n','Diffusion_n','Subdiffusion_n','Static_n', ...
    'Directed_Moving_percent','Diffusive_Moving_percent', ...
    'Subdiffusive_Moving_percent', ...
    'Superdiffusion_percent','Diffusion_percent', ...
    'Subdiffusion_percent','Static_percent'});

%% ============================================================
%  POOLED SUMMARY OVER ALL SHEETS
% ============================================================

rows = {};
for m = 1:N_METHODS

    cls = allTracks.motion_class(allTracks.Method == m);

    nSuper  = sum(strcmp(cls,'Superdiffusion'));
    nDiff   = sum(strcmp(cls,'Diffusion'));
    nSub    = sum(strcmp(cls,'Subdiffusion'));
    nStatic = sum(strcmp(cls,'Static'));
    nTotal  = numel(cls);
    nMoving = nSuper + nDiff + nSub;

    rows(end+1,:) = { ...
        m, nTotal, nMoving, nSuper, nDiff, nSub, nStatic, ...
        pct(nSuper,nMoving), pct(nDiff,nMoving), pct(nSub,nMoving), ...
        pct(nSuper,nTotal),  pct(nDiff,nTotal),  pct(nSub,nTotal), ...
        pct(nStatic,nTotal) };   %#ok<AGROW>

    fprintf('\nMethod %d  (n = %d, moving = %d)\n', m, nTotal, nMoving);
    fprintf('  Superdiffusion %6.2f %% of moving, %6.2f %% of all\n', ...
            pct(nSuper,nMoving), pct(nSuper,nTotal));
    fprintf('  Diffusion      %6.2f %% of moving, %6.2f %% of all\n', ...
            pct(nDiff,nMoving),  pct(nDiff,nTotal));
    fprintf('  Subdiffusion   %6.2f %% of moving, %6.2f %% of all\n', ...
            pct(nSub,nMoving),   pct(nSub,nTotal));
    fprintf('  Static                            %6.2f %% of all\n', ...
            pct(nStatic,nTotal));
end

totalTable = cell2table(rows, 'VariableNames', { ...
    'Method','Total_mRNA','Total_Moving', ...
    'Superdiffusion_n','Diffusion_n','Subdiffusion_n','Static_n', ...
    'Directed_Moving_percent','Diffusive_Moving_percent', ...
    'Subdiffusive_Moving_percent', ...
    'Superdiffusion_percent','Diffusion_percent', ...
    'Subdiffusion_percent','Static_percent'});

%% ============================================================
%  WRITE OUTPUT
% ============================================================

fSummary = fullfile(outputRoot,'Summary.xlsx');
fTracks  = fullfile(outputRoot,'All_tracks_all_methods.xlsx');

if exist(fSummary,'file'); delete(fSummary); end
if exist(fTracks,'file');  delete(fTracks);  end

writetable(totalTable,   fSummary, 'Sheet','Pooled');
writetable(summaryTable, fSummary, 'Sheet','PerSheet');
writetable(allTracks,    fTracks,  'Sheet','AllMethods');

for m = 1:N_METHODS
    writetable(allTracks(allTracks.Method == m,:), ...
               fTracks, 'Sheet', sprintf('Method%d',m));
end

for mm = EXPORT_METHODS

    fSel = fullfile(outputRoot, ...
                    sprintf('Method%d_Selected_20Frames.xlsx', mm));
    if exist(fSel,'file'); delete(fSel); end

    key     = sprintf('m%d', mm);
    written = false;

    classList = {'Diffusion','Subdiffusion','Superdiffusion'};
    sheetList = {'Normaldiffusive','Subdiffusive','Superdiffusive'};

    for cc = 1:numel(classList)
        Tsel = selWin.(key).(classList{cc});
        if ~isempty(Tsel)
            writetable(Tsel, fSel, 'Sheet', sheetList{cc});
            written = true;
            fprintf('Method %d, %-15s : %d tracks\n', ...
                    mm, sheetList{cc}, height(Tsel)/WINDOW);
        end
    end

    if written
        fprintf('Saved: %s\n', fSel);
    end
end

fprintf('\nSaved:\n  %s\n  %s\n', fSummary, fTracks);
fprintf('Excluded, shorter than one %d-frame window: %d\n', ...
        WINDOW, sum(allExcluded));

%% ============================================================
%  LOCAL FUNCTIONS
% ============================================================

function p = pct(a,b)
% Percentage with a guard against division by zero.
    if b > 0
        p = 100*a/b;
    else
        p = NaN;
    end
end


function [ok, cTime, cID, cX, cY] = findColumns(T)
% Locate the four required columns, ignoring case and spaces.
    names = lower(strrep(T.Properties.VariableNames,' ',''));
    cTime = pick(names,{'time','t','frame','postion_t','position_t'});
    cID   = pick(names,{'mrna_id','mrnaid','track_id','trackid','id'});
    cX    = pick(names,{'x','position_x','posx'});
    cY    = pick(names,{'y','position_y','posy'});
    ok    = all([cTime cID cX cY] > 0);
end


function idx = pick(names, candidates)
    idx = 0;
    for i = 1:numel(candidates)
        hit = find(strcmp(names, candidates{i}), 1);
        if ~isempty(hit)
            idx = hit;
            return;
        end
    end
end


function [out, nExcluded, selWin] = classifySheet( ...
    T, cTime, cID, cX, cY, sheet, ...
    dt, window, lagFrac, ...
    alphaTh, r2Th, dotTh, ...
    dStaticTh, netdisp2Th, alphaTrackTh, ...
    locOffset, nMethods, dotMode, ...
    exportMethods, selWin)
% Classify every track of one sheet under all four methods.

    ids  = string(table2array(T(:,cID)));
    time = table2array(T(:,cTime));
    xy   = [table2array(T(:,cX)), table2array(T(:,cY))];

    uid       = unique(ids,'stable');
    nExcluded = 0;

    sheetCol = {}; idCol = {}; methCol = []; nfCol = [];
    clsCol   = {}; aTCol = []; dTCol = []; offCol = [];
    ndCol    = []; nwCol = []; nqCol = []; maxACol = [];

    for u = 1:numel(uid)

        sel = (ids == uid(u));
        [~,order] = sort(time(sel));
        P = xy(sel,:);
        P = P(order,:);
        n = size(P,1);

        if n < 3
            continue;
        end

        [alphaT, Dtrack, offsetT, netd2] = ...
            trackMetrics(P, dt, lagFrac, locOffset);

        isStatic = (isfinite(Dtrack) && Dtrack <= dStaticTh) || ...
                   (isfinite(netd2)  && netd2  <= netdisp2Th);

        % --- static: kept regardless of track length ---------
        if isStatic
            for m = 1:nMethods
                sheetCol{end+1,1} = sheet;      %#ok<AGROW>
                idCol{end+1,1}    = char(uid(u)); %#ok<AGROW>
                methCol(end+1,1)  = m;          %#ok<AGROW>
                nfCol(end+1,1)    = n;          %#ok<AGROW>
                clsCol{end+1,1}   = 'Static';   %#ok<AGROW>
                aTCol(end+1,1)    = alphaT;     %#ok<AGROW>
                dTCol(end+1,1)    = Dtrack;     %#ok<AGROW>
                offCol(end+1,1)   = offsetT;    %#ok<AGROW>
                ndCol(end+1,1)    = netd2;      %#ok<AGROW>
                nwCol(end+1,1)    = 0;          %#ok<AGROW>
                nqCol(end+1,1)    = 0;          %#ok<AGROW>
                maxACol(end+1,1)  = NaN;        %#ok<AGROW>
            end
            continue;
        end

        % --- non-static: no length filter --------------------
        % Any non-static track with at least one full window is
        % classified, whatever its total length.

        nWin = n - window + 1;
        if nWin < 1
            nExcluded = nExcluded + 1;   % shorter than one window
            continue;
        end

        for m = 1:nMethods

            nQual = 0;
            bestA = NaN;

            for w = 1:nWin
                [aWin, r2Win, dotWin] = windowMetrics( ...
                    P(w:w+window-1,:), m, dt, locOffset, dotMode);
                if ~isfinite(aWin)
                    continue;
                end
                if isnan(bestA) || aWin > bestA
                    bestA = aWin;
                end
                if aWin > alphaTh && r2Win >= r2Th && dotWin > dotTh
                    nQual = nQual + 1;
                end
            end

            if nQual > 0
                cls = 'Superdiffusion';
            elseif isfinite(alphaT) && alphaT < alphaTrackTh
                cls = 'Subdiffusion';
            else
                cls = 'Diffusion';
            end

            if ismember(m, exportMethods)
                oneWin = selectWindows(P, uid(u), sheet, cls, m, dt, ...
                    window, alphaT, locOffset, alphaTh, r2Th, dotTh, dotMode);
                if ~isempty(oneWin)
                    key = sprintf('m%d', m);
                    selWin.(key).(cls) = [selWin.(key).(cls); oneWin];
                end
            end

            sheetCol{end+1,1} = sheet;          %#ok<AGROW>
            idCol{end+1,1}    = char(uid(u));   %#ok<AGROW>
            methCol(end+1,1)  = m;              %#ok<AGROW>
            nfCol(end+1,1)    = n;              %#ok<AGROW>
            clsCol{end+1,1}   = cls;            %#ok<AGROW>
            aTCol(end+1,1)    = alphaT;         %#ok<AGROW>
            dTCol(end+1,1)    = Dtrack;         %#ok<AGROW>
            offCol(end+1,1)   = offsetT;        %#ok<AGROW>
            ndCol(end+1,1)    = netd2;          %#ok<AGROW>
            nwCol(end+1,1)    = nWin;           %#ok<AGROW>
            nqCol(end+1,1)    = nQual;          %#ok<AGROW>
            maxACol(end+1,1)  = bestA;          %#ok<AGROW>
        end
    end

    out = table(sheetCol, idCol, methCol, nfCol, clsCol, ...
                aTCol, dTCol, offCol, ndCol, nwCol, nqCol, maxACol, ...
        'VariableNames', {'Sheet','mRNA_ID','Method','n_frames', ...
        'motion_class','alpha_track','D_track','offset_track', ...
        'netdisp2','n_windows','n_qualified','max_alpha_win'});
end


function [alphaTrack, Dtrack, offsetTrack, netdisp2] = ...
    trackMetrics(xy, dt, lagFrac, locOffset)
% Full-track metrics.
%
%   D_track : fitted from the OBSERVED MSD with a free intercept,
%             MSD = 4*D*tau + c. The intercept absorbs the
%             localization error, so the error is NOT subtracted
%             before this fit.
%
%   alpha   : log-log fit of the CORRECTED MSD. A log-log fit has
%             no intercept term to absorb a constant offset, so
%             the error must be removed first.

    n = size(xy,1);
    netdisp2 = sum((xy(end,:) - xy(1,:)).^2);

    nLag = floor(n*lagFrac);
    if nLag < 3
        alphaTrack = NaN; Dtrack = NaN; offsetTrack = NaN;
        return;
    end

    lags = (1:nLag)';
    tau  = lags*dt;

    msdObserved  = timeAveragedMSD(xy, lags);
    msdCorrected = msdObserved - locOffset;

    % alpha from the corrected curve
    alphaTrack = loglogFit(tau, msdCorrected);

    % D from the observed curve, intercept free
    ok = isfinite(msdObserved) & isfinite(tau);
    if sum(ok) >= 3
        p           = polyfit(tau(ok), msdObserved(ok), 1);
        Dtrack      = p(1)/4;
        offsetTrack = p(2);
    else
        Dtrack = NaN; offsetTrack = NaN;
    end
end


function [alphaWin, R2Win, dispOverT] = ...
    windowMetrics(xy, method, dt, locOffset, dotMode)
% Metrics for one sliding window.
%
%   Method 1 uses r2(tau) measured from the window start; it is a
%   single realization, not a time-averaged MSD.
%   Methods 2-4 use the time-averaged MSD over different lag ranges.
%
%   dispOverT is MSD(tau)/tau, either at the largest lag ('last')
%   or as the median across lags ('median').

    w = size(xy,1);

    switch method
        case 1
            lags  = (1:w-1)';
            d     = xy(2:end,:) - xy(1,:);
            curve = sum(d.^2,2);
        case 2
            nLag = w - 1;
            if nLag < 3
                alphaWin = NaN; R2Win = NaN; dispOverT = NaN; return;
            end
            lags  = (1:nLag)';
            curve = timeAveragedMSD(xy, lags);
        case 3
            nLag = floor(w/2);
            if nLag < 3
                alphaWin = NaN; R2Win = NaN; dispOverT = NaN; return;
            end
            lags  = (1:nLag)';
            curve = timeAveragedMSD(xy, lags);
        case 4
            nLag = floor(w/3);
            if nLag < 3
                alphaWin = NaN; R2Win = NaN; dispOverT = NaN; return;
            end
            lags  = (1:nLag)';
            curve = timeAveragedMSD(xy, lags);
        otherwise
            error('Invalid method.');
    end

    tau           = lags*dt;
    curveCorrected = curve - locOffset;

    [alphaWin, R2Win] = loglogFit(tau, curveCorrected);

    ok = isfinite(curveCorrected) & curveCorrected > 0 & ...
         isfinite(tau) & tau > 0;
    if ~any(ok)
        dispOverT = NaN;
        return;
    end

    ratio = curveCorrected(ok) ./ tau(ok);

    if strcmpi(dotMode,'median')
        dispOverT = median(ratio);
    else
        dispOverT = ratio(end);     % MSD(tau_max)/tau_max
    end
end


function msd = timeAveragedMSD(xy, lags)
% Time-averaged MSD at each requested lag.
    n   = size(xy,1);
    msd = nan(numel(lags),1);
    for i = 1:numel(lags)
        k = lags(i);
        if k >= n
            continue;
        end
        d      = xy(1+k:end,:) - xy(1:end-k,:);
        msd(i) = mean(sum(d.^2,2));
    end
end


function [slope, R2] = loglogFit(tau, curve)
% Slope and R^2 of log10(curve) against log10(tau).
    ok = isfinite(curve) & curve > 0 & isfinite(tau) & tau > 0;
    if sum(ok) < 3
        slope = NaN; R2 = NaN; return;
    end

    lt = log10(tau(ok));
    lc = log10(curve(ok));

    p     = polyfit(lt, lc, 1);
    slope = p(1);

    fitted = polyval(p, lt);
    SST    = sum((lc - mean(lc)).^2);
    if SST > 0
        R2 = 1 - sum((lc - fitted).^2)/SST;
    else
        R2 = NaN;
    end
end


function sel = selectWindows(P, uidStr, sheet, cls, method, dt, ...
                             window, alphaTrack, locOffset, ...
                             alphaTh, r2Th, dotTh, dotMode)
% Pick one 20-frame window per track and return its coordinates.
%
%   Superdiffusion : among the windows that satisfy all three
%                    criteria, take the one whose largest
%                    r2(tau) (squared displacement from the
%                    window start) is greatest. Alpha reported
%                    is that window's alpha_win.
%   Diffusion,
%   Subdiffusion   : take the first 20 frames. Alpha reported is
%                    alpha_track.
%
%   Time is reset to 0 .. (window-1)*dt.

    sel  = table();
    n    = size(P,1);
    if n < window
        return;
    end
    nWin = n - window + 1;

    if strcmp(cls,'Superdiffusion')

        bestScore = -inf;
        bestStart = 0;
        bestAlpha = NaN;

        for w = 1:nWin
            xyW = P(w:w+window-1,:);
            [aWin, r2Win, dotWin] = windowMetrics( ...
                xyW, method, dt, locOffset, dotMode);
            if ~isfinite(aWin) || ~isfinite(r2Win) || ~isfinite(dotWin)
                continue;
            end
            if ~(aWin > alphaTh && r2Win >= r2Th && dotWin > dotTh)
                continue;
            end
            % score: largest squared displacement from window start
            d     = xyW(2:end,:) - xyW(1,:);
            score = max(sum(d.^2,2));
            if isfinite(score) && score > bestScore
                bestScore = score;
                bestStart = w;
                bestAlpha = aWin;
            end
        end

        if bestStart == 0
            return;              % no qualifying window
        end
        startFrame = bestStart;
        alphaOut   = bestAlpha;

    else
        startFrame = 1;
        alphaOut   = alphaTrack;
    end

    idx = startFrame:startFrame+window-1;

    sel = table( ...
        (0:window-1)'*dt, ...
        repmat(string(sheet) + "_" + string(uidStr), window, 1), ...
        P(idx,1), ...
        P(idx,2), ...
        repmat(alphaOut, window, 1), ...
        repmat(startFrame, window, 1), ...
        'VariableNames', {'Time','mRNA_ID','x','y','Alpha','StartFrame'});
end
