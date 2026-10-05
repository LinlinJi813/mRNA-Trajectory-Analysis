% motor_traces_V9.m
% Kinesin trajectory generator (stepping model from Ferro et al., eLife 2019).
%
% Model
%   Before each step the motor dwells. With probability 1 - percent the dwell
%   is a pause (mean dwell_time2), otherwise a normal dwell (mean dwell_time1);
%   both are exponential. Each step moves step_size nm along the track.
%   During a pause the particle jitters: discrete steps of jit_s nm in random
%   directions at exponential intervals (mean jit_tau), confined to a disc of
%   radius jit_R nm (reflected at the edge). The offset is kept after the pause.
%   Position is held between events (stepwise). Frames sample the position at
%   0, frame_time, 2*frame_time, ...; Gaussian localization error is added in
%   camera coordinates. Each trace has a random direction (0-360 deg) and
%   starts at the origin.
%
% Length (number of frames per trace)
%   'fixed'  : every trace has n_frames frames.
%   'sample' : the number of frames is drawn with replacement from
%              frame_counts (e.g. frame numbers of real mRNA tracks).
%
% Output
%   CSV with columns Time (s), mRNA_ID, x (um), y (um),
%   readable by classify_trajectories.py.

clear;

%% Parameters
n_traces     = 2000;
seed = 'shuffle';

percent      = 0.6;      % fraction of dwells that are not pauses
dwell_time1  = 0.016241;      % s, mean normal dwell
dwell_time2  = 0.1;       % s, mean pause
step_size    = 60;          % nm

jit_R        = 100;         % nm, confinement radius during pauses (0 = static)
jit_s        = 60;          % nm, jitter step
jit_tau      = 0.016241;      % s, mean interval between jitter steps

frame_time   = 0.016241;   % s
sigma_x      = 15;         % nm
sigma_y      = 13.2;       % nm

length_mode  = 'fixed';    % 'fixed' or 'sample'
n_frames_fix = 45;         % 'fixed': frames per trace
frame_counts = [];         % 'sample': e.g. readmatrix('frame_counts.txt')

outfile      = 'motor_traces_2.csv';

%% Simulation
rng(seed);
if strcmp(length_mode, 'sample') && isempty(frame_counts)
    error('frame_counts is empty');
end

fid = fopen(outfile, 'w');
fprintf(fid, 'Time,mRNA_ID,x,y\n');

for tr = 1:n_traces
    if strcmp(length_mode, 'sample')
        n_frames = frame_counts(randi(numel(frame_counts)));
    else
        n_frames = n_frames_fix;
    end
    t_need = (n_frames - 1) * frame_time;       % time of the last frame

    % events: time, axial position, lateral position (nm); event 1 = origin
    cap = 4096;
    et = zeros(cap, 1); eu = et; ev = et;
    ne = 1;
    t = 0; u = 0; a = 0; b = 0;     % u: track position, (a, b): jitter offset

    while t < t_need
        if rand < 1 - percent                          % pause
            t1 = t - dwell_time2 * log(rand);
            if jit_R > 0 && jit_s > 0
                tj = t - jit_tau * log(rand);
                while tj < t1 && tj < t_need
                    ang = 2 * pi * rand;
                    na = a + jit_s * cos(ang);
                    nb = b + jit_s * sin(ang);
                    r = hypot(na, nb);
                    if r > jit_R                        % reflect into the disc
                        f = (2 * jit_R - r) / r;
                        na = na * f; nb = nb * f;
                        r2 = hypot(na, nb);
                        if r2 > jit_R
                            na = na * jit_R / r2; nb = nb * jit_R / r2;
                        end
                    end
                    a = na; b = nb;
                    ne = ne + 1;
                    if ne > cap, cap = 2 * cap; et(cap) = 0; eu(cap) = 0; ev(cap) = 0; end
                    et(ne) = tj; eu(ne) = u + a; ev(ne) = b;
                    tj = tj - jit_tau * log(rand);
                end
            end
            t = t1;
        else                                           % normal dwell
            t = t - dwell_time1 * log(rand);
        end
        u = u + step_size;                              % step
        ne = ne + 1;
        if ne > cap, cap = 2 * cap; et(cap) = 0; eu(cap) = 0; ev(cap) = 0; end
        et(ne) = t; eu(ne) = u + a; ev(ne) = b;
    end

    % stepwise sampling: last event at or before each frame time
    tf = (0:n_frames - 1)' * frame_time;
    idx = ones(n_frames, 1); k = 1;
    for i = 1:n_frames
        while k < ne && et(k + 1) <= tf(i) + 1e-12
            k = k + 1;
        end
        idx(i) = k;
    end
    uf = eu(idx); vf = ev(idx);

    th = 2 * pi * rand;
    x = (uf * cos(th) - vf * sin(th) + sigma_x * randn(n_frames, 1)) / 1000;
    y = (uf * sin(th) + vf * cos(th) + sigma_y * randn(n_frames, 1)) / 1000;

    fprintf(fid, '%.6f,%d,%.6f,%.6f\n', [tf, repmat(tr, n_frames, 1), x, y]');
end

fclose(fid);
fprintf('%d traces written to %s\n', n_traces, outfile);
