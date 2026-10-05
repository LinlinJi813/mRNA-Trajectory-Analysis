#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
============================================================
2D Brownian Trajectory Simulator - Experimental Joint Mode
============================================================

唯一模式：
------------------------------------------------------------

Experimental D + Frame (Joint)

    D_track 和 n_frames 从实验文件的同一行联合抽样

    保持原始 D-frame 配对关系

    例如实验数据：

        D_track     n_frames
        0.031       35
        0.052       48
        0.081       72
        0.043       51

    如果随机抽到第 3 行：

        D = 0.081
        n_frames = 72

    二者始终来自同一行。

Sampling:
    - with replacement
    - 每条 trajectory 独立抽取一个实验行

Experimental file 支持：
    - XLSX
    - XLS
    - CSV
    - TXT
    - DAT

实验文件必须包含：
    D_track
    n_frames

输出：
    Excel
    CSV
    Tracks CSV
    PNG
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# Brownian trajectory generation
# ============================================================

def generate_brownian_trajectory(
    D,
    n_frames,
    dt,
    sigma_x=0.0,
    sigma_y=0.0,
    x0=0.0,
    y0=0.0
):

    if n_frames < 1:
        raise ValueError(
            "Number of frames must be >= 1."
        )

    if D < 0:
        raise ValueError(
            "D must be >= 0."
        )

    if dt <= 0:
        raise ValueError(
            "dt must be > 0."
        )

    # --------------------------------------------------------
    # Brownian displacement
    # --------------------------------------------------------

    if n_frames > 1:

        step_std = np.sqrt(
            2.0 * D * dt
        )

        dx = np.random.normal(
            loc=0.0,
            scale=step_std,
            size=n_frames - 1
        )

        dy = np.random.normal(
            loc=0.0,
            scale=step_std,
            size=n_frames - 1
        )

        x = np.empty(n_frames)
        y = np.empty(n_frames)

        x[0] = x0
        y[0] = y0

        x[1:] = (
            x0 +
            np.cumsum(dx)
        )

        y[1:] = (
            y0 +
            np.cumsum(dy)
        )

    else:

        x = np.array([x0])
        y = np.array([y0])

    # --------------------------------------------------------
    # Localization error
    # --------------------------------------------------------

    if sigma_x > 0:

        x = x + np.random.normal(
            loc=0.0,
            scale=sigma_x,
            size=n_frames
        )

    if sigma_y > 0:

        y = y + np.random.normal(
            loc=0.0,
            scale=sigma_y,
            size=n_frames
        )

    return x, y


# ============================================================
# Normalize column name
# ============================================================

def normalize_column_name(name):

    return (
        str(name)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("-", "_")
    )


# ============================================================
# Find column
# ============================================================

def find_column(df, possible_names):

    normalized_columns = {}

    for col in df.columns:

        normalized = normalize_column_name(col)

        normalized_columns[
            normalized
        ] = col

    # --------------------------------------------------------
    # Exact match
    # --------------------------------------------------------

    for name in possible_names:

        normalized_name = (
            normalize_column_name(name)
        )

        if normalized_name in normalized_columns:

            return normalized_columns[
                normalized_name
            ]

    # --------------------------------------------------------
    # Partial match
    # --------------------------------------------------------

    for col in df.columns:

        normalized = normalize_column_name(col)

        for name in possible_names:

            normalized_name = (
                normalize_column_name(name)
            )

            if normalized_name in normalized:

                return col

    return None


# ============================================================
# Load experimental D + Frame file
# ============================================================

def load_experimental_D_frame_file(filepath):

    if not filepath:

        raise ValueError(
            "No experimental file selected."
        )

    if not os.path.isfile(filepath):

        raise ValueError(
            f"File does not exist:\n{filepath}"
        )

    ext = os.path.splitext(
        filepath
    )[1].lower()

    # ========================================================
    # Excel
    # ========================================================

    if ext in [".xlsx", ".xls"]:

        try:

            excel_file = pd.ExcelFile(
                filepath
            )

            sheet_names = (
                excel_file.sheet_names
            )

            if len(sheet_names) == 0:

                raise ValueError(
                    "No sheets were found "
                    "in the Excel file."
                )

            sheet_name = sheet_names[0]

            df = pd.read_excel(
                filepath,
                sheet_name=sheet_name
            )

        except Exception as e:

            raise ValueError(
                "Failed to read Excel file.\n\n"
                f"{e}"
            )

    # ========================================================
    # CSV
    # ========================================================

    elif ext == ".csv":

        try:

            df = pd.read_csv(
                filepath
            )

        except Exception:

            try:

                df = pd.read_csv(
                    filepath,
                    header=None
                )

            except Exception as e:

                raise ValueError(
                    "Failed to read CSV file.\n\n"
                    f"{e}"
                )

        sheet_name = None

    # ========================================================
    # TXT / DAT
    # ========================================================

    elif ext in [".txt", ".dat"]:

        try:

            df = pd.read_csv(
                filepath,
                sep=r"[\s,\t]+",
                engine="python"
            )

        except Exception:

            try:

                df = pd.read_csv(
                    filepath,
                    header=None,
                    sep=r"[\s,\t]+",
                    engine="python"
                )

            except Exception as e:

                raise ValueError(
                    "Failed to read TXT/DAT file.\n\n"
                    f"{e}"
                )

        sheet_name = None

    else:

        raise ValueError(
            "Unsupported file format.\n\n"
            "Supported formats:\n"
            "XLSX, XLS, CSV, TXT, DAT"
        )

    # ========================================================
    # Basic validation
    # ========================================================

    if df.empty:

        raise ValueError(
            "The experimental file is empty."
        )

    # ========================================================
    # Find D column
    # ========================================================

    D_column = find_column(
        df,
        [
            "D_track",
            "D",
            "D_um2_per_s",
            "D_um2/s",
            "D_um2",
            "Diffusion",
            "DiffusionCoefficient"
        ]
    )

    # ========================================================
    # Find frame column
    # ========================================================

    frame_column = find_column(
        df,
        [
            "n_frames",
            "N_FRAMES",
            "frames",
            "frame_count",
            "frame"
        ]
    )

    # ========================================================
    # D is REQUIRED
    # ========================================================

    if D_column is None:

        raise ValueError(
            "Could not find the D_track column.\n\n"
            "The experimental file must contain:\n"
            "D_track"
        )

    # ========================================================
    # n_frames is REQUIRED
    # ========================================================

    if frame_column is None:

        raise ValueError(
            "Could not find the n_frames column.\n\n"
            "The experimental file must contain:\n"
            "n_frames"
        )

    # ========================================================
    # Create paired dataframe
    #
    # IMPORTANT:
    # D and n_frames remain in the same row.
    # ========================================================

    paired_df = pd.DataFrame({

        "D_track":
            pd.to_numeric(
                df[D_column],
                errors="coerce"
            ),

        "n_frames":
            pd.to_numeric(
                df[frame_column],
                errors="coerce"
            )
    })

    # ========================================================
    # Remove NaN
    # ========================================================

    paired_df = paired_df.dropna(
        subset=[
            "D_track",
            "n_frames"
        ]
    )

    # ========================================================
    # Remove non-finite values
    # ========================================================

    paired_df = paired_df[
        np.isfinite(
            paired_df["D_track"]
        )
        &
        np.isfinite(
            paired_df["n_frames"]
        )
    ]

    # ========================================================
    # D >= 0
    # ========================================================

    paired_df = paired_df[
        paired_df["D_track"] >= 0
    ]

    # ========================================================
    # Frame > 0
    # ========================================================

    paired_df = paired_df[
        paired_df["n_frames"] > 0
    ]

    # ========================================================
    # Frame must be integer
    # ========================================================

    frame_values = (
        paired_df["n_frames"]
        .to_numpy()
    )

    if not np.all(
        np.isclose(
            frame_values,
            np.round(frame_values)
        )
    ):

        raise ValueError(
            "n_frames contains non-integer values.\n\n"
            "Every n_frames value must be an integer."
        )

    # ========================================================
    # Convert frame to integer
    # ========================================================

    paired_df["n_frames"] = (
        np.round(
            paired_df["n_frames"]
        )
        .astype(int)
    )

    # ========================================================
    # Reset index
    # ========================================================

    paired_df = paired_df.reset_index(
        drop=True
    )

    if len(paired_df) == 0:

        raise ValueError(
            "No valid paired D_track + n_frames "
            "rows were found."
        )

    return paired_df, sheet_name


# ============================================================
# Joint D + Frame sampling
# ============================================================

def sample_joint_D_frame_pairs(
    paired_df,
    n_trajectories
):

    if paired_df is None:

        raise ValueError(
            "Experimental data are not loaded."
        )

    if len(paired_df) == 0:

        raise ValueError(
            "Experimental data are empty."
        )

    # --------------------------------------------------------
    # Randomly select experimental rows
    #
    # WITH replacement
    #
    # D and n_frames are always taken from
    # the SAME selected row.
    # --------------------------------------------------------

    indices = np.random.choice(
        len(paired_df),
        size=n_trajectories,
        replace=True
    )

    sampled_df = (
        paired_df
        .iloc[indices]
        .reset_index(drop=True)
    )

    # --------------------------------------------------------
    # D
    # --------------------------------------------------------

    D_values = (
        sampled_df[
            "D_track"
        ]
        .to_numpy(
            dtype=float
        )
    )

    # --------------------------------------------------------
    # Frames
    # --------------------------------------------------------

    frame_values = (
        sampled_df[
            "n_frames"
        ]
        .to_numpy(
            dtype=int
        )
    )

    return (
        D_values,
        frame_values,
        sampled_df
    )


# ============================================================
# Write Excel
# ============================================================

def write_excel(
    filepath,
    trajectory_records,
    D_records
):

    trajectory_df = pd.DataFrame(
        trajectory_records,
        columns=[
            "Time",
            "mRNA_ID",
            "x",
            "y",
            "z"
        ]
    )

    D_df = pd.DataFrame(
        D_records,
        columns=[
            "mRNA_ID",
            "D_um2_per_s",
            "N_FRAMES",
            "Sampling_Mode"
        ]
    )

    with pd.ExcelWriter(
        filepath,
        engine="openpyxl"
    ) as writer:

        trajectory_df.to_excel(
            writer,
            sheet_name="Trajectories",
            index=False
        )

        D_df.to_excel(
            writer,
            sheet_name="D_parameters",
            index=False
        )


# ============================================================
# Write trajectory CSV
# ============================================================

def write_csv(
    filepath,
    trajectory_records
):

    trajectory_df = pd.DataFrame(
        trajectory_records,
        columns=[
            "Time",
            "mRNA_ID",
            "x",
            "y",
            "z"
        ]
    )

    trajectory_df.to_csv(
        filepath,
        index=False
    )


# ============================================================
# Write track CSV
# ============================================================

def write_tracks_csv(
    filepath,
    D_records
):

    D_df = pd.DataFrame(
        D_records,
        columns=[
            "TRACK_ID",
            "D_um2_per_s",
            "N_FRAMES",
            "Sampling_Mode"
        ]
    )

    D_df.to_csv(
        filepath,
        index=False
    )


# ============================================================
# Plot trajectories
# ============================================================

def plot_trajectories(
    filepath,
    trajectories,
    title
):

    plt.figure(
        figsize=(8, 8)
    )

    for x, y in trajectories:

        plt.plot(
            x,
            y,
            linewidth=0.8,
            alpha=0.75
        )

        plt.scatter(
            x[0],
            y[0],
            s=10
        )

    plt.xlabel(
        "X (µm)"
    )

    plt.ylabel(
        "Y (µm)"
    )

    plt.title(
        title
    )

    plt.axis(
        "equal"
    )

    plt.tight_layout()

    plt.savefig(
        filepath,
        dpi=300
    )

    plt.close()


# ============================================================
# GUI
# ============================================================

class BrownianSimulatorGUI:

    def __init__(
        self,
        root
    ):

        self.root = root

        self.root.title(
            "2D Brownian Trajectory Simulator - Joint D + Frame"
        )

        self.root.geometry(
            "1050x760"
        )

        self.root.minsize(
            950,
            680
        )

        # ====================================================
        # Variables
        # ====================================================

        self.output_dir = tk.StringVar(
            value=os.getcwd()
        )

        self.prefix = tk.StringVar(
            value="Brownian_Joint"
        )

        self.n_trajectories = tk.IntVar(
            value=100
        )

        self.dt = tk.DoubleVar(
            value=0.016241
        )

        self.replicates = tk.IntVar(
            value=1
        )

        self.experimental_D_file = tk.StringVar(
            value=""
        )

        self.sigma_x_nm = tk.DoubleVar(
            value=15.0
        )

        self.sigma_y_nm = tk.DoubleVar(
            value=13.2
        )

        # ====================================================
        # Style
        # ====================================================

        style = ttk.Style()

        try:

            style.theme_use(
                "clam"
            )

        except Exception:

            pass

        # ====================================================
        # Main frame
        # ====================================================

        main = ttk.Frame(
            self.root,
            padding=12
        )

        main.pack(
            fill="both",
            expand=True
        )

        # ====================================================
        # Title
        # ====================================================

        title = ttk.Label(
            main,
            text=(
                "2D Brownian Trajectory Simulator"
            ),
            font=(
                "Arial",
                18,
                "bold"
            )
        )

        title.pack(
            pady=(0, 5)
        )

        subtitle = ttk.Label(
            main,
            text=(
                "Experimental D + Frame (Joint Sampling)"
            ),
            font=(
                "Arial",
                11
            )
        )

        subtitle.pack(
            pady=(0, 10)
        )

        # ====================================================
        # BASIC SETTINGS
        # ====================================================

        basic_frame = ttk.LabelFrame(
            main,
            text="Basic Settings",
            padding=10
        )

        basic_frame.pack(
            fill="x",
            pady=5
        )

        # ----------------------------------------------------
        # Number trajectories
        # ----------------------------------------------------

        ttk.Label(
            basic_frame,
            text="Trajectories:"
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        ttk.Entry(
            basic_frame,
            textvariable=self.n_trajectories,
            width=10
        ).grid(
            row=0,
            column=1,
            padx=5
        )

        # ----------------------------------------------------
        # dt
        # ----------------------------------------------------

        ttk.Label(
            basic_frame,
            text="dt (s/frame):"
        ).grid(
            row=0,
            column=2,
            sticky="w",
            padx=(30, 5)
        )

        ttk.Entry(
            basic_frame,
            textvariable=self.dt,
            width=12
        ).grid(
            row=0,
            column=3,
            padx=5
        )

        # ----------------------------------------------------
        # Replicates
        # ----------------------------------------------------

        ttk.Label(
            basic_frame,
            text="Replicates:"
        ).grid(
            row=0,
            column=4,
            sticky="w",
            padx=(30, 5)
        )

        ttk.Entry(
            basic_frame,
            textvariable=self.replicates,
            width=8
        ).grid(
            row=0,
            column=5,
            padx=5
        )

        # ====================================================
        # JOINT SAMPLING
        # ====================================================

        D_frame = ttk.LabelFrame(
            main,
            text="Experimental D + Frame (Joint)",
            padding=10
        )

        D_frame.pack(
            fill="x",
            pady=5
        )

        # ----------------------------------------------------
        # Sampling mode
        # ----------------------------------------------------

        ttk.Label(
            D_frame,
            text="Sampling mode:"
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        mode_label = ttk.Label(
            D_frame,
            text=(
                "Experimental D + Frame (Joint)"
            ),
            font=(
                "Arial",
                10,
                "bold"
            )
        )

        mode_label.grid(
            row=0,
            column=1,
            columnspan=3,
            sticky="w",
            padx=5
        )

        # ----------------------------------------------------
        # Description
        # ----------------------------------------------------

        description = ttk.Label(
            D_frame,
            text=(
                "D_track and n_frames are sampled "
                "from the SAME experimental row. "
                "The original D-frame pairing is preserved. "
                "Sampling is performed WITH replacement."
            ),
            wraplength=750
        )

        description.grid(
            row=1,
            column=0,
            columnspan=6,
            sticky="w",
            padx=5,
            pady=(0, 10)
        )

        # ----------------------------------------------------
        # Experimental file
        # ----------------------------------------------------

        ttk.Label(
            D_frame,
            text="Experimental file:"
        ).grid(
            row=2,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        self.experimental_entry = ttk.Entry(
            D_frame,
            textvariable=self.experimental_D_file,
            width=70
        )

        self.experimental_entry.grid(
            row=2,
            column=1,
            columnspan=4,
            sticky="ew",
            padx=5
        )

        self.browse_D_button = ttk.Button(
            D_frame,
            text="Browse...",
            command=self.browse_D_file
        )

        self.browse_D_button.grid(
            row=2,
            column=5,
            padx=5
        )

        # ----------------------------------------------------
        # Required columns
        # ----------------------------------------------------

        required_label = ttk.Label(
            D_frame,
            text=(
                "Required columns: D_track + n_frames"
            )
        )

        required_label.grid(
            row=3,
            column=0,
            columnspan=6,
            sticky="w",
            padx=5,
            pady=(3, 5)
        )

        # ====================================================
        # LOCALIZATION ERROR
        # ====================================================

        localization_frame = ttk.LabelFrame(
            main,
            text="Localization Error",
            padding=10
        )

        localization_frame.pack(
            fill="x",
            pady=5
        )

        ttk.Label(
            localization_frame,
            text="Sigma X (nm):"
        ).grid(
            row=0,
            column=0,
            padx=5
        )

        ttk.Entry(
            localization_frame,
            textvariable=self.sigma_x_nm,
            width=10
        ).grid(
            row=0,
            column=1,
            padx=5
        )

        ttk.Label(
            localization_frame,
            text="Sigma Y (nm):"
        ).grid(
            row=0,
            column=2,
            padx=(25, 5)
        )

        ttk.Entry(
            localization_frame,
            textvariable=self.sigma_y_nm,
            width=10
        ).grid(
            row=0,
            column=3,
            padx=5
        )

        ttk.Label(
            localization_frame,
            text="(input in nm)"
        ).grid(
            row=0,
            column=4,
            padx=10
        )

        # ====================================================
        # OUTPUT
        # ====================================================

        output_frame = ttk.LabelFrame(
            main,
            text="Output",
            padding=10
        )

        output_frame.pack(
            fill="x",
            pady=5
        )

        ttk.Label(
            output_frame,
            text="Output folder:"
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=5
        )

        ttk.Entry(
            output_frame,
            textvariable=self.output_dir,
            width=70
        ).grid(
            row=0,
            column=1,
            padx=5,
            sticky="ew"
        )

        ttk.Button(
            output_frame,
            text="Browse...",
            command=self.browse_output
        ).grid(
            row=0,
            column=2,
            padx=5
        )

        ttk.Label(
            output_frame,
            text="Prefix:"
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=5,
            pady=7
        )

        ttk.Entry(
            output_frame,
            textvariable=self.prefix,
            width=25
        ).grid(
            row=1,
            column=1,
            sticky="w",
            padx=5
        )

        # ====================================================
        # BUTTONS
        # ====================================================

        button_frame = ttk.Frame(
            main
        )

        button_frame.pack(
            fill="x",
            pady=10
        )

        self.generate_button = ttk.Button(
            button_frame,
            text="GENERATE TRAJECTORIES",
            command=self.run_simulation
        )

        self.generate_button.pack(
            side="left",
            padx=5,
            ipadx=15,
            ipady=5
        )

        ttk.Button(
            button_frame,
            text="CLEAR LOG",
            command=self.clear_log
        ).pack(
            side="left",
            padx=5,
            ipadx=15,
            ipady=5
        )

        ttk.Button(
            button_frame,
            text="EXIT",
            command=self.root.destroy
        ).pack(
            side="right",
            padx=5,
            ipadx=15,
            ipady=5
        )

        # ====================================================
        # Progress
        # ====================================================

        self.progress = ttk.Progressbar(
            main,
            orient="horizontal",
            mode="determinate"
        )

        self.progress.pack(
            fill="x",
            pady=5
        )

        # ====================================================
        # LOG
        # ====================================================

        log_frame = ttk.LabelFrame(
            main,
            text="Log",
            padding=5
        )

        log_frame.pack(
            fill="both",
            expand=True,
            pady=5
        )

        self.log_text = tk.Text(
            log_frame,
            height=10,
            wrap="word"
        )

        self.log_text.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        self.log_text.configure(
            yscrollcommand=scrollbar.set
        )

    # ========================================================
    # Logging
    # ========================================================

    def log(
        self,
        text
    ):

        self.log_text.insert(
            "end",
            text + "\n"
        )

        self.log_text.see(
            "end"
        )

        self.root.update_idletasks()

    # ========================================================
    # Clear log
    # ========================================================

    def clear_log(
        self
    ):

        self.log_text.delete(
            "1.0",
            "end"
        )

    # ========================================================
    # Browse experimental file
    # ========================================================

    def browse_D_file(
        self
    ):

        filepath = (
            filedialog.askopenfilename(
                title=(
                    "Select experimental "
                    "D_track + n_frames file"
                ),
                filetypes=[
                    (
                        "Excel files",
                        "*.xlsx *.xls"
                    ),
                    (
                        "CSV files",
                        "*.csv"
                    ),
                    (
                        "Text files",
                        "*.txt"
                    ),
                    (
                        "DAT files",
                        "*.dat"
                    ),
                    (
                        "All files",
                        "*.*"
                    )
                ]
            )
        )

        if filepath:

            self.experimental_D_file.set(
                filepath
            )

            self.log(
                "Selected experimental "
                "D_track + n_frames file:"
            )

            self.log(
                filepath
            )

    # ========================================================
    # Browse output folder
    # ========================================================

    def browse_output(
        self
    ):

        folder = (
            filedialog.askdirectory(
                title="Select output folder"
            )
        )

        if folder:

            self.output_dir.set(
                folder
            )

    # ========================================================
    # Run simulation
    # ========================================================

    def run_simulation(
        self
    ):

        try:

            # =================================================
            # Read GUI values
            # =================================================

            n_trajectories = int(
                self.n_trajectories.get()
            )

            dt = float(
                self.dt.get()
            )

            replicates = int(
                self.replicates.get()
            )

            sigma_x_nm = float(
                self.sigma_x_nm.get()
            )

            sigma_y_nm = float(
                self.sigma_y_nm.get()
            )

            sigma_x = (
                sigma_x_nm / 1000.0
            )

            sigma_y = (
                sigma_y_nm / 1000.0
            )

            # =================================================
            # Validate
            # =================================================

            if n_trajectories <= 0:

                raise ValueError(
                    "Number of trajectories "
                    "must be > 0."
                )

            if dt <= 0:

                raise ValueError(
                    "dt must be > 0."
                )

            if replicates <= 0:

                raise ValueError(
                    "Replicates must be > 0."
                )

            if sigma_x_nm < 0:

                raise ValueError(
                    "Sigma X must be >= 0."
                )

            if sigma_y_nm < 0:

                raise ValueError(
                    "Sigma Y must be >= 0."
                )

            # =================================================
            # Experimental file
            # =================================================

            filepath = (
                self.experimental_D_file
                .get()
                .strip()
            )

            if not filepath:

                raise ValueError(
                    "Please select an experimental "
                    "D_track + n_frames file."
                )

            self.log(
                "Reading experimental "
                "D_track + n_frames distribution..."
            )

            (
                paired_data,
                sheet_name
            ) = load_experimental_D_frame_file(
                filepath
            )

            # =================================================
            # Excel sheet
            # =================================================

            if sheet_name is not None:

                self.log(
                    f"Excel sheet used: "
                    f"{sheet_name}"
                )

            # =================================================
            # Data information
            # =================================================

            self.log(
                f"Loaded "
                f"{len(paired_data)} "
                f"valid experimental "
                f"D-frame pairs."
            )

            self.log(
                "D_track range: "
                f"{paired_data['D_track'].min():.6g}"
                f" - "
                f"{paired_data['D_track'].max():.6g} "
                f"µm²/s"
            )

            self.log(
                "n_frames range: "
                f"{paired_data['n_frames'].min()}"
                f" - "
                f"{paired_data['n_frames'].max()}"
            )

            unique_frames = np.sort(
                paired_data[
                    "n_frames"
                ].unique()
            )

            self.log(
                "Unique n_frames values: "
                f"{len(unique_frames)}"
            )

            self.log(
                "Sampling mode: "
                "EXPERIMENTAL D + FRAME (JOINT)"
            )

            self.log(
                "D_track and n_frames "
                "will remain paired."
            )

            self.log(
                "Sampling: "
                "WITH replacement"
            )

            # =================================================
            # Output
            # =================================================

            output_dir = (
                self.output_dir.get()
            )

            if not output_dir:

                raise ValueError(
                    "Please select an output folder."
                )

            os.makedirs(
                output_dir,
                exist_ok=True
            )

            prefix = (
                self.prefix.get()
                .strip()
            )

            if not prefix:

                prefix = "Brownian_Joint"

            # =================================================
            # Prepare GUI
            # =================================================

            self.generate_button.configure(
                state="disabled"
            )

            self.progress["value"] = 0

            total_steps = (
                n_trajectories
                * replicates
            )

            completed = 0

            # =================================================
            # Header
            # =================================================

            self.log("")

            self.log(
                "=" * 65
            )

            self.log(
                "STARTING BROWNIAN "
                "TRAJECTORY SIMULATION"
            )

            self.log(
                "=" * 65
            )

            self.log(
                f"Trajectories / replicate : "
                f"{n_trajectories}"
            )

            self.log(
                "Frames / trajectory      : "
                "EXPERIMENTAL"
            )

            self.log(
                f"dt                       : "
                f"{dt} s/frame"
            )

            self.log(
                f"Replicates               : "
                f"{replicates}"
            )

            self.log(
                "Sampling mode            : "
                "Experimental D + Frame (Joint)"
            )

            self.log(
                f"Localization X           : "
                f"{sigma_x_nm} nm"
            )

            self.log(
                f"Localization Y           : "
                f"{sigma_y_nm} nm"
            )

            self.log("")

            # =================================================
            # Replicates
            # =================================================

            for rep in range(
                1,
                replicates + 1
            ):

                self.log(
                    f"--- Replicate "
                    f"{rep}/{replicates} ---"
                )

                # =================================================
                # JOINT SAMPLING
                # =================================================

                (
                    D_values,
                    frame_values,
                    sampled_df
                ) = sample_joint_D_frame_pairs(
                    paired_df=paired_data,
                    n_trajectories=n_trajectories
                )

                sampling_mode_label = (
                    "Experimental_Joint"
                )

                # =================================================
                # Records
                # =================================================

                trajectory_records = []

                D_records = []

                trajectories_for_plot = []

                # =================================================
                # Generate trajectories
                # =================================================

                for i in range(
                    n_trajectories
                ):

                    mRNA_ID = i + 1

                    D = float(
                        D_values[i]
                    )

                    current_n_frames = int(
                        frame_values[i]
                    )

                    # ------------------------------------------------
                    # Validation
                    # ------------------------------------------------

                    if current_n_frames <= 0:

                        raise ValueError(
                            f"Invalid n_frames for "
                            f"mRNA_ID {mRNA_ID}: "
                            f"{current_n_frames}"
                        )

                    if D < 0:

                        raise ValueError(
                            f"Invalid D for "
                            f"mRNA_ID {mRNA_ID}: "
                            f"{D}"
                        )

                    # ------------------------------------------------
                    # Random starting position
                    # ------------------------------------------------

                    x0 = np.random.uniform(
                        0.0,
                        10.0
                    )

                    y0 = np.random.uniform(
                        0.0,
                        10.0
                    )

                    # ------------------------------------------------
                    # Generate trajectory
                    # ------------------------------------------------

                    x, y = (
                        generate_brownian_trajectory(
                            D=D,
                            n_frames=current_n_frames,
                            dt=dt,
                            sigma_x=sigma_x,
                            sigma_y=sigma_y,
                            x0=x0,
                            y0=y0
                        )
                    )

                    # ------------------------------------------------
                    # Store for plot
                    # ------------------------------------------------

                    trajectories_for_plot.append(
                        (x, y)
                    )

                    # ------------------------------------------------
                    # Store trajectory
                    # ------------------------------------------------

                    for frame in range(
                        current_n_frames
                    ):

                        time = (
                            frame * dt
                        )

                        trajectory_records.append(
                            (
                                time,
                                mRNA_ID,
                                x[frame],
                                y[frame],
                                0.0
                            )
                        )

                    # ------------------------------------------------
                    # Store D parameters
                    # ------------------------------------------------

                    D_records.append(
                        (
                            mRNA_ID,
                            D,
                            current_n_frames,
                            sampling_mode_label
                        )
                    )

                    # ------------------------------------------------
                    # Progress
                    # ------------------------------------------------

                    completed += 1

                    self.progress[
                        "value"
                    ] = (
                        completed
                        /
                        total_steps
                        *
                        100.0
                    )

                    # ------------------------------------------------
                    # Log every 10 trajectories
                    # ------------------------------------------------

                    if (
                        (i + 1) % 10 == 0
                        or
                        i == n_trajectories - 1
                    ):

                        self.log(
                            f"Generated "
                            f"{i + 1}/"
                            f"{n_trajectories} trajectories"
                        )

                    self.root.update_idletasks()

                # =================================================
                # Output filenames
                # =================================================

                base_name = (
                    f"{prefix}_rep{rep}"
                )

                excel_file = os.path.join(
                    output_dir,
                    base_name + ".xlsx"
                )

                csv_file = os.path.join(
                    output_dir,
                    base_name + ".csv"
                )

                tracks_csv_file = os.path.join(
                    output_dir,
                    base_name + "_tracks.csv"
                )

                png_file = os.path.join(
                    output_dir,
                    base_name + "_trajectories.png"
                )

                # =================================================
                # Write Excel
                # =================================================

                self.log(
                    "Writing Excel..."
                )

                write_excel(
                    filepath=excel_file,
                    trajectory_records=(
                        trajectory_records
                    ),
                    D_records=D_records
                )

                # =================================================
                # Write CSV
                # =================================================

                self.log(
                    "Writing trajectory CSV..."
                )

                write_csv(
                    filepath=csv_file,
                    trajectory_records=(
                        trajectory_records
                    )
                )

                # =================================================
                # Write Tracks CSV
                # =================================================

                self.log(
                    "Writing track CSV..."
                )

                write_tracks_csv(
                    filepath=tracks_csv_file,
                    D_records=D_records
                )

                # =================================================
                # Plot
                # =================================================

                self.log(
                    "Creating trajectory image..."
                )

                plot_trajectories(
                    filepath=png_file,
                    trajectories=(
                        trajectories_for_plot
                    ),
                    title=(
                        "2D Brownian Trajectories - "
                        f"Replicate {rep}"
                    )
                )

                # =================================================
                # Statistics
                # =================================================

                actual_frames = np.array(
                    frame_values,
                    dtype=int
                )

                actual_D = np.array(
                    D_values,
                    dtype=float
                )

                self.log(
                    f"D range used: "
                    f"{np.min(actual_D):.6g}"
                    f" - "
                    f"{np.max(actual_D):.6g} "
                    f"µm²/s"
                )

                self.log(
                    f"Mean D used: "
                    f"{np.mean(actual_D):.6g} "
                    f"µm²/s"
                )

                self.log(
                    f"Frame range used: "
                    f"{np.min(actual_frames)}"
                    f" - "
                    f"{np.max(actual_frames)}"
                )

                self.log(
                    f"Mean frames used: "
                    f"{np.mean(actual_frames):.2f}"
                )

                self.log(
                    f"Total trajectory rows: "
                    f"{len(trajectory_records)}"
                )

                # -------------------------------------------------
                # Verify pairing
                # -------------------------------------------------

                pairing_check = np.all(
                    np.isclose(
                        D_values,
                        sampled_df[
                            "D_track"
                        ].to_numpy(
                            dtype=float
                        )
                    )
                    &
                    (
                        frame_values
                        ==
                        sampled_df[
                            "n_frames"
                        ].to_numpy(
                            dtype=int
                        )
                    )
                )

                if pairing_check:

                    self.log(
                        "D-frame pairing check: PASSED"
                    )

                else:

                    raise RuntimeError(
                        "D-frame pairing check FAILED."
                    )

                self.log(
                    f"Replicate {rep} finished."
                )

                self.log(
                    f"Excel: {excel_file}"
                )

                self.log(
                    f"CSV:   {csv_file}"
                )

                self.log(
                    f"Tracks: {tracks_csv_file}"
                )

                self.log(
                    f"PNG:   {png_file}"
                )

                self.log("")

            # =====================================================
            # Complete
            # =====================================================

            self.progress["value"] = 100

            self.log(
                "=" * 65
            )

            self.log(
                "SIMULATION COMPLETED SUCCESSFULLY"
            )

            self.log(
                "=" * 65
            )

            self.log(
                f"Total trajectories: "
                f"{n_trajectories * replicates}"
            )

            self.log(
                "Sampling mode: "
                "Experimental D + Frame (Joint)"
            )

            self.log(
                "D-track/frame pairing: "
                "PRESERVED"
            )

            self.log(
                "Sampling replacement: "
                "WITH replacement"
            )

            self.log(
                f"Output folder: "
                f"{output_dir}"
            )

            # =====================================================
            # Message
            # =====================================================

            messagebox.showinfo(
                "Completed",
                "Trajectory generation "
                "completed successfully!\n\n"
                f"Trajectories per replicate: "
                f"{n_trajectories}\n"
                "D + Frame mode:\n"
                "Experimental Joint Sampling\n\n"
                "D-track/frame pairing:\n"
                "PRESERVED\n\n"
                "Sampling:\n"
                "WITH replacement\n\n"
                f"Replicates: "
                f"{replicates}\n\n"
                f"Output folder:\n"
                f"{output_dir}"
            )

        # =====================================================
        # Error
        # =====================================================

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

            self.log(
                f"ERROR: {e}"
            )

        # =====================================================
        # Always enable button
        # =====================================================

        finally:

            self.generate_button.configure(
                state="normal"
            )


# ============================================================
# Main
# ============================================================

def main():

    root = tk.Tk()

    app = BrownianSimulatorGUI(
        root
    )

    root.mainloop()


if __name__ == "__main__":

    main()