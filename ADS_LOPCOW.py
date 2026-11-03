# ======================================================================
# ADS-LOPCOW
# STAGE B — FULL 18,000-MATRIX SIMULATION
# CURRENT HYBRID MIN-MAX METHODOLOGY
#
# PURPOSE
# -------
# 1. Run classical LOPCOW on all 18,000 matrices.
# 2. Run current ADS-LOPCOW on the same matrices.
# 3. Preserve original LOPCOW min-max normalization for Rule B.
# 4. Apply updated Rule A1 to constant criteria on known semantic scales.
# 5. Retain Rule A2 for unanchored quantitative constant criteria.
# 6. Sweep z = 0.0, 0.1, ..., 1.0.
# 7. Save detailed, aggregate, semantic and validation results.
#
# CURRENT RULE A1
# ---------------
# Benefit:
#     h = (c - L) / (U - L)
#
# Cost:
#     h = (U - c) / (U - L)
#
# For the constant transformed column:
#     RMS = h
#
# Score:
#     S = 1 + RMS
#
# Percentage value:
#     PV = 100 * ln(1 + RMS)
#
# CURRENT RULE B
# --------------
# Original LOPCOW min-max normalization.
#
# COA  = RMS / SD
# DOA  = RMS * SD
# DOA* = 1 + DOA
#
# Score(z) = z*COA + (1-z)*DOA*
# PV       = 100*ln(Score)
#
# IMPORTANT FOR CURRENT SIMULATION
# --------------------------------
# All generated criteria are ratings on a known 1–9 semantic scale.
# Therefore any constant criterion is treated as Rule A1.
#
# All criteria are treated as BENEFIT criteria in this simulation.
# ======================================================================


!pip install -q xlsxwriter openpyxl


# ======================================================================
# 1. IMPORTS
# ======================================================================

import re
import numpy as np
import pandas as pd

from google.colab import files


# ======================================================================
# 2. GLOBAL SETTINGS
# ======================================================================

EPS = 1e-12

SEMANTIC_LOWER = 1.0
SEMANTIC_UPPER = 9.0

Z_VALUES = np.round(
    np.arange(
        0.0,
        1.01,
        0.1
    ),
    1
)

OUTPUT_FILE = (
    "ADS_LOPCOW_18000_CURRENT_FULL_RESULTS.xlsx"
)


# ======================================================================
# 3. CRITERION TYPE CONFIGURATION
# ======================================================================
#
# Current simulation:
# all criteria default to BENEFIT.
#
# If later you want mixed benefit/cost criteria, add explicit lists.
#
# Example:
#
# CRITERION_TYPE_CONFIG = {
#     "5x6": [
#         "benefit",
#         "benefit",
#         "cost",
#         "benefit",
#         "benefit",
#         "cost"
#     ]
# }
#
# ======================================================================

CRITERION_TYPE_CONFIG = {}


# ======================================================================
# 4. UPLOAD THE 18,000-MATRIX WORKBOOK
# ======================================================================

print(
    "Upload ADS_LOPCOW_18000_Matrices.xlsx ..."
)

uploaded = files.upload()

excel_names = [
    name
    for name in uploaded.keys()
    if name.lower().endswith(
        (
            ".xlsx",
            ".xlsm"
        )
    )
]

if not excel_names:

    raise ValueError(
        "No Excel workbook was uploaded."
    )

INPUT_FILE = excel_names[0]

print(
    f"\nLoaded workbook: {INPUT_FILE}"
)


# ======================================================================
# 5. OPEN WORKBOOK
# ======================================================================

excel_file = pd.ExcelFile(
    INPUT_FILE
)

print(
    "\nWorkbook sheets:"
)

print(
    excel_file.sheet_names
)


# ======================================================================
# 6. IDENTIFY MATRIX SHEETS
# ======================================================================

def parse_dimension(
    sheet_name
):

    match = re.fullmatch(
        r"\s*(\d+)\s*[xX]\s*(\d+)\s*",
        str(sheet_name)
    )

    if match is None:

        return None

    return (
        int(
            match.group(1)
        ),
        int(
            match.group(2)
        )
    )


SHEET_DIMENSIONS = {}

for sheet in excel_file.sheet_names:

    dim = parse_dimension(
        sheet
    )

    if dim is not None:

        SHEET_DIMENSIONS[
            sheet
        ] = dim


if not SHEET_DIMENSIONS:

    raise ValueError(
        "No matrix sheets such as "
        "5x6, 6x7, 7x8, 6x9 or 5x8 were found."
    )


print(
    "\nDetected matrix dimensions:"
)

for sheet, (
    m,
    n
) in SHEET_DIMENSIONS.items():

    print(
        f"  {sheet}: "
        f"{m} alternatives × "
        f"{n} criteria"
    )


# ======================================================================
# 7. RECONSTRUCT ALL MATRICES IN A SHEET
# ======================================================================

def reconstruct_batch(
    df,
    n_rows,
    n_cols
):
    """
    Converts flattened columns such as:

        A1_C1
        A1_C2
        ...
        A5_C6

    into:

        raw[matrix, alternative, criterion]
    """

    n_matrices = len(
        df
    )

    raw = np.empty(
        (
            n_matrices,
            n_rows,
            n_cols
        ),
        dtype=float
    )


    for i in range(
        n_rows
    ):

        for j in range(
            n_cols
        ):

            column = (
                f"A{i+1}_C{j+1}"
            )

            if column not in df.columns:

                raise KeyError(
                    f"Missing expected matrix column: {column}"
                )


            raw[
                :,
                i,
                j
            ] = pd.to_numeric(
                df[
                    column
                ],
                errors="raise"
            ).to_numpy(
                dtype=float
            )


    return raw


# ======================================================================
# 8. GET CRITERION TYPES
# ======================================================================

def get_criterion_types(
    sheet_name,
    n_cols
):

    if (
        sheet_name
        in CRITERION_TYPE_CONFIG
    ):

        criterion_types = (
            CRITERION_TYPE_CONFIG[
                sheet_name
            ]
        )

        if (
            len(
                criterion_types
            )
            != n_cols
        ):

            raise ValueError(
                f"Criterion type count "
                f"does not match {sheet_name}."
            )

    else:

        criterion_types = [
            "benefit"
            for _ in range(
                n_cols
            )
        ]


    criterion_types = np.asarray(
        [
            str(x)
            .strip()
            .lower()

            for x
            in criterion_types
        ],
        dtype=object
    )


    valid = np.isin(
        criterion_types,
        [
            "benefit",
            "cost"
        ]
    )

    if not np.all(
        valid
    ):

        raise ValueError(
            f"Invalid criterion type "
            f"in {sheet_name}."
        )


    return criterion_types


# ======================================================================
# 9. CLASSICAL LOPCOW NORMALIZATION
# ======================================================================

def classical_minmax_batch(
    raw,
    criterion_types
):
    """
    Original LOPCOW observed min-max normalization.

    Constant criteria remain undefined because:
        max = min
    and therefore:
        denominator = 0
    """

    x_min = np.min(
        raw,
        axis=1
    )

    x_max = np.max(
        raw,
        axis=1
    )

    ranges = (
        x_max
        - x_min
    )

    constant_mask = (
        np.abs(
            ranges
        )
        <= EPS
    )


    norm = np.full_like(
        raw,
        np.nan,
        dtype=float
    )

    n_cols = (
        raw.shape[2]
    )


    for j in range(
        n_cols
    ):

        valid = (
            ~constant_mask[
                :,
                j
            ]
        )

        if not np.any(
            valid
        ):

            continue


        if (
            criterion_types[j]
            == "benefit"
        ):

            norm[
                valid,
                :,
                j
            ] = (

                raw[
                    valid,
                    :,
                    j
                ]

                -

                x_min[
                    valid,
                    j
                ][:, None]

            ) / (

                ranges[
                    valid,
                    j
                ][:, None]

            )


        elif (
            criterion_types[j]
            == "cost"
        ):

            norm[
                valid,
                :,
                j
            ] = (

                x_max[
                    valid,
                    j
                ][:, None]

                -

                raw[
                    valid,
                    :,
                    j
                ]

            ) / (

                ranges[
                    valid,
                    j
                ][:, None]

            )


        else:

            raise ValueError(
                f"Unknown criterion type: "
                f"{criterion_types[j]}"
            )


    return (
        norm,
        constant_mask,
        ranges
    )


# ======================================================================
# 10. CLASSICAL LOPCOW
# ======================================================================

def classical_lopcow_batch(
    norm,
    constant_mask
):
    """
    Runs classical LOPCOW.

    If any criterion is constant, the full classical
    weight vector is considered unavailable.
    """

    n_matrices = (
        norm.shape[0]
    )

    n_cols = (
        norm.shape[2]
    )


    rms = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )

    sd = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )

    coa = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )

    pv = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )

    weights = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )


    for j in range(
        n_cols
    ):

        valid = (
            ~constant_mask[
                :,
                j
            ]
        )

        if not np.any(
            valid
        ):

            continue


        x = norm[
            valid,
            :,
            j
        ]


        rms[
            valid,
            j
        ] = np.sqrt(
            np.mean(
                x ** 2,
                axis=1
            )
        )


        sd[
            valid,
            j
        ] = np.std(
            x,
            axis=1,
            ddof=1
        )


        valid_sd_local = (
            sd[
                valid,
                j
            ]
            > EPS
        )


        valid_indices = (
            np.where(
                valid
            )[0]
        )


        good_indices = (
            valid_indices[
                valid_sd_local
            ]
        )


        coa[
            good_indices,
            j
        ] = (

            rms[
                good_indices,
                j
            ]

            /

            sd[
                good_indices,
                j
            ]
        )


        pv[
            good_indices,
            j
        ] = (

            np.abs(
                np.log(
                    coa[
                        good_indices,
                        j
                    ]
                )
            )

            * 100.0
        )


    classical_failure = np.any(
        constant_mask,
        axis=1
    )


    valid_matrices = (
        ~classical_failure
    )


    if np.any(
        valid_matrices
    ):

        pv_sum = np.sum(
            pv[
                valid_matrices,
                :
            ],
            axis=1
        )


        positive = (
            pv_sum
            > EPS
        )


        valid_indices = np.where(
            valid_matrices
        )[0]


        good_indices = (
            valid_indices[
                positive
            ]
        )


        weights[
            good_indices,
            :
        ] = (

            pv[
                good_indices,
                :
            ]

            /

            pv_sum[
                positive
            ][:, None]
        )


    return {

        "rms":
            rms,

        "sd":
            sd,

        "coa":
            coa,

        "pv":
            pv,

        "weights":
            weights,

        "failed":
            classical_failure
    }


# ======================================================================
# 11. CLASSIFY ADS RULES
# ======================================================================

def classify_ads_rules_batch(
    raw,
    semantic_anchor_flags
):
    """
    Rule classification occurs BEFORE normalization.

    A1:
        range = 0
        known semantic anchor

    A2:
        range = 0
        no semantic anchor

    B:
        range > 0
    """

    x_min = np.min(
        raw,
        axis=1
    )

    x_max = np.max(
        raw,
        axis=1
    )

    ranges = (
        x_max
        - x_min
    )


    zero_range = (
        np.abs(
            ranges
        )
        <= EPS
    )


    semantic_anchor_flags = np.asarray(
        semantic_anchor_flags,
        dtype=bool
    )


    a1 = (
        zero_range
        &
        semantic_anchor_flags[
            None,
            :
        ]
    )


    a2 = (
        zero_range
        &
        (
            ~semantic_anchor_flags[
                None,
                :
            ]
        )
    )


    b = (
        ~zero_range
    )


    return {

        "x_min":
            x_min,

        "x_max":
            x_max,

        "ranges":
            ranges,

        "a1":
            a1,

        "a2":
            a2,

        "b":
            b
    }


# ======================================================================
# 12. CURRENT ADS COMPONENTS
# ======================================================================

def ads_components_batch(
    raw,
    criterion_types,
    semantic_anchor_flags,
    semantic_lower,
    semantic_upper
):
    """
    Current hybrid min-max ADS-LOPCOW.

    Rule A1
    -------
    Benefit:
        h = (c-L)/(U-L)

    Cost:
        h = (U-c)/(U-L)

    RMS = h

    Rule A2
    -------
    PV = 0

    Rule B
    -------
    Original LOPCOW observed min-max normalization.

    COA = RMS/SD
    DOA = RMS*SD
    DOA* = 1 + DOA
    """

    if (
        semantic_upper
        <= semantic_lower
    ):

        raise ValueError(
            "semantic_upper must be greater than semantic_lower."
        )


    info = classify_ads_rules_batch(
        raw,
        semantic_anchor_flags
    )


    x_min = (
        info[
            "x_min"
        ]
    )

    x_max = (
        info[
            "x_max"
        ]
    )

    ranges = (
        info[
            "ranges"
        ]
    )

    a1 = (
        info[
            "a1"
        ]
    )

    a2 = (
        info[
            "a2"
        ]
    )

    b = (
        info[
            "b"
        ]
    )


    n_matrices = (
        raw.shape[0]
    )

    n_cols = (
        raw.shape[2]
    )


    norm = np.full_like(
        raw,
        np.nan,
        dtype=float
    )


    semantic_relative = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )


    rms = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )


    sd = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )


    coa = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )


    doa = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )


    doa_star = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )


    # ==================================================================
    # RULE A1
    # ==================================================================

    for j in range(
        n_cols
    ):

        mask = (
            a1[
                :,
                j
            ]
        )


        if not np.any(
            mask
        ):

            continue


        c = raw[
            mask,
            0,
            j
        ]


        if (
            criterion_types[j]
            == "benefit"
        ):

            h = (

                c
                - semantic_lower

            ) / (

                semantic_upper
                - semantic_lower

            )


        elif (
            criterion_types[j]
            == "cost"
        ):

            h = (

                semantic_upper
                - c

            ) / (

                semantic_upper
                - semantic_lower

            )


        else:

            raise ValueError(
                f"Unknown criterion type: "
                f"{criterion_types[j]}"
            )


        if np.any(
            h
            < -EPS
        ) or np.any(
            h
            > 1.0 + EPS
        ):

            raise ValueError(
                "A Rule A1 constant value falls outside "
                "the declared semantic scale."
            )


        h = np.clip(
            h,
            0.0,
            1.0
        )


        semantic_relative[
            mask,
            j
        ] = h


        # Constant transformed column:
        # every entry = h
        # therefore RMS = h

        rms[
            mask,
            j
        ] = h


        sd[
            mask,
            j
        ] = 0.0


    # ==================================================================
    # RULE A2
    # ==================================================================

    sd[
        a2
    ] = 0.0


    # ==================================================================
    # RULE B
    # ==================================================================

    for j in range(
        n_cols
    ):

        mask = (
            b[
                :,
                j
            ]
        )


        if not np.any(
            mask
        ):

            continue


        if (
            criterion_types[j]
            == "benefit"
        ):

            norm[
                mask,
                :,
                j
            ] = (

                raw[
                    mask,
                    :,
                    j
                ]

                -

                x_min[
                    mask,
                    j
                ][:, None]

            ) / (

                ranges[
                    mask,
                    j
                ][:, None]

            )


        elif (
            criterion_types[j]
            == "cost"
        ):

            norm[
                mask,
                :,
                j
            ] = (

                x_max[
                    mask,
                    j
                ][:, None]

                -

                raw[
                    mask,
                    :,
                    j
                ]

            ) / (

                ranges[
                    mask,
                    j
                ][:, None]

            )


        else:

            raise ValueError(
                f"Unknown criterion type: "
                f"{criterion_types[j]}"
            )


        x = norm[
            mask,
            :,
            j
        ]


        rms[
            mask,
            j
        ] = np.sqrt(
            np.mean(
                x ** 2,
                axis=1
            )
        )


        sd[
            mask,
            j
        ] = np.std(
            x,
            axis=1,
            ddof=1
        )


        if np.any(
            sd[
                mask,
                j
            ]
            <= EPS
        ):

            raise RuntimeError(
                "Rule B unexpectedly produced zero SD."
            )


        coa[
            mask,
            j
        ] = (

            rms[
                mask,
                j
            ]

            /

            sd[
                mask,
                j
            ]
        )


        doa[
            mask,
            j
        ] = (

            rms[
                mask,
                j
            ]

            *

            sd[
                mask,
                j
            ]
        )


        doa_star[
            mask,
            j
        ] = (

            1.0

            +

            doa[
                mask,
                j
            ]
        )


    return {

        "x_min":
            x_min,

        "x_max":
            x_max,

        "ranges":
            ranges,

        "rule_a1":
            a1,

        "rule_a2":
            a2,

        "rule_b":
            b,

        "norm":
            norm,

        "semantic_relative":
            semantic_relative,

        "rms":
            rms,

        "sd":
            sd,

        "coa":
            coa,

        "doa":
            doa,

        "doa_star":
            doa_star
    }


# ======================================================================
# 13. ADS PV + WEIGHTS FOR ONE z
# ======================================================================

def ads_pv_weights_batch(
    components,
    z
):
    """
    Rule A1:
        Score = 1 + RMS
        PV = 100 ln(1 + RMS)

    Rule A2:
        PV = 0

    Rule B:
        Score(z)
        = z*COA + (1-z)*DOA*

        PV
        = 100 ln(Score)
    """

    a1 = (
        components[
            "rule_a1"
        ]
    )

    a2 = (
        components[
            "rule_a2"
        ]
    )

    b = (
        components[
            "rule_b"
        ]
    )


    n_matrices, n_cols = (
        a1.shape
    )


    score = np.full(
        (
            n_matrices,
            n_cols
        ),
        np.nan
    )


    pv = np.zeros(
        (
            n_matrices,
            n_cols
        ),
        dtype=float
    )


    # ==================================================================
    # RULE A1
    # ==================================================================

    score[
        a1
    ] = (

        1.0

        +

        components[
            "rms"
        ][a1]
    )


    pv[
        a1
    ] = (

        100.0

        *

        np.log(
            score[
                a1
            ]
        )
    )


    # ==================================================================
    # RULE A2
    # ==================================================================

    score[
        a2
    ] = np.nan

    pv[
        a2
    ] = 0.0


    # ==================================================================
    # RULE B
    # ==================================================================

    score[
        b
    ] = (

        z
        *
        components[
            "coa"
        ][b]

        +

        (
            1.0
            - z
        )

        *
        components[
            "doa_star"
        ][b]
    )


    if np.any(
        score[
            b
        ]
        < 1.0 - EPS
    ):

        raise RuntimeError(
            f"Rule B score below 1 at z={z}."
        )


    pv[
        b
    ] = (

        100.0

        *

        np.log(
            score[
                b
            ]
        )
    )


    # Numerical cleanup.

    pv[
        np.abs(
            pv
        )
        < EPS
    ] = 0.0


    if np.any(
        pv
        < -EPS
    ):

        raise RuntimeError(
            f"Negative ADS PV detected at z={z}."
        )


    # ==================================================================
    # NORMALIZE INTO WEIGHTS
    # ==================================================================

    pv_sum = np.sum(
        pv,
        axis=1
    )


    weights = np.full_like(
        pv,
        np.nan,
        dtype=float
    )


    valid = (
        pv_sum
        > EPS
    )


    weights[
        valid,
        :
    ] = (

        pv[
            valid,
            :
        ]

        /

        pv_sum[
            valid
        ][:, None]
    )


    return (
        score,
        pv,
        weights
    )


# ======================================================================
# 14. WEIGHT STATISTICS
# ======================================================================

def weight_statistics(
    weights
):

    valid_rows = np.any(
        np.isfinite(
            weights
        ),
        axis=1
    )


    weight_sum = np.full(
        weights.shape[0],
        np.nan
    )

    max_weight = np.full(
        weights.shape[0],
        np.nan
    )

    min_weight = np.full(
        weights.shape[0],
        np.nan
    )

    hhi = np.full(
        weights.shape[0],
        np.nan
    )

    effective_n = np.full(
        weights.shape[0],
        np.nan
    )

    top_criterion = np.full(
        weights.shape[0],
        None,
        dtype=object
    )


    if np.any(
        valid_rows
    ):

        w = weights[
            valid_rows,
            :
        ]


        weight_sum[
            valid_rows
        ] = np.sum(
            w,
            axis=1
        )


        max_weight[
            valid_rows
        ] = np.max(
            w,
            axis=1
        )


        min_weight[
            valid_rows
        ] = np.min(
            w,
            axis=1
        )


        hhi[
            valid_rows
        ] = np.sum(
            w ** 2,
            axis=1
        )


        effective_n[
            valid_rows
        ] = (

            1.0

            /

            hhi[
                valid_rows
            ]
        )


        top_idx = np.argmax(
            w,
            axis=1
        )


        top_criterion[
            valid_rows
        ] = np.array(
            [
                f"C{i+1}"
                for i
                in top_idx
            ],
            dtype=object
        )


    return (
        weight_sum,
        max_weight,
        min_weight,
        hhi,
        effective_n,
        top_criterion
    )


# ======================================================================
# 15. RESULT CONTAINERS
# ======================================================================

global_summary_parts = []

aggregate_input_parts = []

validation_parts = []

semantic_summary_parts = []

rule_count_parts = []


# ======================================================================
# 16. WRITE FULL OUTPUT WORKBOOK
# ======================================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="xlsxwriter"
) as writer:


    # ==================================================================
    # README
    # ==================================================================

    readme = pd.DataFrame(
        [
            [
                "Method",
                "Current hybrid min-max ADS-LOPCOW"
            ],
            [
                "Expected matrices",
                18000
            ],
            [
                "Semantic scale",
                "1–9"
            ],
            [
                "Current simulation orientation",
                "All criteria treated as benefit criteria"
            ],
            [
                "Rule A1 benefit transformation",
                "h=(c-L)/(U-L)"
            ],
            [
                "Rule A1 cost transformation",
                "h=(U-c)/(U-L)"
            ],
            [
                "Rule A1 RMS",
                "RMS=h"
            ],
            [
                "Rule A1 score",
                "S=1+RMS"
            ],
            [
                "Rule A1 PV",
                "PV=100×ln(1+RMS)"
            ],
            [
                "Rule A2",
                "PV=0"
            ],
            [
                "Rule B",
                (
                    "Original LOPCOW observed min-max "
                    "normalization retained."
                )
            ],
            [
                "COA",
                "RMS/SD"
            ],
            [
                "DOA",
                "RMS×SD"
            ],
            [
                "Shifted DOA",
                "1+DOA"
            ],
            [
                "Rule B score",
                "S(z)=z×COA+(1-z)×(1+DOA)"
            ],
            [
                "Rule B PV",
                "PV=100×ln(S)"
            ],
            [
                "z sweep",
                "0.0 to 1.0 in increments of 0.1"
            ],
            [
                "Endpoint validation",
                (
                    "For normal matrices, ADS at z=1 "
                    "must reproduce classical LOPCOW."
                )
            ]
        ],
        columns=[
            "Item",
            "Description"
        ]
    )


    readme.to_excel(
        writer,
        sheet_name="README",
        index=False
    )


    # ==================================================================
    # PROCESS EACH DIMENSION
    # ==================================================================

    total_matrix_count = 0


    for sheet_name, (
        n_rows,
        n_cols
    ) in SHEET_DIMENSIONS.items():


        print(
            "\n"
            + "=" * 72
        )

        print(
            f"Processing {sheet_name}"
        )

        print(
            "=" * 72
        )


        df = excel_file.parse(
            sheet_name
        )


        total_matrix_count += len(
            df
        )


        raw = reconstruct_batch(
            df,
            n_rows,
            n_cols
        )


        criterion_types = (
            get_criterion_types(
                sheet_name,
                n_cols
            )
        )


        # --------------------------------------------------------------
        # Current simulation:
        # every criterion lies on a known 1–9 semantic scale.
        # --------------------------------------------------------------

        semantic_anchor_flags = np.ones(
            n_cols,
            dtype=bool
        )


        # ==============================================================
        # CLASSICAL LOPCOW
        # ==============================================================

        (
            classical_norm,
            classical_constant_mask,
            classical_ranges
        ) = classical_minmax_batch(
            raw,
            criterion_types
        )


        classical = (
            classical_lopcow_batch(
                classical_norm,
                classical_constant_mask
            )
        )


        # ==============================================================
        # ADS COMPONENTS
        # ==============================================================

        ads = ads_components_batch(

            raw=
                raw,

            criterion_types=
                criterion_types,

            semantic_anchor_flags=
                semantic_anchor_flags,

            semantic_lower=
                SEMANTIC_LOWER,

            semantic_upper=
                SEMANTIC_UPPER
        )


        # ==============================================================
        # RULE COUNTS
        # ==============================================================

        rule_count_parts.append(
            pd.DataFrame(
                {

                    "Dimension":
                        [
                            sheet_name
                        ],

                    "N_Matrices":
                        [
                            len(
                                df
                            )
                        ],

                    "N_A1_Criteria":
                        [
                            int(
                                np.sum(
                                    ads[
                                        "rule_a1"
                                    ]
                                )
                            )
                        ],

                    "N_A2_Criteria":
                        [
                            int(
                                np.sum(
                                    ads[
                                        "rule_a2"
                                    ]
                                )
                            )
                        ],

                    "N_Rule_B_Criteria":
                        [
                            int(
                                np.sum(
                                    ads[
                                        "rule_b"
                                    ]
                                )
                            )
                        ],

                    "Classical_Failed_Matrices":
                        [
                            int(
                                np.sum(
                                    classical[
                                        "failed"
                                    ]
                                )
                            )
                        ]
                }
            )
        )


        # ==============================================================
        # BASE SHEET
        # ==============================================================

        base = pd.DataFrame(
            {

                "Global_Matrix_ID":
                    df[
                        "Global_Matrix_ID"
                    ].to_numpy(),

                "Dimension":
                    sheet_name,

                "Case_Type":
                    df[
                        "Case_Type"
                    ].to_numpy(),

                "Declared_Constant_Criterion":
                    df[
                        "Constant_Criterion"
                    ].to_numpy(),

                "Declared_Constant_Value":
                    df[
                        "Constant_Value"
                    ].to_numpy(),

                "Classical_LOPCOW_Failed":
                    classical[
                        "failed"
                    ],

                "A1_Count":
                    np.sum(
                        ads[
                            "rule_a1"
                        ],
                        axis=1
                    ),

                "A2_Count":
                    np.sum(
                        ads[
                            "rule_a2"
                        ],
                        axis=1
                    ),

                "Rule_B_Count":
                    np.sum(
                        ads[
                            "rule_b"
                        ],
                        axis=1
                    )
            }
        )


        for j in range(
            n_cols
        ):

            c = (
                j + 1
            )


            rule_label = np.where(

                ads[
                    "rule_a1"
                ][
                    :,
                    j
                ],

                "A1",

                np.where(

                    ads[
                        "rule_a2"
                    ][
                        :,
                        j
                    ],

                    "A2",

                    "B"
                )
            )


            base[
                f"C{c}_Type"
            ] = criterion_types[
                j
            ]


            base[
                f"C{c}_Observed_Range"
            ] = ads[
                "ranges"
            ][
                :,
                j
            ]


            base[
                f"C{c}_Rule"
            ] = rule_label


            base[
                f"C{c}_A1_h"
            ] = ads[
                "semantic_relative"
            ][
                :,
                j
            ]


            base[
                f"C{c}_ADS_RMS"
            ] = ads[
                "rms"
            ][
                :,
                j
            ]


            base[
                f"C{c}_ADS_SD"
            ] = ads[
                "sd"
            ][
                :,
                j
            ]


            base[
                f"C{c}_COA"
            ] = ads[
                "coa"
            ][
                :,
                j
            ]


            base[
                f"C{c}_DOA"
            ] = ads[
                "doa"
            ][
                :,
                j
            ]


            base[
                f"C{c}_DOA_Star"
            ] = ads[
                "doa_star"
            ][
                :,
                j
            ]


            base[
                f"C{c}_Classical_PV"
            ] = classical[
                "pv"
            ][
                :,
                j
            ]


            base[
                f"C{c}_Classical_Weight"
            ] = classical[
                "weights"
            ][
                :,
                j
            ]


        base.to_excel(
            writer,
            sheet_name=(
                f"Base_{sheet_name}"
            ),
            index=False
        )


        # ==============================================================
        # z SWEEP
        # ==============================================================

        dimension_results = []


        for z in Z_VALUES:


            (
                score,
                pv,
                weights
            ) = ads_pv_weights_batch(
                ads,
                float(
                    z
                )
            )


            (
                weight_sum,
                max_weight,
                min_weight,
                hhi,
                effective_n,
                top_criterion
            ) = weight_statistics(
                weights
            )


            result = pd.DataFrame(
                {

                    "Global_Matrix_ID":
                        df[
                            "Global_Matrix_ID"
                        ].to_numpy(),

                    "Dimension":
                        sheet_name,

                    "Case_Type":
                        df[
                            "Case_Type"
                        ].to_numpy(),

                    "Constant_Criterion":
                        df[
                            "Constant_Criterion"
                        ].to_numpy(),

                    "Constant_Value":
                        df[
                            "Constant_Value"
                        ].to_numpy(),

                    "z":
                        float(
                            z
                        ),

                    "Weight_Sum":
                        weight_sum,

                    "Max_Weight":
                        max_weight,

                    "Min_Weight":
                        min_weight,

                    "Weight_HHI":
                        hhi,

                    "Effective_N_Criteria":
                        effective_n,

                    "Top_Criterion":
                        top_criterion
                }
            )


            for j in range(
                n_cols
            ):

                c = (
                    j + 1
                )


                result[
                    f"C{c}_Score"
                ] = score[
                    :,
                    j
                ]


                result[
                    f"C{c}_PV"
                ] = pv[
                    :,
                    j
                ]


                result[
                    f"C{c}_Weight"
                ] = weights[
                    :,
                    j
                ]


            dimension_results.append(
                result
            )


            compact = result[
                [
                    "Global_Matrix_ID",
                    "Dimension",
                    "Case_Type",
                    "Constant_Criterion",
                    "Constant_Value",
                    "z",
                    "Weight_Sum",
                    "Max_Weight",
                    "Min_Weight",
                    "Weight_HHI",
                    "Effective_N_Criteria",
                    "Top_Criterion"
                ]
            ].copy()


            global_summary_parts.append(
                compact
            )

            aggregate_input_parts.append(
                compact
            )


        dimension_results_df = pd.concat(
            dimension_results,
            ignore_index=True
        )


        dimension_results_df.to_excel(
            writer,
            sheet_name=(
                f"ADS_{sheet_name}"
            ),
            index=False
        )


        # ==============================================================
        # VALIDATION 1
        # NORMAL MATRICES: ADS z=1 = CLASSICAL LOPCOW
        # ==============================================================

        normal_mask = (
            ~classical[
                "failed"
            ]
        )


        if np.any(
            normal_mask
        ):


            (
                score_z1,
                pv_z1,
                weights_z1
            ) = ads_pv_weights_batch(
                ads,
                1.0
            )


            pv_diff = np.abs(

                pv_z1[
                    normal_mask,
                    :
                ]

                -

                classical[
                    "pv"
                ][
                    normal_mask,
                    :
                ]
            )


            weight_diff = np.abs(

                weights_z1[
                    normal_mask,
                    :
                ]

                -

                classical[
                    "weights"
                ][
                    normal_mask,
                    :
                ]
            )


            validation_parts.append(
                pd.DataFrame(
                    {

                        "Dimension":
                            [
                                sheet_name
                            ],

                        "Validation_Check":
                            [
                                (
                                    "Normal matrices: "
                                    "ADS z=1 equals classical LOPCOW"
                                )
                            ],

                        "N_Matrices":
                            [
                                int(
                                    np.sum(
                                        normal_mask
                                    )
                                )
                            ],

                        "Maximum_PV_Difference":
                            [
                                float(
                                    np.nanmax(
                                        pv_diff
                                    )
                                )
                            ],

                        "Maximum_Weight_Difference":
                            [
                                float(
                                    np.nanmax(
                                        weight_diff
                                    )
                                )
                            ],

                        "Pass":
                            [
                                bool(
                                    np.nanmax(
                                        weight_diff
                                    )
                                    < 1e-10
                                )
                            ]
                    }
                )
            )


        # ==============================================================
        # VALIDATION 2
        # RULE A1 SEMANTIC CASES
        # ==============================================================

        controlled_mask = (
            np.any(
                ads[
                    "rule_a1"
                ],
                axis=1
            )
        )


        semantic_records = []


        for matrix_idx in np.where(
            controlled_mask
        )[0]:


            a1_indices = np.where(
                ads[
                    "rule_a1"
                ][
                    matrix_idx,
                    :
                ]
            )[0]


            for j in a1_indices:


                raw_constant = float(
                    raw[
                        matrix_idx,
                        0,
                        j
                    ]
                )


                h = float(
                    ads[
                        "semantic_relative"
                    ][
                        matrix_idx,
                        j
                    ]
                )


                a1_score = (
                    1.0
                    + h
                )


                a1_pv = (
                    100.0
                    *
                    np.log(
                        a1_score
                    )
                )


                semantic_records.append(
                    {

                        "Global_Matrix_ID":
                            df.iloc[
                                matrix_idx
                            ][
                                "Global_Matrix_ID"
                            ],

                        "Dimension":
                            sheet_name,

                        "Criterion":
                            f"C{j+1}",

                        "Criterion_Type":
                            criterion_types[
                                j
                            ],

                        "Constant_Value":
                            raw_constant,

                        "Scale_Lower":
                            SEMANTIC_LOWER,

                        "Scale_Upper":
                            SEMANTIC_UPPER,

                        "A1_h":
                            h,

                        "A1_RMS":
                            h,

                        "A1_Score_1plusRMS":
                            a1_score,

                        "A1_PV":
                            a1_pv
                    }
                )


        if semantic_records:

            semantic_summary_parts.append(
                pd.DataFrame(
                    semantic_records
                )
            )


        print(
            f"Completed {sheet_name}: "
            f"{len(df):,} matrices × "
            f"{len(Z_VALUES)} z values"
        )


    # ==================================================================
    # GLOBAL SUMMARY
    # ==================================================================

    global_summary = pd.concat(
        global_summary_parts,
        ignore_index=True
    )


    global_summary.to_excel(
        writer,
        sheet_name="Global_Summary",
        index=False
    )


    # ==================================================================
    # AGGREGATE SUMMARY
    # ==================================================================

    aggregate_data = pd.concat(
        aggregate_input_parts,
        ignore_index=True
    )


    aggregate_summary = (

        aggregate_data

        .groupby(
            [
                "Dimension",
                "Case_Type",
                "z"
            ],
            as_index=False
        )

        .agg(

            N_Matrices=(
                "Global_Matrix_ID",
                "count"
            ),

            Mean_Max_Weight=(
                "Max_Weight",
                "mean"
            ),

            SD_Max_Weight=(
                "Max_Weight",
                "std"
            ),

            Mean_Min_Weight=(
                "Min_Weight",
                "mean"
            ),

            Mean_HHI=(
                "Weight_HHI",
                "mean"
            ),

            SD_HHI=(
                "Weight_HHI",
                "std"
            ),

            Mean_Effective_N=(
                "Effective_N_Criteria",
                "mean"
            )
        )
    )


    aggregate_summary.to_excel(
        writer,
        sheet_name="Aggregate_Summary",
        index=False
    )


    # ==================================================================
    # RULE COUNTS
    # ==================================================================

    rule_counts = pd.concat(
        rule_count_parts,
        ignore_index=True
    )


    rule_counts.to_excel(
        writer,
        sheet_name="Rule_Counts",
        index=False
    )


    # ==================================================================
    # RULE A1 SEMANTIC DETAILS
    # ==================================================================

    if semantic_summary_parts:


        semantic_summary = pd.concat(
            semantic_summary_parts,
            ignore_index=True
        )


        semantic_summary.to_excel(
            writer,
            sheet_name="A1_Semantic_Cases",
            index=False
        )


        semantic_by_value = (

            semantic_summary

            .groupby(
                [
                    "Constant_Value"
                ],
                as_index=False
            )

            .agg(

                N_Criteria=(
                    "Criterion",
                    "count"
                ),

                Mean_h=(
                    "A1_h",
                    "mean"
                ),

                Min_h=(
                    "A1_h",
                    "min"
                ),

                Max_h=(
                    "A1_h",
                    "max"
                ),

                Mean_A1_PV=(
                    "A1_PV",
                    "mean"
                ),

                Min_A1_PV=(
                    "A1_PV",
                    "min"
                ),

                Max_A1_PV=(
                    "A1_PV",
                    "max"
                )
            )
        )


        semantic_by_value.to_excel(
            writer,
            sheet_name="A1_By_Value",
            index=False
        )


    # ==================================================================
    # VALIDATION SUMMARY
    # ==================================================================

    if validation_parts:

        validation_df = pd.concat(
            validation_parts,
            ignore_index=True
        )

    else:

        validation_df = pd.DataFrame()


    validation_df.to_excel(
        writer,
        sheet_name="Validation",
        index=False
    )


    # ==================================================================
    # SIMPLE FORMATTING
    # ==================================================================

    workbook = writer.book


    header_format = workbook.add_format(
        {

            "bold":
                True,

            "border":
                1,

            "text_wrap":
                True,

            "valign":
                "top"
        }
    )


    for (
        worksheet_name,
        worksheet
    ) in writer.sheets.items():


        worksheet.freeze_panes(
            1,
            0
        )


        worksheet.set_column(
            0,
            100,
            17
        )


# ======================================================================
# 17. POST-RUN VALIDATION
# ======================================================================

print(
    "\n"
    + "=" * 72
)

print(
    "FULL CURRENT ADS-LOPCOW SIMULATION COMPLETE"
)

print(
    "=" * 72
)


# ----------------------------------------------------------------------
# Matrix count
# ----------------------------------------------------------------------

print(
    f"\nMatrices detected: "
    f"{total_matrix_count:,}"
)


if (
    total_matrix_count
    == 18000
):

    print(
        "18,000 matrix count: PASS"
    )

else:

    print(
        "WARNING: matrix count is not 18,000."
    )


# ----------------------------------------------------------------------
# Matrix-z row count
# ----------------------------------------------------------------------

expected_rows = (
    total_matrix_count
    *
    len(
        Z_VALUES
    )
)

actual_rows = len(
    global_summary
)


print(
    f"\nExpected matrix-z rows: "
    f"{expected_rows:,}"
)

print(
    f"Actual matrix-z rows:   "
    f"{actual_rows:,}"
)


if (
    actual_rows
    == expected_rows
):

    print(
        "Matrix-z row count: PASS"
    )

else:

    print(
        "Matrix-z row count: FAIL"
    )


# ----------------------------------------------------------------------
# Weight-sum validation
# ----------------------------------------------------------------------

valid_sums = (
    global_summary[
        "Weight_Sum"
    ]
    .dropna()
    .to_numpy(
        dtype=float
    )
)


if len(
    valid_sums
) > 0:

    max_sum_error = np.max(
        np.abs(
            valid_sums
            - 1.0
        )
    )

    print(
        f"\nMaximum |sum(weights)-1| = "
        f"{max_sum_error:.12f}"
    )

else:

    max_sum_error = np.nan

    print(
        "\nNo valid weight vectors found."
    )


# ----------------------------------------------------------------------
# Classical endpoint validation
# ----------------------------------------------------------------------

if (
    not validation_df.empty
):

    print(
        "\nClassical LOPCOW vs ADS z=1:"
    )

    print(
        validation_df.to_string(
            index=False
        )
    )


# ----------------------------------------------------------------------
# Rule A1 summary
# ----------------------------------------------------------------------

if semantic_summary_parts:

    print(
        "\nRule A1 summary by constant value:"
    )

    print(
        semantic_by_value.to_string(
            index=False
        )
    )


# ----------------------------------------------------------------------
# Theoretical endpoint checks for current 1–9 benefit simulation
# ----------------------------------------------------------------------

expected_h_1 = 0.0

expected_pv_1 = (
    100.0
    *
    np.log(
        1.0
    )
)

expected_h_9 = 1.0

expected_pv_9 = (
    100.0
    *
    np.log(
        2.0
    )
)


print(
    "\nExpected Rule A1 endpoints:"
)

print(
    f"Constant 1 -> "
    f"h={expected_h_1:.4f}, "
    f"PV={expected_pv_1:.6f}"
)

print(
    f"Constant 9 -> "
    f"h={expected_h_9:.4f}, "
    f"PV={expected_pv_9:.6f}"
)


# ----------------------------------------------------------------------
# Case counts
# ----------------------------------------------------------------------

print(
    "\nMatrix-z observations by Case_Type:"
)

print(
    global_summary[
        "Case_Type"
    ]
    .value_counts()
)


# ----------------------------------------------------------------------
# Output file
# ----------------------------------------------------------------------

print(
    f"\nSaved output workbook as:\n"
    f"{OUTPUT_FILE}"
)


files.download(
    OUTPUT_FILE
)
