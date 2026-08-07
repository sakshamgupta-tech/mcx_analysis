import io

import pandas as pd
import streamlit as st

st.set_page_config(page_title="mcx_analysis", layout="wide")
st.title("mcx_analysis")

uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Upload a file to begin.")
    st.stop()

# ---- Load the 'Legs' sheet ----
try:
    xls = pd.ExcelFile(uploaded_file)
except Exception as e:
    st.error(f"Could not read the file: {e}")
    st.stop()

sheet_names = xls.sheet_names
default_sheet = next((s for s in sheet_names if "leg" in s.lower()), sheet_names[0])

# ---- Column mapping (auto-guess, user can override) ----
def guess_column(columns, candidates):
    for cand in candidates:
        for col in columns:
            if col.strip().lower() == cand.lower():
                return col
    for cand in candidates:
        for col in columns:
            if cand.lower() in col.strip().lower():
                return col
    return columns[0]

def parse_time_column(series):
    """Parse a time-of-day column ("HH:MM:SS.fff" or "HH:MM:SS") so every value
    lands on the SAME reference date (1900-01-01). Mixing parse paths that assign
    different default dates (e.g. one branch defaulting to today's real date) makes
    values incomparable — a value with a real 2026 date would always look "later"
    than a 1900-01-01 value regardless of actual time-of-day, silently corrupting
    any max()/min() comparison across rows. """
    s = series.astype(str)
    parsed = pd.to_datetime(s, format="%H:%M:%S.%f", errors="coerce")

    missing = parsed.isna()
    if missing.any():
        parsed.loc[missing] = pd.to_datetime(s[missing], format="%H:%M:%S", errors="coerce")

    missing = parsed.isna()
    if missing.any():
        generic = pd.to_datetime(s[missing], errors="coerce")
        # Strip whatever date the generic parser assigned and re-anchor to the
        # same 1900-01-01 reference date used by the explicit-format branches above.
        time_only = generic - generic.dt.normalize() + pd.Timestamp("1900-01-01")
        parsed.loc[missing] = time_only

    return parsed


with st.expander("Advanced: sheet & column mapping (auto-detected — override only if needed)"):
    sheet_name = st.selectbox("Sheet", sheet_names, index=sheet_names.index(default_sheet))

    df = xls.parse(sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    cols = list(df.columns)

    st.caption(f"Loaded {len(df):,} rows from '{sheet_name}'.")
    with st.expander("Preview raw data"):
        st.dataframe(df.head(20), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        portfolio_col = st.selectbox(
            "Portfolio Name column",
            cols,
            index=cols.index(guess_column(cols, ["Portfolio Name", "Portfolio"])),
        )
        time_col = st.selectbox(
            "Entry Time column",
            cols,
            index=cols.index(guess_column(cols, ["Entry Time", "Order Entry Time", "Time"])),
        )
        exit_time_col = st.selectbox(
            "Exit Time column",
            cols,
            index=cols.index(guess_column(cols, ["Exit Time"])),
        )
        exit_type_col = st.selectbox(
            "Exit Type column",
            cols,
            index=cols.index(guess_column(cols, ["Exit Type"])),
        )
        status_col = st.selectbox(
            "Status column",
            cols,
            index=cols.index(guess_column(cols, ["Status"])),
        )
    with c2:
        user_id_col = st.selectbox(
            "User ID column",
            cols,
            index=cols.index(guess_column(cols, ["User ID", "UserID"])),
        )
        user_name_col = st.selectbox(
            "User Name column (display name)",
            cols,
            index=cols.index(guess_column(cols, ["User Alias", "User Name", "UserName", "Alias"])),
        )
        pnl_col = st.selectbox(
            "PNL column",
            cols,
            index=cols.index(guess_column(cols, ["PNL", "P&L", "PnL"])),
        )

    status_values = sorted(df[status_col].dropna().astype(str).unique().tolist())
    # Default to both "Completed" and "EntryCompleted" when present — a leg that only
    # got its entry filled (EntryCompleted) still contributes real PNL and should not
    # be silently dropped from the calculation.
    default_statuses = [s for s in status_values if s.strip().lower() in ("completed", "entrycompleted")]
    if not default_statuses and status_values:
        default_statuses = [status_values[0]]
    status_filters = st.multiselect("Only include rows with Status in", status_values, default=default_statuses)
    status_filter = " + ".join(status_filters) if status_filters else ""

# ---- Prepare working data ----
work = df[
    [portfolio_col, user_id_col, user_name_col, time_col, exit_time_col, exit_type_col, pnl_col, status_col]
].copy()
work.columns = ["Portfolio Name", "User ID", "User Name", "Entry Time", "Exit Time", "Exit Type", "PNL", "Status"]
work = work.dropna(subset=["Portfolio Name", "User ID", "Entry Time", "Status"])
# Strip stray whitespace so joins against other sheets/files (Grid Log, MultiLeg Orders)
# match on value, not on incidental leading/trailing spaces.
work["Portfolio Name"] = work["Portfolio Name"].astype(str).str.strip()
work["User ID"] = work["User ID"].astype(str).str.strip()

all_portfolios = sorted(work["Portfolio Name"].unique())

# ---- Filter to the selected Status(es) FIRST ----
status_filters_lower = [s.strip().lower() for s in status_filters]
work = work[work["Status"].astype(str).str.strip().str.lower().isin(status_filters_lower)]

if work.empty:
    st.error(f"No rows with Status in {status_filters}. Check the column mapping / filter above.")
    st.stop()

work["Entry Time Parsed"] = parse_time_column(work["Entry Time"])

bad_rows = work["Entry Time Parsed"].isna().sum()
if bad_rows:
    st.warning(f"{bad_rows} row(s) had an Entry Time that couldn't be parsed and were skipped.")
work = work.dropna(subset=["Entry Time Parsed"])

if work.empty:
    st.error("No usable rows after parsing Entry Time. Check the column mapping above.")
    st.stop()

# ---- Start: per portfolio, earliest Entry Time among Status-filtered rows, and its user ----
idx_min = work.groupby("Portfolio Name")["Entry Time Parsed"].idxmin()
start_found = work.loc[idx_min, ["Portfolio Name", "User ID", "Entry Time Parsed"]].copy()
start_found = start_found.rename(columns={"User ID": "Start User ID", "Entry Time Parsed": "Start Time"})
start_found["Start Time"] = start_found["Start Time"].dt.strftime("%H:%M:%S")

# =========================================================================
# End / sq-off: per portfolio,
#   1. For each user, take their LAST leg (max Exit Time) row — including its Exit Type.
#   2. Across users of that portfolio, take the row with the SMALLEST of those max times.
#      That's the portfolio's End Time — the fastest user to finish
#      exiting all their legs — and that user (End User ID) + Exit Type is shown alongside it.
#   This is independent of Max Portfolio User ID (that's a separate ranking, used only for PnL).
# =========================================================================
exit_work = work.copy()
exit_work["Exit Time Parsed"] = parse_time_column(exit_work["Exit Time"])
exit_work = exit_work.dropna(subset=["Exit Time Parsed"])

end_found = pd.DataFrame(columns=["Portfolio Name", "End User ID", "End Time", "End User Exit Type"])
if not exit_work.empty:
    # Step 1: each user's last-leg row (max Exit Time) per portfolio
    idx_last_leg = exit_work.groupby(["Portfolio Name", "User ID"])["Exit Time Parsed"].idxmax()
    per_user_exit = exit_work.loc[idx_last_leg, ["Portfolio Name", "User ID", "Exit Time Parsed", "Exit Type"]]

    # Step 2: End Time = smallest of those per-user last-exit times, per portfolio
    idx_sqoff = per_user_exit.groupby("Portfolio Name")["Exit Time Parsed"].idxmin()
    end_found = per_user_exit.loc[idx_sqoff, ["Portfolio Name", "User ID", "Exit Time Parsed", "Exit Type"]].copy()
    end_found = end_found.rename(
        columns={"User ID": "End User ID", "Exit Time Parsed": "End Time", "Exit Type": "End User Exit Type"}
    )
    end_found["End Time"] = end_found["End Time"].dt.strftime("%H:%M:%S")

# ---- Combine Start + End into one table: Portfolio Name, Start Time, Start User,
# End Time, End User, End User Exit Type. PnL / Max Portfolio User ID / Exit Type
# (Max Portfolio User's own Exit Type) get merged in later. ----
result = pd.DataFrame({"Portfolio Name": all_portfolios})
result = result.merge(start_found, on="Portfolio Name", how="left")
result = result.merge(end_found, on="Portfolio Name", how="left")
fill_cols = ["Start User ID", "Start Time", "End User ID", "End Time", "End User Exit Type"]
result[fill_cols] = result[fill_cols].fillna("-")

# =========================================================================
# Portfolio PnL by "primary" user:
#   1. From the 'MultiLeg Orders' sheet, filter to the chosen Status
#      (e.g. COMPLETE) and count the unique Portfolio Names — that's the
#      total portfolios considered.
#   2. Count how many of those distinct portfolios each user appears in;
#      rank users by that count, highest first.
#   3. The top-ranked (max-portfolio) user is assigned every portfolio
#      they appear in. Whichever portfolios are left go to the
#      next-ranked (2nd max) user who appears in them, and so on, until
#      every portfolio is assigned to exactly one user.
#   4. That assigned (Max Portfolio) user's PnL for each portfolio is the SUM
#      of their leg PNL from the 'Legs' sheet (Status-filtered) for that
#      portfolio. (End Time / End User ID / Exit Type are computed separately
#      above, independent of this ranking.)
# =========================================================================
mlo_sheet_name = next((s for s in sheet_names if "multileg" in s.lower()), None)

if mlo_sheet_name is None:
    st.error("Could not find a 'MultiLeg Orders' sheet in this file.")
else:
    mlo_df = xls.parse(mlo_sheet_name)
    mlo_df.columns = [str(c).strip() for c in mlo_df.columns]
    mlo_cols = list(mlo_df.columns)

    with st.expander("MultiLeg Orders sheet — column mapping"):
        mc1, mc2 = st.columns(2)
        with mc1:
            mlo_portfolio_col = st.selectbox(
                "Portfolio Name column",
                mlo_cols,
                index=mlo_cols.index(guess_column(mlo_cols, ["Portfolio Name", "Portfolio"])),
                key="mlo_portfolio_col",
            )
            mlo_user_col = st.selectbox(
                "User ID column",
                mlo_cols,
                index=mlo_cols.index(guess_column(mlo_cols, ["User ID", "UserID"])),
                key="mlo_user_col",
            )
        with mc2:
            mlo_status_col = st.selectbox(
                "Status column",
                mlo_cols,
                index=mlo_cols.index(guess_column(mlo_cols, ["Status"])),
                key="mlo_status_col",
            )
            mlo_status_values = sorted(mlo_df[mlo_status_col].dropna().astype(str).unique().tolist())
            mlo_default_status = next(
                (s for s in mlo_status_values if s.strip().lower() in ("complete", "completed")),
                mlo_status_values[0] if mlo_status_values else None,
            )
            mlo_status_filter = st.selectbox(
                "Only include rows with Status =",
                mlo_status_values,
                index=mlo_status_values.index(mlo_default_status) if mlo_default_status else 0,
                key="mlo_status_filter",
            )

    mlo_work = mlo_df[[mlo_portfolio_col, mlo_user_col, mlo_status_col]].copy()
    mlo_work.columns = ["Portfolio Name", "User ID", "Status"]
    mlo_work = mlo_work.dropna(subset=["Portfolio Name", "User ID", "Status"])
    mlo_work["Portfolio Name"] = mlo_work["Portfolio Name"].astype(str).str.strip()
    mlo_work["User ID"] = mlo_work["User ID"].astype(str).str.strip()

    mlo_all_portfolios = sorted(mlo_work["Portfolio Name"].unique())
    mlo_ran = mlo_work[mlo_work["Status"].astype(str).str.strip().str.lower() == mlo_status_filter.strip().lower()]

    coverage = mlo_ran.groupby("User ID")["Portfolio Name"].apply(set).to_dict()
    user_portfolio_counts = {u: len(p) for u, p in coverage.items()}
    ranked_users = sorted(user_portfolio_counts, key=lambda u: (-user_portfolio_counts[u], u))

    unassigned = set(mlo_ran["Portfolio Name"].unique())
    assignment_rows = []
    user_of_portfolio = {}
    for rank, u in enumerate(ranked_users, start=1):
        got = coverage[u] & unassigned
        if got:
            assignment_rows.append({"Rank": rank, "User ID": u, "Portfolios Assigned": len(got)})
            for p in got:
                user_of_portfolio[p] = u
        unassigned -= got

    with st.expander("How portfolios were assigned to a primary user"):
        st.dataframe(pd.DataFrame(assignment_rows), use_container_width=True)

    # ---- Look up each assigned user's PNL for that portfolio by summing their leg PNL
    # from the 'Legs' sheet (the Status-filtered `work` table built earlier already has
    # Portfolio Name, User ID and PNL for every leg). ----
    leg_pnl_by_portfolio_user = work.groupby(["Portfolio Name", "User ID"])["PNL"].sum()
    pnl_lookup = leg_pnl_by_portfolio_user.to_dict()

    pnl_rows = []
    for p in mlo_all_portfolios:
        u = user_of_portfolio.get(p)
        if u is None:
            pnl_rows.append({"Portfolio Name": p, "Max Portfolio User ID": "-", "PnL": 0})
        else:
            pnl_rows.append({"Portfolio Name": p, "Max Portfolio User ID": u, "PnL": pnl_lookup.get((p, u), 0.0)})

    pnl_result = pd.DataFrame(pnl_rows)

    # =====================================================================
    # One combined table, keyed on the unique Portfolio Name: Start/End time
    # info (independent "fastest to finish" logic) + PnL by primary user,
    # merged side by side.
    # =====================================================================
    combined = result.merge(pnl_result, on="Portfolio Name", how="outer")

    fill_cols_combined = ["Start Time", "Start User ID", "End Time", "End User ID", "End User Exit Type", "Max Portfolio User ID"]
    combined["_missing"] = combined["Start Time"].isna() & combined["Max Portfolio User ID"].isna()
    combined_missing_count = combined["_missing"].sum()
    combined[fill_cols_combined] = combined[fill_cols_combined].fillna("-")
    combined["PnL"] = combined["PnL"].fillna(0)

    # ---- Exit Type = the Max Portfolio User's own last leg (max Exit Time) Exit
    # Type for that portfolio — same user as Max Portfolio User ID / PnL, not End
    # User ID's. End User Exit Type (a separate column) keeps the End User ID's own
    # value for comparison. (per_user_exit was built earlier in the End/sq-off section.) ----
    max_user_exit_type = per_user_exit.set_index(["Portfolio Name", "User ID"])["Exit Type"].to_dict()

    def base_exit_type(row):
        return max_user_exit_type.get((row["Portfolio Name"], row["Max Portfolio User ID"]), "-")

    combined["Exit Type"] = combined.apply(base_exit_type, axis=1)

    combined = combined.sort_values(["_missing", "Portfolio Name"]).drop(columns="_missing").reset_index(drop=True)
    combined = combined[
        [
            "Portfolio Name", "Start User ID", "Start Time", "End User ID", "End Time", "End User Exit Type",
            "Max Portfolio User ID", "PnL", "Exit Type",
        ]
    ]

    # =====================================================================
    # Grid Log (separate file): if the log has a "Combined SL" entry for a
    # portfolio's End User ID, that overrides the Exit Type shown for that
    # portfolio with "Combined SL".
    # =====================================================================
    st.subheader("Grid Log — Combined SL override (optional)")
    grid_log_file = st.file_uploader("Upload the Grid Log CSV file", type=["csv"], key="grid_log_file")

    if grid_log_file is not None:
        try:
            grid_df = pd.read_csv(grid_log_file)
        except Exception as e:
            st.error(f"Could not read the Grid Log file: {e}")
            grid_df = None

        if grid_df is not None:
            grid_df.columns = [str(c).strip() for c in grid_df.columns]
            grid_cols = list(grid_df.columns)

            with st.expander("Grid Log — column mapping"):
                gc1, gc2 = st.columns(2)
                with gc1:
                    grid_message_col = st.selectbox(
                        "Message column (e.g. 'Combined SL: -2500 hit for User: ...; Option Portfolio: ...;')",
                        grid_cols,
                        index=grid_cols.index(
                            guess_column(grid_cols, ["Message", "Remark", "Remarks", "Log", "Description", "Event", "Reason"])
                        ),
                        key="grid_message_col",
                    )
                    grid_time_col = st.selectbox(
                        "Log time column (e.g. '15:30:21:267')",
                        grid_cols,
                        index=grid_cols.index(guess_column(grid_cols, ["Time", "Timestamp", "Log Time", "Date Time"])),
                        key="grid_time_col",
                    )
                with gc2:
                    grid_user_col = st.selectbox(
                        "User ID column",
                        grid_cols,
                        index=grid_cols.index(guess_column(grid_cols, ["UserID", "User ID"])),
                        key="grid_user_col",
                    )
                    grid_portfolio_col = st.selectbox(
                        "Portfolio column",
                        grid_cols,
                        index=grid_cols.index(guess_column(grid_cols, ["Option Portfolio", "Portfolio Name", "Portfolio"])),
                        key="grid_portfolio_col",
                    )

            grid_work = grid_df[[grid_time_col, grid_message_col, grid_user_col, grid_portfolio_col]].copy()
            grid_work.columns = ["Time", "Message", "User ID", "Portfolio Name"]
            grid_work = grid_work.dropna(subset=["Time", "Message", "User ID", "Portfolio Name"])
            grid_work["Message"] = grid_work["Message"].astype(str)
            grid_work["User ID"] = grid_work["User ID"].astype(str).str.strip()
            grid_work["Portfolio Name"] = grid_work["Portfolio Name"].astype(str).str.strip()

            # ---- Grid Log times are "HH:MM:SS:fff" (colon before the milliseconds,
            # not a dot) — normalize to "HH:MM:SS.fff" so parse_time_column can read it. ----
            grid_time_str = grid_work["Time"].astype(str).str.strip()
            four_part = grid_time_str.str.count(":") == 3
            grid_time_str.loc[four_part] = grid_time_str[four_part].str.replace(r":(\d+)$", r".\1", regex=True)
            # ---- The report only ever displays/compares time down to whole seconds
            # (e.g. "16:20:25", no milliseconds) — so floor the Grid Log timestamp to
            # the same second-level precision before comparing, instead of comparing
            # at millisecond precision the report never actually shows. ----
            grid_work["Time Parsed"] = parse_time_column(grid_time_str).dt.floor("s")

            # ---- Only rows where "Combined SL" is at the START of the message — this
            # excludes unrelated lines like "Start Monitoring ... FOR Combined SL" where
            # the phrase only appears at the end. Portfolio Name / User ID come directly
            # from their own CSV columns, not parsed out of the message text. ----
            combined_sl_rows = grid_work[grid_work["Message"].str.strip().str.match(r"(?i)^combined sl")].copy()

            # ---- Only a valid trigger if it happened BETWEEN that specific user's own
            # entry into the portfolio and their own exit (their own first leg's Entry
            # Time and last leg's Exit Time, from the Legs sheet) — not the portfolio's
            # overall Start/End Time, which may belong to a different user. A Grid Log
            # line timestamped before the user entered, OR after they'd already exited,
            # can't be the real reason for that exit. Both sides are floored to whole
            # seconds to match the precision the report actually shows. ----
            user_entry_time = work.groupby(["Portfolio Name", "User ID"])["Entry Time Parsed"].min().dt.floor("s").to_dict()
            user_exit_time = exit_work.groupby(["Portfolio Name", "User ID"])["Exit Time Parsed"].max().dt.floor("s").to_dict()
            combined_sl_rows["User Entry Time"] = combined_sl_rows.apply(
                lambda r: user_entry_time.get((r["Portfolio Name"], r["User ID"])), axis=1
            )
            combined_sl_rows["User Exit Time"] = combined_sl_rows.apply(
                lambda r: user_exit_time.get((r["Portfolio Name"], r["User ID"])), axis=1
            )
            combined_sl_rows = combined_sl_rows[
                combined_sl_rows["User Entry Time"].notna()
                & combined_sl_rows["User Exit Time"].notna()
                & (combined_sl_rows["Time Parsed"] >= combined_sl_rows["User Entry Time"])
                & (combined_sl_rows["Time Parsed"] <= combined_sl_rows["User Exit Time"])
            ]
            combined_sl_rows = combined_sl_rows[["Portfolio Name", "User ID", "Time", "Time Parsed", "Message"]]

            # ---- If a (Portfolio, User) has more than one valid Combined SL line, use
            # the LATEST one — the trigger closest to their actual exit — for the
            # timestamp shown alongside "Combined SL" in the report. ----
            idx_latest_sl = combined_sl_rows.groupby(["Portfolio Name", "User ID"])["Time Parsed"].idxmax()
            latest_sl = combined_sl_rows.loc[idx_latest_sl]
            combined_sl_timestamp = {
                (row["Portfolio Name"], row["User ID"]): row["Time"] for _, row in latest_sl.iterrows()
            }
            combined_sl_pairs = set(combined_sl_timestamp.keys())

            combined_sl_rows = combined_sl_rows.drop(columns="Time Parsed")

            def make_combined_sl_applier(user_col, value_col):
                def apply_combined_sl(row):
                    key = (row["Portfolio Name"], row[user_col])
                    if key in combined_sl_pairs:
                        return f"Combined SL ({combined_sl_timestamp[key]})"
                    return row[value_col]
                return apply_combined_sl

            # Exit Type = Max Portfolio User ID's own Combined SL check (already in place).
            combined["Exit Type"] = combined.apply(
                make_combined_sl_applier("Max Portfolio User ID", "Exit Type"), axis=1
            )
            # End User Exit Type = the SAME Combined SL check, but against End User ID instead.
            combined["End User Exit Type"] = combined.apply(
                make_combined_sl_applier("End User ID", "End User Exit Type"), axis=1
            )

            # ---- A separate file/table with ONLY the Grid Log lines that are Combined SL ----
            with st.expander(f"Grid Log rows — Combined SL only ({len(combined_sl_pairs)} found)"):
                st.dataframe(combined_sl_rows.reset_index(drop=True), use_container_width=True)

                combined_sl_csv_bytes = combined_sl_rows.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "Download Combined SL grid log rows as CSV",
                    data=combined_sl_csv_bytes,
                    file_name="grid_log_combined_sl_only.csv",
                    mime="text/csv",
                )

    st.subheader("Portfolio report — start/end time + PnL by primary user")
    if combined_missing_count:
        st.warning(f"{combined_missing_count} portfolio(s) had no data under the current filters and are marked accordingly.")
    st.dataframe(combined, use_container_width=True)

    # ---- Excel download: Sheet 1 = the combined report, Sheet 2 = the ranking that
    # decided each portfolio's Max Portfolio User ID. ----
    rank_df = pd.DataFrame(assignment_rows)

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        combined.to_excel(writer, sheet_name="Portfolio Report", index=False)
        rank_df.to_excel(writer, sheet_name="Rank", index=False)

    st.download_button(
        "Download combined report as Excel",
        data=excel_buffer.getvalue(),
        file_name="portfolio_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
