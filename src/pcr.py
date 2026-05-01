"""
PCR — Performance Calculator and Ranking

Multi-variable run-scoring method. Each run is scored against its peer
group via percentile-rank on each metric, weighted-summed into a 0–100
composite score, then bucketed into bands.

Methodology: docs/pcr.md
Config:      config/pcr.yaml

Public entry point:
    rank_runs(df, config) -> DataFrame with PCR columns appended
"""

from __future__ import annotations

import pandas as pd
import numpy as np


# =====================================================================
# Preprocessing
# =====================================================================

def add_footage_bucket(df, config):
    """Add FOOTAGE_BUCKET column based on Phase_CALC + TOTAL_DRILL cutoffs.
    Multi-phase rows (or unknown phases) get null bucket."""
    by_phase = config.get("footage_bucket", {}).get("by_phase", {})

    def bucket(row):
        phase = row.get("Phase_CALC")
        drill = row.get("TOTAL_DRILL")
        if pd.isna(phase) or pd.isna(drill):
            return None
        cutoffs = by_phase.get(phase)
        if not cutoffs:
            return None  # multi-phase or unknown
        if drill < cutoffs["short_max"]:
            return "short"
        if drill <= cutoffs["long_min"]:
            return "medium"
        return "long"

    df = df.copy()
    df["FOOTAGE_BUCKET"] = df.apply(bucket, axis=1)
    return df


def filter_pool(df, config):
    """Restrict to the absolute baseline pool and drop rows with required-null columns."""
    abs_baseline = config.get("absolute_baseline", {})
    if abs_baseline.get("scope") == "from_date":
        start_date = pd.Timestamp(abs_baseline["start_date"])
        if "DATE_OUT" in df.columns:
            df = df[df["DATE_OUT"] >= start_date].copy()

    filters = config.get("filters", {}).get("exclude_when", {})
    for col in filters.get("null_columns", []):
        if col in df.columns:
            df = df[df[col].notna()].copy()

    sources_excl = filters.get("sources_to_exclude", [])
    if sources_excl and "SOURCE" in df.columns:
        df = df[~df["SOURCE"].isin(sources_excl)].copy()

    return df


# =====================================================================
# Peer-group lookup
# =====================================================================

def find_peer_group(target_idx, pool, config):
    """Walk the fallback chain. Return (peer_indices, level_id, n_peers).
    Empty Index, None, 0 if no level produces enough peers."""
    target = pool.loc[target_idx]
    levels = config.get("peer_group", {}).get("fallback_levels", [])
    min_n = config.get("peer_group", {}).get("min_peer_runs", 10)

    for level in levels:
        keys = level["keys"]
        if any(pd.isna(target.get(k)) for k in keys):
            continue
        mask = pd.Series(True, index=pool.index)
        for k in keys:
            mask = mask & (pool[k] == target.get(k))
        peers_idx = pool.index[mask]
        peers_idx = peers_idx[peers_idx != target_idx]
        if len(peers_idx) >= min_n:
            return peers_idx, level["id"], len(peers_idx)
    return pd.Index([]), None, 0


# =====================================================================
# Scoring
# =====================================================================

def percentile_rank(value, peer_values, direction="higher_is_better"):
    """Percentile rank (0-100) of `value` among `peer_values`.
    Uses the average method: % strictly less + 0.5 * % equal."""
    if pd.isna(value):
        return None
    peers = pd.to_numeric(peer_values, errors="coerce").dropna()
    if len(peers) == 0:
        return None
    n = len(peers)
    less = (peers < value).sum()
    equal = (peers == value).sum()
    pct = float((less + 0.5 * equal) / n * 100)
    if direction == "lower_is_better":
        pct = 100.0 - pct
    return pct


def _active_weights_for(target, variables):
    """Return a dict {metric_key: normalized_weight} for metrics not skipped on target.
    Weights are renormalized so the active set sums to 1.0."""
    active = {}
    for k, var in variables.items():
        skip = var.get("skip_when")
        if skip:
            col = skip.get("column")
            eq = skip.get("equals")
            if col in target.index and target.get(col) == eq:
                continue
        active[k] = float(var.get("weight", 0))
    total = sum(active.values())
    if total == 0:
        return {}
    return {k: w / total for k, w in active.items()}


def assign_band(score, all_scores_sorted, bands_config):
    """Assign a band to a single score based on its percentile within the
    full set of scores."""
    if pd.isna(score):
        return None
    n = len(all_scores_sorted)
    if n == 0:
        return None
    # rank percentile = fraction of scores strictly less than this score
    pos = np.searchsorted(all_scores_sorted, score, side="left")
    rank_pct = pos / n
    if rank_pct >= bands_config.get("excellent", 0.85):
        return "Excellent"
    if rank_pct >= bands_config.get("good", 0.65):
        return "Good"
    if rank_pct >= bands_config.get("average", 0.35):
        return "Average"
    return "Poor"


# =====================================================================
# Main entry point
# =====================================================================

def rank_runs(df, config):
    """Score every eligible run in df. Return a DataFrame with PCR columns appended."""
    output_cfg = config.get("output", {})
    score_col = output_cfg.get("score_col", "pcr_score")
    rank_col = output_cfg.get("rank_col", "pcr_rank")
    band_col = output_cfg.get("band_col", "pcr_band")
    peer_level_col = output_cfg.get("peer_level_col", "pcr_peer_level")
    peer_n_col = output_cfg.get("peer_n_col", "pcr_peer_n")
    metric_prefix = output_cfg.get("per_metric_prefix", "pcr_")
    rolling_avg_col = output_cfg.get("rolling_avg_col", "peer_avg_pcr_365d")
    delta_col = output_cfg.get("delta_col", "delta_vs_365d")

    if "DATE_OUT" in df.columns and not pd.api.types.is_datetime64_any_dtype(df["DATE_OUT"]):
        df = df.copy()
        df["DATE_OUT"] = pd.to_datetime(df["DATE_OUT"], errors="coerce")

    pool = add_footage_bucket(df, config)
    pool = filter_pool(pool, config)

    variables = config.get("variables", {})
    metric_keys = list(variables.keys())

    pcr_scores = {}
    peer_levels = {}
    peer_ns = {}
    metric_pcts = {k: {} for k in metric_keys}

    for target_idx in pool.index:
        target = pool.loc[target_idx]
        peers_idx, level_id, n_peers = find_peer_group(target_idx, pool, config)
        if n_peers == 0:
            continue
        peers = pool.loc[peers_idx]

        weights = _active_weights_for(target, variables)
        if not weights:
            continue

        composite = 0.0
        any_metric_scored = False
        for k, w in weights.items():
            var = variables[k]
            col = var["column"]
            direction = var.get("direction", "higher_is_better")
            value = target.get(col)
            peer_vals = peers[col] if col in peers.columns else pd.Series(dtype=float)
            pct = percentile_rank(value, peer_vals, direction)
            metric_pcts[k][target_idx] = pct
            if pct is not None:
                composite += w * pct
                any_metric_scored = True

        if not any_metric_scored:
            continue

        pcr_scores[target_idx] = composite
        peer_levels[target_idx] = level_id
        peer_ns[target_idx] = n_peers

    out = pool.copy()
    out[score_col] = out.index.map(pcr_scores)
    out[peer_level_col] = out.index.map(peer_levels)
    out[peer_n_col] = out.index.map(peer_ns)
    for k in metric_keys:
        out[f"{metric_prefix}{k}_pct"] = out.index.map(metric_pcts[k])

    out[rank_col] = out[score_col].rank(method="min", ascending=False)

    # Bands
    sorted_scores = np.sort(out[score_col].dropna().values)
    out[band_col] = out[score_col].apply(
        lambda s: assign_band(s, sorted_scores, config.get("bands", {}))
    )

    # Rolling 365-day peer-avg PCR
    rolling_cfg = config.get("rolling_baseline", {})
    window_days = rolling_cfg.get("window_days", 365)
    min_rolling_peers = rolling_cfg.get("min_peer_runs_for_rolling", 3)

    rolling_avgs = {}
    for target_idx in out.index:
        if pd.isna(out.at[target_idx, score_col]):
            continue
        target = out.loc[target_idx]
        peers_idx, _, _ = find_peer_group(target_idx, pool, config)
        if len(peers_idx) == 0:
            continue
        peers = out.loc[peers_idx]
        target_date = target.get("DATE_OUT")
        if "DATE_OUT" not in peers.columns or pd.isna(target_date):
            continue
        cutoff = target_date - pd.Timedelta(days=window_days)
        peers_in_window = peers[
            (peers["DATE_OUT"] >= cutoff)
            & (peers["DATE_OUT"] <= target_date)
            & peers[score_col].notna()
        ]
        if len(peers_in_window) >= min_rolling_peers:
            rolling_avgs[target_idx] = float(peers_in_window[score_col].mean())

    out[rolling_avg_col] = out.index.map(rolling_avgs)
    out[delta_col] = out[score_col] - out[rolling_avg_col]

    return out
