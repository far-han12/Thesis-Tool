import json
from pathlib import Path
from collections import Counter, defaultdict
import statistics
import pandas as pd

RESULTS_DIR = Path("sample_results")
REPORT_DIR = Path("reports")
REPORT_DIR.mkdir(parents=True, exist_ok=True)

OUT_FILE = REPORT_DIR / "raw_json_transparency_report.xlsx"

TABLE_ROWS = {
    "region": 5,
    "nation": 25,
    "supplier": 1000,
    "customer": 15000,
    "part": 20000,
    "partsupp": 80000,
    "orders": 150000,
    "lineitem": 600000,
}

TARGET_FILES = [
    "autoadmin_balanced10_w2_b300_FIXED.json",
    "autoadmin_balanced25_w2_b500.json",
    "autoadmin_queryformer10_w1_b300.json",
    "drop_balanced10_w2_b300.json",
    "drop_balanced25_w2_b500.json",
    "drop_queryformer10_w1_b300.json",
    "extend_balanced10_w2_b300.json",
    "extend_balanced25_w2_b500.json",
    "extend_balanced42_w1_b500.json",
    "extend_perturb10_w2_b300.json",
    "extend_perturb25_w2_b500.json",
    "extend_queryformer10_w1_b300.json",
]


def load_raw_result(path):
    raw = json.load(open(path))
    top_key = list(raw.keys())[0]
    return top_key, raw[top_key]


def infer_advisor(filename):
    name = filename.lower()
    if "autoadmin" in name or "auto_admin" in name:
        return "AutoAdmin"
    if "extend" in name:
        return "Extend"
    if "drop" in name:
        return "Drop"
    return "Unknown"


def infer_run_type(filename):
    name = filename.lower()
    if "queryformer" in name:
        return "QueryFormer"
    if "perturb" in name:
        return "Perturbed"
    return "Optimizer"


def parse_index(index_text):
    if "#" not in index_text:
        return "", []
    table, cols = index_text.split("#", 1)
    return table, cols.split(",")


def cost_reduction(data):
    total_no = data.get("total_no_cost", sum(data.get("no_cost", [])))
    total_ind = data.get("total_ind_cost", sum(data.get("ind_cost", [])))
    if total_no == 0:
        return 0
    return (total_no - total_ind) / total_no * 100


def risk_level(index_count, row_count):
    if row_count <= 0:
        return "UNKNOWN", 0

    ratio = index_count / row_count * 1000

    if ratio > 100:
        return "CRITICAL", ratio
    if ratio > 10:
        return "HIGH", ratio
    if ratio > 1:
        return "MEDIUM", ratio
    return "LOW", ratio


def metric_1_query_impact(file_name, data):
    no_cost = data.get("no_cost", [])
    ind_cost = data.get("ind_cost", [])

    rows = []
    improvements = []

    for i, (nc, ic) in enumerate(zip(no_cost, ind_cost), start=1):
        improvement = ((nc - ic) / nc * 100) if nc else 0
        saving = nc - ic
        improvements.append(improvement)

        rows.append({
            "file_name": file_name,
            "advisor": infer_advisor(file_name),
            "run_type": infer_run_type(file_name),
            "query_no": i,
            "no_cost": nc,
            "ind_cost": ic,
            "absolute_saving": saving,
            "improvement_percent": improvement,
        })

    summary = {
        "query_count": len(rows),
        "queries_improved_gt_50": sum(x > 50 for x in improvements),
        "queries_improved_gt_25": sum(x > 25 for x in improvements),
        "queries_improved_gt_10": sum(x > 10 for x in improvements),
        "queries_improved_0_to_10": sum(0 < x <= 10 for x in improvements),
        "queries_no_improvement": sum(x <= 0 for x in improvements),
        "avg_query_improvement_percent": statistics.mean(improvements) if improvements else 0,
        "median_query_improvement_percent": statistics.median(improvements) if improvements else 0,
        "best_query_improvement_percent": max(improvements) if improvements else 0,
        "worst_query_improvement_percent": min(improvements) if improvements else 0,
    }

    return summary, rows


def metric_2_write_overhead(file_name, data):
    indexes = data.get("indexes", [])
    table_counts = Counter()

    for idx in indexes:
        table, cols = parse_index(idx)
        if table:
            table_counts[table] += 1

    rows = []

    for table, count in table_counts.most_common():
        row_count = TABLE_ROWS.get(table, 0)
        risk, ratio = risk_level(count, row_count)

        rows.append({
            "file_name": file_name,
            "advisor": infer_advisor(file_name),
            "run_type": infer_run_type(file_name),
            "table": table,
            "index_count": count,
            "row_count": row_count,
            "indexes_per_1000_rows": ratio,
            "risk_level": risk,
        })

    summary = {
        "total_indexes": len(indexes),
        "tables_affected": len(table_counts),
        "most_indexed_table": table_counts.most_common(1)[0][0] if table_counts else "",
        "most_indexed_table_count": table_counts.most_common(1)[0][1] if table_counts else 0,
        "critical_tables": sum(1 for r in rows if r["risk_level"] == "CRITICAL"),
        "high_risk_tables": sum(1 for r in rows if r["risk_level"] == "HIGH"),
        "medium_risk_tables": sum(1 for r in rows if r["risk_level"] == "MEDIUM"),
        "low_risk_tables": sum(1 for r in rows if r["risk_level"] == "LOW"),
    }

    return summary, rows


def metric_3_redundancy(file_name, data):
    indexes = data.get("indexes", [])

    exact_counts = Counter(indexes)
    exact_duplicates = sum(v - 1 for v in exact_counts.values() if v > 1)

    permutation_groups = defaultdict(list)

    for idx in indexes:
        table, cols = parse_index(idx)
        if not table:
            continue
        key = (table, tuple(sorted(cols)))
        permutation_groups[key].append(idx)

    rows = []
    redundant_count = 0

    for (table, sorted_cols), group in permutation_groups.items():
        if len(group) > 1:
            redundant = len(group) - 1
            redundant_count += redundant

            rows.append({
                "file_name": file_name,
                "advisor": infer_advisor(file_name),
                "run_type": infer_run_type(file_name),
                "table": table,
                "column_set_sorted": ",".join(sorted_cols),
                "group_size": len(group),
                "redundant_count_estimate": redundant,
                "indexes": " | ".join(group),
            })

    total_indexes = len(indexes)

    summary = {
        "exact_duplicate_indexes": exact_duplicates,
        "redundant_permutation_indexes": redundant_count,
        "essential_indexes_estimate": total_indexes - redundant_count,
        "redundancy_percent": (redundant_count / total_indexes * 100) if total_indexes else 0,
        "redundancy_groups": len(rows),
    }

    return summary, rows


def index_width_distribution(data):
    indexes = data.get("indexes", [])
    width_counts = Counter()

    for idx in indexes:
        table, cols = parse_index(idx)
        if table:
            width_counts[len(cols)] += 1

    return {
        "single_column_indexes": width_counts.get(1, 0),
        "two_column_indexes": width_counts.get(2, 0),
        "three_column_indexes": width_counts.get(3, 0),
        "width_distribution": "; ".join([f"w{k}:{v}" for k, v in sorted(width_counts.items())]),
    }


def sel_info_summary(data):
    sel_info = data.get("sel_info", {})

    cache_hits = sel_info.get("cache_hits", "")
    cost_requests = sel_info.get("cost_requests", "")
    steps = sel_info.get("step", [])

    if isinstance(cache_hits, list):
        cache_hits = cache_hits[-1] if cache_hits else 0

    if isinstance(cost_requests, list):
        cost_requests = cost_requests[-1] if cost_requests else 0

    if isinstance(steps, list):
        step_count = len(steps)
    else:
        step_count = ""

    return {
        "cache_hits": cache_hits,
        "cost_requests": cost_requests,
        "selection_steps_count": step_count,
    }


def metric_4_drift(original_name, original_data, perturbed_name, perturbed_data):
    original_indexes = set(original_data.get("indexes", []))
    perturbed_indexes = set(perturbed_data.get("indexes", []))

    common = original_indexes & perturbed_indexes
    only_original = original_indexes - perturbed_indexes
    only_perturbed = perturbed_indexes - original_indexes
    union = original_indexes | perturbed_indexes

    jaccard = len(common) / len(union) * 100 if union else 0
    original_basis = len(common) / len(original_indexes) * 100 if original_indexes else 0
    perturbed_basis = len(common) / len(perturbed_indexes) * 100 if perturbed_indexes else 0

    summary = {
        "original_file": original_name,
        "perturbed_file": perturbed_name,
        "original_index_count": len(original_indexes),
        "perturbed_index_count": len(perturbed_indexes),
        "common_index_count": len(common),
        "only_original_count": len(only_original),
        "only_perturbed_count": len(only_perturbed),
        "union_index_count": len(union),
        "jaccard_similarity_percent": jaccard,
        "original_basis_stability_percent": original_basis,
        "perturbed_basis_stability_percent": perturbed_basis,
        "drift_sensitivity_percent": 100 - jaccard,
    }

    rows = []

    for idx in sorted(common):
        rows.append({
            "original_file": original_name,
            "perturbed_file": perturbed_name,
            "group": "common",
            "index": idx,
        })

    for idx in sorted(only_original):
        rows.append({
            "original_file": original_name,
            "perturbed_file": perturbed_name,
            "group": "only_original",
            "index": idx,
        })

    for idx in sorted(only_perturbed):
        rows.append({
            "original_file": original_name,
            "perturbed_file": perturbed_name,
            "group": "only_perturbed",
            "index": idx,
        })

    return summary, rows


def main():
    all_run_summary = []
    all_query_rows = []
    all_write_rows = []
    all_redundancy_rows = []

    loaded_results = {}

    for file_name in TARGET_FILES:
        path = RESULTS_DIR / file_name

        if not path.exists():
            print("Missing:", file_name)
            continue

        print("Reading raw JSON:", file_name)

        try:
            top_key, data = load_raw_result(path)
        except Exception as e:
            print("Failed:", file_name, e)
            continue

        loaded_results[file_name] = data

        config = data.get("config", {})

        m1_summary, m1_rows = metric_1_query_impact(file_name, data)
        m2_summary, m2_rows = metric_2_write_overhead(file_name, data)
        m3_summary, m3_rows = metric_3_redundancy(file_name, data)
        width_summary = index_width_distribution(data)
        sel_summary = sel_info_summary(data)

        total_no = data.get("total_no_cost", sum(data.get("no_cost", [])))
        total_ind = data.get("total_ind_cost", sum(data.get("ind_cost", [])))

        run_summary = {
            "file_name": file_name,
            "top_key": top_key,
            "advisor": infer_advisor(file_name),
            "run_type": infer_run_type(file_name),
            "budget_MB": config.get("budget_MB", ""),
            "max_index_width": config.get("max_index_width", ""),
            "max_indexes": config.get("max_indexes", ""),
            "constraint": config.get("constraint", ""),
            "cost_estimation": config.get("cost_estimation", "optimizer"),
            "workload_count": len(data.get("workload", [])),
            "query_cost_count": len(data.get("no_cost", [])),
            "total_no_cost": total_no,
            "total_ind_cost": total_ind,
            "total_cost_saving": total_no - total_ind,
            "aggregate_cost_reduction_percent": cost_reduction(data),
            **m1_summary,
            **m2_summary,
            **m3_summary,
            **width_summary,
            **sel_summary,
        }

        all_run_summary.append(run_summary)
        all_query_rows.extend(m1_rows)
        all_write_rows.extend(m2_rows)
        all_redundancy_rows.extend(m3_rows)

    drift_summaries = []
    drift_rows = []

    drift_pairs = [
        (
            "extend_balanced10_w2_b300.json",
            "extend_perturb10_w2_b300.json",
        ),
        (
            "extend_balanced25_w2_b500.json",
            "extend_perturb25_w2_b500.json",
        ),
    ]

    for original, perturbed in drift_pairs:
        if original in loaded_results and perturbed in loaded_results:
            summary, rows = metric_4_drift(
                original,
                loaded_results[original],
                perturbed,
                loaded_results[perturbed],
            )
            drift_summaries.append(summary)
            drift_rows.extend(rows)

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame(all_run_summary).to_excel(writer, sheet_name="All Runs Summary", index=False)
        pd.DataFrame(all_query_rows).to_excel(writer, sheet_name="Metric1_QueryImpact", index=False)
        pd.DataFrame(all_write_rows).to_excel(writer, sheet_name="Metric2_WriteRisk", index=False)
        pd.DataFrame(all_redundancy_rows).to_excel(writer, sheet_name="Metric3_Redundancy", index=False)
        pd.DataFrame(drift_summaries).to_excel(writer, sheet_name="Metric4_DriftSummary", index=False)
        pd.DataFrame(drift_rows).to_excel(writer, sheet_name="Metric4_DriftIndexes", index=False)

    print()
    print("Created:", OUT_FILE)
    print("Runs processed:", len(all_run_summary))
    print("Query impact rows:", len(all_query_rows))
    print("Write risk rows:", len(all_write_rows))
    print("Redundancy rows:", len(all_redundancy_rows))
    print("Drift comparisons:", len(drift_summaries))


if __name__ == "__main__":
    main()
