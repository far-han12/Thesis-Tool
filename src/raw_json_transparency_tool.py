import json
from pathlib import Path
from collections import Counter, defaultdict
import statistics
import pandas as pd

RESULTS_DIR = Path("sample_results")
REPORT_DIR = Path("reports")
REPORT_DIR.mkdir(parents=True, exist_ok=True)

TEXT_REPORT = REPORT_DIR / "dba_transparency_report.txt"
EXCEL_REPORT = REPORT_DIR / "cross_advisor_comparison.xlsx"

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


def r2(value):
    try:
        return round(float(value), 2)
    except Exception:
        return value


def pct(value):
    try:
        return f"{float(value):.2f}%"
    except Exception:
        return str(value)


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


def infer_run_group(filename):
    name = filename.lower()

    if "balanced10" in name and "w2" in name and "b300" in name:
        return "10q_w2_b300"

    if "balanced25" in name and "w2" in name and "b500" in name:
        return "25q_w2_b500"

    if "queryformer10" in name and "w1" in name and "b300" in name:
        return "queryformer_10q_w1_b300"

    if "perturb10" in name and "w2" in name and "b300" in name:
        return "perturb_10q_w2_b300"

    if "perturb25" in name and "w2" in name and "b500" in name:
        return "perturb_25q_w2_b500"

    if "balanced42" in name:
        return "42q_w1_b500_baseline"

    return "other"


def infer_run_type(filename, config):
    name = filename.lower()

    if config.get("cost_estimation") == "queryformer" or "queryformer" in name:
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


def query_impact(data):
    no_cost = data.get("no_cost", [])
    ind_cost = data.get("ind_cost", [])

    rows = []
    improvements = []

    for i, (nc, ic) in enumerate(zip(no_cost, ind_cost), start=1):
        improvement = ((nc - ic) / nc * 100) if nc else 0
        improvements.append(improvement)

        rows.append({
            "query_no": i,
            "no_cost": r2(nc),
            "ind_cost": r2(ic),
            "absolute_saving": r2(nc - ic),
            "improvement_percent": r2(improvement),
        })

    zero_benefit = [row["query_no"] for row in rows if row["improvement_percent"] <= 0]
    top_beneficiaries = sorted(rows, key=lambda x: x["improvement_percent"], reverse=True)[:5]

    summary = {
        "query_count": len(rows),
        "queries_gt_50": sum(x > 50 for x in improvements),
        "queries_10_to_50": sum(10 < x <= 50 for x in improvements),
        "queries_0_to_10": sum(0 < x <= 10 for x in improvements),
        "queries_zero_benefit": sum(x <= 0 for x in improvements),
        "avg_improvement": r2(statistics.mean(improvements)) if improvements else 0,
        "median_improvement": r2(statistics.median(improvements)) if improvements else 0,
        "best_improvement": r2(max(improvements)) if improvements else 0,
        "worst_improvement": r2(min(improvements)) if improvements else 0,
        "zero_benefit_queries": zero_benefit,
        "top_beneficiaries": top_beneficiaries,
        "rows": rows,
    }

    return summary


def write_overhead(data):
    indexes = data.get("indexes", [])
    table_counts = Counter()

    for idx in indexes:
        table, _ = parse_index(idx)
        if table:
            table_counts[table] += 1

    rows = []

    for table, count in table_counts.most_common():
        row_count = TABLE_ROWS.get(table, 0)
        risk, ratio = risk_level(count, row_count)

        rows.append({
            "table": table,
            "index_count": count,
            "row_count": row_count,
            "indexes_per_1000_rows": r2(ratio),
            "risk_level": risk,
        })

    summary = {
        "total_indexes": len(indexes),
        "tables_affected": len(table_counts),
        "critical_tables": sum(1 for x in rows if x["risk_level"] == "CRITICAL"),
        "high_tables": sum(1 for x in rows if x["risk_level"] == "HIGH"),
        "medium_tables": sum(1 for x in rows if x["risk_level"] == "MEDIUM"),
        "low_tables": sum(1 for x in rows if x["risk_level"] == "LOW"),
        "most_indexed_table": table_counts.most_common(1)[0][0] if table_counts else "",
        "most_indexed_table_count": table_counts.most_common(1)[0][1] if table_counts else 0,
        "rows": rows,
    }

    return summary


def redundancy_analysis(data):
    indexes = data.get("indexes", [])

    exact_counts = Counter(indexes)
    exact_duplicate_count = sum(v - 1 for v in exact_counts.values() if v > 1)

    permutation_groups = defaultdict(list)

    for idx in indexes:
        table, cols = parse_index(idx)
        if not table:
            continue
        key = (table, tuple(sorted(cols)))
        permutation_groups[key].append(idx)

    permutation_rows = []
    permutation_redundant = 0

    for (table, sorted_cols), group in permutation_groups.items():
        if len(group) > 1:
            redundant = len(group) - 1
            permutation_redundant += redundant

            permutation_rows.append({
                "type": "PERMUTATION",
                "table": table,
                "column_set": ",".join(sorted_cols),
                "group_size": len(group),
                "redundant_count": redundant,
                "indexes": " | ".join(group),
            })

    # Prefix redundancy:
    # If table(a,b) exists, table(a) is prefix-redundant.
    parsed = []

    for idx in indexes:
        table, cols = parse_index(idx)
        if table:
            parsed.append((idx, table, tuple(cols)))

    prefix_rows = []
    prefix_redundant_indexes = set()

    for idx_a, table_a, cols_a in parsed:
        for idx_b, table_b, cols_b in parsed:
            if idx_a == idx_b:
                continue

            if table_a != table_b:
                continue

            if len(cols_a) >= len(cols_b):
                continue

            if cols_b[:len(cols_a)] == cols_a:
                prefix_redundant_indexes.add(idx_a)
                prefix_rows.append({
                    "type": "PREFIX",
                    "table": table_a,
                    "redundant_index": idx_a,
                    "covered_by_index": idx_b,
                    "reason": f"{idx_b} has prefix {idx_a}",
                })
                break

    safe_to_remove = set()

    for row in permutation_rows:
        indexes_in_group = row["indexes"].split(" | ")
        safe_to_remove.update(indexes_in_group[1:])

    safe_to_remove.update(prefix_redundant_indexes)

    total_indexes = len(indexes)
    safe_remove_count = len(safe_to_remove)
    essential_count = total_indexes - safe_remove_count

    summary = {
        "total_indexes": total_indexes,
        "exact_duplicate_indexes": exact_duplicate_count,
        "redundant_permutation_indexes": permutation_redundant,
        "prefix_redundant_indexes": len(prefix_redundant_indexes),
        "safe_to_remove_indexes": safe_remove_count,
        "essential_indexes": essential_count,
        "redundancy_percent": r2((safe_remove_count / total_indexes * 100) if total_indexes else 0),
        "permutation_rows": permutation_rows,
        "prefix_rows": prefix_rows,
        "safe_to_remove_list": sorted(safe_to_remove),
    }

    return summary


def drift_analysis(original_name, original_data, perturbed_name, perturbed_data):
    original_indexes = set(original_data.get("indexes", []))
    perturbed_indexes = set(perturbed_data.get("indexes", []))

    stable = original_indexes & perturbed_indexes
    fragile = original_indexes - perturbed_indexes
    new = perturbed_indexes - original_indexes
    union = original_indexes | perturbed_indexes

    jaccard = len(stable) / len(union) * 100 if union else 0
    drift = 100 - jaccard

    summary = {
        "original_file": original_name,
        "perturbed_file": perturbed_name,
        "original_index_count": len(original_indexes),
        "perturbed_index_count": len(perturbed_indexes),
        "stable_indexes": len(stable),
        "fragile_indexes": len(fragile),
        "new_indexes": len(new),
        "jaccard_similarity_percent": r2(jaccard),
        "drift_sensitivity_percent": r2(drift),
    }

    rows = []

    for idx in sorted(stable):
        rows.append({"label": "STABLE", "index": idx})

    for idx in sorted(fragile):
        rows.append({"label": "FRAGILE", "index": idx})

    for idx in sorted(new):
        rows.append({"label": "NEW", "index": idx})

    return summary, rows


def advisor_report_text(file_name, data, drift_lookup=None):
    config = data.get("config", {})
    advisor = infer_advisor(file_name)
    run_type = infer_run_type(file_name, config)

    q = query_impact(data)
    w = write_overhead(data)
    r = redundancy_analysis(data)

    aggregate = cost_reduction(data)
    workload_count = len(data.get("workload", []))
    width = config.get("max_index_width", "")
    budget = config.get("budget_MB", "")

    lines = []

    lines.append("=" * 70)
    lines.append("TRANSPARENCY REPORT")
    lines.append(f"Advisor: {advisor} | Run type: {run_type} | Workload: {workload_count} queries | Width: {width} | Budget: {budget}MB")
    lines.append("=" * 70)
    lines.append("")
    lines.append("METRIC 1 - QUERY IMPACT")
    lines.append(f"Aggregate cost reduction: {pct(aggregate)}")
    lines.append(f"Queries with >50% benefit: {q['queries_gt_50']} of {q['query_count']}")
    lines.append(f"Queries with 10-50% benefit: {q['queries_10_to_50']} of {q['query_count']}")
    lines.append(f"Queries with 0-10% benefit: {q['queries_0_to_10']} of {q['query_count']}")
    lines.append(f"Queries with zero benefit: {q['queries_zero_benefit']} of {q['query_count']}")

    if q["top_beneficiaries"]:
        lines.append("Top beneficiaries:")
        for row in q["top_beneficiaries"][:3]:
            lines.append(f"  Query {row['query_no']}: {pct(row['improvement_percent'])} improvement")

    if q["zero_benefit_queries"]:
        zero_list = ", ".join([f"Query {x}" for x in q["zero_benefit_queries"][:10]])
        lines.append(f"No-benefit queries: {zero_list}")
    else:
        lines.append("No-benefit queries: none")

    lines.append("")
    lines.append("METRIC 2 - WRITE OVERHEAD RISK")
    lines.append(f"Total indexes recommended: {w['total_indexes']}")

    risky_rows = [x for x in w["rows"] if x["risk_level"] in ["CRITICAL", "HIGH"]]

    if risky_rows:
        for row in risky_rows:
            symbol = "WARNING"
            lines.append(
                f"{symbol}: {row['risk_level']} risk on {row['table']} - "
                f"{row['index_count']} indexes, {row['indexes_per_1000_rows']} indexes per 1000 rows"
            )
    else:
        lines.append("No HIGH or CRITICAL write-risk tables detected.")

    lines.append("Interpretation: these tables must maintain extra index structures on every INSERT/UPDATE/DELETE.")
    lines.append("The original advisor output does not surface this write overhead warning.")

    lines.append("")
    lines.append("METRIC 3 - REDUNDANCY")
    lines.append(f"Redundant permutation indexes: {r['redundant_permutation_indexes']} of {r['total_indexes']}")
    lines.append(f"Prefix-redundant indexes: {r['prefix_redundant_indexes']} of {r['total_indexes']}")
    lines.append(f"Essential indexes: {r['essential_indexes']} of {r['total_indexes']}")
    lines.append(f"Safe-to-remove estimate: {r['safe_to_remove_indexes']} indexes")
    lines.append(f"Redundancy percentage: {pct(r['redundancy_percent'])}")
    lines.append("Interpretation: redundant indexes may increase storage and write cost without adding proportional benefit.")

    if drift_lookup and file_name in drift_lookup:
        d = drift_lookup[file_name]
        lines.append("")
        lines.append("METRIC 4 - DRIFT SENSITIVITY")
        lines.append(f"Compared with: {d['perturbed_file']}")
        lines.append(f"Jaccard similarity: {pct(d['jaccard_similarity_percent'])}")
        lines.append(f"Drift sensitivity score: {pct(d['drift_sensitivity_percent'])}")
        lines.append(f"Stable indexes: {d['stable_indexes']}")
        lines.append(f"Fragile indexes: {d['fragile_indexes']}")
        lines.append(f"New indexes after perturbation: {d['new_indexes']}")
        lines.append("Recommendation: do not over-invest in fragile indexes without DBA review.")

    lines.append("")
    lines.append("WHAT THE ADVISOR TOLD YOU:")
    lines.append(f"- {pct(aggregate)} aggregate optimizer-estimated cost reduction")
    lines.append("")
    lines.append("WHAT IT DID NOT TELL YOU:")
    lines.append(f"- {q['queries_zero_benefit']} queries get zero benefit")
    lines.append(f"- {w['critical_tables']} tables face CRITICAL write risk")
    lines.append(f"- {r['safe_to_remove_indexes']} indexes are potentially redundant")
    if drift_lookup and file_name in drift_lookup:
        d = drift_lookup[file_name]
        lines.append(f"- {pct(d['drift_sensitivity_percent'])} recommendation drift risk")
    lines.append("=" * 70)
    lines.append("")

    return "\n".join(lines)


def load_all_results():
    loaded = {}

    for file_name in TARGET_FILES:
        path = RESULTS_DIR / file_name

        if not path.exists():
            print(f"Missing: {file_name}")
            continue

        _, data = load_raw_result(path)
        loaded[file_name] = data

    return loaded


def build_cross_advisor_rows(loaded):
    rows = []

    comparison_groups = {
        "10q_w2_b300": [
            "extend_balanced10_w2_b300.json",
            "drop_balanced10_w2_b300.json",
            "autoadmin_balanced10_w2_b300_FIXED.json",
        ],
        "25q_w2_b500": [
            "extend_balanced25_w2_b500.json",
            "drop_balanced25_w2_b500.json",
            "autoadmin_balanced25_w2_b500.json",
        ],
        "queryformer_10q_w1_b300": [
            "extend_queryformer10_w1_b300.json",
            "drop_queryformer10_w1_b300.json",
            "autoadmin_queryformer10_w1_b300.json",
        ],
    }

    for group_name, files in comparison_groups.items():
        metrics = {}

        for file_name in files:
            if file_name not in loaded:
                continue

            data = loaded[file_name]
            advisor = infer_advisor(file_name)

            q = query_impact(data)
            w = write_overhead(data)
            r = redundancy_analysis(data)

            metrics[advisor] = {
                "Aggregate cost reduction": pct(cost_reduction(data)),
                "Total indexes": w["total_indexes"],
                "Queries with zero benefit": q["queries_zero_benefit"],
                "CRITICAL risk tables": w["critical_tables"],
                "HIGH risk tables": w["high_tables"],
                "Safe-to-remove redundant indexes": r["safe_to_remove_indexes"],
                "Redundancy %": pct(r["redundancy_percent"]),
                "Most indexed table": w["most_indexed_table"],
            }

        metric_names = [
            "Aggregate cost reduction",
            "Total indexes",
            "Queries with zero benefit",
            "CRITICAL risk tables",
            "HIGH risk tables",
            "Safe-to-remove redundant indexes",
            "Redundancy %",
            "Most indexed table",
        ]

        for metric in metric_names:
            rows.append({
                "Comparison group": group_name,
                "Metric": metric,
                "Extend": metrics.get("Extend", {}).get(metric, "-"),
                "Drop": metrics.get("Drop", {}).get(metric, "-"),
                "AutoAdmin": metrics.get("AutoAdmin", {}).get(metric, "-"),
            })

    return rows


def main():
    loaded = load_all_results()

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

    drift_lookup = {}
    drift_summary_rows = []
    drift_index_rows = []

    for original, perturbed in drift_pairs:
        if original in loaded and perturbed in loaded:
            summary, rows = drift_analysis(original, loaded[original], perturbed, loaded[perturbed])
            drift_lookup[original] = summary
            drift_summary_rows.append(summary)

            for row in rows:
                row["original_file"] = original
                row["perturbed_file"] = perturbed
                drift_index_rows.append(row)

    text_sections = []

    for file_name, data in loaded.items():
        text_sections.append(advisor_report_text(file_name, data, drift_lookup=drift_lookup))

    TEXT_REPORT.write_text("\n".join(text_sections))

    cross_rows = build_cross_advisor_rows(loaded)

    with pd.ExcelWriter(EXCEL_REPORT, engine="openpyxl") as writer:
        pd.DataFrame(cross_rows).to_excel(writer, sheet_name="Cross Advisor Comparison", index=False)
        pd.DataFrame(drift_summary_rows).to_excel(writer, sheet_name="Drift Summary", index=False)
        pd.DataFrame(drift_index_rows).to_excel(writer, sheet_name="Drift Indexes", index=False)

    print("Created:", TEXT_REPORT)
    print("Created:", EXCEL_REPORT)
    print("Runs processed:", len(loaded))
    print("Cross-advisor rows:", len(cross_rows))
    print("Drift comparisons:", len(drift_summary_rows))


if __name__ == "__main__":
    main()
