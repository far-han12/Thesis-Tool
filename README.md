# Index_EAB Transparency Tool

This repository contains a Python-based transparency analysis tool for Index_EAB result JSON files.

## Thesis Context

This tool is part of the undergraduate thesis:

**HCI-Based What-If Index Advisor: Hypothetical Planning and Metric Evaluation for Database Optimization**

The tool does not generate new index recommendations. Instead, it reads existing Index_EAB result JSON files and translates them into DBA-facing operational metrics.

## Research Purpose

Existing index advisors usually report aggregate cost reduction and recommended indexes. However, this hides operationally important information a DBA needs before deployment.

This tool extracts four transparency metrics:

1. Per-query cost delta distribution
2. Write overhead risk per table
3. Index redundancy analysis
4. Workload drift sensitivity score

## Folder Structure

```text
src/
  raw_json_transparency_tool.py

sample_results/
  Index_EAB result JSON files

reports/
  Generated Excel reports
# Thesis-Tool
