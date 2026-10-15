#!/usr/bin/env bash
# Rebuilds the synthetic dataset and the formula-driven Excel model.
set -e
python3 step1_build_model.py
python3 step2_customer_profitability.py
rm -f tmp.xlsx state.pkl
echo "Built Commercial_Pricing_Profitability.xlsx (open in Excel; formulas recalculate on open)"
