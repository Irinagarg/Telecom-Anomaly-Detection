# 4G Network Intelligence System
Vodafone Idea (Vi) — Punjab Region | Nokia 4G KPI Analysis

## Project Overview
End-to-end ML pipeline for anomaly detection, health scoring,
KPI failure tracking, and natural language querying of Nokia
4G network KPI data across 300+ cells.

## ML Techniques
- Isolation Forest, One-Class SVM, LOF, Ensemble Voting
- SHAP Explainability
- Holt-Winters + ARIMA Forecasting
- FAISS + RAG + LLaMA 3.3 70B

## KPI Thresholds (Vi Network SLA)
| KPI | Threshold | Weight |
|-----|-----------|--------|
| ERAB Drop Rate | <= 0.5% | 30 pts |
| E2E Call Setup SR | >= 99% | 25 pts |
| VoLTE CSSR | >= 99% | 20 pts |
| CQI | >= 7 | 15 pts |
| Avg DL Throughput | >= 3 Mbps | 10 pts |

## Run KPI Tracker Web App
docker build -t kpi-tracker .
docker run -p 5000:5000 kpi-tracker

## Tech Stack
Python, scikit-learn, pandas, Flask, Docker, FAISS,
sentence-transformers, Groq API, openpyxl, statsmodels
