# Annual Technology Report 2025

## Executive Summary

This report provides a comprehensive overview of the technology initiatives undertaken by **ACME Corporation** during the fiscal year 2025. It covers strategic investments, key milestones, infrastructure upgrades, and recommendations for the upcoming period.

> Our digital transformation journey has delivered measurable business value across every department, reducing operational costs by **23 %** and improving customer satisfaction scores to an all-time high.

---

## 1. Strategic Objectives

### 1.1 Cloud Migration Programme

The organisation completed the migration of **87 %** of on-premises workloads to the cloud platform during Q3. Key achievements include:

- Decommissioned 14 legacy data-centre racks
- Reduced infrastructure spend by $1.2 M annually
- Achieved 99.97 % uptime SLA across all migrated services
- Onboarded 6 additional business units onto the unified cloud tenancy

### 1.2 Cybersecurity Posture

Following the external audit conducted in January, the Security Operations team implemented the following controls:

1. Zero-Trust Network Architecture rolled out to all remote workers
2. Multi-Factor Authentication enforced across 100 % of user accounts
3. Endpoint Detection & Response (EDR) deployed to 4,200 devices
4. Quarterly penetration testing programme established with a certified third party

### 1.3 Data & Analytics Platform

#### Data Governance

A new enterprise data catalogue was commissioned in H1, bringing structured metadata management to over **2,800 data assets**. Access policies are now enforced via attribute-based controls tied to the HR system.

#### Real-Time Streaming

The data engineering team delivered a Kafka-based streaming platform capable of processing 850,000 events per second. Downstream consumers include the fraud detection engine and the customer-360 personalisation service.

---

## 2. Infrastructure & Operations

### 2.1 Network Refresh

| Region       | Sites Updated | Bandwidth Increase | Completion |
|--------------|:-------------:|-------------------:|:----------:|
| North America | 12           | 10 Gbps → 40 Gbps  | Q1 2025    |
| Europe        | 8            | 1 Gbps → 10 Gbps   | Q2 2025    |
| Asia-Pacific  | 5            | 1 Gbps → 10 Gbps   | Q3 2025    |
| Latin America | 3            | 100 Mbps → 1 Gbps  | Q4 2025    |

### 2.2 DevOps Maturity

The platform engineering team standardised on a *GitOps* deployment model across all product squads. The resulting improvements are documented below.

**Before GitOps:**

- Average deployment frequency: 1.2 releases / week
- Mean time to recovery (MTTR): 4.5 hours
- Change failure rate: 18 %

**After GitOps:**

- Average deployment frequency: **14.7 releases / week**
- Mean time to recovery (MTTR): **22 minutes**
- Change failure rate: **3.1 %**

---

## 3. Software Development

### 3.1 API Standards

All new APIs must conform to the internal `REST-v2` specification. A sample endpoint definition is shown below:

```json
{
  "endpoint": "POST /api/v2/documents/process",
  "auth": "Bearer <token>",
  "request": {
    "template": "<base64-encoded .docx>",
    "markdown": "# My Document\n\nContent here..."
  },
  "response": {
    "documentId": "a3f1c9e2-...",
    "filename": "output_20250915_143022.docx",
    "outputSize": 48230
  }
}
```

### 3.2 Technology Radar

The following technologies were evaluated and classified this year:

- **Adopt**: Rust for systems programming, WebAssembly for edge compute, OpenTelemetry for observability
- **Trial**: LLM-assisted code review, WASI, Turso (libSQL)
- **Assess**: Deno Deploy, Bun runtime, Effect-TS
- **Hold**: Legacy SOAP services, jQuery for new projects, self-managed Kubernetes without GitOps

---

## 4. Talent & Culture

### 4.1 Headcount

The Technology division grew from **312** to **389** full-time employees during 2025, with a focus on:

1. Site Reliability Engineering (18 new hires)
2. Data Science & ML Engineering (22 new hires)
3. Cloud Infrastructure (15 new hires)
4. Product Security (12 new hires)

### 4.2 Training & Certification

Over **1,400 certifications** were achieved across the division, including:

- AWS Certified Solutions Architect – Professional: 47
- Google Cloud Professional Data Engineer: 31
- Certified Kubernetes Administrator (CKA): 28
- Certified Information Systems Security Professional (CISSP): 19

---

## 5. Financial Summary

| Category                   | Budget (USD)  | Actual (USD)  | Variance  |
|----------------------------|:-------------:|:-------------:|:---------:|
| Cloud Infrastructure       | 4,200,000     | 3,980,000     | +5.2 %    |
| Software Licences          | 1,800,000     | 1,920,000     | -6.7 %    |
| Personnel & Contractors    | 38,000,000    | 37,400,000    | +1.6 %    |
| Security & Compliance      | 2,100,000     | 2,350,000     | -11.9 %   |
| Training & Certifications  | 450,000       | 392,000       | +12.9 %   |
| **Total**                  | **46,550,000**| **46,042,000**| **+1.1 %**|

---

## 6. Risks & Issues

> **High Priority** — The end-of-life of the Oracle EBS ERP system in December 2026 requires an accelerated migration plan to be approved by the board no later than Q1 2026.

The following risks have been logged in the enterprise risk register:

1. **ERP Migration Delay** — High likelihood, critical impact
2. **Third-Party Supply Chain Attack** — Medium likelihood, high impact
3. **Talent Attrition in ML Team** — Medium likelihood, medium impact
4. **Regulatory Changes (AI Act)** — Low likelihood, high impact

---

## 7. Recommendations

Based on the analysis above, the Technology Leadership team recommends the following actions for FY2026:

1. **Approve** the ERP replacement programme budget of $8.4 M
2. **Expand** the Platform Engineering function by an additional 10 engineers
3. **Mandate** OpenTelemetry adoption across all product services by Q2 2026
4. **Establish** an AI Centre of Excellence to govern LLM usage and AI risk
5. **Commission** an independent review of the third-party software supply chain

---

## Appendix A — Glossary

**EDR** — Endpoint Detection & Response
**GitOps** — A set of practices that use Git pull requests as the single source of truth for declarative infrastructure and applications
**LLM** — Large Language Model
**MTTR** — Mean Time To Recovery
**SLA** — Service Level Agreement
**WASI** — WebAssembly System Interface

---

*Prepared by the Office of the Chief Technology Officer — ACME Corporation*
*Classification: Internal Use Only*
