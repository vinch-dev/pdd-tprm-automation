# TPRM Last-Mile Automation: SAP-Excel-Outlook
Conceptual framework for an automation system designed to bridge ERP data silos with automated compliance reporting


## Project Overview
In Third-Party Risk Management, manual data entry is a significant bottleneck. This project documents a "last-mile" automation workflow designed to bridge the gap between **SAP (ERP)** and **Outlook (Reporting)**.

## Workflow Diagram 
```mermaid
graph TD
    subgraph "External Systems (SAP ERP)"
        A[Button Trigger] --> B(VBA Keyword Processing)
        B --> C(SAP Script Execution)
    end
    subgraph "Local Processing (Excel RPA)"
        C --> D(Custom Report Generation)
        D --> E{File Already Open?}
        E -->|Yes| G[Prompt: Close File]
        G --> E
        E -->|No| H(Pricing Logic Check)
    end
    subgraph "Compliance & Output"
        H --> I{Tariff Available?}
        I -->|Yes| J[Apply Price]
        I -->|No| K[Flag for Manual Review]
        J --> L[Final Excel Review]
        K --> L
    end
    subgraph "Mail List & Report"
        L--> M(Outlook mail template draft)
        M--> N{Vendor recorded?}
        N -->|Yes| O[Apply emails receiver]
        N -->|No| P[Flag for Manual Review]
        O --> Q[Final Outlook draft Review]
        P --> Q
    end
```
## Business Impact
* **Efficiency:** Reduce 80% of manual workflow filing.
* **Accuracy:** Eliminated copy-paste errors in sensitive pricing.
* **Compliance:** Ensured real-time notification for non-agreed rates.

---
*Note: This repository contains logic frameworks and pseudocode to demonstrate technical proficiency. Proprietary company code is not included.*
