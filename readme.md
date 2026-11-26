# MailFlow AX

### AI-Powered Mailbox Exchange

MailFlow AX is an enterprise-oriented mailbox automation application designed to monitor Outlook mailboxes, classify incoming emails using a local AI engine, allocate emails to appropriate users or folders, move emails automatically, and track subsequent actions.

The solution combines **Outlook automation, local AI classification, automated mailbox allocation, action tracking, and Excel-based logging** into a single workflow.

---

## Overview

MailFlow AX follows the workflow:

**Monitor → Classify → Allocate → Move → Track → Log**

The application continuously monitors selected Outlook folders, identifies new or recovered emails, sends relevant email content to a local AI classifier, determines the appropriate processing path, performs mailbox movement/allocation, and records the resulting activity.

---

## Key Capabilities

### 🤖 AI-Based Email Classification
Uses a local AI engine to classify mailbox emails without requiring cloud-based AI processing.

### 📬 Automated Mail Allocation
Routes classified emails to the appropriate users, queues, or target folders according to the configured allocation logic.

### 📁 Automated Mail Movement
Moves processed emails into their designated Outlook folders after classification and allocation.

### ✅ Action Tracking
Tracks downstream email actions including:

- Reply
- Reply-All
- Forward

### 📊 Operations Dashboard
Provides a consolidated view of mailbox-processing activity, including:

- Total Emails
- Allocated
- Actioned
- Pending
- Action Rate
- AI Confidence

### ⚡ Live Activity Monitoring
Displays real-time processing activity, including mailbox monitoring, classification, recovery scans, and email-processing events.

### 📑 Excel Logging & Reporting
Maintains processing and action information in an Excel-based log for operational tracking and reporting.

### 🔄 Recovery & Checkpoint Processing
Supports recovery scanning and checkpoint-based processing to reduce the risk of missing emails after interruptions or restarts.

---

## Application Flow

```text
Outlook Mailbox
      │
      ▼
Email Monitoring
      │
      ▼
AI Classification
      │
      ▼
Allocation Logic
      │
      ▼
Move to Target Folder
      │
      ▼
Action Tracking
      │
      ▼
Excel Logging & Dashboard
```

---

## Technology Stack

- **Python**
- **Microsoft Outlook / Outlook COM**
- **Local AI / Ollama**
- **OpenPyXL**
- **Pandas**
- **PyWin32**
- **Excel**
- **Desktop GUI**
- **Power BI / reporting integration where applicable**

---

## Architecture

MailFlow AX is designed around a modular processing pipeline:

### Mailbox Connector
Connects to the local Outlook environment and monitors configured mailbox folders.

### AI Classification Layer
Processes email information using a locally hosted AI model.

### Allocation Engine
Applies allocation rules and determines the destination user or folder.

### Mail Movement Layer
Performs the required Outlook folder movement.

### Action Tracker
Monitors whether processed emails receive a Reply, Reply-All, or Forward action.

### Logging Layer
Stores processing and action information for auditability, reporting, and dashboard consumption.

### Recovery Layer
Uses checkpoints and processed-message tracking to support restart and recovery scenarios.

---

## Current Interface

The application provides two primary views:

### Dashboard
A high-level operational view containing mailbox-processing KPIs and system status.

### Live Activity
A real-time execution view showing monitoring events, classification activity, recovery information, and processing logs.

---

## Example Metrics

The dashboard can surface operational indicators such as:

| Metric | Description |
|---|---|
| Total Emails | Emails identified by the processing pipeline |
| Allocated | Emails successfully allocated |
| Actioned | Emails with a tracked downstream action |
| Pending | Emails still awaiting action |
| Action Rate | Percentage of processed emails with an action |
| AI Confidence | Confidence returned by the classification process |

---

## Project Structure

The repository contains the MailFlow AX application source, supporting resources, configuration files, documentation, and demonstration assets.

> The exact repository tree may evolve as additional modules and deployment resources are added.

---

## Requirements

Typical requirements include:

- Windows environment
- Microsoft Outlook desktop application
- Python 3.x
- Required Python packages listed in the project dependency configuration
- Ollama or another supported local AI runtime
- Access to the configured Outlook mailbox
- Microsoft Excel for log/report access

---

## Getting Started

### 1. Clone the repository

```bash
git clone <repository-url>
cd MailFlow-AX
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure the local AI environment

Ensure the required Ollama model is installed and available locally.

### 4. Configure Outlook and project settings

Update the project configuration according to the mailbox, allocation folders, users, and processing requirements.

### 5. Run the application

```bash
python <entry-point>.py
```

> The exact entry-point filename should match the current repository structure.

---

## Important Notes

MailFlow AX is currently a project/application implementation and may contain environment-specific configuration.

Before deploying in another environment, review:

- Outlook mailbox configuration
- Target folders
- Allocation rules
- AI model configuration
- Excel log locations
- User mappings
- Recovery/checkpoint settings

---

## Security & Privacy

MailFlow AX can process mailbox content and operational email information.

Do **not** commit the following to a public repository:

- Real mailbox credentials
- Access tokens
- API keys
- Passwords
- Private configuration files
- Production email logs
- Confidential customer/company data
- Personally identifiable information

Use sanitized sample data when publishing screenshots, logs, demonstrations, or test files.

---

## Roadmap

Potential future enhancements include:

- Microsoft Graph integration for production deployments
- Expanded allocation/configuration management
- Improved AI classification and confidence handling
- Enhanced analytics and reporting
- Additional mailbox providers/connectors
- Centralized configuration management
- Deployment packaging and installer support

---

## Project Status

**Active Development**

MailFlow AX is being continuously enhanced with new automation, AI, monitoring, reporting, and operational capabilities.

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

## Author

**Varad Patil**

Built as part of ongoing work in **automation, AI-enabled operations, mailbox processing, and process improvement**.
