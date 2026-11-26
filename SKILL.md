---
name: mailflow-ax
description: Automate Outlook mailbox processing by monitoring email folders, classifying messages with a local AI model, allocating emails to configured users or folders, moving messages, tracking Reply/Reply-All/Forward actions, and maintaining operational logs.
summary: AI-powered Outlook mailbox monitoring, classification, allocation, movement, action tracking, and Excel logging.
version: "1.0.0"
license: MIT
tags:
  - email-automation
  - outlook
  - mailbox
  - ai
  - classification
  - automation
  - python
keywords:
  - Outlook automation
  - email classification
  - mailbox monitoring
  - email allocation
  - mail movement
  - action tracking
  - Ollama
  - Excel logging
use_cases:
  - Monitor Outlook inbox and sent items for new or recovered emails
  - Classify emails using a local AI model
  - Allocate emails to configured users or folders
  - Move processed emails automatically
  - Track Reply, Reply-All, and Forward actions
  - Maintain operational email-processing logs
---

# MailFlow AX

## Purpose

Use this skill when the user needs to automate operational mailbox processing through Outlook.

MailFlow AX follows the processing model:

**Monitor → Classify → Allocate → Move → Track → Log**

The goal is to reduce repetitive mailbox handling while maintaining visibility into processing status, allocation, downstream actions, and recovery events.

---

## When to Use This Skill

Activate this skill when the task involves:

- Monitoring Outlook mailboxes
- Detecting new or recovered emails
- Classifying email content
- Routing or allocating emails
- Moving emails between Outlook folders
- Tracking whether users Reply, Reply-All, or Forward
- Maintaining Excel-based processing logs
- Recovering from interrupted mailbox-processing sessions
- Reviewing MailFlow AX processing activity
- Extending or maintaining the MailFlow AX automation

Do not activate this skill for unrelated email composition, generic Outlook usage, or general-purpose data processing unless the task specifically involves the MailFlow AX workflow.

---

## Core Workflow

### 1. Monitor

Connect to the configured Outlook mailbox and monitor the required folders.

Typical monitored locations include:

- Inbox
- Sent Items

Identify new messages and messages that may have arrived while the application was stopped.

Maintain message-level tracking so previously processed messages are not unnecessarily processed again.

---

### 2. Normalize and Identify

Before classification:

- Normalize message timestamps.
- Identify the message uniquely.
- Check whether the message already exists in the processing log.
- Preserve relevant metadata such as subject, sender, recipient information, received time, and processing status.

Do not duplicate processing when a message is already recorded as successfully processed unless the user explicitly requests reprocessing.

---

### 3. Classify

Use the configured local AI engine to classify the email.

The preferred architecture is local/offline AI through an available Ollama-compatible model.

The classifier should determine the category or processing route required by the configured MailFlow AX logic.

Record classification output and confidence where available.

Do not invent a classification when the model result is unavailable or ambiguous.

---

### 4. Allocate

Use the configured allocation rules to determine the appropriate destination.

Allocation may be based on:

- Email category
- Business rule
- User mapping
- Target folder
- Queue
- Round-robin allocation
- Other project-specific routing configuration

Before allocating, validate that the destination exists and that the required mapping is available.

If allocation cannot be resolved safely, mark the message as an exception/pending item rather than silently assigning it to an arbitrary destination.

---

### 5. Move

Move the email to the configured Outlook destination after successful classification and allocation.

The operation should be traceable in the processing log.

Do not treat an email as successfully allocated merely because a destination was selected; verify the movement operation where technically possible.

---

### 6. Track Actions

After allocation, track downstream user actions.

Supported actions include:

- Reply
- Reply-All
- Forward

Do not treat simply opening or reading an email as a completed business action unless the project configuration explicitly requires that behavior.

The action status should remain pending until an applicable action is detected.

---

### 7. Log

Maintain an operational log containing enough information to understand the lifecycle of each processed email.

Typical information can include:

- Message identifier
- Subject
- Sender
- Received timestamp
- Classification
- AI confidence
- Allocation destination
- Processing status
- Movement status
- Action status
- Action type
- Action timestamp
- Error or exception information

Excel may be used as the operational logging and reporting layer.

---

## Dashboard Expectations

The MailFlow AX interface can expose operational indicators such as:

- Total Emails
- Allocated
- Actioned
- Pending
- Action Rate
- AI Confidence

These values should be derived from the processing records rather than manually entered.

The dashboard should provide a quick operational overview without hiding underlying processing events.

---

## Live Activity

The live activity view should communicate processing events such as:

- Mailbox monitoring started
- Inbox monitoring enabled
- Sent Items monitoring enabled
- Recovery checkpoint loaded
- Recovery scan executed
- New email detected
- Email classified
- Allocation completed
- Email moved
- Action detected
- Processing exception encountered

Use clear status messages so an operator can determine what the automation is currently doing.

---

## Recovery and Checkpoints

MailFlow AX should support restart-safe processing.

When possible:

1. Load previously processed message identifiers.
2. Load the latest recovery checkpoint.
3. Scan for messages received after the checkpoint.
4. Compare discovered messages with already processed messages.
5. Process only messages that require action.
6. Update the checkpoint after successful processing.

A recovery scan should not blindly reprocess the entire mailbox.

---

## Error Handling

When an operation fails:

- Preserve the original message information.
- Record the failure in the log.
- Keep the message in a recoverable state when possible.
- Do not falsely report the message as completed.
- Continue processing independent messages when safe.
- Surface exceptions clearly to the operator.

Typical failure areas include:

- Outlook connection issues
- Missing folders
- Invalid allocation mappings
- AI engine unavailable
- Classification failure
- Excel/log access problems
- Message movement failure
- Duplicate processing detection

---

## Configuration Principles

Keep environment-specific configuration separate from the core processing logic wherever possible.

Configuration may include:

- Outlook folder names
- Target allocation folders
- User mappings
- Allocation rules
- AI model name
- Excel log location
- Recovery settings
- Processing intervals

Do not hard-code credentials, access tokens, passwords, or confidential production values.

---

## Execution Guidance

When modifying or extending this project:

1. Inspect the existing repository structure before changing files.
2. Identify the current application entry point.
3. Reuse existing modules and configuration where possible.
4. Preserve the existing processing pipeline.
5. Make the smallest safe change required.
6. Validate the affected workflow before declaring success.
7. Update documentation when behavior changes.

Do not rewrite the complete application merely to implement a small feature.

---

## Data Privacy

MailFlow AX processes mailbox content and may handle business-sensitive email information.

Never expose or commit:

- Credentials
- Tokens
- Passwords
- Production mailbox data
- Confidential customer information
- Private email content
- Real access configuration
- Unnecessary personal information

Use sanitized sample data for demonstrations and repository examples.

---

## Expected Agent Behavior

When an agent uses this skill, it should:

- Understand the MailFlow AX processing lifecycle.
- Inspect existing code before making changes.
- Prefer the current implementation over rebuilding equivalent functionality.
- Preserve existing logging and recovery behavior.
- Validate allocation before moving messages.
- Avoid destructive mailbox operations unless explicitly requested.
- Report uncertainty instead of inventing successful processing.
- Keep user/company-sensitive data out of generated examples and commits.

---

## Typical Requests

Examples of tasks that should activate this skill:

> "Monitor my Outlook inbox and classify new emails."

> "Route these emails to the correct users."

> "Add Reply/Forward tracking to MailFlow AX."

> "Fix the recovery scan."

> "Add a new allocation rule."

> "Update the MailFlow dashboard with another processing metric."

> "Log the AI confidence into Excel."

> "Move classified emails into the correct Outlook folders."

---

## Out of Scope

This skill does not define:

- Generic email writing
- Generic Outlook configuration
- Marketing email automation
- Spam generation
- Bulk unsolicited messaging
- Credential harvesting
- Unauthorized mailbox access
- Destructive mailbox operations without explicit authorization

---

## Success Criteria

A MailFlow AX task should only be considered complete when:

1. The requested processing behavior is implemented.
2. Existing processing and recovery behavior remains intact.
3. Relevant logging is updated.
4. Errors are handled appropriately.
5. The affected workflow has been tested or otherwise verified.
6. No confidential credentials or production secrets were introduced.
