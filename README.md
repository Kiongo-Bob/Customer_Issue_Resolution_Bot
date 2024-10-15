### Customer Issue Resolution Bot ###

## Overview
This AI-driven automation is designed to streamline customer service by handling incoming queries through Outlook. It uses **OpenAI** to generate quick, real-time responses to customer emails and records interactions using an Excel-based logging system.

If the issue remains unresolved, the automation escalates the query to a human agent and logs the escalation in Excel. The solution prevents duplicate handling and provides insights based on customer interactions.

---

## Features
* **Email Query Handling**: Automatically retrieves customer queries from Outlook.
* **OpenAI-Powered Responses**: Generates quick fix solutions using OpenAI's API for efficient customer support.
* **Excel Logging**: Tracks all customer interactions, escalations, and AI responses using Excel.
* **Escalation Management**: Escalates unresolved issues to human agents for manual intervention.
* **No Orchestrator**: This automation runs without UiPath Orchestrator and uses local Excel files, DataTables, and DataRows for data storage and manipulation.

---

## Pre-Requisites
Before running this automation, ensure the following:

1. **Outlook Account**:
   * You must have a valid Outlook account to retrieve emails.

2. **OpenAI Account**:
   * Ensure that the account signed into your UiPath Cloud has an active OpenAI account to enable AI-generated responses.

3. **Excel File**:
   * An Excel file will be used for logging interactions, tracking customer issues, escalations, and AI responses.

---

## Setup Instructions

### 1. Configure Outlook Integration
* Open the **Retrieve Emails** workflow.
* Update the **email address** in the workflow to match the Outlook account that will be retrieving customer queries.
* Follow the UiPath instructions to ensure your Outlook account is integrated properly.
<img  title="Outlook Configuration" src ="Data\Temp\Screenshot 2024-10-15 135529.png">

### 2. Configure Email for Bob
* The automation is initially set up to handle emails from **Bob**. To edit this:
  * In the **Retrieve Emails** workflow, locate the section where the emails are filtered by sender.
  * Update the email filter to match the sender's email (e.g., "bob@example.com").
  * If required, update the recipient's email (who receives complaints) in the same workflow to match the Outlook account handling customer service.
  <img  title="Client Email" src ="Data\Temp\Screenshot 2024-10-15 135545.png">

### 3. Excel File Setup
* The automation relies on an Excel file to log customer interactions. Ensure that the Excel file is created and structured with the following columns:
  * **Date** (Timestamp of the interaction)
  * **Customer Email**
  * **Email Subject**
  * **Category** (e.g., Billing, Technical Support, General Inquiry)
  * **Solution** (AI-generated response)
  * **Escalation Status** (Resolved/Unresolved)

### 4. Running the Automation
* Run the automation without Orchestrator. The bot will retrieve emails from Outlook, classify the issues, generate responses, and log interactions in the Excel sheet.
* If an issue remains unresolved, it will escalate the complaint to a human agent and log the escalation in Excel.

---

## Key Considerations
* This automation does **not use UiPath Orchestrator**; all data is handled locally using **Excel**, **DataTables**, and **DataRows**.
* Ensure both your **Outlook account** and **OpenAI account** are properly configured for the bot to work successfully.

---

## Customization
* **Edit Email Sender/Recipient**: Update the email sender (e.g., Bob) and recipient in the **Retrieve Emails** workflow as per your use case.
* **Customizing Excel File**: Modify the Excel logging sheet structure to include any additional fields or data as needed.

---

## Support
For any issues or further customization, contact the support team or refer to UiPath documentation for integrating Outlook and OpenAI.
