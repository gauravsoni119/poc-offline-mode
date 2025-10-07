# POC: Offline Mode Issue with `OnSendAddinsEnabled` in New Outlook

This repository demonstrates a reproducible issue with **Outlook add-ins** when using the `OnSendAddinsEnabled` policy flag in the **new Outlook** (Windows).

The project is based on the official [Microsoft Outlook Add-in Quick Start (Yo Office)](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart-yo) sample and serves as a **proof of concept (POC)** to show that **offline mode is not working as expected** when the `OnSendAddinsEnabled` flag is applied.

---

## 🎯 Objective

Microsoft documentation claims that **offline mode is supported** in the new Outlook with the `OnSendAddinsEnabled` flag enabled.  
However, during testing, it was observed that:

- The `OnMessageSend` event **does not trigger** when Outlook is offline and the message always goes to the Outbox folder. As per [documentation](https://learn.microsoft.com/en-gb/office/dev/add-ins/outlook/one-outlook#add-in-availability-when-offline), depending upon situation, the message should either go to drafts folder(instead of Outbox folder) or smart alerts should run to check the compliance. 

- The expected behavior (as documented) is **that the message either should go to draft folder(if outlook launched without internet) or add-in should execute(if connection established after launching outlook while offline)**, provided the add-in is configured with the `OnSendAddinsEnabled` policy.

This repository provides a minimal reproducible setup to demonstrate the issue.

---

## ⚙️ Setup Instructions

### 1. Prerequisites
- Node.js (v18 or above)

### 2. Clone and Install
- git clone https://github.com/gauravsoni119/poc-offline-mode.git
- cd poc-offline-mode
- npm install

### 3. Run the Add-in
- npm start



