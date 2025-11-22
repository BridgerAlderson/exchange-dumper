# Exchange Dumper

![Python](https://img.shields.io/badge/Python-3.x-blue?style=for-the-badge&logo=python)
![Category](https://img.shields.io/badge/Category-Red%20Team-red?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Operational-success?style=for-the-badge)

**Advanced Exchange Server Data Extraction & Access Validation Tool**

**ExDumper** is a specialized offensive security tool designed for Red Team operations. It interacts with Microsoft Exchange Servers via EWS (Exchange Web Services) to validate credential access rights (Impersonation vs. Delegate) and perform targeted data extraction from specific mailboxes.

Designed to be stealthy, robust, and fault-tolerant against missing metadata (e.g., undated Drafts).



## ⚡ Features

* **Auto-Detection of Access Rights:** Automatically checks if the compromised account has `ApplicationImpersonation` or `Delegate` rights.
* **Robust Data Dumping:** Extracts data from:
    * 📂 Emails (Inbox, Sent, Drafts, Junk, Trash, Archive)
    * 📅 Calendar Items
    * 📇 Contacts
    * 📝 Tasks
* **Attachment Harvesting:** Automatically identifies and downloads attachments from emails.
* **Fault Tolerant:** Includes fallback logic for missing timestamps (common in *Drafts*) and handles corrupt metadata gracefully.
* **Clean Output:** Organizes dumped data into structured directories with readable `.txt` files.


## 📦 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/BridgerAlderson/exchange-dumper.git
   ```
   
   ```bash
   cd exchange-dumper
   ```
   
3. Install dependencies:
   ```bash
   pip3 install -r requirements.txt
   ```
## 🚀 Usage
ExDumper requires the target domain, a valid service account (or compromised user), and the target mailbox you wish to dump.

```bash
python3 ExDumper.py --domain <DOMAIN> --user <USERNAME> --password '<PASSWORD>' --server <EXCHANGE_SERVER_IP> --target <TARGET_EMAIL>
```
**Output Structure**
The tool creates a timestamped directory for each session:
```bash
ExDump_administrator_20251121_163433/
├── Attachments/     # Binary files extracted from emails
├── Calendar/        # Meeting details and appointments
├── Contacts/        # VCard-like contact dumps
├── Drafts/          # Unsent messages (High Value Loot!)
├── Inbox/           # Incoming mail
├── Sent/            # Outgoing mail (Check for sent credentials)
├── Tasks/           # To-Do lists
└── ...
```

**EDUCATIONAL USE ONLY.**

This tool is developed for Red Team operations, Penetration Testing, and Security Research contexts.
. Do not use this tool on networks or systems you do not have explicit permission to test.
. The author is not responsible for any misuse or damage caused by this program.
. Use responsibly and adhere to local laws and regulations.
