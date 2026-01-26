# Student Connect — Google Apps Script 

Google Apps Script bound to a Google Spreadsheet that manages the lifecycle of eSIM and SIM lines across multiple countries (UK, Spain, France).

The script integrates with carrier APIs, internal proxy services, Google Drive and Gmail to handle provisioning, pre-activation, activation, suspension, porting, QR generation and customer notifications.

---

## Supported countries

### UK
- Uses eSIM-Go APIs.
- Full programmatic lifecycle: pre-activation, activation, suspension and porting.
- Retrieves ICCID, MSISDN, matching ID and generates LPA/QR codes.

### Spain
- Mix of automated and semi-manual flows depending on operator.
- Uses direct operator gateways and proxy endpoints.
- Supports activation, suspension, reactivation and duplication flows.

### France
- Uses a proxy service to handle carrier-specific operations.
- Distinguishes between physical SIM and eSIM flows.
- eSIM activations rely on internal stock management and ticket-based confirmation.

---

## Data model (Spreadsheet)

The script expects specific sheets and columns to exist.

### Main sheet: `Activaciones`
Key columns (1-based):
- A: Country (`uk`, `spain`, `france`)
- B: ICCID
- C: MSISDN
- D: Product / Plan
- E: Type (`SIM` / `eSIM`)
- F: Activation date
- G: Expiry date
- I: Status (`WIP`, `PreActivada`, `Activa`, `Suspendida`, `Error`)
- K: Customer name
- L: Customer email
- T (20): Error / debug message
- X (24): Country-specific identifier (LPA, ticket, matching ID)

Additional sheets may be required depending on country:
- `ICC France Airmob`
- `UK eSIM Go Portabilidad`
- `Mailchimp Export`
- `Mailchimp Sync Logs`

---

## Lifecycle flows

- **Pre-activation**  
  Reserves or assigns an eSIM/SIM, retrieves identifiers, updates the sheet and sends preliminary customer instructions.

- **Activation**  
  Applies the selected plan or bundle, confirms activation, sends final QR/PDF instructions and updates status.

- **Suspension / Reactivation**  
  Revokes or restores active assignments via carrier or proxy endpoints.

- **Porting (UK)**  
  Requests PAC codes and stores results in the porting sheet.

---

## Notifications

- QR codes are generated dynamically and sent via email.
- PDFs stored in Google Drive may be attached.
- All emails are sent using `GmailApp`.

---

## Configuration

Sensitive values must be stored in **Script Properties**:
- API keys
- Proxy base URLs
- Sender email
- Drive file IDs (if applicable)

Hardcoded credentials should not be committed.

---

## Error handling

- All external API calls capture and log errors.
- Error details are written back to the spreadsheet.
- Script execution state is reflected directly in the `Status` column.

---

## Usage notes

- The script is operated via spreadsheet rows and Apps Script functions.
- Some maintenance tasks run via time-driven triggers.
- Activations can be triggered manually for a specific row when required.

---

⚠️ This repository contains **source code only**.  
The production environment is the bound Google Spreadsheet and its associated Apps Script project.


