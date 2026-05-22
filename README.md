# 📦 Delivery Note Automation — AI-Powered Logistics Operations Tool

### *Built by a Logistics Professional, Not a Programmer — Powered by Domain Knowledge \& Prompt Engineering*

<br>

> \*\*⚠️ Sanitization Notice:\*\* This script has been prepared for public sharing. All real company names, employee email addresses, and internal network drive paths have been replaced with generic placeholders such as `\[Company\_Name]` and `\[recipient@example.com]`. You \*\*must\*\* substitute these with your own data before using the script. See \[Setup \& Configuration](#-setup--configuration) for the full list.

<br>

\---

## 🚩 The Business Problem

As a logistics intern handling **Delivery Notes (DN)** daily, I ran into the same bottleneck every single morning.

Every time I exported a DN spreadsheet from the system, the raw data was unformatted and unusable as-is. Before it could be shared with anyone, I had to:

* Manually **reformat the table** — adding borders, inserting rows, applying the right color scheme for each customer, and adding a Total row with a SUM formula
* **Navigate to the correct network folder** — each of 14+ customers had a completely different directory structure — and save the file with a precisely formatted filename based on the date
* **Draft a separate Outlook email** for each customer, manually typing the subject line, copying the right recipients from memory, and pasting the formatted table into the email body

Doing this for 4–5 customers every day consumed roughly **30 minutes** — time that couldn't be spent on actual logistics work.

The process was repetitive, rule-driven, and followed the same logic every time. It was exactly the kind of work that should not require a human to do manually.

<br>

\---

## 💡 My Solution

I don't write code. But I understand the business process deeply — every rule, every exception, every customer-specific requirement.

So I used that domain knowledge to **direct an AI (Large Language Model) to write the automation for me**, through structured and iterative prompt engineering.

The result is `Master\_FormatDN\_And\_Save\_Ultimate` — a VBA macro that runs inside Microsoft Excel and handles the entire end-to-end process in **under 10 seconds**.

||Before|After|
|-|-|-|
|⏱️ Time per day|\~30 minutes|\~10 seconds|
|🔁 Manual steps|4 multi-step tasks per customer|1 button click|
|⚠️ Risk of error|High — wrong folder, wrong recipient, wrong color|Near-zero — all logic is automated|
|📈 Scalability|Adding a new customer = more manual work|Adding a new customer = one new code block|

> I didn't need to know how to code. I needed to know the business rules well enough to explain them precisely. That's where my value was.

<br>

\---

## 🧠 How I Guided the AI

This script was not generated from a single prompt. It was built through **multiple rounds of structured, specific, real-world instructions** — the same way a business analyst would brief a developer, except the developer was an AI.

Here's how I approached it:

### I Translated Business Rules Into Structured Specs

Rather than asking the AI to "make a macro for delivery notes," I gave it detailed, tabular business rules — the same way you would write a requirements document:

```
CUST-01  Customer A (Plant 1)   → Header color: #44B3E1 | Date goes in column I separately
CUST-02  Customer B (Site X)    → Header color: #FFFF00 | Date goes in column I separately
CUST-03  Customer B (Site Y)    → Header color: #44B3E1 | Merge columns I, J, K | Remove gridlines
CUST-13  Customer C             → Two tables (AM 09:30 / PM 21:30) | Two different colors
```

I knew that vague input produces vague output. Structured input produces precise code.

### I Defined the Date \& Folder Logic Explicitly

The most complex part was that each customer's save path and filename depended on the **actual date inside the data file** — not today's date on the computer. I was explicit about this:

> \*"The date in all file paths, folder names, and email subjects must come from Column C of the raw data only. If Column C says 14-Mar-2026, then the folder must be `...\\2026\\3.Mar` and the filename must reflect `14-Mar-26`. Never use the system clock."\*

This level of specificity is what separated a working script from a broken one.

### I Caught and Fixed Real Bugs Through Testing

After each version, I tested the macro against real data and reported exactly what went wrong:

* **The Outlook Signature Bug:** When the macro injected the email body, the user's default Outlook signature disappeared. I identified this and told the AI: *"Call `.Display` first to load the signature, then inject the HTML content — not the other way around."*
* **The Milk-Run Split Problem:** One customer (Dual-table customer) had two delivery windows in the same file — 09:30 AM and 21:30 PM. A single table was wrong. I described exactly what the output should look like: two separate tables, two separate color-coded headers, with DN numbers listed under each time window in the email body.
* **The Formatting Gap Problem:** Line breaks in the email body created ugly oversized gaps. I diagnosed it and instructed: *"Use HTML non-breaking spaces (`\&nbsp;`) instead of tab characters or line breaks for indentation inside the email body."*

### I Sanitized the Code Before Publishing

Before uploading to GitHub, I prompted the AI to do a full security review — replacing every real employee name, email address, network path, and company name with clearly labeled placeholders. Publishing clean code that others can safely use and adapt is part of being a responsible professional.

<br>

\---

## ✨ Key Features

### ✅ Data Validation — Catches Problems Before Anything Runs

The macro will not proceed if the data doesn't meet the required standards. It checks:

* All 11 column headers are present and correctly named
* Every customer code exists in the approved list — unknown codes are flagged
* All rows in a file belong to the same customer — mixed files are rejected
* All dates are valid and within a sensible year range (2020–2030)
* All quantities are positive numbers
* No Delivery Note number fields are blank

Errors are highlighted in red in column L, and the macro stops with a clear message.

### 🎨 Intelligent Table Formatting — Different Rules for Every Customer

Once data passes validation, the macro transforms the raw export into a clean, professional-looking document:

* Applies the correct **color-coded header** for each customer (15 customer configurations supported)
* Handles three different **layout styles** automatically: date-in-column-I, full three-column merge, or the special Dual-table customer milk-run layout
* Adds borders, center alignment, Calibri font, and auto-fit column widths throughout
* Adds a **Total row** with a live SUM formula, matched to the customer's header color
* Automatically **removes the Invoice column** if it contains no data
* Formats dates consistently as `dd-Mmm-yy`

### 📁 Automatic File Routing \& Saving — The Right Folder, Every Time

The macro builds the save path and filename from the document's own data:

* Extracts the date from Column C and constructs the exact folder structure each customer requires (e.g., `...\\2026\\3.Mar\\16-Mar`)
* **Creates missing folders automatically** — no manual folder creation needed
* **Prevents accidental overwrites** — if a file already exists, it saves as `filename - Copy00`, `Copy01`, and so on
* Saves as standard `.xlsx` format

After saving, it writes two clickable hyperlinks back onto the sheet:

* `\[1] Data Folder` — links directly to where the file was saved
* `\[2] Attachment Folder` — links to the DN attachment folder; displayed in red with a fallback shortcut if the folder doesn't exist yet

### 📧 Outlook Email Drafting — Pre-Written, Ready to Review

For supported customers, the macro opens a pre-filled Outlook draft:

* **To** and **CC** fields are pre-populated with the correct recipients per customer
* **Subject line** is auto-generated from the customer name and document date
* **Email body** includes the correct greeting and message template, with dates in bold
* The formatted Excel table is converted to HTML and **embedded directly into the email body** — preserving colors, borders, and layout
* For Dual-table customer (milk-run), DN numbers are listed under the correct time window (09:30 AM / 21:30 PM) in the email body automatically
* The email **opens as a draft only** — nothing is ever sent automatically

<br>

\---

## ⚙️ Setup \& Configuration

> ⚠️ \*\*Read this before running the script.\*\* The file contains placeholder values throughout. You need to replace them with your own data first.

Open the `.bas` file in Excel's VBA Editor (`Alt+F11`) and use **Find \& Replace** (`Ctrl+H`) to update each item below.

\---

### Step 1 — Set Your Sheet Name

Find this line near the top of the code and replace `YourSheetName` with the exact name of your worksheet tab:

```vba
Set ws = ThisWorkbook.Sheets("YourSheetName")
```

\---

### Step 2 — Replace Company \& Location Names

These appear throughout the formatting and email sections:

|Placeholder|What to Replace It With|
|-|-|
|`\[Company\_Name]`|Your customer's company name|
|`\[Company\_Code]`|A short code or abbreviation for that customer|
|`\[Location\_1]`, `\[Location\_2]`|Site or plant names (e.g., `Prachinburi`, `Bangplee`)|
|`\[Location\_3]`, `\[Location\_4]`|Additional sites as needed|
|`\[Location\_Code]`|Short location identifier used in file names|
|`\[Country\_Name]`, `\[Country\_Code]`|For international shipments|
|`\[Our\_Company\_Name]`|**Your** company's name (the sender)|

\---

### Step 3 — Replace Email Addresses

These appear in the email drafting section for each customer:

|Placeholder|What to Replace It With|
|-|-|
|`\[recipient\_to@example.com]`|Primary To recipients|
|`\[recipient\_cc@example.com]`|CC recipients|
|`\[Contact\_Name\_1]` through `\[Contact\_Name\_5]`|Contact names used in the greeting line|
|`\[Internal\_Team]`|Your internal team name (e.g., `FG Team`)|
|`\[Client\_Team]`|The customer's team name|

\---

### Step 4 — Replace File Paths

These appear as `basePath` and `attachBasePath` in the customer routing section:

```vba
' Replace the generic structure:
basePath = "C:\\Users\\\[Username]\\Documents\\\[Shipment\_Folder]\\\[YYYY]\\\[Location\_1]\\\[M].\[MMM]"

' With your actual network path, for example:
basePath = "Z:\\Internal\_Drive\\Logistics\_Dept\\Shipment\_Data\\\[YYYY]\\\[M].\[MMM]"
```

|Placeholder|What to Replace It With|
|-|-|
|`\[Username]`|Your Windows username|
|`\[Shipment\_Folder]`|Your root folder for shipment files|
|`\[DN\_Folder]`|Your root folder for DN attachment files|

> \*\*Important:\*\* Do \*\*not\*\* replace the date tokens: `\[YYYY]`, `\[YY]`, `\[MM]`, `\[M]`, `\[MMM]`, `\[DD]`. These are automatically filled in at runtime from the document date.

\---

### Step 5 — Update the Customer Code List

The validation step checks each row's customer code against this list. Update it to match your actual customers:

```vba
validCustomers = "|CUST-01|CUST-02|CUST-03|..."
```

Keep the pipe character `|` before and after each code. Add a matching `Case "CUST-XX"` block in the formatting section for any new customers you add.

<br>

\---

## ▶️ How to Run

**1. Import the module**

* Open your Excel workbook
* Press `Alt+F11` to open the VBA Editor
* Right-click your workbook → **Import File** → select `Master\_FormatDN.bas`
* Close the editor and save the workbook as `.xlsm` (macro-enabled format)

**2. Prepare your data**

* Paste your raw DN export into the configured worksheet
* Headers go in row 1, data starts from row 2
* The sheet must have exactly 11 columns (A through K) — see the table below

**3. Run the macro**

* Press `Alt+F8`
* Select `Master\_FormatDN\_And\_Save\_Ultimate`
* Click **Run**

**4. Fix any validation errors**

* If rows are flagged in column L, correct the data and run again
* The macro won't proceed past the validation check if errors exist

**5. Review and send the email**

* The Outlook draft will open automatically
* Review all fields before clicking Send — **the macro never sends automatically**

\---

### Required Column Layout

Your worksheet must have these 11 columns in order before running:

|Column|Header|Notes|
|-|-|-|
|A|`Delivery Note NO`|Cannot be empty|
|B|`INVOICE NO`|Optional — column auto-removed if blank throughout|
|C|`Document Date`|Must be a valid date, 2020–2030|
|D|`Ship to Code`|Must match your `validCustomers` list|
|E|`Ship to Name`|Free text|
|F|`DN item`|Free text|
|G|`PURCHASE ORDER`|Used for routing sub-logic in some customers|
|H|`SO`|Free text|
|I|`Material Code`|Free text|
|J|`Material Description`|Free text|
|K|`QTY`|Must be a positive number|

All rows must belong to the **same customer** in a single run. Mixed-customer files will be rejected.

<br>

\---

## 📁 Files in This Repository

```
/
├── Master\_FormatDN.bas    # The VBA module — import this into your Excel workbook
└── README.md              # This file
```

<br>

\---

## 📄 License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT) — free to use, adapt, and build on for any purpose.

\---

<br>

<div align="center">

*You don't need to know how to code to solve a problem with code.
You need to understand the problem well enough to explain it clearly.*

</div>

