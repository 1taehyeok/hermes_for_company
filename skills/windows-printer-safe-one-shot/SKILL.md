---
name: windows-printer-safe-one-shot
description: Safe one-shot print verification workflow for Windows printers from WSL. Use when duplicate/accumulated printing is suspected and only one controlled live test should be submitted.
version: 1.0.0
author: Hermes Agent
license: MIT
metadata:
  hermes:
    tags: [printing, windows, wsl, acrobat, printservice, verification, safety]
---

# Windows Printer Safe One-Shot

## When to use
Use this skill only when:
- Hermes is running in WSL
- output must go to a Windows printer
- duplicate / accumulated / runaway printing is suspected
- the user wants exactly one controlled live print test

## Goal
Prove whether **one print submission results in one actual print**.

## Safety rules
- Do not send repeated live print tests back-to-back.
- After one live submission, stop and ask the user for physical printer confirmation.
- Do not enable automatic fallback printing on a physical printer.
- If no file modification is needed, prefer direct printing over creating unnecessary copies.

## Standard sequence
1. Confirm the exact target printer.
2. Confirm the exact target file.
3. Check current queue state with `Get-PrintJob`.
4. Check that `Microsoft-Windows-PrintService/Operational` is enabled.
5. If possible, start from a clean baseline (no current queue entries and no new events after the log was enabled).
6. Submit exactly **one** print command.
7. Immediately inspect Operational events.
8. Stop and get user confirmation about the physical output.
9. Only then decide whether another test is needed.

## Recommended command checks
### Printer / queue
```bash
/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe -NoProfile -Command \
  "Get-PrintJob -PrinterName 'Samsung X7600 Series (10.69.16.2)' -ErrorAction SilentlyContinue | \
    Select-Object ID,DocumentName,JobStatus,PagesPrinted,TotalPages,SubmittedTime | Format-Table -AutoSize"
```

### Operational log state
```bash
/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe -NoProfile -Command \
  "Get-WinEvent -ListLog 'Microsoft-Windows-PrintService/Operational' | \
    Select-Object LogName,IsEnabled,RecordCount,LastWriteTime | Format-List"
```

### Recent Operational events
```bash
/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe -NoProfile -Command \
  "Get-WinEvent -LogName 'Microsoft-Windows-PrintService/Operational' -MaxEvents 10 | \
    Select-Object TimeCreated,Id,LevelDisplayName,Message | Format-List"
```

## Preferred live test for PDF
When a known-good English-named PDF working copy already exists, Acrobat `/t` is acceptable for a one-shot verification:

```powershell
Start-Process -FilePath "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe" `
  -ArgumentList @("/t", "C:\path\analysis-note-final.pdf", "Samsung X7600 Series (10.69.16.2)")
```

## Interpreting the result
A healthy one-shot flow usually looks like:
- queue empty before test
- exactly one submission
- Operational log records a normal sequence such as 800 / 801 / 842 / 805 / 307
- user confirms the printer produced one correct output

Interpretation:
- one log flow + one physical print = command path is behaving normally
- multiple log flows after one submission = duplicate invocation / submission problem
- one log flow but multiple physical prints = driver / spooler / printer-side issue more likely
- no useful log flow and no print = submission path likely failed

## Proven cases in this environment
On 2026-04-22, the following one-shot test succeeded:
- file: `C:\Users\Public\Documents\hermes-print-jobs\20260421_acrobat_compare\analysis-note-final.pdf`
- printer: `Samsung X7600 Series (10.69.16.2)`
- engine: Acrobat `/t`
- result: one normal physical print
- observed event IDs: 800, 801, 842, 805, 307

Additional comparison result from the same environment:
- Korean-named / Korean-UNC-path kickboard PDF reached normal Windows print events but did **not** produce physical output
- English working copy of that same PDF did produce physical output
- Operational conclusion: for PDF printing in this environment, treat **Korean path/filename as a confirmed risk** and prefer creating an English-named working copy before live printing
- After the job is confirmed complete, the temporary English working copy may be removed
- Final user-facing reporting should stay minimal: file, printer, range (full or pages), copies

## Reporting template
Report only these essentials:
- target file
- target printer
- engine used
- whether queue was clear before test
- key log event IDs observed
- user-confirmed physical result
- whether another live test should be blocked pending investigation

## Pitfalls
- Do not assume “no queue entry seen” means failure.
- Do not assume “process exited 0” means physical print succeeded.
- Do not run fallback image printing automatically.
- Do not continue testing before user feedback when live paper output is involved.
