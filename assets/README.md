# Evidence assets

Capture only material that the intended Friday Talk audience is authorised to view. Crop or blur names, client references, emails, URLs, attachment filenames, environment identifiers, and any sensitive product data.

## Required assets

1. `01-slack-report.png`
   - A redacted original report from `#pedaxis-ext-bug-report` or `#internal-redlens-bug-report`.
   - Keep the message and enough thread context to show why the report is actionable.

2. `02-slack-linear-creation.png`
   - The same thread showing the bot marker, structured request to Linear, and Linear’s creation reply.
   - Prefer the Pedaxis `DFPT-2027` example because it leads to a complete flow.

3. `03-linear-issue.png`
   - The Linear issue showing title, project, milestone, labels, priority, and final state.
   - Redact assignee, customer identifiers, URLs, and unrelated issue details.

4. `04-delivery-evidence.png`
   - A redacted GitLab merge request or CI result that shows the work reached review and test validation.
   - Use only a case whose test state is unambiguous.

5. `05-qa-thread.png`
   - The relevant QA confirmation in the linked Slack thread, such as the Pedaxis `PASS` event.

## Asset order in the talk

Use the assets as a story, not as a gallery:

1. Report in Slack
2. Structured creation in the same thread
3. Linear record
4. Delivery evidence
5. QA and final thread update
