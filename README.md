# Keeping a Platform Alive

Source package for a 15-minute Friday Talk about the feedback and delivery loop used by RedLens and Pedaxis.

## Central question

How can a platform keep listening after release, turn customer feedback into prioritised work, and return a validated outcome to the original conversation?

The talk explains the operating loop behind the answer:

```text
Slack report → intake automation → Linear issue → triage and delivery → QA → thread update
```

Slack holds the conversation and its evolving evidence. Linear holds the structured work record. The local development harness turns an approved issue into an isolated, reviewed, independently tested change. The original Slack thread remains the place where people can add context and receive the outcome.

## Final deliverables

- [`output/keeping-a-platform-alive.pdf`](output/keeping-a-platform-alive.pdf) — final 14-slide deck, exported from Canva.
- [`Keeping a Platform Alive Speaker Guide.docx`](Keeping%20a%20Platform%20Alive%20Speaker%20Guide.docx) — bilingual Portuguese and English talk track.
- [`slides.md`](slides.md) — final slide-by-slide plan and timing.
- [`speaker-notes.md`](speaker-notes.md) — concise presenter notes and narrative anchors.
- [`assets/README.md`](assets/README.md) — capture and redaction rules for supporting evidence.

## Canonical design

The editable deck lives in [Canva](https://www.canva.com/design/DAHaYbrwtek/Q07sS0qsvagsclHXSBFmMQ/edit). The PDF above is the version exported from that design for this repository.

## Audience outcome

The audience should understand that the value comes from a connected feedback loop, rather than from a ticket-creation bot alone. The same pattern can apply to any product that needs a visible route from user signal to validated delivery.

## Sources

The talk is grounded in the currently configured RedLens and Pedaxis Slack-to-Linear automations, the Pedaxis delivery harness contract, and redacted operational examples. Do not commit customer-identifying screenshots or private Slack content without approval.
