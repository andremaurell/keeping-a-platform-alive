from pathlib import Path

from docx import Document


GUIDE = Path(__file__).resolve().parents[1] / "Keeping a Platform Alive Speaker Guide.docx"

ENGLISH = {
    "Slide 10": (
        "Here, we start by estimating cost, effort, and risk. Once the estimate is approved by "
        "the client, we enter the development process for the agreed change. Next, that "
        "implementation is reviewed and approved. Finally, the team moves into the commit layer, "
        "which in this case means a delivery commitment, not only a Git commit. This flow prevents "
        "a good idea from becoming a promise before there is a minimum understanding of the work "
        "required."
    ),
    "Slide 11": (
        "Before the result is sent to the customer, we have our own QA stage. We reproduce the "
        "reported bug, isolate the context from the development and review agents, and run tests "
        "based directly on the requirements of the created task and on the messages sent in the "
        "Slack thread. Only after this round of end-to-end tests passes is the solution actually "
        "delivered to the reporter."
    ),
}


def remove(paragraph):
    paragraph._element.getparent().remove(paragraph._element)


def main():
    doc = Document(GUIDE)
    paragraphs = doc.paragraphs

    for title_prefix, english in ENGLISH.items():
        start = next(i for i, p in enumerate(paragraphs) if p.text.startswith(title_prefix))
        end = next(
            (i for i, p in enumerate(paragraphs[start + 1 :], start + 1) if p.text.startswith("Slide ")),
            len(paragraphs),
        )
        section = paragraphs[start:end]
        english_label = next(p for p in section if p.text == "English")
        label_index = next(i for i, p in enumerate(doc.paragraphs) if p._p is english_label._p)
        english_body = doc.paragraphs[label_index + 1]
        english_body.clear()
        english_body.add_run(english)

        # Remove an obsolete standalone label left after the transition text.
        for paragraph in list(section):
            if paragraph.text == "Transition":
                remove(paragraph)

    doc.save(GUIDE)
    print(GUIDE)


if __name__ == "__main__":
    main()
