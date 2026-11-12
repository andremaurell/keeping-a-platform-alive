from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Pt


GUIDE = Path(__file__).resolve().parents[1] / "Keeping a Platform Alive Speaker Guide.docx"

# Portuguese is intentionally not stored here: the saved guide is the source of truth
# for it.  These English versions were translated from the current Portuguese talk track.
SLIDES = {
    "01": {
        "en": (
            "Hello everyone, and thank you for being here. Today I will show how to keep a "
            "platform alive: by allowing users and customers, when they encounter a problem or "
            "have an idea or feature request, to report it directly to the team responsible for "
            "development. This presentation shows the loop that connects conversation, decision, "
            "and delivery."
        ),
        "pt_transition": "Primeiro, vale olhar para como essa ideia começou.",
        "en_transition": "First, let’s look at how this idea started.",
    },
    "02": {
        "en": (
            "The first idea came from the RedLens project: a feedback page inside the platform, "
            "with a floating button called “bug report.” It would identify the current route, "
            "capture a screenshot, and allow people to attach evidence. It was a good idea because "
            "it kept feedback close to the problem. But it also required building and maintaining "
            "another product surface. The current workflow preserves the same principle, but uses "
            "the place where the team already talks every day: Slack."
        ),
        "pt_transition": "Esse lugar é o Slack.",
        "en_transition": "That place is Slack.",
    },
    "03": {
        "en": (
            "Slack is already our standard communication channel inside the company, used daily by "
            "the whole team. That makes it a better entry point than an isolated form. In this "
            "anonymised example, someone reports a problem, the bot creates the structured request, "
            "and Linear creates the issue. The conversation does not disappear when the issue is "
            "created; it remains available to add context. For this to work, the initial report "
            "does not need to be perfect. It only needs to be useful."
        ),
        "pt_transition": "Para isso funcionar, o report inicial não precisa ser perfeito. Ele só precisa ser útil.",
        "en_transition": "For this to work, the initial report does not need to be perfect. It only needs to be useful.",
    },
    "04": {
        "en": (
            "A useful report carries four things: the observation, the context, the impact, and "
            "some evidence. The observation describes what happened without trying to diagnose it. "
            "The context shows where and under which conditions it happened. The impact explains why "
            "it is worth looking at now. Evidence can be a screenshot, a log, or a document. This "
            "gives the team a starting point without requiring the user to complete a long form. "
            "From that signal, automation handles the repetitive work."
        ),
        "pt_transition": "A partir desse sinal, a automação faz o trabalho repetitivo.",
        "en_transition": "From that signal, automation handles the repetitive work.",
    },
    "05": {
        "en": (
            "The automation follows four basic principles. First, it listens: it reads the thread "
            "that was created. During triage, it searches for matching tickets to avoid duplication. "
            "Then comes structuring, or enrichment: it turns the full conversation into a structured "
            "record in the board. If the ticket is a duplicate, it enriches the existing ticket with "
            "more context. Finally, it creates the task."
        ),
        "pt_transition": "O princípio é o mesmo nos dois produtos, embora os harness de desenvolvimento de ambos sejam diferentes.",
        "en_transition": "The principle is the same in both products, even though their development harnesses are different.",
    },
    "06": {
        "en": (
            "RedLens and Pedaxis are the two projects where this automation is used today. Each has "
            "its own channels, harnesses, and teams. Even so, both flows follow the same principle: "
            "capture the signal in context, create a ticket, assign someone who can act, and return "
            "the result to the original conversation. These examples use similar environments, with "
            "Slack and Linear, but the model can also be used with other platforms without having to "
            "reinvent the whole approach."
        ),
        "pt_transition": "O que mantém esse modelo unido é a própria thread.",
        "en_transition": "What keeps this model together is the thread itself.",
    },
    "07": {
        "en": (
            "A thread is not a ticket queue; each thread has its own ticket. In this example, an "
            "issue has already been created and the reporter adds a new observation. That update is "
            "important because, in real cases, context rarely arrives in a single message. The "
            "automation keeps detecting new messages so it can enrich the ticket that was created."
        ),
        "pt_transition": "A conversa evolui no Slack. A decisão fica registrada no Linear.",
        "en_transition": "The conversation evolves in Slack. The decision is recorded in Linear.",
    },
    "08": {
        "en": (
            "In Linear, we already have the operational decision, structured by the automation from "
            "the Slack conversation. There are four important elements: the assignee is responsible "
            "for the next step; the scope defines the change that will be made; the state shows the "
            "current status of the task; and the evidence proves closure. Beside it, there is an "
            "example of a real issue, showing that Linear records the decision the team will execute. "
            "After the task is created, we need to triage it."
        ),
        "pt_transition": "Depois da criação da tarefa, precisamos fazer uma triagem dessa tarefa.",
        "en_transition": "After the task is created, we need to triage it.",
    },
    "09": {
        "en": (
            "At this point, triage is not a classification table. It is a simple sequence of "
            "decisions. First, we collect the signal: we read the report, the thread, and the "
            "evidence. Then we assess the impact: urgency, scope, and whether the behaviour repeats. "
            "Finally, we route the work by setting an assignee, priority, and destination in Linear. "
            "The output of triage is a decision that someone can execute, not only a label."
        ),
        "pt_transition": "Com a tarefa criada e filtrada, devemos entrar no fluxo de desenvolvimento dela.",
        "en_transition": "With the task created and filtered, we can move into its development flow.",
    },
    "10": {
        "en": (
            "Here, I use commit to mean a delivery commitment, not only a Git commit. We start by "
            "estimating cost, effort, and risk. Once the estimate is approved by the client, we enter "
            "the development process for the agreed change. Next, that implementation is reviewed and "
            "approved. Finally, the team moves into the commit layer, with a clear owner and clear "
            "expectations. This flow prevents a good idea from becoming a promise before there is a "
            "minimum understanding of the work required."
        ),
        "pt_transition": "Assumir a entrega não basta. Também precisamos ter nossos meios de garantir que a solução está realmente pronta antes de prosseguir.",
        "en_transition": "Committing to delivery is not enough. We also need ways to ensure the solution is truly ready before we move forward.",
    },
    "11": {
        "en": (
            "This slide answers the question: how do we know that a fix is really ready? Before the "
            "result goes back to the customer, we have our own QA stage. We reproduce the reported "
            "bug, isolate the context from the development and review agents, and run tests based "
            "directly on the requirements of the created task and on the messages in the Slack thread. "
            "Only after this round of end-to-end tests passes is the solution actually delivered to "
            "the reporter. This means the workflow does not rely only on a feeling that the problem "
            "has been solved."
        ),
        "pt_transition": "",
        "en_transition": "After validation, the delivery needs to return to the person who raised the signal.",
    },
    "12": {
        "en": (
            "A completed issue is not, by itself, the end of the process. The person who reported it "
            "needs to be able to see what happened. That is why the result returns to the thread with "
            "the status, testing evidence, and, when useful, a short explanation of the decision. "
            "This avoids a black-box effect and shows that the signal was heard, even when the answer "
            "is not simply an immediate change."
        ),
        "pt_transition": "Automação conecta o fluxo, mas algumas decisões continuam sendo humanas.",
        "en_transition": "Automation connects the flow, but some decisions remain human.",
    },
    "13": {
        "en": (
            "The idea of automation is to reduce repetitive processes and speed up the whole bug "
            "report resolution process, but it does not replace human judgement. Product priority, "
            "the evaluation of delivery trade-offs, and acceptance of the result still require human "
            "approval before the flow can continue."
        ),
        "pt_transition": "Com isso, podemos fechar com o workflow completo.",
        "en_transition": "With that in place, we can close with the complete workflow.",
    },
    "14": {
        "en": (
            "I would like to finish by thanking everyone who took the time to listen today. There are "
            "still many next steps ahead, such as adding Playwright to generate videos during the "
            "end-to-end testing stage and creating fully automated triage. A DefensePoint BookStack "
            "page has already been created with a step-by-step guide for building this workflow. If "
            "you have any questions or ideas, I am available to discuss them. Thank you very much."
        ),
        "pt_transition": "",
        "en_transition": "",
    },
}


def insert_after(paragraph):
    element = OxmlElement("w:p")
    paragraph._p.addnext(element)
    return Paragraph(element, paragraph._parent)


def remove(paragraph):
    paragraph._element.getparent().remove(paragraph._element)


def set_text(paragraph, text, *, lead=None):
    paragraph.clear()
    if lead:
        run = paragraph.add_run(lead)
        run.bold = True
    paragraph.add_run(text)


def paragraph_index(document, paragraph):
    return next(i for i, candidate in enumerate(document.paragraphs) if candidate._p is paragraph._p)


def upsert_transition(document, anchor, text, prefix):
    if not text:
        return
    paragraphs = document.paragraphs
    next_index = paragraph_index(document, anchor) + 1
    following = paragraphs[next_index] if next_index < len(paragraphs) else None
    if following and following.text.startswith(prefix):
        set_text(following, text, lead=prefix)
        return
    transition = insert_after(anchor)
    transition.style = "Normal"
    transition.paragraph_format.space_before = Pt(2)
    set_text(transition, text, lead=prefix)


def main():
    doc = Document(GUIDE)
    # Remove the former combined PT/EN transition paragraphs before placing the
    # two language-specific transitions beside their respective scripts.
    for paragraph in list(doc.paragraphs):
        if (
            paragraph.text.startswith("PT  ")
            or paragraph.text.startswith("Transição  ")
            or paragraph.text.startswith("Transition  ")
        ):
            remove(paragraph)

    paragraphs = doc.paragraphs
    headings = [
        (i, p, p.text.split()[1])
        for i, p in enumerate(paragraphs)
        if p.style.name == "Heading 1" and p.text.startswith("Slide ")
    ]

    slide_sections = []
    for heading_position, (start, heading, slide_id) in enumerate(headings):
        end = headings[heading_position + 1][0] if heading_position + 1 < len(headings) else len(paragraphs)
        section = paragraphs[start:end]
        portuguese_label = next(p for p in section if p.text == "Português")
        english_label = next(p for p in section if p.text == "English")
        slide_sections.append((slide_id, portuguese_label, english_label))

    for slide_id, portuguese_label, english_label in slide_sections:
        content = SLIDES[slide_id]
        current_paragraphs = doc.paragraphs
        pt_body = current_paragraphs[paragraph_index(doc, portuguese_label) + 1]
        en_body = current_paragraphs[paragraph_index(doc, english_label) + 1]
        set_text(en_body, content["en"])

        upsert_transition(doc, pt_body, content["pt_transition"], "Transição  ")
        upsert_transition(doc, en_body, content["en_transition"], "Transition  ")

    doc.save(GUIDE)
    print(GUIDE)


if __name__ == "__main__":
    main()
