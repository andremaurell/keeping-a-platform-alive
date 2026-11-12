from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


OUT = "/home/andremaurell/presentation-source-of-truth/gitlab-talk/Keeping a Platform Alive Speaker Guide.docx"


SLIDES = [
    {
        "n": "01", "title": "Keeping a Platform Alive", "time": "0:00–0:45",
        "purpose_pt": "Abrir com o problema que o fluxo resolve: como manter a plataforma atenta ao que acontece depois do release.",
        "pt": "Quero começar com uma ideia simples: uma plataforma continua viva depois que vai para produção. Usuários e clientes encontram problemas, fazem perguntas e têm ideias. O desafio não é apenas receber esse sinal. É transformar esse sinal em uma decisão visível, priorizada e testável, e depois retornar com uma resposta. Esta apresentação mostra o loop que conecta conversa, decisão e entrega.",
        "en": "I want to start with a simple idea: a platform stays alive after it goes to production. Users and customers find problems, ask questions, and have ideas. The challenge is not only receiving that signal. It is turning it into a visible, prioritised, and testable decision, then returning with an answer. This presentation shows the loop that connects conversation, decision, and delivery.",
        "transition_pt": "Primeiro, vale olhar para como essa ideia começou.",
        "transition_en": "First, it is useful to look at how this idea started.",
    },
    {
        "n": "02", "title": "Feedback in Context", "time": "0:45–1:45",
        "purpose_pt": "Contrastar a primeira ideia de feedback dentro do produto com o fluxo atual centrado na conversa.",
        "pt": "A primeira proposta era uma página de feedback dentro da plataforma, com um botão flutuante. Ele identificaria a rota atual, capturaria uma tela e permitiria anexar evidências. Era uma boa ideia porque mantinha o feedback perto do problema. Mas ela também exigia construir e manter mais uma superfície do produto. O fluxo atual preserva o mesmo princípio, usando o lugar onde o time já conversa todos os dias.",
        "en": "The first proposal was a feedback page inside the platform, with a floating button. It would identify the current route, capture a screenshot, and let people attach evidence. It was a good idea because it kept feedback close to the problem. But it also required building and maintaining another product surface. The current workflow keeps the same principle by using the place where the team already talks every day.",
        "transition_pt": "Esse lugar é o Slack.",
        "transition_en": "That place is Slack.",
    },
    {
        "n": "03", "title": "Slack as the Feedback Surface", "time": "1:45–2:45",
        "purpose_pt": "Mostrar o Slack como ponto de entrada natural e contextual para bug reports e feature requests.",
        "pt": "O Slack já contém as pessoas, a urgência, as perguntas de esclarecimento e as evidências em volta de um report. Por isso ele é um ponto de entrada melhor do que um formulário isolado. Neste exemplo anonimizado, alguém relata o problema, o bot cria a solicitação estruturada e o Linear cria a issue. A conversa não some quando a issue nasce; ela continua disponível para complementar o contexto.",
        "en": "Slack already contains the people, urgency, clarifying questions, and evidence around a report. That makes it a better entry point than an isolated form. In this anonymised example, someone reports the problem, the bot creates a structured request, and Linear creates the issue. The conversation does not disappear when the issue is created; it stays available to add context.",
        "transition_pt": "Para isso funcionar, o report inicial não precisa ser perfeito. Ele só precisa ser útil.",
        "transition_en": "For this to work, the initial report does not need to be perfect. It only needs to be useful.",
    },
    {
        "n": "04", "title": "What a Useful Report Contains", "time": "2:45–3:30",
        "purpose_pt": "Explicar o mínimo de informação que ajuda o time a agir sem criar burocracia para quem reporta.",
        "pt": "Um report útil carrega quatro coisas: a observação, o contexto, o impacto e alguma evidência. A observação descreve o que aconteceu sem tentar diagnosticar. O contexto mostra onde e em quais condições. O impacto explica por que vale olhar agora. E a evidência pode ser uma captura, um trace ou uma conversa. Isso dá ao time um ponto de partida sem exigir que o usuário complete um formulário longo.",
        "en": "A useful report carries four things: the observation, the context, the impact, and some evidence. The observation describes what happened without trying to diagnose it. The context shows where and under which conditions it happened. The impact explains why it matters now. Evidence can be a capture, a trace, or a conversation. This gives the team a starting point without requiring the user to complete a long form.",
        "transition_pt": "A partir desse sinal, a automação faz o trabalho repetitivo.",
        "transition_en": "From that signal, automation handles the repetitive work.",
    },
    {
        "n": "05", "title": "Intake Automation", "time": "3:30–4:30",
        "purpose_pt": "Explicar a passagem automatizada de uma thread para uma issue estruturada, com controle de duplicidade e contexto.",
        "pt": "A automação lê o report e a thread, procura possíveis duplicidades e cria ou enriquece uma issue no Linear. O ponto importante é que ela não tenta tomar todas as decisões de produto. Ela transforma a conversa em um registro estruturado e deixa explícito de onde aquele registro veio. Isso reduz trabalho manual e evita que o mesmo problema vire tickets desconectados.",
        "en": "The automation reads the report and the thread, looks for possible duplicates, and creates or enriches an issue in Linear. The important point is that it does not try to make every product decision. It turns the conversation into a structured record and makes the source of that record explicit. This reduces manual work and prevents the same problem from becoming disconnected tickets.",
        "transition_pt": "O princípio é o mesmo nos dois produtos, embora a triagem seja diferente.",
        "transition_en": "The operating principle is the same in both products, even though triage differs.",
    },
    {
        "n": "06", "title": "Two Flows One Operating Principle", "time": "4:30–5:30",
        "purpose_pt": "Comparar RedLens e Pedaxis sem perder a mensagem central comum.",
        "pt": "RedLens e Pedaxis têm canais, cadências e regras de encaminhamento próprias. Ainda assim, os dois fluxos seguem o mesmo princípio: capturar o sinal no contexto, criar uma decisão rastreável, encaminhar para quem pode agir e devolver o resultado para a conversa original. Essa separação permite adaptar a operação ao produto sem reinventar o modelo inteiro a cada vez.",
        "en": "RedLens and Pedaxis have their own channels, cadences, and routing rules. Even so, both flows follow the same principle: capture the signal in context, create a traceable decision, route it to the people who can act, and return the result to the original conversation. This separation lets us adapt the operation to each product without reinventing the whole model every time.",
        "transition_pt": "O que mantém esse modelo unido é a própria thread.",
        "transition_en": "What keeps this model connected is the thread itself.",
    },
    {
        "n": "07", "title": "The Thread Becomes Living Context", "time": "5:30–6:30",
        "purpose_pt": "Demonstrar que a thread acumula contexto depois da criação da issue.",
        "pt": "A thread não é uma fila de tickets. Ela é a camada narrativa em volta do registro de decisão. Neste mock, a issue já foi criada e o reporter adiciona uma nova observação. Essa atualização é importante porque o contexto real quase sempre chega em partes. Em vez de abrir uma segunda issue ou perder o detalhe em uma mensagem solta, a conversa continua ligada ao mesmo trabalho.",
        "en": "A thread is not a ticket queue. It is the narrative layer around the decision record. In this mock, the issue has already been created and the reporter adds a new observation. That update matters because real context almost always arrives in parts. Instead of opening a second issue or losing the detail in a separate message, the conversation stays connected to the same work.",
        "transition_pt": "A conversa evolui no Slack. A decisão fica registrada no Linear.",
        "transition_en": "The conversation evolves in Slack. The decision is recorded in Linear.",
    },
    {
        "n": "08", "title": "Linear is the Decision Record", "time": "6:30–7:30",
        "purpose_pt": "Separar claramente a camada de conversa da camada de decisão operacional.",
        "pt": "No Linear, a issue deixa clara a decisão operacional. O assignee mostra quem responde pelo próximo passo. O scope delimita a mudança. O state mostra o compromisso atual. E a evidence comprova o encerramento. O screenshot mostra como esses campos convivem em uma issue real, mas a mensagem principal é simples: Slack preserva a conversa; Linear registra a decisão que o time vai executar.",
        "en": "In Linear, the issue makes the operational decision clear. The assignee identifies the person responsible for the next step. The scope bounds the change. The state shows the current commitment. And the evidence proves closure. The screenshot shows how these fields sit together in a real issue, but the main message is simple: Slack preserves the conversation; Linear records the decision the team will execute.",
        "transition_pt": "Depois de registrar a decisão, precisamos escolher o próximo passo certo.",
        "transition_en": "After recording the decision, we need to choose the right next step.",
    },
    {
        "n": "09", "title": "Triage Turns a Report into a Decision", "time": "7:30–9:00",
        "purpose_pt": "Mostrar como a triagem transforma uma conversa aberta em uma decisão operacional com um próximo passo acionável.",
        "pt": "Neste ponto, triage não é uma tabela de classificações. É uma sequência simples de decisões. Primeiro, coletamos o sinal: lemos o report, a thread e as evidências. Depois, avaliamos o impacto: urgência, escopo e se o comportamento se repete. Por fim, encaminhamos o trabalho, definindo assignee, prioridade e destino no Linear. A saída da triagem é uma decisão que alguém consegue executar, não apenas uma etiqueta.",
        "en": "At this point, triage is not just a classification table. It is a simple sequence of decisions. First, we collect the signal: we read the report, the thread, and the evidence. Then we assess the impact: urgency, scope, and whether the behaviour repeats. Finally, we route the work by setting an assignee, priority, and destination in Linear. The output of triage is a decision that someone can act on, not only a label.",
        "transition_pt": "Quando existe trabalho a fazer, o próximo passo é transformar intenção em compromisso de entrega.",
        "transition_en": "When there is work to do, the next step is turning intent into a delivery commitment.",
    },
    {
        "n": "10", "title": "Estimation and Approval", "time": "9:00–10:00",
        "purpose_pt": "Explicar o fluxo que leva uma necessidade triada a uma entrega assumida pelo time.",
        "pt": "Aqui eu uso commit no sentido de compromisso de entrega, não apenas de um commit de Git. Primeiro estimamos escopo, esforço e risco. Depois desenvolvemos a mudança que foi combinada. Em seguida, alguém revisa e aprova a implementação. Por fim, o time assume o compromisso de entrega com owner e expectativa claros. Esse fluxo evita que uma boa ideia vire promessa antes de existir um entendimento mínimo do trabalho.",
        "en": "Here I use commit to mean a delivery commitment, not only a Git commit. First, we estimate scope, effort, and risk. Then we develop the agreed change. Next, someone reviews and approves the implementation. Finally, the team makes a delivery commitment with a clear owner and clear expectations. This flow prevents a good idea from becoming a promise before there is a minimum understanding of the work.",
        "transition_pt": "Assumir a entrega não basta. Também precisamos saber quando uma correção está pronta para seguir.",
        "transition_en": "Making the commitment is not enough. We also need to know when a fix is ready to move forward.",
    },
    {
        "n": "11", "title": "Before Shipping a Fix", "time": "10:00–11:00",
        "purpose_pt": "Definir objetivamente o que deixa uma correção pronta para envio.",
        "pt": "Este slide responde à pergunta: como sabemos que uma correção está realmente pronta? Nós reproduzimos o report, isolamos a causa, fazemos a mudança, verificamos o comportamento, revisamos e só então enviamos. O critério curto é: o report reproduz, os testes passam e a revisão confirma a mudança. Assim, o fluxo não depende apenas de uma sensação de que o problema foi resolvido.",
        "en": "This slide answers the question: how do we know that a fix is really ready? We reproduce the report, isolate the cause, make the change, verify the behaviour, review it, and only then ship it. The short criterion is: the report reproduces, the tests pass, and the review confirms the change. This means the workflow does not depend only on a feeling that the problem is solved.",
        "transition_pt": "Depois da validação, a entrega precisa voltar para quem levantou o sinal.",
        "transition_en": "After validation, the delivery needs to return to the person who raised the signal.",
    },
    {
        "n": "12", "title": "RedLens Delivery Harness", "time": "10:30–11:30",
        "purpose_pt": "Aplicar o modelo de prontidão ao fluxo local do RedLens.",
        "pt": "No RedLens, o harness resume o ciclo em quatro movimentos: reproduzir, mudar, verificar e retornar. A thread vira um caso local confiável. A alteração fica pequena e observável. A verificação testa o comportamento, não só a implementação. E o retorno fecha o ciclo com evidência na conversa original. É o mesmo princípio do slide anterior, aplicado de forma mais direta ao contexto do RedLens.",
        "en": "In RedLens, the harness summarises the cycle in four moves: reproduce, change, verify, and return. The thread becomes a reliable local case. The change stays small and observable. Verification tests behaviour, not only implementation. And the return closes the loop with evidence in the original conversation. It is the same principle from the previous slide, applied more directly to the RedLens context.",
        "transition_pt": "Depois da validação, a entrega precisa voltar para quem levantou o sinal.",
        "transition_en": "After validation, the delivery needs to return to the person who raised the signal.",
    },
    {
        "n": "13", "title": "Closing the Loop", "time": "11:00–12:00",
        "purpose_pt": "Mostrar que a entrega só se completa quando o resultado é comunicado no contexto original.",
        "pt": "Uma issue concluída não é, por si só, o fim do processo. A pessoa que reportou precisa conseguir ver o que aconteceu. Por isso o resultado volta para a thread com o status, a evidência de teste e, quando fizer sentido, uma explicação curta da decisão. Isso evita o efeito de caixa preta e mostra que o sinal foi ouvido, mesmo quando a resposta não é simplesmente uma mudança imediata.",
        "en": "A completed issue is not, by itself, the end of the process. The person who reported it needs to be able to see what happened. That is why the result returns to the thread with the status, testing evidence, and, when useful, a short explanation of the decision. This avoids a black-box effect and shows that the signal was heard, even when the answer is not an immediate change.",
        "transition_pt": "Automação conecta o fluxo, mas algumas decisões continuam sendo humanas.",
        "transition_en": "Automation connects the flow, but some decisions remain human.",
    },
    {
        "n": "14", "title": "Human Checkpoints", "time": "12:00–13:00",
        "purpose_pt": "Deixar claro quais decisões não devem ser delegadas ao bot.",
        "pt": "O bot reduz coordenação repetitiva, mas não substitui julgamento. Pessoas continuam decidindo prioridade de produto, avaliando o trade-off de uma entrega, aprovando compromisso e aceitando o resultado. Essa fronteira é importante: automação deve preparar contexto e acelerar a operação, enquanto o time mantém responsabilidade pelas escolhas que têm impacto no produto e no cliente.",
        "en": "The bot reduces repetitive coordination, but it does not replace judgement. People still decide product priority, assess delivery trade-offs, approve a commitment, and accept the result. This boundary matters: automation should prepare context and speed up the operation, while the team keeps responsibility for choices that affect the product and the customer.",
        "transition_pt": "Com isso, podemos fechar com o padrão completo.",
        "transition_en": "With that in place, we can close with the full pattern.",
    },
    {
        "n": "15", "title": "The Operating Loop", "time": "13:00–15:00",
        "purpose_pt": "Fechar a apresentação reforçando o padrão reutilizável e o benefício operacional.",
        "pt": "O padrão completo é um loop. O Slack guarda a conversa e o contexto vivo. O Linear registra a decisão estruturada. A triagem dá um próximo passo claro. O harness transforma a decisão em uma mudança testada. E o resultado retorna para a mesma conversa onde o sinal começou. O valor não está apenas em criar tickets automaticamente. Está em manter uma rota visível do report até uma entrega validada.",
        "en": "The complete pattern is a loop. Slack keeps the conversation and its living context. Linear records the structured decision. Triage gives a clear next step. The harness turns the decision into a tested change. And the result returns to the same conversation where the signal began. The value is not only in creating tickets automatically. It is in keeping a visible route from the report to a validated delivery.",
        "transition_pt": "Obrigado. Posso detalhar qualquer parte do fluxo ou mostrar um exemplo real anonimizado.",
        "transition_en": "Thank you. I can go deeper into any part of the flow or show an anonymised real example.",
    },
]

# The final deck omits the previous RedLens Delivery Harness page.
REMOVED_SLIDE_IDS = {"12"}


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    elem = OxmlElement("w:tblHeader")
    elem.set(qn("w:val"), "true")
    tr_pr.append(elem)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Page ")
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    paragraph._p.append(fld)


def set_margins(section):
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)


def style_document(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Aptos"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.space_after = Pt(4)

    title = styles["Title"]
    title.font.name = "Aptos Display"
    title._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos Display")
    title.font.size = Pt(26)
    title.font.bold = True
    title.font.color.rgb = RGBColor(0, 0, 0)

    for name, size in [("Heading 1", 17), ("Heading 2", 12)]:
        style = styles[name]
        style.font.name = "Aptos Display"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos Display")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(8)
        style.paragraph_format.space_after = Pt(4)

    if "Guide label" not in styles:
        lbl = styles.add_style("Guide label", WD_STYLE_TYPE.PARAGRAPH)
        lbl.font.name = "Aptos"
        lbl.font.size = Pt(9.5)
        lbl.font.bold = True
        lbl.font.color.rgb = RGBColor(0, 0, 0)
        lbl.paragraph_format.space_before = Pt(5)
        lbl.paragraph_format.space_after = Pt(2)


def add_label(doc, text):
    p = doc.add_paragraph(style="Guide label")
    p.add_run(text)
    return p


def add_slide(doc, slide):
    doc.add_heading(f"Slide {slide['n']}  {slide['title']}", level=1)
    meta = doc.add_paragraph()
    meta.paragraph_format.space_after = Pt(8)
    r = meta.add_run(f"Suggested time  {slide['time']}")
    r.bold = True
    r.font.size = Pt(10)

    add_label(doc, "Goal")
    doc.add_paragraph(slide["purpose_pt"])

    add_label(doc, "Português")
    doc.add_paragraph(slide["pt"])
    pt_transition = doc.add_paragraph()
    pt_transition.paragraph_format.space_before = Pt(2)
    pt_transition.add_run("Transição  ").bold = True
    pt_transition.add_run(slide["transition_pt"])

    add_label(doc, "English")
    doc.add_paragraph(slide["en"])
    en_transition = doc.add_paragraph()
    en_transition.paragraph_format.space_before = Pt(2)
    en_transition.add_run("Transition  ").bold = True
    en_transition.add_run(slide["transition_en"])


def main():
    doc = Document()
    section = doc.sections[0]
    set_margins(section)
    style_document(doc)

    header = section.header.paragraphs[0]
    header.text = "KEEPING A PLATFORM ALIVE  |  SPEAKER GUIDE"
    header.style = doc.styles["Normal"]
    for run in header.runs:
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = RGBColor(90, 90, 90)
    add_page_number(section.footer.paragraphs[0])
    for run in section.footer.paragraphs[0].runs:
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(90, 90, 90)

    doc.add_paragraph("Keeping a Platform Alive Speaker Guide", style="Title")
    sub = doc.add_paragraph()
    sub.add_run("Bilingual Portuguese and English talk track for a 15 minute presentation").italic = True
    sub.paragraph_format.space_after = Pt(14)

    doc.add_heading("How to use this guide", level=1)
    doc.add_paragraph(
        "This is a rehearsal guide, not a script to read word for word. For each slide, choose one language for the room and use the other version to prepare. The suggested timing totals approximately 15 minutes. Keep the transitions short so the deck feels like one connected story."
    )
    doc.add_paragraph(
        "The central message is consistent throughout: Slack keeps the conversation and its evolving evidence. Linear records the accountable decision. The delivery harness turns that decision into a tested change. The outcome returns to the original thread."
    )
    doc.add_page_break()

    remaining_slides = [slide for slide in SLIDES if slide["n"] not in REMOVED_SLIDE_IDS]
    for new_number, source_slide in enumerate(remaining_slides, start=1):
        slide = dict(source_slide)
        slide["n"] = f"{new_number:02d}"
        add_slide(doc, slide)

    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
