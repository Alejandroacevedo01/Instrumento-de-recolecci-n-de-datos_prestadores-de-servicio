from __future__ import annotations

import html
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.document import Document as DocumentObject
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P


ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
DOCX_FILES = sorted(ROOT.glob("*.docx"))


@dataclass
class Option:
    text: str
    multiple: bool = False


@dataclass
class ScaleItem:
    text: str
    values: list[str]


@dataclass
class Question:
    qid: str
    text: str
    qtype: str = "texto_largo"
    required: bool = False
    logic: str = ""
    note: str = ""
    options: list[Option] = field(default_factory=list)
    scale_items: list[ScaleItem] = field(default_factory=list)
    max_choices: int | None = None


@dataclass
class Section:
    title: str
    description: str = ""
    questions: list[Question] = field(default_factory=list)


def iter_blocks(parent: DocumentObject) -> Iterable[Paragraph | Table]:
    for child in parent.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def clean(text: str) -> str:
    return " ".join(text.replace("\u00a0", " ").split())


def parse_meta(text: str) -> dict[str, str]:
    text = re.sub(r"^CONFIGURACIÓN:\s*", "", text.strip(), flags=re.I)
    result: dict[str, str] = {}
    logic_parts: list[str] = []
    for part in (p.strip() for p in text.split(";") if p.strip()):
        if "=" in part:
            key, value = part.split("=", 1)
            key = key.strip().lower()
            if key in {"id", "tipo", "obligatoria"}:
                result[key] = value.strip()
            else:
                logic_parts.append(part)
        else:
            logic_parts.append(part)
    result["lógica"] = "; ".join(logic_parts)
    return result


def question_id_and_text(text: str) -> tuple[str, str]:
    match = re.match(r"^([A-Z]\d+[A-Z]?)\.\s*(.+)$", clean(text))
    if not match:
        return "Q", clean(text)
    return match.group(1), match.group(2)


def max_choices_from_note(note: str) -> int | None:
    match = re.search(r"máximo\s+(\w+|\d+)", note, flags=re.I)
    if not match:
        return None
    word = match.group(1).lower()
    numbers = {
        "una": 1,
        "uno": 1,
        "dos": 2,
        "tres": 3,
        "cuatro": 4,
        "cinco": 5,
        "seis": 6,
    }
    return int(word) if word.isdigit() else numbers.get(word)


def parse_document(path: Path) -> tuple[list[str], list[Section], str]:
    doc = Document(path)
    consent: list[str] = []
    sections: list[Section] = []
    current_section: Section | None = None
    current_question: Question | None = None
    in_consent = False
    in_survey = False
    confirmation = (
        "Gracias por participar. Sus respuestas fueron registradas. "
        "Si desea participar en una entrevista de seguimiento, podrá usar "
        "el formulario independiente indicado al final."
    )

    for block in iter_blocks(doc):
        if isinstance(block, Table):
            if current_question and current_question.qtype == "escala_por_item":
                rows = block.rows
                if not rows:
                    continue
                values = [clean(c.text) for c in rows[0].cells[1:]]
                for row in rows[1:]:
                    item = clean(row.cells[0].text)
                    if item:
                        current_question.scale_items.append(ScaleItem(item, values))
            continue

        text = clean(block.text)
        style = block.style.name if block.style else ""
        if not text:
            continue

        if text == "CONSENTIMIENTO INFORMADO":
            in_consent = True
            continue
        if text == "ENCUESTA SOBRE INTERPRETACIÓN JUDICIAL EN COLOMBIA":
            in_survey = True
            in_consent = False
            continue
        if text == "MENSAJE DE CONFIRMACIÓN":
            in_survey = False
            current_question = None
            continue
        if text == "ESPECIFICACIONES PARA PROGRAMACIÓN DIGITAL":
            break
        if text.startswith("Texto para mostrar después del envío:"):
            confirmation = text.split(":", 1)[1].strip()
            continue

        if in_consent and style != "Question":
            if style == "Heading 2":
                consent.append(f"<h2>{html.escape(text)}</h2>")
            elif style == "List Bullet":
                consent.append(f"<p class=\"bullet\">• {html.escape(text)}</p>")
            elif style not in {"Question Meta", "Form Option", "Form Note"}:
                consent.append(f"<p>{html.escape(text)}</p>")

        if style == "Heading 1" and text.startswith("SECCIÓN"):
            current_section = Section(text)
            sections.append(current_section)
            current_question = None
            continue

        if in_survey and current_section and style == "Normal" and not current_section.questions:
            if not text.startswith("_") and not text.startswith("Instrucción general"):
                current_section.description = text
            continue

        if style == "Question":
            qid, qtext = question_id_and_text(text)
            current_question = Question(qid=qid, text=qtext)
            if qid == "C01":
                consent_section = Section("Consentimiento informado", questions=[current_question])
                sections.insert(0, consent_section)
            elif current_section:
                current_section.questions.append(current_question)
            continue

        if not current_question:
            continue

        if style == "Question Meta":
            meta = parse_meta(text)
            current_question.qid = meta.get("id", current_question.qid)
            current_question.qtype = meta.get("tipo", current_question.qtype)
            current_question.required = meta.get("obligatoria", "no").lower() in {"sí", "si", "true"}
            current_question.logic = meta.get("lógica", "")
        elif style == "Form Note":
            current_question.note = text
            current_question.max_choices = max_choices_from_note(text)
        elif style == "Form Option":
            multiple = text.startswith("□")
            option_text = re.sub(r"^[○□]\s*", "", text)
            current_question.options.append(Option(option_text, multiple))

    return consent, sections, confirmation


def esc(value: str) -> str:
    return html.escape(value, quote=True)


def condition_for(qid: str) -> str:
    return {
        "P07A": "P07=Sí",
        "P14": "P13=No",
        "P15": "P13=Sí",
        "P16": "P13=Sí",
        "P17": "P13=Sí",
        "P18": "P13=Sí",
        "P26": "P25=No",
        "P27": "P25=Sí",
    }.get(qid, "")


def is_exclusive_option(text: str) -> bool:
    low = text.lower()
    return low.startswith(("prefiero no", "no cuento", "no he trabajado", "aún no he", "no aplica"))


def needs_detail(text: str) -> bool:
    low = text.lower().strip(". ")
    return "especifique" in low or low in {"otra", "otro", "otra acreditación o certificación"}


def render_choice_question(q: Question) -> str:
    multiple = q.qtype == "selección_multiple"
    input_type = "checkbox" if multiple else "radio"
    parts = []
    if q.note:
        parts.append(f'<p class="question-note" id="{esc(q.qid)}-note">{esc(q.note)}</p>')
    if q.max_choices:
        parts.append(f'<p class="selection-status" id="{esc(q.qid)}-status" aria-live="polite"></p>')
    parts.append('<div class="options">')
    for index, option in enumerate(q.options, 1):
        oid = f"{q.qid}-{index}"
        attrs = []
        if multiple and is_exclusive_option(option.text):
            attrs.append('data-exclusive="true"')
        parts.append('<div class="option-row">')
        parts.append(
            f'<input type="{input_type}" id="{esc(oid)}" name="{esc(q.qid)}" '
            f'value="{esc(option.text)}" {' '.join(attrs)}>'
        )
        parts.append(f'<label for="{esc(oid)}">{esc(option.text)}</label>')
        if needs_detail(option.text):
            parts.append(
                f'<label class="visually-hidden" for="{esc(oid)}-detail">Especifique: {esc(option.text)}</label>'
                f'<input class="detail-input" id="{esc(oid)}-detail" name="{esc(q.qid)}_detalle_{index}" '
                f'type="text" placeholder="Especifique (opcional)" disabled>'
            )
        parts.append('</div>')
    parts.append('</div>')
    return "".join(parts)


def render_scale_question(q: Question) -> str:
    parts = []
    if q.note:
        parts.append(f'<p class="question-note">{esc(q.note)}</p>')
    for row_index, item in enumerate(q.scale_items, 1):
        name = f"{q.qid}_{row_index}"
        parts.append(f'<fieldset class="scale-item"><legend>{esc(item.text)}</legend><div class="scale-options">')
        for value_index, value in enumerate(item.values, 1):
            oid = f"{name}-{value_index}"
            parts.append(
                f'<span><input type="radio" id="{esc(oid)}" name="{esc(name)}" value="{esc(value)}">'
                f'<label for="{esc(oid)}">{esc(value)}</label></span>'
            )
        parts.append('</div></fieldset>')
    return "".join(parts)


def render_question(q: Question) -> str:
    condition = condition_for(q.qid)
    attrs = [f'id="question-{esc(q.qid)}"', 'class="question-block"']
    if condition:
        attrs.append(f'data-condition="{esc(condition)}"')
        attrs.append('hidden')
    if q.max_choices:
        attrs.append(f'data-max-choices="{q.max_choices}"')
    if q.required:
        attrs.append('data-required="true"')

    required_text = '<span class="required">Obligatoria</span>' if q.required else '<span class="optional">Opcional</span>'
    legend = f'<legend><span class="question-id">{esc(q.qid)}</span> {esc(q.text)} {required_text}</legend>'
    if q.qtype in {"selección_unica", "selección_multiple"}:
        body = render_choice_question(q)
    elif q.qtype == "escala_por_item":
        body = render_scale_question(q)
    else:
        body = (
            f'<label class="visually-hidden" for="{esc(q.qid)}">{esc(q.text)}</label>'
            f'<textarea id="{esc(q.qid)}" name="{esc(q.qid)}" rows="5"></textarea>'
        )
    return f'<fieldset {' '.join(attrs)}>{legend}{body}<p class="field-error" id="error-{esc(q.qid)}"></p></fieldset>'


def display_section_title(title: str) -> str:
    if title.upper().startswith("SECCIÓN") and "." in title:
        prefix, remainder = title.split(".", 1)
        return f"{prefix.title()}. {remainder.strip().capitalize()}"
    return title


def render_section(section: Section, index: int, total: int, consent: list[str]) -> str:
    is_consent = index == 0
    body = []
    if is_consent:
        body.extend(consent)
    elif section.description:
        body.append(f'<p class="section-description">{esc(section.description)}</p>')
    body.extend(render_question(q) for q in section.questions)

    back = '' if index == 0 else '<button type="button" class="secondary back-button">Anterior</button>'
    if index == total - 1:
        forward = '<button type="submit" class="primary">Enviar respuestas</button>'
    else:
        forward = '<button type="button" class="primary next-button">Continuar</button>'
    controls = f'<div class="step-controls">{back}{forward}</div>'
    hidden = '' if index == 0 else ' hidden'
    return (
        f'<section class="form-step" data-step="{index}" aria-labelledby="step-title-{index}"{hidden}>'
        f'<p class="step-count">Paso {index + 1} de {total}</p>'
        f'<h1 id="step-title-{index}" tabindex="-1">{esc(display_section_title(section.title))}</h1>'
        f'{"".join(body)}{controls}</section>'
    )


def build_html(consent: list[str], sections: list[Section], confirmation: str) -> str:
    endpoint = os.environ.get("FORM_ENDPOINT", "").strip()
    contact_url = os.environ.get("CONTACT_FORM_URL", "").strip()
    endpoint_attr = esc(endpoint)
    preview = not endpoint
    sections_html = "".join(render_section(section, i, len(sections), consent) for i, section in enumerate(sections))
    confirmation_text = confirmation.replace("[ENLACE AL FORMULARIO DE CONTACTO]", "")
    contact = (
        f'<p><a class="contact-link" href="{esc(contact_url)}" rel="noopener">Registrar datos para la entrevista de seguimiento</a></p>'
        if contact_url
        else '<p class="preview-only">El enlace independiente para entrevistas se agregará antes de la publicación final.</p>'
    )
    preview_banner = (
        '<div class="preview-banner" role="status"><strong>Modo de prueba:</strong> '
        'puede recorrer y validar el formulario, pero las respuestas aún no se almacenan.</div>'
        if preview
        else ""
    )
    return f'''<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="description" content="Encuesta académica sobre interpretación judicial en Colombia">
  <title>Encuesta sobre interpretación judicial en Colombia</title>
  <link rel="stylesheet" href="styles.css">
</head>
<body data-preview="{str(preview).lower()}">
  <a class="skip-link" href="#survey-form">Ir al formulario</a>
  <header class="site-header">
    <p class="institution">Universidad de Antioquia · Escuela de Idiomas</p>
    <p>Maestría en Traducción con Énfasis en Interpretación</p>
  </header>
  <main>
    {preview_banner}
    <div id="error-summary" class="error-summary" role="alert" tabindex="-1" hidden>
      <h2>Revise la información</h2><ul></ul>
    </div>
    <form id="survey-form" action="{endpoint_attr}" method="post" target="submission-target" novalidate>
      <input type="hidden" name="instrumento" value="interpretacion_judicial_colombia">
      {sections_html}
    </form>
    <section id="declined" class="result-panel" tabindex="-1" hidden>
      <h1>Participación finalizada</h1>
      <p>Ha indicado que no acepta participar. No se enviaron respuestas del cuestionario.</p>
    </section>
    <section id="confirmation" class="result-panel" tabindex="-1" hidden>
      <h1>Gracias por participar</h1>
      <p>{esc(confirmation_text)}</p>
      {contact}
    </section>
    <iframe name="submission-target" id="submission-target" title="Destino de envío" hidden></iframe>
  </main>
  <footer><p>Investigador: Alejandro Acevedo · alejandro.acevedo@udea.edu.co</p></footer>
  <script src="app.js"></script>
</body>
</html>'''


CSS = r'''
:root { color-scheme: light; --green:#006b3c; --green-dark:#004d2c; --cream:#f7f8f5; --ink:#17221c; --muted:#526158; --border:#b8c3bc; --error:#a40000; }
* { box-sizing:border-box; }
body { margin:0; color:var(--ink); background:var(--cream); font-family:Arial, Helvetica, sans-serif; line-height:1.55; }
.skip-link { position:absolute; left:-9999px; top:0; padding:.75rem; background:#fff; color:#000; z-index:10; }
.skip-link:focus { left:1rem; top:1rem; }
.site-header, footer { background:var(--green-dark); color:#fff; padding:1.25rem max(1rem, calc((100% - 860px)/2)); }
.site-header p, footer p { margin:.15rem 0; }
.institution { font-weight:700; }
main { width:min(860px, calc(100% - 2rem)); margin:2rem auto; }
.preview-banner { border:2px solid #856404; background:#fff3cd; color:#4d3b00; padding:1rem; margin-bottom:1.5rem; }
.form-step { background:#fff; border:1px solid var(--border); border-radius:.75rem; padding:clamp(1rem, 3vw, 2rem); box-shadow:0 2px 10px rgba(0,0,0,.06); }
.step-count { color:var(--muted); font-weight:700; margin:0; }
h1 { line-height:1.2; margin:.4rem 0 1rem; }
h2 { font-size:1.15rem; margin-top:1.6rem; }
.section-description { border-left:4px solid var(--green); padding-left:1rem; }
.bullet { padding-left:1rem; }
.question-block, .scale-item { border:0; border-top:1px solid #dfe5e1; padding:1.5rem 0; margin:0; min-inline-size:0; }
.question-block > legend, .scale-item > legend { font-weight:700; padding:0 .3rem 0 0; max-width:100%; }
.question-id { color:var(--green-dark); }
.required, .optional { display:inline-block; margin-left:.5rem; font-size:.78rem; font-weight:700; padding:.12rem .4rem; border-radius:.25rem; }
.required { color:#fff; background:var(--green-dark); }
.optional { color:var(--muted); border:1px solid var(--border); }
.question-note, .selection-status { color:var(--muted); margin:.5rem 0 1rem; }
.option-row { display:flex; flex-wrap:wrap; align-items:flex-start; gap:.55rem; margin:.7rem 0; }
input[type=radio], input[type=checkbox] { inline-size:1.2rem; block-size:1.2rem; margin-top:.2rem; flex:0 0 auto; }
.option-row > label { flex:1 1 70%; }
textarea, input[type=text] { width:100%; border:2px solid #718078; border-radius:.35rem; padding:.7rem; font:inherit; background:#fff; color:inherit; }
.detail-input { margin-left:1.75rem; width:min(30rem, calc(100% - 1.75rem)); }
.scale-item { margin-left:.75rem; }
.scale-options { display:flex; flex-wrap:wrap; gap:.6rem 1rem; margin-top:.65rem; }
.scale-options span { display:flex; align-items:center; gap:.35rem; min-width:3.3rem; }
button { border-radius:.35rem; padding:.75rem 1.15rem; font:inherit; font-weight:700; cursor:pointer; }
.primary { color:#fff; background:var(--green); border:2px solid var(--green); }
.secondary { color:var(--green-dark); background:#fff; border:2px solid var(--green-dark); }
button:hover { filter:brightness(.92); }
button:focus-visible, input:focus-visible, textarea:focus-visible, a:focus-visible { outline:4px solid #f3b700; outline-offset:3px; }
.step-controls { display:flex; justify-content:space-between; gap:1rem; margin-top:1.5rem; }
.field-error { color:var(--error); font-weight:700; margin:.5rem 0 0; }
.has-error { border-left:4px solid var(--error); padding-left:1rem; }
.error-summary { background:#fff; border:3px solid var(--error); padding:1rem 1.25rem; margin-bottom:1.5rem; }
.error-summary h2 { margin-top:0; }
.error-summary a { color:var(--error); font-weight:700; }
.result-panel { background:#fff; border:1px solid var(--border); border-radius:.75rem; padding:2rem; }
.preview-only { font-weight:700; color:#6b5100; }
.visually-hidden { position:absolute!important; width:1px!important; height:1px!important; padding:0!important; margin:-1px!important; overflow:hidden!important; clip:rect(0,0,0,0)!important; white-space:nowrap!important; border:0!important; }
[hidden] { display:none!important; }
@media (max-width:560px) { main { width:min(100% - 1rem, 860px); margin:1rem auto; } .form-step { border-radius:.4rem; padding:1rem; } .step-controls { flex-direction:column-reverse; } button { width:100%; } .scale-options { display:grid; grid-template-columns:repeat(3, 1fr); } }
'''


JS = r'''
(() => {
  const form = document.getElementById('survey-form');
  const steps = [...document.querySelectorAll('.form-step')];
  const summary = document.getElementById('error-summary');
  const summaryList = summary.querySelector('ul');
  const preview = document.body.dataset.preview === 'true';
  let current = 0;
  let iframeLoaded = false;

  const fieldValues = (name) => [...form.querySelectorAll(`[name="${CSS.escape(name)}"]:checked`)].map(el => el.value);
  const singleValue = (name) => fieldValues(name)[0] || '';

  function updateConditions() {
    document.querySelectorAll('[data-condition]').forEach(block => {
      const [source, expected] = block.dataset.condition.split('=');
      const visible = singleValue(source) === expected;
      block.hidden = !visible;
      block.querySelectorAll('input, textarea, select').forEach(control => {
        control.disabled = !visible;
      });
    });
  }

  function updateDetailInputs(source) {
    const row = source.closest('.option-row');
    if (!row) return;
    const detail = row.querySelector('.detail-input');
    if (detail) {
      detail.disabled = !source.checked;
      if (!source.checked) detail.value = '';
    }
  }

  function enforceExclusive(source) {
    if (source.type !== 'checkbox') return;
    const group = [...form.querySelectorAll(`input[type="checkbox"][name="${CSS.escape(source.name)}"]`)];
    if (source.dataset.exclusive === 'true' && source.checked) {
      group.filter(el => el !== source).forEach(el => { el.checked = false; updateDetailInputs(el); });
    } else if (source.checked) {
      group.filter(el => el.dataset.exclusive === 'true').forEach(el => { el.checked = false; });
    }
  }

  function updateLimit(block) {
    const max = Number(block.dataset.maxChoices || 0);
    if (!max) return true;
    const checked = [...block.querySelectorAll('input[type="checkbox"]:checked')];
    const status = block.querySelector('.selection-status');
    if (status) status.textContent = `${checked.length} de máximo ${max} opciones seleccionadas.`;
    const valid = checked.length <= max;
    block.classList.toggle('has-error', !valid);
    block.querySelector('.field-error').textContent = valid ? '' : `Seleccione máximo ${max} opciones.`;
    return valid;
  }

  function clearErrors(step) {
    step.querySelectorAll('.has-error').forEach(el => el.classList.remove('has-error'));
    step.querySelectorAll('.field-error').forEach(el => { el.textContent = ''; });
    summary.hidden = true;
    summaryList.innerHTML = '';
  }

  function validateStep(step) {
    clearErrors(step);
    const errors = [];
    step.querySelectorAll('.question-block:not([hidden])').forEach(block => {
      const qid = block.id.replace('question-', '');
      if (block.dataset.required === 'true') {
        const hasChoice = block.querySelector('input[type="radio"]:checked, input[type="checkbox"]:checked');
        const text = block.querySelector('textarea, input[type="text"]');
        if (!hasChoice && !(text && text.value.trim())) {
          block.classList.add('has-error');
          block.querySelector('.field-error').textContent = 'Esta pregunta es obligatoria.';
          errors.push({ block, message: `${qid}: responda esta pregunta.` });
        }
      }
      if (!updateLimit(block)) errors.push({ block, message: `${qid}: reduzca el número de opciones.` });
    });
    if (errors.length) {
      errors.forEach(({block, message}) => {
        const li = document.createElement('li');
        const a = document.createElement('a');
        a.href = `#${block.id}`;
        a.textContent = message;
        li.appendChild(a); summaryList.appendChild(li);
      });
      summary.hidden = false; summary.focus();
      return false;
    }
    return true;
  }

  function showStep(index) {
    steps.forEach((step, i) => { step.hidden = i !== index; });
    current = index;
    steps[index].querySelector('h1').focus({preventScroll:true});
    window.scrollTo({top:0, behavior:'smooth'});
  }

  form.addEventListener('change', event => {
    const source = event.target;
    if (!(source instanceof HTMLInputElement)) return;
    enforceExclusive(source); updateDetailInputs(source); updateConditions();
    const block = source.closest('.question-block');
    if (block) updateLimit(block);
  });

  document.querySelectorAll('.next-button').forEach(button => button.addEventListener('click', () => {
    const step = steps[current];
    if (!validateStep(step)) return;
    if (current === 0 && singleValue('C01') === 'No acepto participar.') {
      form.hidden = true;
      const declined = document.getElementById('declined'); declined.hidden = false; declined.focus();
      return;
    }
    showStep(Math.min(current + 1, steps.length - 1));
  }));

  document.querySelectorAll('.back-button').forEach(button => button.addEventListener('click', () => showStep(Math.max(0, current - 1))));

  form.addEventListener('submit', event => {
    if (!validateStep(steps[current])) { event.preventDefault(); return; }
    if (preview) {
      event.preventDefault();
      form.hidden = true;
      const confirmation = document.getElementById('confirmation');
      confirmation.querySelector('p').textContent = 'Prueba completada correctamente. Las respuestas no fueron almacenadas porque el formulario continúa en modo de prueba.';
      confirmation.hidden = false; confirmation.focus();
      return;
    }
    iframeLoaded = false;
    setTimeout(() => { iframeLoaded = true; }, 250);
  });

  document.getElementById('submission-target').addEventListener('load', () => {
    if (preview || !iframeLoaded) return;
    form.hidden = true;
    const confirmation = document.getElementById('confirmation'); confirmation.hidden = false; confirmation.focus();
  });

  updateConditions();
  document.querySelectorAll('[data-max-choices]').forEach(updateLimit);
})();
'''


def main() -> None:
    if not DOCX_FILES:
        raise SystemExit("No se encontró un archivo .docx en la raíz del repositorio.")
    source = DOCX_FILES[0]
    consent, sections, confirmation = parse_document(source)
    if len(sections) < 8:
        raise SystemExit(f"Conversión incompleta: solo se detectaron {len(sections)} secciones.")
    question_count = sum(len(section.questions) for section in sections)
    if question_count < 43:
        raise SystemExit(f"Conversión incompleta: solo se detectaron {question_count} preguntas.")

    DIST.mkdir(parents=True, exist_ok=True)
    (DIST / "index.html").write_text(build_html(consent, sections, confirmation), encoding="utf-8")
    (DIST / "styles.css").write_text(CSS.strip() + "\n", encoding="utf-8")
    (DIST / "app.js").write_text(JS.strip() + "\n", encoding="utf-8")
    (DIST / ".nojekyll").write_text("", encoding="utf-8")
    print(f"Formulario generado desde {source.name}: {len(sections)} pasos, {question_count} preguntas.")


if __name__ == "__main__":
    main()
